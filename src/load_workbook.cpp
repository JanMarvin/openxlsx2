#include <set>
#include "openxlsx2.h"

// [[Rcpp::export]]
Rcpp::DataFrame col_to_df(XPtrXML doc) {

  check_xptr_validity(doc);

  std::set<std::string> col_nams{"bestFit", "collapsed",    "customWidth", "hidden", "max",
                                 "min",     "outlineLevel", "phonetic",    "style",  "width"};

  R_xlen_t nn = std::distance(doc->begin(), doc->end());
  R_xlen_t kk = static_cast<R_xlen_t>(col_nams.size());

  // 1. create the list
  Rcpp::List df(kk);
  for (R_xlen_t i = 0; i < kk; ++i) {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
  }

  // 2. fill the list
  // <row ...>
  auto itr = 0;
  for (auto col : doc->children("col")) {
    for (auto attrs : col.attributes()) {
      std::string attr_name = attrs.name();
      std::string attr_value = attrs.value();
      auto find_res = col_nams.find(attr_name);

      // check if name is already known
      if (col_nams.count(attr_name) == 0) {
        Rcpp::Rcout << attr_name << ": not found in col name table" << std::endl;
      } else {
        R_xlen_t mtc = std::distance(col_nams.begin(), find_res);
        Rcpp::as<Rcpp::CharacterVector>(df[mtc])[itr] = attr_value;
      }
    }

    ++itr;
  }

  // 3. Create a data.frame
  df.attr("row.names") = Rcpp::IntegerVector::create(NA_INTEGER, -nn);
  df.attr("names") = col_nams;
  df.attr("class") = "data.frame";

  return df;
}

inline Rcpp::DataFrame row_to_df(XPtrXML doc) {

  check_xptr_validity(doc);

  auto ws = doc->child("worksheet").child("sheetData");

  std::vector<std::string> row_nams = {
    "r",
    "spans",
    "s",
    "ht",
    "hidden",
    "collapsed",
    "customFormat",
    "customHeight",
    "x14ac:dyDescent",
    "outlineLevel",
    "ph",
    "thickBot",
    "thickTop"
  };
  std::unordered_map<std::string, R_xlen_t> name_to_index;
  for (size_t i = 0; i < row_nams.size(); ++i)
    name_to_index[row_nams[i]] = static_cast<R_xlen_t>(i);

  R_xlen_t nn = std::distance(ws.children("row").begin(), ws.children("row").end());
  R_xlen_t kk = static_cast<R_xlen_t>(row_nams.size());

  // 1. create the list
  Rcpp::List df(kk);
  for (R_xlen_t i = 0; i < kk; ++i) {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
  }

  // 2. fill the list
  // <row ...>
  R_xlen_t row_idx = 0;
  for (auto row : ws.children("row")) {
    bool has_rowname = false;

    for (auto attrs : row.attributes()) {
      std::string attr_name = attrs.name();
      std::string attr_value = attrs.value();

      auto it = name_to_index.find(attr_name);
      if (it != name_to_index.end()) {
        Rcpp::CharacterVector col = df[it->second];
        col[row_idx] = attr_value;
        if (attr_name == "r") has_rowname = true;
      } else {
        Rcpp::Rcout << attr_name << ": not found in row name table" << std::endl;
      }
    }

    if (!has_rowname) {
      Rcpp::CharacterVector col = df[name_to_index["r"]];
      col[row_idx] = std::to_string(row_idx + 1);
    }

    ++row_idx;
  }

  // 3. Create a data.frame
  df.attr("row.names") = Rcpp::IntegerVector::create(NA_INTEGER, -nn);
  df.attr("names") = row_nams;
  df.attr("class") = "data.frame";

  return df;
}

// this function imports the data from the dataset and returns row_attr and cc
// [[Rcpp::export]]
void loadvals(Rcpp::Environment sheet_data, XPtrXML doc) {

  check_xptr_validity(doc);

  auto ws = doc->child("worksheet").child("sheetData");

  bool has_cm = false, has_ph = false, has_vm = false;

  // character
  Rcpp::DataFrame row_attributes;

  R_xlen_t total_cells_count = 0;
  for (auto worksheet_row : ws.children("row")) {
    total_cells_count += std::distance(worksheet_row.children("c").begin(), worksheet_row.children("c").end());
  }

  std::vector<xml_col> xml_cols;
  if (total_cells_count > 0) {
    xml_cols.reserve(static_cast<size_t>(total_cells_count));
  }

  // we check against these
  const std::string f_str = "f";
  const std::string r_str = "r";
  const std::string s_str = "s";
  const std::string t_str = "t";
  const std::string v_str = "v";
  const std::string cm_str = "cm";
  const std::string is_str = "is";
  const std::string ph_str = "ph";
  const std::string vm_str = "vm";

  /*****************************************************************************
   * Row information is returned as list of lists returning as much as possible.
   *
   * Col information is returned as dataframe returning only a fraction of known
   * tags and attributes.
   ****************************************************************************/
  row_attributes = row_to_df(doc);

  R_xlen_t idx = 0, itr_rows = 0;
  for (auto worksheet : ws.children("row")) {
    /* ---------------------------------------------------------------------- */
    /* read cval, and ctyp -------------------------------------------------- */
    /* ---------------------------------------------------------------------- */

    // buffer is string buf is SEXP
    std::string buffer, attr_name, val_name;

    uint32_t itr_cols = 0;
    for (auto col : worksheet.children("c")) {
      checkInterrupt(idx);

      // contains all values of a col
      xml_col single_xml_col;

      // get number of children and attributes
      auto nn = std::distance(col.children().begin(), col.children().end());

      // typ: attribute ------------------------------------------------------
      bool has_colname = false;
      for (auto attr : col.attributes()) {
        buffer = attr.value();
        attr_name = attr.name();

        if (attr_name == r_str) {
          // get r attr e.g. "A1" and return colnames "A"
          single_xml_col.r = buffer;

          // get col name
          single_xml_col.c_r = rm_rownum(buffer);
          has_colname = true;

          // get colnum
          single_xml_col.row_r = rm_colnum(buffer);

          // if some cells of the workbook have colnames but other dont,
          // this will increase itr_cols and avoid duplicates in cc
          itr_cols = static_cast<uint32_t>(uint_col_to_int(single_xml_col.c_r) - 1);
        }

        if (attr_name == s_str) single_xml_col.c_s = buffer;
        if (attr_name == t_str) single_xml_col.c_t = buffer;
        if (attr_name == cm_str) {
          has_cm = true;
          single_xml_col.c_cm = buffer;
        }
        if (attr_name == ph_str) {
          has_ph = true;
          single_xml_col.c_ph = buffer;
        }
        if (attr_name == vm_str) {
          has_vm = true;
          single_xml_col.c_vm = buffer;
        }
      }

      // some files have no colnames. in this case we need to add c_r and row_r
      // if the file provides dimensions, they could be fixed later
      if (!has_colname) {
        single_xml_col.c_r = int_to_col(itr_cols + 1);
        single_xml_col.row_r = std::to_string(itr_rows + 1);
        single_xml_col.r = single_xml_col.c_r + single_xml_col.row_r;
      }

      // val ------------------------------------------------------------------
      if (nn > 0) {
        for (auto val : col.children()) {
          val_name = val.name();

          // <is>
          if (val_name == is_str) {
            std::ostringstream oss;
            val.print(oss, " ", pugi::format_raw | pugi::format_no_escapes);
            single_xml_col.is = oss.str();
          }  // </is>

          if (val_name == f_str) {  // <f>
            // Store the content of <f> as single_xml_col.f
            single_xml_col.f = val.text().get();

            // Serialize the attributes of <f> as single_xml_col.f_attr
            std::ostringstream attr_stream;
            for (auto f_attr : val.attributes()) {
              attr_stream << f_attr.name() << "=\"" << f_attr.value() << "\" ";
            }
            single_xml_col.f_attr = attr_stream.str();
          }  // </f>

          // <v>
          if (val_name == v_str)  single_xml_col.v = val.text().get();
        }

        /* row is done */
      }

      xml_cols.push_back(single_xml_col);

      ++itr_cols;
    }

    /* ---------------------------------------------------------------------- */

    ++itr_rows;
  }

  // Rcpp::Rcout << has_cm << ": " << has_ph << ": " << has_vm << std::endl;

  sheet_data["row_attr"] = row_attributes;
  sheet_data["cc"] = xml_cols_to_df(xml_cols, has_cm, has_ph, has_vm);
}
