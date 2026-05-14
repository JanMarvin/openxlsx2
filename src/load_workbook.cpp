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
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(nn));
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

// Names of the attributes captured into the row_attr data.frame.
static const char* const ROW_ATTR_NAMES[] = {
  "r", "spans", "s", "ht", "hidden", "collapsed", "customFormat",
  "customHeight", "x14ac:dyDescent", "outlineLevel", "ph", "thickBot", "thickTop"
};
static const R_xlen_t ROW_ATTR_KK =
  static_cast<R_xlen_t>(sizeof(ROW_ATTR_NAMES) / sizeof(ROW_ATTR_NAMES[0]));

// Allocate the row_attr data.frame sized to n_rows. The columns are filled
// later, inline in loadvals()'s main row loop, so the <row> tree is only
// traversed once. Returns the data.frame; col_ptrs is filled with the column
// SEXPs (so the caller can SET_STRING_ELT directly without re-resolving).
inline Rcpp::DataFrame alloc_row_df(R_xlen_t n_rows, std::vector<SEXP>& col_ptrs) {
  Rcpp::List df(ROW_ATTR_KK);
  col_ptrs.resize(static_cast<size_t>(ROW_ATTR_KK));
  for (R_xlen_t i = 0; i < ROW_ATTR_KK; ++i) {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(n_rows));
    col_ptrs[static_cast<size_t>(i)] = VECTOR_ELT(df, i);
  }

  Rcpp::CharacterVector nms(ROW_ATTR_KK);
  for (R_xlen_t i = 0; i < ROW_ATTR_KK; ++i) nms[i] = ROW_ATTR_NAMES[i];

  df.attr("row.names") = Rcpp::IntegerVector::create(NA_INTEGER, -n_rows);
  df.attr("names") = nms;
  df.attr("class") = "data.frame";
  return df;
}

// Fill one row's attributes into the pre-allocated row_attr columns.
// col_ptrs / name lookup come from alloc_row_df(). Mirrors the old row_to_df
// body for a single <row>.
inline void fill_row_attr(const pugi::xml_node& row, R_xlen_t row_idx,
                          const std::vector<SEXP>& col_ptrs) {
  bool has_rowname = false;

  for (auto attrs : row.attributes()) {
    const char* attr_name = attrs.name();

    // linear scan over ROW_ATTR_NAMES: the list is short (13) and this avoids
    // constructing a std::string per attribute for a map lookup
    R_xlen_t hit = -1;
    for (R_xlen_t k = 0; k < ROW_ATTR_KK; ++k) {
      if (std::strcmp(attr_name, ROW_ATTR_NAMES[k]) == 0) { hit = k; break; }
    }

    if (hit >= 0) {
      SET_STRING_ELT(col_ptrs[static_cast<size_t>(hit)], row_idx,
                     Rf_mkChar(attrs.value()));
      if (hit == 0) has_rowname = true; // "r" is index 0
    } else {
      Rcpp::Rcout << attr_name << ": not found in row name table" << std::endl;
    }
  }

  if (!has_rowname) {
    SET_STRING_ELT(col_ptrs[0], row_idx,
                   Rf_mkChar(std::to_string(row_idx + 1).c_str()));
  }
}

// this function imports the data from the dataset and returns row_attr and cc
// [[Rcpp::export]]
void loadvals(Rcpp::Environment sheet_data, XPtrXML doc) {

  check_xptr_validity(doc);

  auto ws = doc->child("worksheet").child("sheetData");

  bool has_cm = false, has_ph = false, has_vm = false;

  // One cheap pass over <row> to size both the row_attr data.frame and the
  // cell vector reserve. Previously the <row>/<c> tree was walked four times
  // (row count + row fill in row_to_df, cell count here, then the main loop);
  // now it is walked once for sizing and once to fill everything.
  R_xlen_t n_rows = 0;
  R_xlen_t total_cells_count = 0;
  for (auto worksheet_row : ws.children("row")) {
    ++n_rows;
    total_cells_count += std::distance(worksheet_row.children("c").begin(),
                                       worksheet_row.children("c").end());
  }

  std::vector<xml_col> xml_cols;
  if (total_cells_count > 0) {
    xml_cols.reserve(static_cast<size_t>(total_cells_count));
  }

  // we check against these
  const char* const f_str = "f";
  const char* const r_str = "r";
  const char* const s_str = "s";
  const char* const t_str = "t";
  const char* const v_str = "v";
  const char* const cm_str = "cm";
  const char* const is_str = "is";
  const char* const ph_str = "ph";
  const char* const vm_str = "vm";

  /*****************************************************************************
   * Row information is returned as list of lists returning as much as possible.
   *
   * Col information is returned as dataframe returning only a fraction of known
   * tags and attributes.
   ****************************************************************************/
  std::vector<SEXP> row_attr_cols;
  Rcpp::DataFrame row_attributes = alloc_row_df(n_rows, row_attr_cols);
  xml_col single_xml_col;

  R_xlen_t idx = 0, itr_rows = 0;
  for (auto worksheet : ws.children("row")) {

    // fill this row's attributes into row_attr inline (was a separate
    // traversal in row_to_df)
    fill_row_attr(worksheet, itr_rows, row_attr_cols);

    /* ---------------------------------------------------------------------- */
    /* read cval, and ctyp -------------------------------------------------- */
    /* ---------------------------------------------------------------------- */

    uint32_t itr_cols = 0;
    for (auto col : worksheet.children("c")) {
      checkInterrupt(idx);

      // contains all values of a col
      single_xml_col.clear();

      // typ: attribute ------------------------------------------------------
      bool has_colname = false;
      for (auto attr : col.attributes()) {

        if (std::strcmp(attr.name(), r_str) == 0) {
          // get r attr e.g. "A1" and return colnames "A"
          single_xml_col.r = attr.value();

          // get col name
          single_xml_col.c_r = filter_digits(attr.value(), false);
          has_colname = true;

          // get colnum
          single_xml_col.row_r = filter_digits(attr.value(), true);

          // if some cells of the workbook have colnames but other dont,
          // this will increase itr_cols and avoid duplicates in cc
          itr_cols = static_cast<uint32_t>(uint_col_to_int(single_xml_col.c_r) - 1);
        }

        else if (std::strcmp(attr.name(), s_str) == 0) single_xml_col.c_s = attr.value();
        else if (std::strcmp(attr.name(), t_str) == 0) single_xml_col.c_t = attr.value();
        else if (std::strcmp(attr.name(), cm_str) == 0) {
          has_cm = true;
          single_xml_col.c_cm = attr.value();
        }
        else if (std::strcmp(attr.name(), ph_str) == 0) {
          has_ph = true;
          single_xml_col.c_ph = attr.value();
        }
        else if (std::strcmp(attr.name(), vm_str) == 0) {
          has_vm = true;
          single_xml_col.c_vm = attr.value();
        }
      }

      // some files have no colnames. in this case we need to add c_r and row_r
      // if the file provides dimensions, they could be fixed later
      if (!has_colname) {
        single_xml_col.c_r = int_to_col(itr_cols + 1);
        single_xml_col.row_r = std::to_string(itr_rows + 1);
        single_xml_col.r = single_xml_col.c_r + single_xml_col.row_r;
      }

      single_xml_col.key = static_cast<double>(std::atoi(single_xml_col.row_r.c_str())) * 16384L +
        uint_col_to_int(single_xml_col.c_r);

      // val ------------------------------------------------------------------
      if (col.first_child()) {
        for (auto val : col.children()) {

          // <v> -- the default
          if (std::strcmp(val.name(), v_str) == 0)  single_xml_col.v = val.text().get();

          // <is>
          else if (std::strcmp(val.name(), is_str) == 0) {
            xml_string_writer writer;
            val.print(writer, " ", pugi::format_raw | pugi::format_no_escapes);
            single_xml_col.is = std::move(writer.result);
          }  // </is>

          else if (std::strcmp(val.name(), f_str) == 0) {  // <f>
            // Store the content of <f> as single_xml_col.f
            single_xml_col.f = val.text().get();

            // Serialize the attributes of <f> as single_xml_col.f_attr
            for (auto f_attr : val.attributes()) {
              single_xml_col.f_attr.append(f_attr.name());
              single_xml_col.f_attr.append("=\"");
              single_xml_col.f_attr.append(f_attr.value());
              single_xml_col.f_attr.append("\" ");
            }
          }  // </f>

        }

        /* row is done */
      }

      xml_cols.push_back(std::move(single_xml_col));

      ++itr_cols;
    }

    /* ---------------------------------------------------------------------- */

    ++itr_rows;
  }

  // Rcpp::Rcout << has_cm << ": " << has_ph << ": " << has_vm << std::endl;

  sheet_data["row_attr"] = row_attributes;
  sheet_data["cc"] = xml_cols_to_df(xml_cols, has_cm, has_ph, has_vm);
}
