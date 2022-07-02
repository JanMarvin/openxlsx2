#include "openxlsx2.h"
#include <set>

// [[Rcpp::export]]
Rcpp::DataFrame col_to_df(XPtrXML doc) {

  std::set<std::string> col_nams= {
    "bestFit",
    "collapsed",
    "customWidth",
    "hidden",
    "max",
    "min",
    "outlineLevel",
    "phonetic",
    "style",
    "width"
  };

  R_xlen_t nn = std::distance(doc->begin(), doc->end());
  R_xlen_t kk = col_nams.size();

  Rcpp::CharacterVector rvec(nn);

  // 1. create the list
  Rcpp::List df(kk);
  for (R_xlen_t i = 0; i < kk; ++i)
  {
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
        auto mtc = std::distance(col_nams.begin(), find_res);
        Rcpp::as<Rcpp::CharacterVector>(df[mtc])[itr] = attr_value;
      }

    }

    // rownames as character vectors matching to <c s= ...>
    rvec[itr] = std::to_string(itr);

    ++itr;
  }

  // 3. Create a data.frame
  df.attr("row.names") = rvec;
  df.attr("names") = col_nams;
  df.attr("class") = "data.frame";

  return df;
}


// [[Rcpp::export]]
Rcpp::CharacterVector df_to_xml(std::string name, Rcpp::DataFrame df_col) {

  auto n = df_col.nrow();
  Rcpp::CharacterVector z(n);

  for (auto i = 0; i < n; ++i) {
    pugi::xml_document doc;
    Rcpp::CharacterVector attrnams = df_col.names();

    pugi::xml_node col = doc.append_child(name.c_str());

    for (auto j = 0; j < df_col.ncol(); ++j) {
      Rcpp::CharacterVector cv_s = "";
      cv_s = Rcpp::as<Rcpp::CharacterVector>(df_col[j])[i];

      // only write attributes where cv_s has a value
      if (cv_s[0] != "") {
        // Rf_PrintValue(cv_s);
        const std::string val_strl = Rcpp::as<std::string>(cv_s);
        col.append_attribute(attrnams[j]) = val_strl.c_str();
      }
    }

    std::ostringstream oss;
    doc.print(oss, " ", pugi::format_raw);

    z[i] = oss.str();
  }

  return z;
}


Rcpp::DataFrame row_to_df(XPtrXML doc) {

  auto ws = doc->child("worksheet").child("sheetData");

  std::set<std::string> row_nams {
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

  size_t nn = std::distance(ws.children("row").begin(), ws.children("row").end());
  size_t kk = row_nams.size();

  Rcpp::CharacterVector rvec(nn);

  // 1. create the list
  Rcpp::List df(kk);
  for (size_t i = 0; i < kk; ++i)
  {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
  }

  // 2. fill the list
  // <row ...>
  auto itr = 0;
  for (auto row : ws.children("row")) {
    bool has_rowname = false;
    for (auto attrs : row.attributes()) {

      std::string attr_name = attrs.name();
      std::string attr_value = attrs.value();

      // mimic which
      auto find_res = row_nams.find(attr_name);

      // check if name is already known
      if (row_nams.count(attr_name) == 0) {
        Rcpp::Rcout << attr_name << ": not found in row name table" << std::endl;
      } else {
        auto mtc = std::distance(row_nams.begin(), find_res);
        Rcpp::as<Rcpp::CharacterVector>(df[mtc])[itr] = attr_value;
        if (attr_name == "r") has_rowname = true;
      }

    }

    // some files have no row name in this case, we add one
    if (!has_rowname) {
      std::string attr_name = "r";
      auto find_res = row_nams.find(attr_name);
      auto mtc = std::distance(row_nams.begin(), find_res);
      Rcpp::as<Rcpp::CharacterVector>(df[mtc])[itr] = std::to_string(itr + 1);
    }

    // rownames as character vectors matching to <c s= ...>
    rvec[itr] = std::to_string(itr);

    ++itr;
  }

  // 3. Create a data.frame
  df.attr("row.names") = rvec;
  df.attr("names") = row_nams;
  df.attr("class") = "data.frame";

  return df;
}

// this function imports the data from the dataset and returns row_attr and cc
// [[Rcpp::export]]
void loadvals(Rcpp::Environment sheet_data, XPtrXML doc) {

  auto ws = doc->child("worksheet").child("sheetData");

  // character
  Rcpp::DataFrame row_attributes;

  std::vector<xml_col> xml_cols;

  // we check against these
  const std::string f_str = "f";
  const std::string r_str = "r";
  const std::string s_str = "s";
  const std::string t_str = "t";
  const std::string v_str = "v";
  const std::string ca_str = "ca";
  const std::string cm_str = "cm";
  const std::string is_str = "is";
  const std::string ph_str = "ph";
  const std::string si_str = "si";
  const std::string vm_str = "vm";
  const std::string ref_str = "ref";



  /*****************************************************************************
   * Row information is returned as list of lists returning as much as possible.
   *
   * Col information is returned as dataframe returning only a fraction of known
   * tags and attributes.
   ****************************************************************************/
  row_attributes = row_to_df(doc);

  auto itr_rows = 0;
  for (auto worksheet: ws.children("row")) {
    /* ---------------------------------------------------------------------- */
    /* read cval, and ctyp -------------------------------------------------- */
    /* ---------------------------------------------------------------------- */

    // buffer is string buf is SEXP
    std::string buffer, attr_name, val_name, cattr_name;

    auto itr_cols = 0;
    for (auto col : worksheet.children("c")) {

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
          std::string colrow = buffer;
          colrow.erase(std::remove_if(colrow.begin(),
                                      colrow.end(),
                                      &isdigit),
                                      colrow.end());
          single_xml_col.c_r = colrow;
          has_colname = true;

          // get colnum
          colrow = buffer;
          // remove numeric from string
          colrow.erase(std::remove_if(colrow.begin(),
                                      colrow.end(),
                                      &isalpha),
                                      colrow.end());
          single_xml_col.row_r = colrow;

          // if some cells of the workbook have colnames but other dont,
          // this will increase itr_cols and avoid duplicates in cc
          itr_cols = uint_col_to_int(single_xml_col.c_r) - 1;

        }

        if (attr_name == s_str) single_xml_col.c_s = buffer;
        if (attr_name == t_str) single_xml_col.c_t = buffer;
        if (attr_name == cm_str) single_xml_col.c_cm = buffer;
        if (attr_name == ph_str) single_xml_col.c_ph = buffer;
        if (attr_name == vm_str) single_xml_col.c_vm = buffer;

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
        auto val_itr = 0;
        for (auto val: col.children()) {

          val_name = val.name();

          // <is>
          if (val_name == is_str) {
            std::ostringstream oss;
            val.print(oss, " ", pugi::format_raw);
            single_xml_col.is = oss.str();
          } // </is>

          // <f>
          if (val_name == f_str) {

            single_xml_col.f = val.child_value();

            // additional attributes to <f>
            // This currently handles
            //  * t=
            //  * ref=
            //  * ca=
            //  * si=
            for (auto cattr : val.attributes())
            {
              buffer = cattr.value();
              cattr_name = cattr.name();
              if (cattr_name == t_str) single_xml_col.f_t = buffer;
              if (cattr_name == ref_str) single_xml_col.f_ref = buffer;
              if (cattr_name == ca_str) single_xml_col.f_ca = buffer;
              if (cattr_name == si_str) single_xml_col.f_si = buffer;
            }

          } // </f>

          // <v>
          if (val_name == v_str)  single_xml_col.v = val.child_value();

          ++val_itr;
        }

        /* row is done */
      }

      xml_cols.push_back(single_xml_col);

      ++itr_cols;
    }


    /* ---------------------------------------------------------------------- */

    ++itr_rows;
  }

  sheet_data["row_attr"] = row_attributes;
  sheet_data["cc"] = Rcpp::wrap(xml_cols);
}
