#include "openxlsx2.h"

// [[Rcpp::export]]
Rcpp::DataFrame col_to_df(XPtrXML doc) {

  Rcpp::CharacterVector col_nams= {
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

  auto nn = std::distance(doc->begin(), doc->end());
  auto kk = col_nams.length();

  Rcpp::CharacterVector rvec(nn);

  // 1. create the list
  Rcpp::List df(kk);
  for (auto i = 0; i < kk; ++i)
  {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
  }

  // 2. fill the list
  // <row ...>
  auto itr = 0;
  for (auto col : doc->children("col")) {
    for (auto attrs : col.attributes()) {

      Rcpp::CharacterVector attr_name = attrs.name();
      std::string attr_value = attrs.value();

      // mimic which
      Rcpp::IntegerVector mtc = Rcpp::match(col_nams, attr_name);
      Rcpp::IntegerVector idx = Rcpp::seq(0, mtc.length()-1);

      // check if name is already known
      if (all(Rcpp::is_na(mtc))) {
        Rcpp::Rcout << attr_name << ": not found in col name table" << std::endl;
      } else {
        size_t ii = Rcpp::as<size_t>(idx[!Rcpp::is_na(mtc)]);
        Rcpp::as<Rcpp::CharacterVector>(df[ii])[itr] = attr_value;
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
Rcpp::CharacterVector df_to_col(Rcpp::DataFrame df_col) {

  auto n = df_col.nrow();
  Rcpp::CharacterVector z(n);

  for (auto i = 0; i < n; ++i) {
    pugi::xml_document doc;
    Rcpp::CharacterVector attrnams = df_col.names();

    pugi::xml_node col = doc.append_child("col");

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

  Rcpp::CharacterVector row_nams= {
    "collapsed",
    "customFormat",
    "customHeight",
    "x14ac:dyDescent",
    "ht",
    "hidden",
    "outlineLevel",
    "r",
    "ph",
    "spans",
    "s",
    "thickBot",
    "thickTop"
  };

  auto nn = std::distance(ws.children("row").begin(), ws.children("row").end());
  auto kk = row_nams.length();

  Rcpp::CharacterVector rvec(nn);

  // 1. create the list
  Rcpp::List df(kk);
  for (auto i = 0; i < kk; ++i)
  {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
  }

  // 2. fill the list
  // <row ...>
  auto itr = 0;
  for (auto row : ws.children("row")) {
    bool has_rowname = false;
    for (auto attrs : row.attributes()) {

      Rcpp::CharacterVector attr_name = attrs.name();
      std::string attr_value = attrs.value();

      // mimic which
      Rcpp::IntegerVector mtc = Rcpp::match(row_nams, attr_name);
      Rcpp::IntegerVector idx = Rcpp::seq(0, mtc.length()-1);

      // check if name is already known
      if (all(Rcpp::is_na(mtc))) {
        Rcpp::Rcout << attr_name << ": not found in row name table" << std::endl;
      } else {
        size_t ii = Rcpp::as<size_t>(idx[!Rcpp::is_na(mtc)]);
        Rcpp::as<Rcpp::CharacterVector>(df[ii])[itr] = attr_value;
        if (attr_name[0] == "r") has_rowname = true;
      }

    }

    // some files have no row name in this case, we add one
    if (!has_rowname) {
      Rcpp::CharacterVector attr_name = {"r"};
      Rcpp::IntegerVector mtc = Rcpp::match(row_nams, attr_name);
      Rcpp::IntegerVector idx = Rcpp::seq(0, mtc.length()-1);
      size_t ii = Rcpp::as<size_t>(idx[!Rcpp::is_na(mtc)]);
      Rcpp::as<Rcpp::CharacterVector>(df[ii])[itr] = std::to_string(itr + 1);
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

//' @import Rcpp
// this function imports the data from the dataset and returns row_attr and cc
// [[Rcpp::export]]
void loadvals(Rcpp::Reference wb, XPtrXML doc) {

  auto ws = doc->child("worksheet").child("sheetData");

  size_t n = std::distance(ws.begin(), ws.end());

  // character
  Rcpp::DataFrame row_attributes;
  Rcpp::Shield<SEXP> rownames(Rf_allocVector(STRSXP, n));

  std::vector<xml_col> xml_cols;

  // we check against these
  const std::string f_str = "f";
  const std::string r_str = "r";
  const std::string s_str = "s";
  const std::string t_str = "t";
  const std::string v_str = "v";
  const std::string is_str = "is";
  const std::string si_str = "si";
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
    std::string buffer;

    auto itr_cols = 0;
    for (auto col : worksheet.children("c")) {

      // contains all values of a col
      xml_col single_xml_col;

      // get number of children and attributes
      auto nn = std::distance(col.children().begin(), col.children().end());

      // typ: attribute ------------------------------------------------------
      bool has_colname = false;
      auto attr_itr = 0;
      for (auto attr : col.attributes()) {

        // buffer = attr.name();

        buffer = attr.value();

        if (attr.name() == r_str) {
          // get r attr e.g. "A1" and return colnames "A"
          // get col name
          std::string colrow = attr.value();
          colrow.erase(std::remove_if(colrow.begin(),
                                      colrow.end(),
                                      &isdigit),
                                      colrow.end());
          single_xml_col.c_r = colrow;
          has_colname = true;

          // get colnum
          colrow = attr.value();
          // remove numeric from string
          colrow.erase(std::remove_if(colrow.begin(),
                                      colrow.end(),
                                      &isalpha),
                                      colrow.end());
          single_xml_col.row_r = colrow;

        }

        if (attr.name() == s_str) single_xml_col.c_s = buffer;
        if (attr.name() == t_str) single_xml_col.c_t = buffer;

        ++attr_itr;
      }
      // some files have no colnames. in this case we need to add c_r and row_r
      // if the file provides dimensions, they could be fixed later
      if (!has_colname) {
        Rcpp::IntegerVector itr_vec(1);
        itr_vec[0] = itr_cols +1;
        std::string tmp_colname= Rcpp::as<std::string>(int_2_cell_ref(itr_vec));
        single_xml_col.c_r = tmp_colname;
        single_xml_col.row_r = std::to_string(itr_rows + 1);
      }

      // val ------------------------------------------------------------------
      if (nn > 0) {
        auto val_itr = 0;
        for (auto val: col.children()) {

          // <is>
          if (val.name() == is_str) {
            std::ostringstream oss;
            val.print(oss, " ", pugi::format_raw);
            single_xml_col.is = oss.str();
          } // </is>

          // <v>
          if (val.name() == v_str)  single_xml_col.v = val.child_value();

          // <f>
          if (val.name() == f_str)  {

            single_xml_col.f = val.child_value();

            // additional attributes to <f>
            // This currently handles
            //  * t=
            //  * ref=
            //  * si=
            for (auto cattr : val.attributes())
            {
              buffer = cattr.value();
              if (cattr.name() == t_str) single_xml_col.f_t = buffer;
              if (cattr.name() == si_str) single_xml_col.f_si = buffer;
              if (cattr.name() == ref_str) single_xml_col.f_ref = buffer;
            }

          } // </f>

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

  wb.field("row_attr") = row_attributes;
  wb.field("cc")  = Rcpp::wrap(xml_cols);

}

// converts sharedstrings xml tree to R-Character Vector
// [[Rcpp::export]]
SEXP si_to_txt(XPtrXML doc) {

  auto sst = doc->child("sst");
  auto n = std::distance(sst.begin(), sst.end());

  Rcpp::CharacterVector res(Rcpp::no_init(n));

  auto i = 0;
  for (auto si : doc->child("sst").children("si"))
  {
    // text to export
    std::string text = "";

    // has only t node
    for (auto t : si.children("t")) {
      text = t.child_value();
    }

    // has r node with t node
    // linebreaks and spaces are handled in the nodes
    for (auto r : si.children("r")) {
      for (auto t :r.children("t")) {
        text += t.child_value();
      }
    }

    // push everything back
    res[i] = text;
    ++i;
  }

  return res;
}

// converts inlineStr xml tree to R-Character Vector
// [[Rcpp::export]]
SEXP is_to_txt(Rcpp::CharacterVector is_vec) {

  auto n = is_vec.length();
  Rcpp::CharacterVector res(Rcpp::no_init(n));

  for (auto i = 0; i < n; ++i) {

    std::string tmp = Rcpp::as<std::string>(is_vec[i]);

    pugi::xml_document doc;
    pugi::xml_parse_result result = doc.load_string(tmp.c_str(), pugi::parse_default | pugi::parse_escapes);

    if (!result) {
      Rcpp::stop("inlineStr xml import unsuccessfull");
    }

    for (auto is : doc.children("is"))
    {
      // text to export
      std::string text = "";

      // has only t node
      for (auto t : is.children("t")) {
        text = t.child_value();
      }

      // has r node with t node
      // phoneticPr (Phonetic Properties)
      // r (Rich Text Run)
      // rPr (Run Properties)
      // rPh (Phonetic Run)
      // t (Text)
      // linebreaks and spaces are handled in the nodes
      for (auto r : is.children("r")) {
        for (auto t :r.children("t")) {
          text += t.child_value();
        }
      }

      // push everything back
      res[i] = text;
    }

  }

  return res;
}

// similar to dcast converts cc dataframe to z dataframe
// [[Rcpp::export]]
void long_to_wide(Rcpp::DataFrame z, Rcpp::DataFrame tt,  Rcpp::DataFrame zz) {

  auto n = zz.nrow();

  Rcpp::IntegerVector rows = zz["rows"];
  Rcpp::IntegerVector cols = zz["cols"];
  Rcpp::CharacterVector vals = zz["val"];
  Rcpp::CharacterVector typs = zz["typ"];

  for (auto i = 0; i < n; ++i) {
    Rcpp::as<Rcpp::CharacterVector>(z[cols[i]])[rows[i]] = vals[i];
    Rcpp::as<Rcpp::CharacterVector>(tt[cols[i]])[rows[i]] = typs[i];
  }
}


// [[Rcpp::export]]
int cell_ref_to_col( std::string x ){

  // This function converts the Excel column letter to an integer
  char A = 'A';
  int a_value = (int)A - 1;
  int sum = 0;

  // remove digits from string
  x.erase(std::remove_if(x.begin()+1, x.end(), ::isdigit),x.end());
  int k = x.length();

  for (int j = 0; j < k; j++){
    sum *= 26;
    sum += (x[j] - a_value);
  }

  return sum;

}


// [[Rcpp::export]]
Rcpp::CharacterVector int_2_cell_ref(Rcpp::IntegerVector cols){

  std::vector<std::string> LETTERS = get_letters();

  int n = cols.size();
  Rcpp::CharacterVector res(n);
  std::fill(res.begin(), res.end(), NA_STRING);

  int x;
  int modulo;


  for(int i = 0; i < n; i++){

    if(!Rcpp::IntegerVector::is_na(cols[i])){

      std::string columnName;
      x = cols[i];
      while(x > 0){
        modulo = (x - 1) % 26;
        columnName = LETTERS[modulo] + columnName;
        x = (x - modulo) / 26;
      }
      res[i] = columnName;
    }

  }

  return res ;

}

