#include "openxlsx2.h"

//' @import Rcpp
// this function imports the data from the dataset and returns row_attr and cc
// [[Rcpp::export]]
void loadvals(Rcpp::Reference wb, XPtrXML doc) {

  auto ws = doc->child("worksheet").child("sheetData");

  size_t n = std::distance(ws.begin(), ws.end());

  // character
  Rcpp::Shield<SEXP> row_attributes(Rf_allocVector(VECSXP, n));
  Rcpp::Shield<SEXP> rownames(Rf_allocVector(STRSXP, n));

  std::vector<xml_col> xml_cols;

  // we check against these
  const std::string f_str = "f";
  const std::string r_str = "r";
  const std::string s_str = "s";
  const std::string t_str = "t";
  const std::string v_str = "v";
  const std::string si_str = "si";
  const std::string ref_str = "ref";

  /*****************************************************************************
   * Row information is returned as list of lists returning as much as possible.
   *
   * Col information is returned as dataframe returning only a fraction of known
   * tags and attributes.
   ****************************************************************************/
  auto itr_rows = 0;
  for (auto worksheet: ws.children("row")) {

    size_t k = std::distance(worksheet.begin(), worksheet.end());

    Rcpp::Shield<SEXP> cc_r(Rf_allocVector(VECSXP, k));
    Rcpp::Shield<SEXP> colnames(Rf_allocVector(STRSXP, k));

    // buffer is string buf is SEXP
    std::string buffer;

    /* row attributes ------------------------------------------------------- */
    auto nn = std::distance(worksheet.attributes_begin(), worksheet.attributes_end());

    Rcpp::Shield<SEXP> row_attr(Rf_allocVector(VECSXP, nn));
    Rcpp::Shield<SEXP> row_attr_nam(Rf_allocVector(STRSXP, nn));

    bool has_rowname = false;
    auto attr_itr = 0;
    for (auto attr : worksheet.attributes()) {

      buffer = attr.name();
      SET_STRING_ELT(row_attr_nam, attr_itr, Rf_mkChar(buffer.c_str()));

      buffer = attr.value();
      Rcpp::Shield<SEXP> buf(Rf_allocVector(STRSXP, 1));
      SET_STRING_ELT(buf, 0, Rf_mkChar(buffer.c_str()));
      SET_VECTOR_ELT(row_attr, attr_itr, buf);
      ++attr_itr;

      // push row name back (will assign it to list)
      if (attr.name() == r_str) {
        buffer = attr.value();
        SET_STRING_ELT(rownames, itr_rows, Rf_mkChar(buffer.c_str()));
        has_rowname = true;
      }

    }
    if(!has_rowname) {
      for (auto i = 0; i < n; ++i) {
        buffer = std::to_string(i+1);
        SET_STRING_ELT(rownames, i, Rf_mkChar(buffer.c_str()));
      }
    }

    // assign names and push back
    ::Rf_setAttrib(row_attr, R_NamesSymbol, row_attr_nam);
    SET_VECTOR_ELT(row_attributes, itr_rows, row_attr);

    /* ---------------------------------------------------------------------- */
    /* read cval, and ctyp -------------------------------------------------- */
    /* ---------------------------------------------------------------------- */

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
      // some files have no colnames. This used to work, check again
      if (!has_colname) {
        Rcpp::IntegerVector itr_vec(1);
        itr_vec[0] = itr_cols +1;
        std::string tmp_colname= Rcpp::as<std::string>(int_2_cell_ref(itr_vec));
        single_xml_col.c_r = tmp_colname;
      }

      // val ------------------------------------------------------------------
      if (nn > 0) {
        auto val_itr = 0;
        for (auto val: col.children()) {


          auto ff_itr = 0;

          // additional attributes to <f t="shared" ...>
          for (auto cattr : val.attributes())
          {
            // buffer = cattr.name();

            buffer = cattr.value();
            if (cattr.name() == t_str) single_xml_col.f_t = buffer;
            if (cattr.name() == si_str) single_xml_col.f_si = buffer;
            if (cattr.name() == ref_str) single_xml_col.f_ref = buffer;

            ++ff_itr;
          }

          buffer = val.name();

          // <is> nodes contain additional <t> node.
          // TODO: check if multiple t nodes are possible, for now return one.
          // the t-node can bring its very own attributes
          // Rcpp::Shield<SEXP> buf(Rf_allocVector(STRSXP, 1));
          if (val.child("t")) {
            buffer = val.child("t").child_value();
            single_xml_col.t = buffer;
          } else {
            buffer = val.child_value();
            // val.name() == v or f
            if (val.name() == v_str) single_xml_col.v = buffer;
            if (val.name() == f_str) single_xml_col.f = buffer;
          }

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

  ::Rf_setAttrib(row_attributes, R_NamesSymbol, rownames);

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

// mimics the R function which used below
Rcpp::IntegerVector rcpp_which(Rcpp::IntegerVector x) {
    Rcpp::IntegerVector v = Rcpp::seq(0, x.size()-1);
    return v[!Rcpp::is_na(x)];
}

// similar to dcast converts cc dataframe to z dataframe
// [[Rcpp::export]]
void long_to_wide(Rcpp::DataFrame z, Rcpp::DataFrame tt,  Rcpp::DataFrame cc, Rcpp::List dn) {

  auto n = cc.nrow();

  Rcpp::CharacterVector row_r = cc["row_r"];
  Rcpp::CharacterVector c_r   = cc["c_r"];
  Rcpp::CharacterVector val   = cc["val"];
  Rcpp::CharacterVector typ   = cc["typ"];

  Rcpp::CharacterVector row_names = dn[0];
  Rcpp::CharacterVector col_names = dn[1];

  for (auto i = 0; i < n; ++i) {

    Rcpp::CharacterVector row_r_i = Rcpp::as<Rcpp::CharacterVector>(row_r[i]);
    Rcpp::CharacterVector c_r_i   = Rcpp::as<Rcpp::CharacterVector>(c_r[i]);
    std::string val_i   = Rcpp::as<std::string>(val[i]);
    std::string val_tt   = Rcpp::as<std::string>(typ[i]);

    Rcpp::IntegerVector s1 = Rcpp::match(row_names, row_r_i);
    int64_t sel_row = Rcpp::as<int64_t>(rcpp_which(s1));

    Rcpp::IntegerVector s2 = Rcpp::match(col_names, c_r_i);
    int64_t sel_col = Rcpp::as<int64_t>(rcpp_which(s2));

    // Rcpp::Rcout << sel_row << " " << sel_col << " " << val_i << std::endl;

    // only update if not missing in the xml input
    if (val_i.compare("_openxlsx_NA_") != 0) {
      Rcpp::as<Rcpp::CharacterVector>(z[sel_col])[sel_row] = val_i;
      Rcpp::as<Rcpp::CharacterVector>(tt[sel_col])[sel_row] = val_tt;
    }

  }
}


// [[Rcpp::export]]
SEXP getOpenClosedNode(std::string xml, std::string open_tag, std::string close_tag){

  if(xml.length() == 0)
    return Rcpp::wrap(NA_STRING);

  xml = " " + xml;
  size_t pos = 0;
  size_t endPos = 0;

  size_t k = open_tag.length();
  size_t l = close_tag.length();

  std::vector<std::string> r;

  while(1){

    pos = xml.find(open_tag, pos+1);
    endPos = xml.find(close_tag, pos+k);

    if((pos == std::string::npos) | (endPos == std::string::npos))
      break;

    r.push_back(xml.substr(pos, endPos-pos+l).c_str());

  }

  Rcpp::CharacterVector out = Rcpp::wrap(r);
  return markUTF8(out);
}


// [[Rcpp::export]]
SEXP getAttr(Rcpp::CharacterVector x, std::string tag){

  size_t n = x.size();
  size_t k = tag.length();

  if(n == 0)
    return Rcpp::wrap(-1);

  std::string xml;
  Rcpp::CharacterVector r(n);
  size_t pos = 0;
  size_t endPos = 0;
  std::string rtagEnd = "\"";

  for(size_t i = 0; i < n; i++){

    // find opening tag
    xml = x[i];
    pos = xml.find(tag, 0);

    if(pos == std::string::npos){
      r[i] = NA_STRING;
    }else{
      endPos = xml.find(rtagEnd, pos+k);
      r[i] = xml.substr(pos+k, endPos-pos-k).c_str();
    }
  }

  return markUTF8(r);   // no need to wrap as r is already a CharacterVector

}


// [[Rcpp::export]]
Rcpp::CharacterVector get_extLst_Major(std::string xml){

  // find page margin or pagesetup then take the extLst after that

  if(xml.length() == 0)
    return Rcpp::wrap(NA_STRING);

  std::vector<std::string> r;
  std::string tagEnd = "</extLst>";
  size_t endPos = 0;
  std::string node;


  size_t pos = xml.find("<pageSetup ", 0);
  if(pos == std::string::npos)
    pos = xml.find("<pageMargins ", 0);

  if(pos == std::string::npos)
    pos = xml.find("</conditionalFormatting>", 0);

  if(pos == std::string::npos)
    return Rcpp::wrap(NA_STRING);

  while(1){

    pos = xml.find("<extLst>", pos + 1);
    if(pos == std::string::npos)
      break;

    endPos = xml.find(tagEnd, pos + 8);

    node = xml.substr(pos + 8, endPos - pos - 8);
    //pos = xml.find("conditionalFormattings", pos + 1);
    //if(pos == std::string::npos)
    //  break;

    r.push_back(node.c_str());

  }

  Rcpp::CharacterVector out = Rcpp::wrap(r);
  return markUTF8(out);

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
