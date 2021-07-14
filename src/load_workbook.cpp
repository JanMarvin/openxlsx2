#include "openxlsx.h"
#include "openxlsx2_types.h"


// loadvals(wb$worksheets[[i]]$sheet_data, worksheet_xml, "worksheet", "sheetData", "row", "c")
// [[Rcpp::export]]
void loadvals(Rcpp::Reference wb, XPtrXML doc) {
  
  auto ws = doc->child("worksheet").child("sheetData");
  
  size_t n = std::distance(ws.begin(), ws.end());
  
  // list
  Rcpp::Shield<SEXP> cc(Rf_allocVector(VECSXP, n));
  
  // character
  Rcpp::Shield<SEXP> row_attributes(Rf_allocVector(VECSXP, n));
  Rcpp::Shield<SEXP> rownames(Rf_allocVector(STRSXP, n));
  
  
  auto itr_rows = 0;
  for (auto worksheet: ws.children("row")) {
    
    size_t k = std::distance(worksheet.begin(), worksheet.end());
    
    Rcpp::Shield<SEXP> cc_r(Rf_allocVector(VECSXP, k));
    Rcpp::Shield<SEXP> colnames(Rf_allocVector(STRSXP, k));
    
    // buffer is string buf is SEXP
    std::string buffer;
    
    /* row attributes ------------------------------------------------------- */
    auto nn = std::distance(worksheet.attributes_begin(), worksheet.attributes_end());
    
    std::string r_str = "r";
    Rcpp::Shield<SEXP> row_attr(Rf_allocVector(VECSXP, nn));
    Rcpp::Shield<SEXP> row_attr_nam(Rf_allocVector(STRSXP, nn));
    
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
      
      // get number of children and attributes
      auto nn = std::distance(col.children().begin(), col.children().end());
      auto aa = std::distance(col.attributes_begin(), col.attributes_end());
      // check if there are f attributes
      auto ff = std::distance(col.child("f").attributes_begin(), col.child("f").attributes_end());
      // prevent dying on a 0 node
      auto tt = nn; if (tt == 0) ++tt;
      
      // cc_cell_nam vector
      std::string val_str = "val", typ_str = "typ", attr_str = "attr";
      Rcpp::Shield<SEXP> cc_cell_nam(Rf_allocVector(STRSXP, 2+ff));
      SET_STRING_ELT(cc_cell_nam, 0, Rf_mkChar(val_str.c_str()));
      SET_STRING_ELT(cc_cell_nam, 1, Rf_mkChar(typ_str.c_str()));
      if (ff > 0) SET_STRING_ELT(cc_cell_nam, 2, Rf_mkChar(attr_str.c_str()));
      
      // lists for cc_cell, v_c, t_c and a_c
      Rcpp::Shield<SEXP> cc_cell(Rf_allocVector(VECSXP, 2+ff));
      Rcpp::Shield<SEXP> v_c(Rf_allocVector(VECSXP, tt));
      Rcpp::Shield<SEXP> t_c(Rf_allocVector(VECSXP, aa));
      Rcpp::Shield<SEXP> a_c(Rf_allocVector(VECSXP, ff));
      
      // character vectors val_name, typ_name, atr_name
      Rcpp::Shield<SEXP> val_name(Rf_allocVector(STRSXP, tt));
      Rcpp::Shield<SEXP> typ_name(Rf_allocVector(STRSXP, aa));
      Rcpp::Shield<SEXP> atr_name(Rf_allocVector(STRSXP, ff));
      
      // typ: attribute ------------------------------------------------------
      auto attr_itr = 0;
      for (auto attr : col.attributes()) {
        
        buffer = attr.name();
        SET_STRING_ELT(typ_name, attr_itr, Rf_mkChar(buffer.c_str()));
        
        buffer = attr.value();
        Rcpp::Shield<SEXP> buf(Rf_allocVector(STRSXP, 1));
        SET_STRING_ELT(buf, 0, Rf_mkChar(buffer.c_str()));
        SET_VECTOR_ELT(t_c, attr_itr, buf);
        
        if (attr.name() == r_str) {
          // get r attr e.g. "A1" and return colnames "A"
          std::string colrow = attr.value();
          // remove numeric from string
          colrow.erase(std::remove_if(colrow.begin(),
                                      colrow.end(),
                                      &isdigit),
                                      colrow.end());
          SET_STRING_ELT(colnames, itr_cols, Rf_mkChar(colrow.c_str()));
        }
        
        ++attr_itr;
      }
      
      // val ------------------------------------------------------------------
      if (nn > 0) {
        auto val_itr = 0;
        for (auto val: col.children()) {
          
          
          auto ff_itr = 0;
          
          // additional attributes to <f t="shared" ...>
          for (auto cattr : val.attributes())
          {
            buffer = cattr.name();
            SET_STRING_ELT(atr_name, ff_itr, Rf_mkChar(buffer.c_str()));
            
            buffer = cattr.value();
            Rcpp::Shield<SEXP> buf(Rf_allocVector(STRSXP, 1));
            SET_STRING_ELT(buf, 0, Rf_mkChar(buffer.c_str()));
            SET_VECTOR_ELT(a_c, ff_itr, buf);
            
            ++ff_itr;
          }
          
          buffer = val.name();
          SET_STRING_ELT(val_name, val_itr, Rf_mkChar(buffer.c_str()));
          
          // is nodes contain additional t node.
          // TODO: check if multiple t nodes are possible, for now return one.
          // the t-node can bring its very own attributes
          Rcpp::Shield<SEXP> buf(Rf_allocVector(STRSXP, 1));
          if (val.child("t")) {
            buffer = val.child("t").child_value();
          } else {
            buffer = val.child_value();
          }
          SET_STRING_ELT(buf, 0, Rf_mkChar(buffer.c_str()));
          SET_VECTOR_ELT(v_c, val_itr, buf);
          
          ++val_itr;
        }
        
        /* row is done */
      }
      
      // assign names
      ::Rf_setAttrib(v_c, R_NamesSymbol, val_name);
      ::Rf_setAttrib(t_c, R_NamesSymbol, typ_name);
      if(ff > 0) ::Rf_setAttrib(a_c, R_NamesSymbol, atr_name);
      
      ::Rf_setAttrib(cc_cell, R_NamesSymbol, cc_cell_nam);
      
      // assign everything to cc_cell
      SET_VECTOR_ELT(cc_cell, 0, v_c);
      SET_VECTOR_ELT(cc_cell, 1, t_c);
      if(ff > 0) SET_VECTOR_ELT(cc_cell, 2, a_c);
      
      // assign cc_cell to cc_r
      SET_VECTOR_ELT(cc_r, itr_cols, cc_cell);
      
      ++itr_cols;
    }
    
    
    /* ---------------------------------------------------------------------- */
    
    ::Rf_setAttrib(cc_r, R_NamesSymbol, colnames);
    SET_VECTOR_ELT(cc, itr_rows, cc_r);
    
    ++itr_rows;
  }
  
  ::Rf_setAttrib(row_attributes, R_NamesSymbol, rownames);
  ::Rf_setAttrib(cc, R_NamesSymbol, rownames);
  
  wb.field("row_attr") = row_attributes;
  wb.field("cc")  = cc;
  
}

// [[Rcpp::export]]
SEXP si_to_txt(XPtrXML doc) {
  
  auto sst = doc->child("sst");
  auto n = std::distance(sst.begin(), sst.end());
  
  Rcpp::CharacterVector res(n);
  
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


// [[Rcpp::export]]
SEXP getNodes(std::string xml, std::string tagIn){
  
  // This function loops over all characters in xml, looking for tag
  // tag should look liked <tag>
  // tagEnd is then generated to be <tag/>
  
  
  if(xml.length() == 0)
    return Rcpp::wrap(NA_STRING);
  
  xml = " " + xml;
  std::vector<std::string> r;
  size_t pos = 0;
  size_t endPos = 0;
  std::string tag = tagIn;
  std::string tagEnd = tagIn.insert(1,"/");
  
  size_t k = tag.length();
  size_t l = tagEnd.length();
  
  while(1){
    
    pos = xml.find(tag, pos+1);
    endPos = xml.find(tagEnd, pos+k);
    
    if((pos == std::string::npos) | (endPos == std::string::npos))
      break;
    
    r.push_back(xml.substr(pos, endPos-pos+l).c_str());
    
  }  
  
  Rcpp::CharacterVector out = Rcpp::wrap(r);  
  return markUTF8(out);
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
std::vector<std::string> getChildlessNode_ss(std::string xml, std::string tag){
  
  size_t k = tag.length();
  std::vector<std::string> r;
  size_t pos = 0;
  size_t endPos = 0;
  std::string tagEnd = "/>";
  
  while(1){
    
    pos = xml.find(tag, pos+1);    
    if(pos == std::string::npos)
      break;
    
    endPos = xml.find(tagEnd, pos+k);
    
    r.push_back(xml.substr(pos, endPos-pos+2).c_str());
    
  }
  
  return r ;  
  
}


// [[Rcpp::export]]
Rcpp::CharacterVector getChildlessNode(std::string xml, std::string tag) {
  
  size_t k = tag.length();
  if(xml.length() == 0)
    return Rcpp::wrap(NA_STRING);
  
  size_t begPos = 0, endPos = 0;
  
  std::vector<std::string> r;
  std::string res = "";
  
  // check "<tag "
  std::string begTag = "<" + tag + " ";
  std::string endTag = ">";
  
  // initial check, which kind of tags to expect
  begPos = xml.find(begTag, begPos);
  
  // if begTag was found
  if(begPos != std::string::npos) {
    
    endPos = xml.find(endTag, begPos);
    res = xml.substr(begPos, (endPos - begPos) + endTag.length());
    
    // check if last 2 characters are "/>"
    // <foo/> or <foo></foo>
    if (res.substr( res.length() - 2 ).compare("/>") != 0) {
      // check </tag>
      endTag = "</" + tag + ">";
    }
    
    // try with <foo ... />
    while( 1 ) {
      
      begPos = xml.find(begTag, begPos);
      endPos = xml.find(endTag, begPos);
      
      if(begPos == std::string::npos) 
        break;
      
      // read from initial "<" to final ">"
      res = xml.substr(begPos, (endPos - begPos) + endTag.length());
      
      begPos = endPos + endTag.length();
      r.push_back(res);
    }
  }
  
  
  Rcpp::CharacterVector out = Rcpp::wrap(r);  
  return markUTF8(out);
  
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
