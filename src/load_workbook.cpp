#include "openxlsx.h"
#include "openxlsx2_types.h"


// loadvals(wb$worksheets[[i]]$sheet_data, worksheet_xml, "worksheet", "sheetData", "row", "c")
// [[Rcpp::export]]
void loadvals(Rcpp::Reference wb, XPtrXML doc) {
  
  auto ws = doc->child("worksheet").child("sheetData");
  
  size_t n = std::distance(ws.begin(), ws.end());
  
  std::string r_str = "r";
  
  Rcpp::List cc(n);
  Rcpp::List row_attributes(n);
  Rcpp::CharacterVector rownames(n);
  
  
  auto itr_rows = 0;
  for (auto worksheet: ws.children("row"))
  {
    size_t k = std::distance(worksheet.begin(), worksheet.end());
    
    Rcpp::List cc_r(k);
    Rcpp::CharacterVector colnames(k);
    
    
    /* ---------------------------------------------------------------------- */
    /* read cval, and ctyp -------------------------------------------------- */
    /* ---------------------------------------------------------------------- */
    
    auto itr_cols = 0;
    for (auto col : worksheet.children("c")) {
      
      auto nn = std::distance(col.children().begin(), col.children().end());
      auto aa = std::distance(col.attributes_begin(), col.attributes_end());
      // 2 per default, 3 maximum
      auto ff = std::distance(col.child("f").attributes_begin(), col.child("f").attributes_end());
      auto tt = nn; if (tt == 0) ++tt;
      
      
      Rcpp::List cc_cell(2+ff);
      Rcpp::List v_c(tt), t_c(aa), a_c(ff);
      Rcpp::CharacterVector val_name(tt), typ_name(aa), atr_name(ff);
      
      
      // typ: attribute ------------------------------------------------------
      auto attr_itr = 0;
      for (auto attr : col.attributes())
      {
        typ_name[attr_itr] = attr.name();
        t_c[attr_itr] = attr.value();
        
        if (attr.name() == r_str) {
          // get r attr e.g. "A1" and return colnames "A"
          std::string colrow = attr.value();
          // remove numeric from string
          colrow.erase(std::remove_if(colrow.begin(),
                                      colrow.end(),
                                      &isdigit),
                                      colrow.end());
          colnames[itr_cols]= colrow;
        }
        
        ++attr_itr;
      }
      
      // val -------------------------------------------------------------------
      
      if (nn > 0) {
        auto val_itr = 0;
        for (auto val: col.children())
        {
          
          std::string val_s = "";
          std::string val_n = "";
          
          auto ff_itr = 0;
          
          // additional attributes to <f t="shared" ...>
          for (auto cattr : val.attributes())
          {
            atr_name[ff_itr] = cattr.name();
            a_c[ff_itr] = cattr.value();
            ++ff_itr;
          }
          
          val_n = val.name();
          
          // is nodes contain additional t node.
          // TODO: check if multiple t nodes are possible, for now return one.
          if (val.child("t")) {
            pugi::xml_node tval = val.child("t");
            val_s = tval.child_value();
          } else {
            val_s = val.child_value();
          }
          
          val_name[val_itr] = val_n;
          v_c[val_itr] = val_s;
          ++val_itr;
        }
      }
      
      v_c.attr("names") = val_name;
      t_c.attr("names") = typ_name;
      if(ff > 0) a_c.attr("names") = atr_name;
      
      cc_cell[0] = v_c;
      cc_cell[1] = t_c;
      if(ff > 0) cc_cell[2] = a_c;
      
      Rcpp::CharacterVector cc_cell_nam = {"val", "typ"};
      if (ff > 0) cc_cell_nam = {"val", "typ", "attr"};
      cc_cell.attr("names") = cc_cell_nam;
      
      cc_r[itr_cols] = cc_cell;
      
      /* row is done */
      ++itr_cols;
    }
    
    
    /* row attributes ------------------------------------------------------- */
    
    auto nn = std::distance(worksheet.attributes_begin(), worksheet.attributes_end());
    Rcpp::List row_attr(nn);
    Rcpp::CharacterVector row_attr_nam(nn);
    
    auto attr_itr = 0;
    for (auto attr : worksheet.attributes())
    {
      row_attr_nam[attr_itr] = attr.name();
      row_attr[attr_itr] = attr.value();
      ++attr_itr;
      
      // push row name back (will assign it to list)
      if (attr.name() == r_str)
        rownames[itr_rows] = attr.value();
      
    }
    row_attr.attr("names") = row_attr_nam;
    
    row_attributes[itr_rows] = row_attr;
    
    /* ---------------------------------------------------------------------- */
    
    cc_r.attr("names") = colnames;
    cc[itr_rows]  = cc_r;
    
    ++itr_rows;
  }
  
  row_attributes.attr("names") = rownames;
  
  cc.attr("names") = rownames;
  
  wb.field("row_attr") = row_attributes;
  wb.field("cc")  = cc;
  
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
