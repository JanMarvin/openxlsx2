
#include "openxlsx2.h"


//' @import Rcpp
// [[Rcpp::export]]
SEXP write_worksheet_xml_2( std::string prior,
                            std::string post, 
                            Rcpp::Reference sheet_data,
                            Rcpp::CharacterVector cols_attr,
                            Rcpp::List rows_attr,
                            Rcpp::Nullable<Rcpp::CharacterVector> row_heights_ = R_NilValue,
                            Rcpp::Nullable<Rcpp::CharacterVector> outline_levels_ = R_NilValue,
                            std::string R_fileName = "output"){
  
  
  // open file and write header XML
  const char * s = R_fileName.c_str();
  std::ofstream xmlFile;
  xmlFile.open (s);
  xmlFile << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
  xmlFile << prior;
  
  
  // sheet_data will be in order, just need to check for row_heights
  // CharacterVector cell_col = int_2_cell_ref(sheet_data.field("cols"));
  Rcpp::List cc = sheet_data.field("cc");
  
  xmlFile << "<sheetData>";
  
  for (size_t i = 0; i < rows_attr.length(); ++i) {
    
    xmlFile << set_row(rows_attr[i], cc[i]);
    
  }
  
  // write closing tag and XML post data
  xmlFile << "</sheetData>";
  xmlFile << post;
  
  //close file
  xmlFile.close();
  
  return Rcpp::wrap(0);
  
}












// [[Rcpp::export]]
SEXP buildMatrixNumeric(Rcpp::CharacterVector v, Rcpp::IntegerVector rowInd, Rcpp::IntegerVector colInd,
                        Rcpp::CharacterVector colNames, int nRows, int nCols){
  
  Rcpp::LogicalVector isNA_element = is_na(v);
  if(is_true(any(isNA_element))){
    
    v = v[!isNA_element];
    rowInd = rowInd[!isNA_element];
    colInd = colInd[!isNA_element];
    
  }
  
  int k = v.size();
  Rcpp::NumericMatrix m(nRows, nCols);
  std::fill(m.begin(), m.end(), NA_REAL);
  
  for(int i = 0; i < k; i++)
    m(rowInd[i], colInd[i]) = atof(v[i]);
  
  Rcpp::List dfList(nCols);
  for(int i=0; i < nCols; ++i)
    dfList[i] = m(Rcpp::_,i);
  
  std::vector<int> rowNames(nRows);
  for(int i = 0;i < nRows; ++i)
    rowNames[i] = i+1;
  
  dfList.attr("names") = colNames;
  dfList.attr("row.names") = rowNames;
  dfList.attr("class") = "data.frame";
  
  return Rcpp::wrap(dfList);
  
  
}




// [[Rcpp::export]]
SEXP buildMatrixMixed(Rcpp::CharacterVector v,
                      Rcpp::IntegerVector rowInd,
                      Rcpp::IntegerVector colInd,
                      Rcpp::CharacterVector colNames,
                      int nRows,
                      int nCols,
                      Rcpp::IntegerVector charCols,
                      Rcpp::IntegerVector dateCols){
  
  
  /* List d(10);
   d[0] = v;
   d[1] = vn;
   d[2] = rowInd;
   d[3] = colInd;
   d[4] = colNames;
   d[5] = nRows;
   d[6] = nCols;
   d[7] = charCols;
   d[8] = dateCols;
   d[9] = originAdj;
   return(wrap(d));
   */
  
  int k = v.size();
  std::string dt_str;
  
  // create and fill matrix
  Rcpp::CharacterMatrix m(nRows, nCols);
  std::fill(m.begin(), m.end(), NA_STRING);
  
  for(int i = 0;i < k; i++)
    m(rowInd[i], colInd[i]) = v[i];
  
  
  
  // this will be the return data.frame
  Rcpp::List dfList(nCols); 
  
  
  // loop over each column and check type
  for(int i = 0; i < nCols; i++){
    
    Rcpp::CharacterVector tmp(nRows);
    
    for(int ri = 0; ri < nRows; ri++)
      tmp[ri] = m(ri,i);
    
    Rcpp::LogicalVector notNAElements = !is_na(tmp);
    
    
    // If column is date class and no strings exist in column
    if( (std::find(dateCols.begin(), dateCols.end(), i) != dateCols.end()) &&
        (std::find(charCols.begin(), charCols.end(), i) == charCols.end()) ){
      
      // these are all dates and no characters --> safe to convert numerics
      
      Rcpp::DateVector datetmp(nRows);
      for(int ri=0; ri < nRows; ri++){
        if(!notNAElements[ri]){
          datetmp[ri] = NA_REAL; //IF TRUE, TRUE else FALSE
        }else{
          // dt_str = as<std::string>(m(ri,i));
          dt_str = m(ri,i);
          datetmp[ri] = Rcpp::Date(atoi(dt_str.substr(5,2).c_str()), atoi(dt_str.substr(8,2).c_str()), atoi(dt_str.substr(0,4).c_str()) );
          //datetmp[ri] = Date(atoi(m(ri,i)) - originAdj);
          //datetmp[ri] = Date(as<std::string>(m(ri,i)));
        }
      }
      
      dfList[i] = datetmp;
      
      
      // character columns
    }else if(std::find(charCols.begin(), charCols.end(), i) != charCols.end()){
      
      // determine if column is logical or date
      bool logCol = true;
      for(int ri = 0; ri < nRows; ri++){
        if(notNAElements[ri]){
          if((m(ri, i) != "TRUE") & (m(ri, i) != "FALSE")){
            logCol = false;
            break;
          }
        }
      }
      
      if(logCol){
        
        Rcpp::LogicalVector logtmp(nRows);
        for(int ri=0; ri < nRows; ri++){
          if(!notNAElements[ri]){
            logtmp[ri] = NA_LOGICAL; //IF TRUE, TRUE else FALSE
          }else{
            logtmp[ri] = (tmp[ri] == "TRUE");
          }
        }
        
        dfList[i] = logtmp;
        
      }else{
        
        dfList[i] = tmp;
        
      }
      
    }else{ // else if column NOT character class (thus numeric)
      
      Rcpp::NumericVector ntmp(nRows);
      for(int ri = 0; ri < nRows; ri++){
        if(notNAElements[ri]){
          ntmp[ri] = atof(m(ri, i)); 
        }else{
          ntmp[ri] = NA_REAL; 
        }
      }
      
      dfList[i] = ntmp;
      
    }
    
  }
  
  std::vector<int> rowNames(nRows);
  for(int i = 0;i < nRows; ++i)
    rowNames[i] = i+1;
  
  dfList.attr("names") = colNames;
  dfList.attr("row.names") = rowNames;
  dfList.attr("class") = "data.frame";
  
  return wrap(dfList);
  
}


// [[Rcpp::export]]
Rcpp::IntegerVector matrixRowInds(Rcpp::IntegerVector indices) {
  
  int n = indices.size();
  Rcpp::LogicalVector notDup = !duplicated(indices);
  Rcpp::IntegerVector res(n);
  
  int j = -1;
  for(int i =0; i < n; i ++){
    if(notDup[i])
      j++;
    res[i] = j;
  }
  
  return wrap(res);
  
}





// [[Rcpp::export]]
Rcpp::CharacterVector build_table_xml(std::string table, std::string tableStyleXML, std::string ref, std::vector<std::string> colNames, bool showColNames, bool withFilter){
  
  int n = colNames.size();
  std::string tableCols;
  table += " totalsRowShown=\"0\">";
  
  if(withFilter)
    table += "<autoFilter ref=\"" + ref + "\"/>";
  
  
  for(int i = 0; i < n; i ++){
    tableCols += "<tableColumn id=\"" + itos(i+1) + "\" name=\"" + colNames[i] + "\"/>";
  }
  
  tableCols = "<tableColumns count=\"" + itos(n) + "\">" + tableCols + "</tableColumns>"; 
  
  table = table + tableCols + tableStyleXML + "</table>";
  
  
  Rcpp::CharacterVector out = Rcpp::wrap(table);  
  return markUTF8(out);
  
}
