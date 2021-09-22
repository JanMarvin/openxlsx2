

#include "openxlsx.h"



Rcpp::IntegerVector which_cpp(Rcpp::LogicalVector x) {
  Rcpp::IntegerVector v = Rcpp::seq(0, x.size() - 1);
  return v[x];
}


// [[Rcpp::export]]
SEXP read_workbook(Rcpp::IntegerVector cols_in,
                   Rcpp::IntegerVector rows_in,
                   Rcpp::CharacterVector v,
                   
                   Rcpp::IntegerVector string_inds,
                   Rcpp::LogicalVector is_date,
                   bool hasColNames,
                   char hasSepNames,
                   bool skipEmptyRows,
                   bool skipEmptyCols,
                   int nRows,
                   Rcpp::Function clean_names
){
  
  
  Rcpp::IntegerVector cols = clone(cols_in);
  Rcpp::IntegerVector rows = clone(rows_in);
  
  int nCells = rows.size();
  int nDates = is_date.size();
  
  /* do we have any dates */
  bool has_date;
  if(nDates == 1){
    if(is_true(any(is_na(is_date)))){
      has_date = false;
    }else{
      has_date = true;
    }
  }else if(nDates == nCells){
    has_date = true;
  }else{
    has_date = false;
  }
  
  bool has_strings = true;
  Rcpp::IntegerVector st_inds0 (1);
  st_inds0[0] = string_inds[0];
  if(is_true(all(is_na(st_inds0))))
    has_strings = false;
  
  
  
  Rcpp::IntegerVector uni_cols = sort_unique(cols);
  if(!skipEmptyCols){  // want to keep all columns - just create a sequence from 1:max(cols)
    uni_cols = seq(1, max(uni_cols));
    cols = cols - 1;
  }else{
    cols = match(cols, uni_cols) - 1;
  }
  
  // scale columns from i:j to 1:(j-i+1)
  int nCols = *std::max_element(cols.begin(), cols.end()) + 1;
  
  // scale rows from i:j to 1:(j-i+1)
  Rcpp::IntegerVector uni_rows = sort_unique(rows);
  
  if(skipEmptyRows){
    rows = match(rows, uni_rows) - 1;
    //int nRows = *std::max_element(rows.begin(), rows.end()) + 1;
  }else{
    rows = rows - rows[0];
  }
  
  // Check if first row are all strings
  //get first row number
  
  Rcpp::CharacterVector col_names(nCols);
  Rcpp::IntegerVector removeFlag;
  int pos = 0;
  
  // If we are told col_names exist take the first row and fill any gaps with X.i
  if(hasColNames){
    
    int row_1 = rows[0];
    char name[6];
    
    Rcpp::IntegerVector row1_inds = which_cpp(rows == row_1);
    Rcpp::IntegerVector header_cols = cols[row1_inds];
    Rcpp::IntegerVector header_inds = match(Rcpp::seq(0, nCols), na_omit(header_cols));
    Rcpp::LogicalVector missing_header = is_na(header_inds);
    
    // looping over each column
    for(unsigned short i=0; i < nCols; i++){
      
      if(missing_header[i]){  // a missing header element
        
        sprintf(&(name[0]), "X%hu", i+1);
        // sprintf(&(name[0]), "X%u", i+1);
        // snprintf(&(name[0]), sizeof(&(name[0])), "X%d", i+1);
        // snprintf(&(name[0]), 10, "X%d", i+1);
        col_names[i] = name;
        
      }else{  // this is a header elements 
        
        col_names[i] = v[pos];
        if(col_names[i] == "NA"){
          sprintf(&(name[0]), "X%hu", i+1);
          // sprintf(&(name[0]), "X%du", i+1);
          // snprintf(&(name[0]), sizeof(&(name[0])), "X%d", i+1);
          // snprintf(&(name[0]), 10, "X%d", i+1);
          col_names[i] = name;
        }
        
        pos++;
        
      }
      
    }
    
    // tidy up column names
    col_names = clean_names(col_names,hasSepNames);
    
    //--------------------------------------------------------------------------------
    // Remove elements from rows, cols, v that have been used as headers
    
    // I've used the first pos elements as headers
    // stringInds contains the indexes of v which are strings
    // string_inds <- string_inds[string_inds > pos]
    if(has_strings){
      string_inds = string_inds[string_inds > pos];
      string_inds = string_inds - pos;
    }
    
    
    rows.erase (rows.begin(), rows.begin() + pos);
    rows = rows - 1;
    v.erase (v.begin(), v.begin() + pos);
    
    //If nothing left return a data.frame with 0 rows
    if(rows.size() == 0){
      
      Rcpp::List dfList(nCols);
      Rcpp::IntegerVector rowNames(0);
      
      for(int i = 0; i < nCols; i++){
        dfList[i] = Rcpp::LogicalVector(0); // this is what read.table does (bool type)
      }
      
      dfList.attr("names") = col_names;
      dfList.attr("row.names") = rowNames;
      dfList.attr("class") = "data.frame";
      return wrap(dfList);
    }
    
    cols.erase(cols.begin(), cols.begin() + pos);
    nRows--; // decrement number of rows as first row is now being used as col_names
    nCells = nCells - pos;
    
    // End Remove elements from rows, cols, v that have been used as headers
    //--------------------------------------------------------------------------------
    
    
    
  }else{ // else col_names is FALSE
    char name[6];
    for(int i =0; i < nCols; i++){
      sprintf(&(name[0]), "X%hu", i+1);
       // snprintf(&(name[0]), sizeof(&(name[0])), "X%d", i+1);
      // sprintf(&(name[0]), "X%u", i+1);
      col_names[i] = name;
    }
  }
  
  
  // ------------------ column names complete
  
  
  
  
  
  // Possible there are no string_inds to begin with and value of string_inds is 0
  // Possible we have string_inds but they have now all been used up by headers
  bool allNumeric = false;
  if((string_inds.size() == 0) | all(is_na(string_inds)))
    allNumeric = true;
  
  if(has_date){
    if(is_true(any(is_date)))
      allNumeric = false;
  }
  
  // If we have colnames some elements where used to create these -so we remove the corresponding number of elements
  if(hasColNames & has_date)
    is_date.erase(is_date.begin(), is_date.begin() + pos);
  
  
  
  //Intialise return data.frame
  SEXP m; 
  
  // for(int i = 0; i < rows.size(); i++)
  //   Rcout << "rows[i]: " << rows[i] << endl;
  // 
  // Rcout << "nRows " << nRows << endl;
  // Rcout << "nCols: " << nCols << endl;
  // Rcout << "cols.size(): " << cols.size() << endl;
  // Rcout << "rows.size(): " << rows.size() << endl;
  // Rcout << "is_date.size(): " << is_date.size() << endl;
  // Rcout << "v.size(): " << v.size() << endl;
  // Rcout << "has_date: " << has_date << endl;
  
  if(allNumeric){
    
    m = buildMatrixNumeric(v, rows, cols, col_names, nRows, nCols);
    
  }else{
    
    // If it contains any strings it will be a character column
    Rcpp::IntegerVector char_cols_unique;
    if(all(is_na(string_inds))){
      char_cols_unique = -1;
    }else{
      
      Rcpp::IntegerVector columns_which_are_characters = cols[string_inds - 1];
      char_cols_unique = unique(columns_which_are_characters);
      
    }
    
    //date columns
    Rcpp::IntegerVector date_columns(1);
    if(has_date){
      
      date_columns = cols[is_date];
      date_columns = sort_unique(date_columns);
      
    }else{
      date_columns[0] = -1;
    }
    
    // List d(10);
    // d[0] = v;
    // d[2] = rows;
    // d[3] = cols;
    // d[4] = col_names;
    // d[5] = nRows;
    // d[6] = nCols;
    // d[7] = char_cols_unique;
    // d[8] = date_columns;
    // return(wrap(d));
    // Rcout << "Running buildMatrixMixed" << endl;
    
    m = buildMatrixMixed(v, rows, cols, col_names, nRows, nCols, char_cols_unique, date_columns);
    
  }
  
  return Rcpp::wrap(m) ;
  
  
}







// [[Rcpp::export]]  
int calc_number_rows(Rcpp::CharacterVector x, bool skipEmptyRows){
  
  int n = x.size();
  if(n == 0)
    return(0);
  
  int nRows;
  
  if(skipEmptyRows){
    
    Rcpp::CharacterVector res(n);
    std::string r;
    for(int i = 0; i < n; i++){
      r = x[i];
      r.erase(std::remove_if(r.begin(), r.end(), ::isalpha), r.end());
      res[i] = r;
    }
    
    Rcpp::CharacterVector uRes = unique(res);
    nRows = uRes.size();
    
  }else{
    
    std::string fRef = Rcpp::as<std::string>(x[0]);
    std::string lRef = Rcpp::as<std::string>(x[n-1]);
    fRef.erase(std::remove_if(fRef.begin(), fRef.end(), ::isalpha), fRef.end());
    lRef.erase(std::remove_if(lRef.begin(), lRef.end(), ::isalpha), lRef.end());
    int firstRow = atoi(fRef.c_str());
    int lastRow = atoi(lRef.c_str());
    nRows = lastRow - firstRow + 1;
    
  }
  
  return(nRows);
  
}
