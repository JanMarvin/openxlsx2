
#include "openxlsx2.h"






// [[Rcpp::export]]
SEXP calc_column_widths(Rcpp::Reference sheet_data
                          , std::vector<std::string> sharedStrings
                          , Rcpp::IntegerVector autoColumns
                          , Rcpp::NumericVector widths
                          , float baseFontCharWidth
                          , float minW
                          , float maxW){
  
  
  int n = sheet_data.field("n_elements");
  Rcpp::IntegerVector cell_types = sheet_data.field("t");
  Rcpp::StringVector cell_values(sheet_data.field("v"));
  Rcpp::IntegerVector cell_cols = sheet_data.field("cols");
  
  Rcpp::NumericVector cell_n_character(n);
  Rcpp::CharacterVector r(n);
  int nLen;
  
  std::string tmp;

  // get widths of all values
  for(int i = 0; i < n; i++){

    if(cell_types[i] == 1){ // "s"
      cell_n_character[i] = sharedStrings[atoi(cell_values[i])].length() - 37; //-37 for shared string tags around text
    }else{
      tmp = cell_values[i];
      nLen = tmp.length();
      cell_n_character[i] = std::min(nLen, 11); // For numerics - max width is 11
    }
    
  }
  
  
  // get column for each value

  // reducing to only the columns that are auto
  Rcpp::LogicalVector notNA = !is_na(match(cell_cols, autoColumns));
  cell_cols = cell_cols[notNA];
  cell_n_character = cell_n_character[notNA];
  widths = widths[notNA];
  Rcpp::IntegerVector unique_cell_cols = sort_unique(cell_cols);
  
  size_t k = unique_cell_cols.size();
  Rcpp::NumericVector column_widths(k);

  
  // for each unique column, get all widths for that column and take max
  for(size_t i = 0; i < k; i++){
    Rcpp::NumericVector wTmp = cell_n_character[cell_cols == unique_cell_cols[i]];
    Rcpp::NumericVector thisColWidths = widths[cell_cols == unique_cell_cols[i]];
    column_widths[i] = max(wTmp * thisColWidths / baseFontCharWidth); 
  }
  
  column_widths[column_widths < minW] = minW;
  column_widths[column_widths > maxW] = maxW;    
  
  // assign column names
  column_widths.attr("names") = unique_cell_cols;
  
  return(wrap(column_widths));
  
}





// [[Rcpp::export]]
SEXP convert_to_excel_ref(Rcpp::IntegerVector cols, std::vector<std::string> LETTERS){
  
  int n = cols.size();  
  Rcpp::CharacterVector res(n);
  
  int x;
  int modulo;
  for(int i = 0; i < n; i++){
    x = cols[i];
    std::string columnName;
    
    while(x > 0){  
      modulo = (x - 1) % 26;
      columnName = LETTERS[modulo] + columnName;
      x = (x - modulo) / 26;
    }
    res[i] = columnName;
  }
  
  return res ;
  
}

// [[Rcpp::export]]
Rcpp::IntegerVector convert_from_excel_ref( Rcpp::CharacterVector x ){
  
  // This function converts the Excel column letter to an integer
  
  std::vector<std::string> r = Rcpp::as<std::vector<std::string> >(x);
  int n = r.size();
  int k;
  
  std::string a;
  Rcpp::IntegerVector colNums(n);
  char A = 'A';
  int aVal = (int)A - 1;
  
  for(int i = 0; i < n; i++){
    a = r[i];
    
    // remove digits from string
    a.erase(std::remove_if(a.begin()+1, a.end(), ::isdigit), a.end());
    
    int sum = 0;
    k = a.length();
    
    for (int j = 0; j < k; j++){
      sum *= 26;
      sum += (a[j] - aVal);
      
    }
    colNums[i] = sum;
  }
  
  return colNums;
  
}





// [[Rcpp::export]]
SEXP convert_to_excel_ref_expand(const std::vector<int>& cols, const std::vector<std::string>& LETTERS, const std::vector<std::string>& rows){
  
  int n = cols.size();  
  int nRows = rows.size();
  std::vector<std::string> res(n);
  
  //Convert col number to excel col letters
  size_t x;
  size_t modulo;
  for(int i = 0; i < n; i++){
    x = cols[i];
    std::string columnName;
    
    while(x > 0){  
      modulo = (x - 1) % 26;
      columnName = LETTERS[modulo] + columnName;
      x = (x - modulo) / 26;
    }
    res[i] = columnName;
  }
  
  Rcpp::CharacterVector r(n*nRows);
  Rcpp::CharacterVector names(n*nRows);
  size_t c = 0;
  for(int i=0; i < nRows; i++)
    for(int j=0; j < n; j++){
      r[c] = res[j] + rows[i];
      names[c] = rows[i];
      c++;
    }
    
  r.attr("names") = names;
  return wrap(r) ;
  
}

std::string itos(int i){
  
  // convert int to string
  std::stringstream s;
  s << i;
  return s.str();
  
}

// [[Rcpp::export]]
std::vector<std::string> get_letters(){
  
  std::vector<std::string> LETTERS(26);
  
  LETTERS[0] = "A";
  LETTERS[1] = "B";
  LETTERS[2] = "C";
  LETTERS[3] = "D";
  LETTERS[4] = "E";
  LETTERS[5] = "F";
  LETTERS[6] = "G";
  LETTERS[7] = "H";
  LETTERS[8] = "I";
  LETTERS[9] = "J";
  LETTERS[10] = "K";
  LETTERS[11] = "L";
  LETTERS[12] = "M";
  LETTERS[13] = "N";
  LETTERS[14] = "O";
  LETTERS[15] = "P";
  LETTERS[16] = "Q";
  LETTERS[17] = "R";
  LETTERS[18] = "S";
  LETTERS[19] = "T";
  LETTERS[20] = "U";
  LETTERS[21] = "V";
  LETTERS[22] = "W";
  LETTERS[23] = "X";
  LETTERS[24] = "Y";
  LETTERS[25] = "Z";
  
  return(LETTERS);
  
}


// [[Rcpp::export]]
Rcpp::CharacterVector markUTF8(Rcpp::CharacterVector x, bool clone) {
  Rcpp::CharacterVector out;
  if (clone) {
    out = Rcpp::clone(x);
  } else {
    out = x;
  }
  const size_t n = x.size();
  for (size_t i = 0; i < n; ++i) {
    out[i] = Rf_mkCharCE(x[i], CE_UTF8);
  }
  return out;
}
