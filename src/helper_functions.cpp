
#include "openxlsx2.h"






// [[Rcpp::export]]
SEXP calc_column_widths(Rcpp::Environment sheet_data
                          , std::vector<std::string> sharedStrings
                          , Rcpp::IntegerVector autoColumns
                          , Rcpp::NumericVector widths
                          , float baseFontCharWidth
                          , float minW
                          , float maxW){


  int n = sheet_data["n_elements"];
  Rcpp::IntegerVector cell_types = sheet_data["t"];
  Rcpp::StringVector cell_values(sheet_data["v"]);
  Rcpp::IntegerVector cell_cols = sheet_data["cols"];

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


// provide a basic rbindlist for lists of named characters
// [[Rcpp::export]]
SEXP rbindlist(Rcpp::List x) {

  auto nn = x.length();
  std::vector<std::string> all_names;

  for (auto i = 0; i < nn; ++i) {
    std::vector<std::string> name_i = Rcpp::as<Rcpp::CharacterVector>(x[i]).attr("names");
    std::copy(name_i.begin(), name_i.end(), std::back_inserter(all_names));
  }

  Rcpp::CharacterVector all_nams = Rcpp::wrap(all_names);
  Rcpp::CharacterVector unique_names = Rcpp::unique(all_nams).sort();

  auto kk = unique_names.length();

  // 1. create the list
  Rcpp::List df(kk);
  for (auto i = 0; i < kk; ++i)
  {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
  }

  for (auto i = 0; i < nn; ++i) {

    Rcpp::CharacterVector values = Rcpp::as<Rcpp::CharacterVector>(x[i]);
    Rcpp::CharacterVector names = values.attr("names");

    // mimic which
    Rcpp::IntegerVector mtc = Rcpp::match(names, unique_names);
    std::vector<size_t> ii = Rcpp::as<std::vector<size_t>>(mtc[!Rcpp::is_na(mtc)]);

    for (auto j = 0; j < ii.size(); ++j) {
      Rcpp::as<Rcpp::CharacterVector>(df[ii[j] -1 ])[i] = values[j];
    }

  }

  // 3. Create a data.frame
  R_xlen_t nrows = Rf_length(df[0]);
  df.attr("row.names") = Rcpp::IntegerVector::create(NA_INTEGER, nrows);
  df.attr("names") = unique_names;
  df.attr("class") = "data.frame";

  return df;
}
