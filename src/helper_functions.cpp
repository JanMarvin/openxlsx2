#include "openxlsx2.h"


// [[Rcpp::export]]
std::string int_to_col(uint32_t cell) {
  std::string col_name = "";

  while (cell > 0)
  {
    auto modulo = (cell - 1) % 26;
    col_name = (char)('A' + modulo) + col_name;
    cell = (cell - modulo) / 26;
  }

  return col_name;
}


// [[Rcpp::export]]
Rcpp::IntegerVector col_to_int(Rcpp::CharacterVector x ){

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
