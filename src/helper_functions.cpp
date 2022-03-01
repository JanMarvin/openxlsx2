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
    transform(a.begin(), a.end(), a.begin(), ::toupper);

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

// similar to dcast converts cc dataframe to z dataframe
// [[Rcpp::export]]
void long_to_wide(Rcpp::DataFrame z, Rcpp::DataFrame tt, Rcpp::DataFrame zz) {

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

// similar to dcast converts cc dataframe to z dataframe
// [[Rcpp::export]]
void wide_to_long(Rcpp::DataFrame z, Rcpp::IntegerVector vtyps, Rcpp::DataFrame zz,
                  bool ColNames, int32_t start_col, int32_t start_row) {

  auto n = z.nrow();
  auto m = z.ncol();

  auto startcol = start_col;
  for (auto i = 0; i < m; ++i) {

    auto startrow = start_row;
    for (auto j = 0; j < n; ++j) {

      int8_t vtyp = vtyps[i];
      // if colname is provided, the first row is always a character
      if (ColNames & (j == 0)) vtyp = character;

      std::string vals = Rcpp::as<std::string>(Rcpp::as<Rcpp::CharacterVector>(z[i])[j]);
      std::string row = std::to_string(startrow);
      std::string col = int_to_col(startcol);

      auto pos = (j * m) + i;

      // create struct
      celltyp cell;

      switch(vtyp)
      {
      case short_date:
      case long_date:
      case accounting:
      case percentage:
      case scientific:
      case comma:
      case numeric:
        cell.v   = vals;
        cell.c_s = "_openxlsx_NA_";
        cell.c_t = "_openxlsx_NA_";
        cell.is  = "_openxlsx_NA_";
        cell.f   = "_openxlsx_NA_";
        break;
      case boolean:
        cell.v   = vals;
        cell.c_s = "_openxlsx_NA_";
        cell.c_t = "b";
        cell.is  = "_openxlsx_NA_";
        cell.f   = "_openxlsx_NA_";
        break;
      case character:
        cell.v   = "_openxlsx_NA_";
        cell.c_s = "_openxlsx_NA_";
        cell.c_t = "inlineStr";
        cell.is  = "<is><t>" + vals + "</t></is>";
        cell.f   = "_openxlsx_NA_";
        break;
      case formula:
        cell.v   = "_openxlsx_NA_";
        cell.c_s = "_openxlsx_NA_";
        cell.c_t = "str";
        cell.is  = "_openxlsx_NA_";
        cell.f   = vals;
        break;
      }

      cell.typ = std::to_string(vtyp);
      cell.r =  col + row;

      // TODO change only if not "_openxlsx_NA_"
      Rcpp::as<Rcpp::CharacterVector>(zz["row_r"])[pos] = row;
      Rcpp::as<Rcpp::CharacterVector>(zz["c_r"])[pos]   = col;
      if (cell.v   != "_openxlsx_NA_") Rcpp::as<Rcpp::CharacterVector>(zz["v"])[pos]   = cell.v;
      if (cell.c_s != "_openxlsx_NA_") Rcpp::as<Rcpp::CharacterVector>(zz["c_s"])[pos] = cell.c_s;
      if (cell.c_t != "_openxlsx_NA_") Rcpp::as<Rcpp::CharacterVector>(zz["c_t"])[pos] = cell.c_t;
      if (cell.is  != "_openxlsx_NA_") Rcpp::as<Rcpp::CharacterVector>(zz["is"])[pos]  = cell.is;
      if (cell.f   != "_openxlsx_NA_") Rcpp::as<Rcpp::CharacterVector>(zz["f"])[pos]   = cell.f;
      if (cell.typ != "_openxlsx_NA_") Rcpp::as<Rcpp::CharacterVector>(zz["typ"])[pos] = cell.typ;
      if (cell.r   != "_openxlsx_NA_") Rcpp::as<Rcpp::CharacterVector>(zz["r"])[pos]   = cell.r;

      ++startrow;
    }
    ++startcol;
  }
}
