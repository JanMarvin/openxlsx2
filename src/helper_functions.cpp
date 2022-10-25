#include "openxlsx2.h"
#include <algorithm>

// [[Rcpp::export]]
SEXP openxlsx2_type(SEXP x) {

  const SEXP names = Rf_getAttrib(x, R_NamesSymbol);
  auto ncol = Rf_length(x);

  Rcpp::IntegerVector type(ncol);
  if (!Rf_isNull(names)) type.attr("names") = names;

  for (auto i = 0; i < ncol; ++i) {

    // check if dim != NULL
    SEXP z;
    if (Rf_isNull(names)) {
      z = x;
    } else {
      z = VECTOR_ELT(x, i);
    }

    switch (TYPEOF(z)) {

    // logical
    case LGLSXP:
      type[i] =  3;
      break;

      // character, formula, hyperlink, array_formula
    case CPLXSXP:
    case STRSXP:
      if (Rf_inherits(z, "formula")) {
        type[i] = 5;
      } else if (Rf_inherits(z, "hyperlink")) {
        type[i] = 10;
      } else if (Rf_inherits(z, "array_formula")) {
        type[i] = 11;
      } else {
        type[i] = 4;
      }
      break;

      // raw, integer, numeric, Date, POSIXct, accounting,
      //  percentage, scientific, comma
    case RAWSXP:
    case INTSXP:
    case REALSXP: {
      if (Rf_inherits(z, "Date")) {
      type[i] = 0;
    } else if (Rf_inherits(z, "POSIXct")) {
      type[i] = 1;
    } else if (Rf_inherits(z, "accounting")) {
      type[i] = 6;
    } else if (Rf_inherits(z, "percentage")) {
      type[i] = 7;
    } else if (Rf_inherits(z, "scientific")) {
      type[i] = 8;
    } else if (Rf_inherits(z, "comma")) {
      type[i] = 9;
    } else if (Rf_inherits(z, "factor")) {
      type[i] = 12;
    } else {
      type[i] = 2; // numeric or integer
    }
    break;

    }

    // whatever is not covered from above
    default: {
      type[i] = 4;
      break;
    }

    }

  }

  return type;
}

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

// driver function for col_to_int
uint32_t uint_col_to_int(std::string& a) {

  char A = 'A';
  int aVal = (int)A - 1;
  int sum = 0;
  int k = a.length();

  for (int j = 0; j < k; j++) {
    sum *= 26;
    sum += (a[j] - aVal);
  }

  return sum;
}

// [[Rcpp::export]]
Rcpp::IntegerVector col_to_int(Rcpp::CharacterVector x ) {

  // This function converts the Excel column letter to an integer

  std::vector<std::string> r = Rcpp::as<std::vector<std::string> >(x);
  int n = r.size();

  std::string a;
  Rcpp::IntegerVector colNums(n);

  for (int i = 0; i < n; i++) {
    a = r[i];

    // check if the value is digit only, if yes, add it and continue the loop
    // at the top. This avoids slow:
    // suppressWarnings(isTRUE(as.character(as.numeric(x)) == x))
    if (std::all_of(a.begin(), a.end(), ::isdigit))
    {
      colNums[i] = std::stoi(a);
      continue;
    }

    // remove digits from string
    a.erase(std::remove_if(a.begin()+1, a.end(), ::isdigit), a.end());
    transform(a.begin(), a.end(), a.begin(), ::toupper);

    // return index from column name
    colNums[i] = uint_col_to_int(a);
  }

  return colNums;

}

// provide a basic rbindlist for lists of named characters
// xlsxFile <- system.file("extdata", "readTest.xlsx", package = "openxlsx2")
// wb <- wb_load(xlsxFile)
// openxlsx2:::rbindlist(xml_attr(wb$styles_mgr$styles$cellXfs, "xf"))
// [[Rcpp::export]]
SEXP rbindlist(Rcpp::List x) {

  size_t nn = x.size();
  std::vector<std::string> all_names;

  // get unique names and create set
  for (size_t i = 0; i < nn; ++i) {
    if (Rf_isNull(x[i])) continue;
    std::vector<std::string> name_i = Rcpp::as<Rcpp::CharacterVector>(x[i]).attr("names");
    std::unique_copy(name_i.begin(), name_i.end(), std::back_inserter(all_names));
  }

  std::sort(all_names.begin(), all_names.end());
  std::set<std::string> unique_names(std::make_move_iterator(all_names.begin()),
                                     std::make_move_iterator(all_names.end()));

  auto kk = unique_names.size();

  // 1. create the list
  Rcpp::List df(kk);
  for (size_t i = 0; i < kk; ++i)
  {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
  }

  for (size_t i = 0; i < nn; ++i) {
    if (Rf_isNull(x[i])) continue;

    std::vector<std::string> values = Rcpp::as<std::vector<std::string>>(x[i]);
    std::vector<std::string> names = Rcpp::as<Rcpp::CharacterVector>(x[i]).attr("names");

    for (size_t j = 0; j < names.size(); ++j) {
      auto find_res = unique_names.find(names[j]);
      auto mtc = std::distance(unique_names.begin(), find_res);

      Rcpp::as<Rcpp::CharacterVector>(df[mtc])[i] = Rcpp::String(values[j]);
    }

  }

  // 3. Create a data.frame
  df.attr("row.names") = Rcpp::IntegerVector::create(NA_INTEGER, nn);
  df.attr("names") = unique_names;
  df.attr("class") = "data.frame";

  return df;
}

// [[Rcpp::export]]
SEXP copy(SEXP x) {
  return Rf_duplicate(x);
}

// provide a basic rbindlist for lists of named characters
// [[Rcpp::export]]
SEXP dims_to_df(Rcpp::IntegerVector rows, std::vector<std::string> cols, bool fill) {

  size_t kk = cols.size();
  size_t nn = rows.size();

  // 1. create the list
  Rcpp::List df(kk);
  for (size_t i = 0; i < kk; ++i)
  {
    if (fill)
      SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
    else
      SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(nn, NA_STRING));
  }

  for (size_t i = 0; i < nn; ++i) {
    for (size_t j = 0; j < kk; ++j) {
      if (fill)
        Rcpp::as<Rcpp::CharacterVector>(df[j])[i] = cols[j] + std::to_string(rows[i]);
    }
  }

  // 3. Create a data.frame
  df.attr("row.names") = rows;
  df.attr("names") = cols;
  df.attr("class") = "data.frame";

  return df;
}

// similar to dcast converts cc dataframe to z dataframe
// [[Rcpp::export]]
void long_to_wide(Rcpp::DataFrame z, Rcpp::DataFrame tt, Rcpp::DataFrame zz) {

  size_t n = zz.nrow();
  size_t z_cols = z.cols(), z_rows = z.rows();
  size_t col = 0, row = 0;

  Rcpp::IntegerVector rows = zz["rows"];
  Rcpp::IntegerVector cols = zz["cols"];
  Rcpp::CharacterVector vals = zz["val"];
  Rcpp::CharacterVector typs = zz["typ"];

  for (size_t i = 0; i < n; ++i) {
    col = cols[i];
    row = rows[i];

    // need to check for missing values
    if ((col <= z_cols) && (row <= z_rows)) {
      Rcpp::as<Rcpp::CharacterVector>(z[col])[row] = vals[i];
      Rcpp::as<Rcpp::CharacterVector>(tt[col])[row] = typs[i];
    }
  }
}

// similar to dcast converts cc dataframe to z dataframe
// [[Rcpp::export]]
void wide_to_long(Rcpp::DataFrame z, Rcpp::IntegerVector vtyps, Rcpp::DataFrame zz,
                  bool ColNames, int32_t start_col, int32_t start_row,
                  Rcpp::CharacterVector ref) {

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
        break;
      case logical:
        cell.v   = vals;
        cell.c_t = "b";
        break;
      case character:
        cell.c_t = "inlineStr";
        cell.is  = txt_to_is(vals, 0, 1, 1);
        break;
      case hyperlink:
      case formula:
        cell.c_t = "str";
        cell.f   = vals;
        break;
      case array_formula:
        cell.f   = vals;
        cell.f_t = "array";
        cell.f_ref = ref[i];
        break;
      }

      cell.typ = std::to_string(vtyp);
      cell.r =  col + row;

      Rcpp::as<Rcpp::CharacterVector>(zz["row_r"])[pos] = row;
      Rcpp::as<Rcpp::CharacterVector>(zz["c_r"])[pos]   = col;
      if (!cell.v.empty())     Rcpp::as<Rcpp::CharacterVector>(zz["v"])[pos]     = cell.v;
      if (!cell.c_s.empty())   Rcpp::as<Rcpp::CharacterVector>(zz["c_s"])[pos]   = cell.c_s;
      if (!cell.c_t.empty())   Rcpp::as<Rcpp::CharacterVector>(zz["c_t"])[pos]   = cell.c_t;
      if (!cell.is.empty())    Rcpp::as<Rcpp::CharacterVector>(zz["is"])[pos]    = cell.is;
      if (!cell.f.empty())     Rcpp::as<Rcpp::CharacterVector>(zz["f"])[pos]     = cell.f;
      if (!cell.f_t.empty())   Rcpp::as<Rcpp::CharacterVector>(zz["f_t"])[pos]   = cell.f_t;
      if (!cell.f_ref.empty()) Rcpp::as<Rcpp::CharacterVector>(zz["f_ref"])[pos] = cell.f_ref;
      if (!cell.typ.empty())   Rcpp::as<Rcpp::CharacterVector>(zz["typ"])[pos]   = cell.typ;
      if (!cell.r.empty())     Rcpp::as<Rcpp::CharacterVector>(zz["r"])[pos]     = cell.r;

      ++startrow;
    }
    ++startcol;
  }
}

// simple helper function to create a data frame of type character
//' @param colnames a vector of the names of the data frame
//' @param n the length of the data frame
//' @noRd
// [[Rcpp::export]]
Rcpp::DataFrame create_char_dataframe(Rcpp::CharacterVector colnames, R_xlen_t n) {

  R_xlen_t kk = colnames.size();

  // 1. create the list
  Rcpp::List df(kk);
  for (R_xlen_t i = 0; i < kk; ++i)
  {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(n)));
  }

  // 3. Create a data.frame
  df.attr("row.names") = Rcpp::IntegerVector::create(NA_INTEGER, n);
  df.attr("names") = colnames;
  df.attr("class") = "data.frame";

  return df;
}
