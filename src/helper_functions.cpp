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

//' this returns the index of the character vector that matches
//' @param x x
//' @param row row
//' @param col col
//' @keywords internal
//' @noRd
R_xlen_t select_rows(vec_string &x, std::string row, std::string col) {
  return std::distance(x.begin(), find(x.begin(), x.end(), col + row));
}

//' convert "TRUE"/"FALSE" to "1"/"0"
//' @param input input
//' @keywords internal
//' @noRd
std::string to_int(Rcpp::String input) {
  if (input == "TRUE") return ("1");
  else return("0");
}

//' update loop used in update_cell(), when writing on worksheet data
//' @param cc cc
//' @param x x
//' @param data_class data_class
//' @param rows rows
//' @param cols cols
//' @param colNames colNames
//' @param removeCellStyle removeCellStyle
//' @param cell cell
//' @param hyperlinkstyle hyperlinkstyle
//' @param no_na_strings no_na_strings
//' @param na_strings_ na_strings
//' @keywords internal
//' @noRd
// [[Rcpp::export]]
void update_cell_loop(
    Rcpp::DataFrame cc,
    Rcpp::DataFrame x,
    Rcpp::IntegerVector data_class,
    vec_string rows,
    vec_string cols,
    bool colNames,
    bool removeCellStyle,
    std::string cell,
    bool no_na_strings,
    Rcpp::Nullable<Rcpp::String> na_strings_ = R_NilValue,
    Rcpp::Nullable<Rcpp::String> hyperlinkstyle_ = R_NilValue
) {

  vec_string cc_r = Rcpp::as<vec_string>(cc["r"]);

  auto m = 0;
  for (auto &col : cols) {

    auto dc = data_class[m];

    auto n = 0;
    for (auto &row : rows) {

      // get the initial data from the new data frame
      Rcpp::String value = "";
      value = Rcpp::wrap(Rcpp::as<Rcpp::CharacterVector>(x[m])[n]);

      R_xlen_t sel = select_rows(cc_r, row, col);

      xml_col uu;
      uu.c_s   = "";
      uu.c_t   = "";
      uu.c_cm  = "";
      uu.c_ph  = "";
      uu.c_vm  = "";
      uu.v     = "";
      uu.f     = "";
      uu.f     = "";
      uu.f_ref = "";
      uu.f_ca  = "";
      uu.f_si  = "";
      uu.is    = "";

      // handle NA_STRING for all input types
      if (value == NA_STRING) {

        if (no_na_strings) {
          uu.v   = "#N/A";
          uu.c_t = "e";
        } else {
          if (na_strings_.isNull()) {
            // do not add any value: <c/>
          } else {
            Rcpp::String na_strings(na_strings_);

            uu.c_t = "inlineStr";
            uu.is  = txt_to_is(na_strings, 0, 1, 1);
          }
        }

      } else {

        // either character or column name and first row
        if ((dc == character) || ((colNames) && (n == 0))) {
          uu.c_t = "inlineStr";
          uu.is = txt_to_is(value.get_cstring(), 0, 1, 1);
        } else if (dc == formula) {
          uu.c_t = "str";
          uu.f = value.get_cstring();
        } else if (dc == array_formula) {
          uu.f = value.get_cstring();
          uu.f_t = "array";
          uu.f_ref = cell.c_str();
        } else if (dc == hyperlink) {
          uu.f = value.get_cstring();
          //FIXME always assign the hyperlink style. This might not be
          // desired. We should provide an option to prevent this.
          if (hyperlinkstyle_.isNotNull()) {
            Rcpp::String hyperlinkstyle(hyperlinkstyle_);
            uu.c_s = hyperlinkstyle.get_cstring();
          }
        } else if (dc == logical) {
          uu.v   = to_int(value);
          uu.c_t = "b";
        } else { // numerics, dates, openxlsx custom styles
          if (value == "NaN") {
            uu.v   = "#VALUE!";
            uu.c_t = "e";
          } else if (value == "-Inf" || value == "Inf") {
            uu.v   = "#NUM!";
            uu.c_t = "e";
          } else {
            uu.v = value.get_cstring();
          }
        }

      }

      // write the xml_col to the data frame
      if (removeCellStyle) Rcpp::as<Rcpp::CharacterVector>(cc["c_s"])[sel] = uu.c_s;
      Rcpp::as<Rcpp::CharacterVector>(cc["c_t"])[sel]   = uu.c_t;
      Rcpp::as<Rcpp::CharacterVector>(cc["c_cm"])[sel]  = uu.c_cm;
      Rcpp::as<Rcpp::CharacterVector>(cc["c_ph"])[sel]  = uu.c_ph;
      Rcpp::as<Rcpp::CharacterVector>(cc["c_vm"])[sel]  = uu.c_vm;
      Rcpp::as<Rcpp::CharacterVector>(cc["v"])[sel]     = uu.v;
      Rcpp::as<Rcpp::CharacterVector>(cc["f"])[sel]     = uu.f;
      Rcpp::as<Rcpp::CharacterVector>(cc["f_t"])[sel]   = uu.f_t;
      Rcpp::as<Rcpp::CharacterVector>(cc["f_ref"])[sel] = uu.f_ref;
      Rcpp::as<Rcpp::CharacterVector>(cc["f_ca"])[sel]  = uu.f_ca;
      Rcpp::as<Rcpp::CharacterVector>(cc["f_si"])[sel]  = uu.f_si;
      Rcpp::as<Rcpp::CharacterVector>(cc["is"])[sel]    = uu.is;

      // cell is done
      ++n;
    }
    ++m;
  }
}

// [[Rcpp::export]]
Rcpp::List build_cell_merges(Rcpp::List comps) {

  size_t nMerges = comps.size();
  Rcpp::List res(nMerges);

  for (size_t i =0; i < nMerges; i++) {
    Rcpp::IntegerVector col = col_to_int(comps[i]);
    Rcpp::CharacterVector comp = comps[i];
    Rcpp::IntegerVector row(2);

    for (size_t j = 0; j < 2; j++) {
      std::string rt(comp[j]);
      rt.erase(std::remove_if(rt.begin(), rt.end(), ::isalpha), rt.end());
      row[j] = atoi(rt.c_str());
    }

    size_t ca(col[0]);
    size_t ck = size_t(col[1]) - ca + 1;

    std::vector<int> v(ck) ;
    for (size_t j = 0; j < ck; j++)
      v[j] = j + ca;

    size_t ra(row[0]);

    size_t rk = int(row[1]) - ra + 1;
    std::vector<int> r(rk) ;
    for (size_t j = 0; j < rk; j++)
      r[j] = j + ra;

    Rcpp::CharacterVector M(ck*rk);
    int ind = 0;
    for (size_t j = 0; j < ck; j++) {
      for (size_t k = 0; k < rk; k++) {
        char name[30];
        sprintf(&(name[0]), "%d-%d", r[k], v[j]);
        M(ind) = name;
        ind++;
      }
    }

    res[i] = M;
  }

  return wrap(res) ;

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
