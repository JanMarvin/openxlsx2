#include "openxlsx2.h"
#include <algorithm>

// For R-devel 4.3 character length on Windows was modified. This caused an
// error when passing large character strings to file.exist() when called in
// read_xml(). We prevent bailing, by checking if the input is to long to be
// a path.
// Most likely this is only required until the dust has settled in R-devel, but
// CRAN checks must succed on R-devel too.
#ifndef R_PATH_MAX
#define R_PATH_MAX PATH_MAX
#endif

//' Check if path is to long to be an R file path
//' @param path the file path used in file.exists()
//' @noRd
// [[Rcpp::export]]
bool to_long(std::string path) {
 return path.size() > R_PATH_MAX;
}

// [[Rcpp::export]]
SEXP openxlsx2_type(SEXP x) {

  const SEXP names = Rf_getAttrib(x, R_NamesSymbol);
  R_xlen_t ncol = Rf_length(x);

  Rcpp::IntegerVector type(ncol);
  if (!Rf_isNull(names)) type.attr("names") = names;

  for (R_xlen_t i = 0; i < ncol; ++i) {

    // check if dim != NULL
    SEXP z;
    if (Rf_isNull(names)) {
      z = x;
    } else {
      z = VECTOR_ELT(x, i);
    }

    SEXP Rclass = Rf_getAttrib(z, R_ClassSymbol);

    switch (TYPEOF(z)) {

    // logical
    case LGLSXP:
      if (Rf_isNull(Rclass)) {
        type[i] = logical; // logical
      } else {
        type[i] = factor; // probably some custom class
      };
      break;

      // character, formula, hyperlink, array_formula
    case CPLXSXP:
    case STRSXP:
      if (Rf_inherits(z, "formula")) {
        type[i] = formula;
      } else if (Rf_inherits(z, "hyperlink")) {
        type[i] = hyperlink;
      } else if (Rf_inherits(z, "array_formula")) {
        type[i] = array_formula;
      } else if (Rf_inherits(z, "cm_formula")) {
        type[i] = cm_formula;
      } else {
        type[i] = character;
      }
      break;

      // raw, integer, numeric, Date, POSIXct, accounting,
      //  percentage, scientific, comma
    case RAWSXP:
    case INTSXP:
    case REALSXP: {
      if (Rf_inherits(z, "Date")) {
        type[i] = short_date;
      } else if (Rf_inherits(z, "POSIXct")) {
        type[i] = long_date;
      } else if (Rf_inherits(z, "accounting")) {
        type[i] = accounting;
      } else if (Rf_inherits(z, "percentage")) {
        type[i] = percentage;
      } else if (Rf_inherits(z, "scientific")) {
        type[i] = scientific;
      } else if (Rf_inherits(z, "comma")) {
        type[i] = comma;
      } else if (Rf_inherits(z, "factor") || !Rf_isNull(Rf_getAttrib(z, Rf_install("labels")))) {
        type[i] = factor;
      } else if (Rf_inherits(z, "hms")) {
        type[i] = hms_time;
      } else if (Rf_inherits(z, "currency")) {
        type[i] = currency;
      } else {
        if (Rf_isNull(Rclass)) {
          type[i] = numeric; // numeric and integer
        } else {
          type[i] = factor; // probably some custom class
        }
      }
      break;
    }

    case VECSXP:
      type[i] = list;
      break;


    // whatever is not covered from above
    default: {
      type[i] = character;
      break;
    }

    }

  }

  return type;
}


// [[Rcpp::export]]
Rcpp::IntegerVector col_to_int(Rcpp::CharacterVector x) {

  // This function converts the Excel column letter to an integer
  R_xlen_t n = static_cast<R_xlen_t>(x.size());

  std::string a;
  Rcpp::IntegerVector colNums(n);

  for (R_xlen_t i = 0; i < n; i++) {
    a = x[i];

    // check if the value is digit only, if yes, add it and continue the loop
    // at the top. This avoids slow:
    // suppressWarnings(isTRUE(as.character(as.numeric(x)) == x))
    if (std::all_of(a.begin(), a.end(), ::isdigit))
    {
      colNums[i] = std::stoi(a);
      continue;
    }

    // return index from column name
    colNums[i] = cell_to_colint(a);
  }

  return colNums;

}

// [[Rcpp::export]]
std::string ox_int_to_col(int32_t cell) {
  uint32_t cell_u32 = static_cast<uint32_t>(cell);
  return int_to_col(cell_u32);
}

// provide a basic rbindlist for lists of named characters
// [[Rcpp::export]]
SEXP rbindlist(Rcpp::List x) {

  R_xlen_t nn = static_cast<R_xlen_t>(x.size());
  std::vector<std::string> all_names;

  // get unique names and create set
  for (R_xlen_t i = 0; i < nn; ++i) {
    if (Rf_isNull(x[i])) continue;
    std::vector<std::string> name_i = Rcpp::as<Rcpp::CharacterVector>(x[i]).attr("names");
    std::unique_copy(name_i.begin(), name_i.end(), std::back_inserter(all_names));
  }

  std::sort(all_names.begin(), all_names.end());
  std::set<std::string> unique_names(std::make_move_iterator(all_names.begin()),
                                     std::make_move_iterator(all_names.end()));

  R_xlen_t kk = static_cast<R_xlen_t>(unique_names.size());

  // 1. create the list
  Rcpp::List df(kk);
  for (R_xlen_t i = 0; i < kk; ++i)
  {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
  }

  for (R_xlen_t i = 0; i < nn; ++i) {
    if (Rf_isNull(x[i])) continue;

    std::vector<std::string> values = Rcpp::as<std::vector<std::string>>(x[i]);
    std::vector<std::string> names = Rcpp::as<Rcpp::CharacterVector>(x[i]).attr("names");

    for (size_t j = 0; j < names.size(); ++j) {
      auto find_res = unique_names.find(names[j]);
      R_xlen_t mtc = std::distance(unique_names.begin(), find_res);

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

// [[Rcpp::export]]
Rcpp::CharacterVector needed_cells(const std::string& range) {
  std::vector<std::string> cells;

  // Parse the input range
  std::string startCell, endCell;
  size_t colonPos = range.find(':');
  if (colonPos != std::string::npos) {
    startCell = range.substr(0, colonPos);
    endCell = range.substr(colonPos + 1);
  } else {
    startCell = range;
    endCell = range;
  }

  if (!validate_dims(startCell) || !validate_dims(endCell)) {
    Rcpp::stop("Invalid input: dims must be something like A1 or A1:B2.");
  }

  // Extract column and row numbers from start and end cells
  uint32_t startRow, endRow;
  int32_t startCol = 0, endCol = 0;

  startCol = cell_to_colint(startCell);
  endCol   = cell_to_colint(endCell);

  startRow = cell_to_rowint(startCell);
  endRow   = cell_to_rowint(endCell);

  // Generate spreadsheet cell references
  for (int32_t col = startCol; col <= endCol; ++col) {
    for (uint32_t row = startRow; row <= endRow; ++row) {
      std::string cell = int_to_col(col);
      cell += std::to_string(row);
      cells.push_back(cell);
    }
  }

  return Rcpp::wrap(cells);
}

// provide a basic rbindlist for lists of named characters
// [[Rcpp::export]]
SEXP dims_to_df(Rcpp::IntegerVector rows, Rcpp::CharacterVector cols, Rcpp::Nullable<Rcpp::CharacterVector> filled, bool fill,
                Rcpp::Nullable<Rcpp::IntegerVector> fcols) {

  R_xlen_t kk = static_cast<R_xlen_t>(cols.size());
  R_xlen_t nn = static_cast<R_xlen_t>(rows.size());

  bool has_fcols  = fcols.isNotNull();
  bool has_filled = filled.isNotNull();

  // 1. create the list
  Rcpp::List df(kk);
  for (R_xlen_t i = 0; i < kk; ++i)
  {
    if (fill)
      SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
    else
      SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(nn, NA_STRING));
  }

  if (fill) {
    if (has_filled) {

      std::vector<std::string> flld = Rcpp::as<std::vector<std::string>>(filled.get());
      std::unordered_set<std::string> flls(flld.begin(), flld.end());

      // with has_filled we always have to run this loop
      for (R_xlen_t i = 0; i < kk; ++i) {
        Rcpp::CharacterVector cvec = Rcpp::as<Rcpp::CharacterVector>(df[i]);
        std::string coli = Rcpp::as<std::string>(cols[i]);
        for (R_xlen_t j = 0; j < nn; ++j) {
          std::string cell = coli + std::to_string(rows[j]);
          if (has_cell(cell, flls))
            cvec[j] = cell;
        }
      }

    } else { // insert cells into data frame

      std::vector<size_t> fcls;
      if (has_fcols) {
        fcls = Rcpp::as<std::vector<size_t>>(fcols.get());
      }

      for (R_xlen_t i = 0; i < kk; ++i) {
        if (has_fcols && std::find(fcls.begin(), fcls.end(), i) == fcls.end())
          continue;
        Rcpp::CharacterVector cvec = Rcpp::as<Rcpp::CharacterVector>(df[i]);
        std::string coli = Rcpp::as<std::string>(cols[i]);
        for (R_xlen_t j = 0; j < nn; ++j) {
          cvec[j] = coli + std::to_string(rows[j]);
        }
      }
    }

  } // else return data frame filled with NA_character_

  // 3. Create a data.frame
  df.attr("row.names") = rows;
  df.attr("names") = cols;
  df.attr("class") = "data.frame";

  return df;
}

// similar to dcast converts cc dataframe to z dataframe
// [[Rcpp::export]]
void long_to_wide(Rcpp::DataFrame z, Rcpp::DataFrame tt, Rcpp::DataFrame zz) {

  R_xlen_t n = static_cast<R_xlen_t>(zz.nrow());
  R_xlen_t col = 0, row = 0;

  Rcpp::IntegerVector rows = zz["rows"];
  Rcpp::IntegerVector cols = zz["cols"];
  Rcpp::CharacterVector vals = zz["val"];
  Rcpp::CharacterVector typs = zz["typ"];

  for (R_xlen_t i = 0; i < n; ++i) {
    col = cols[i];
    row = rows[i];

    // need to check for missing values
    if (row != NA_INTEGER && col != NA_INTEGER) {
      SET_STRING_ELT(Rcpp::as<Rcpp::CharacterVector>(z[col]), row, STRING_ELT(vals, i));
      SET_STRING_ELT(Rcpp::as<Rcpp::CharacterVector>(tt[col]), row, STRING_ELT(typs, i));
    }
  }
}

// function to apply on vector
// @param x a character vector as input
// [[Rcpp::export]]
Rcpp::LogicalVector is_charnum(Rcpp::CharacterVector x) {
  Rcpp::LogicalVector out(x.size());
  for (R_xlen_t i = 0; i < x.size(); ++i) {
    out[i] = is_double(Rcpp::as<std::string>(x[i]));
  }
  return out;
}

// similar to dcast converts cc dataframe to z dataframe
// [[Rcpp::export]]
void wide_to_long(
    Rcpp::DataFrame z,
    std::vector<int32_t> vtyps,
    Rcpp::DataFrame zz,
    bool ColNames,
    int32_t start_col,
    int32_t start_row,
    Rcpp::Nullable<Rcpp::CharacterVector> refed,
    int32_t string_nums,
    bool na_null,
    bool na_missing,
    std::string na_strings,
    bool inline_strings,
    std::string c_cm,
    std::vector<std::string> dims
) {

  R_xlen_t n = static_cast<R_xlen_t>(z.nrow());
  R_xlen_t m = static_cast<R_xlen_t>(z.ncol());
  bool has_dims = static_cast<R_xlen_t>(dims.size()) == (n * m);

  std::vector<std::string> srows(static_cast<size_t>(n));
  for (size_t j = 0; j < static_cast<size_t>(n); ++j) {
    srows[j] = std::to_string(static_cast<size_t>(start_row) + j);
  }

  std::vector<std::string> scols(static_cast<size_t>(m));
  for (size_t i = 0; i < static_cast<size_t>(m); ++i) {
    scols[i] = int_to_col(static_cast<size_t>(start_col) + i);
  }

  bool has_refs  = refed.isNotNull();

  std::vector<std::string> ref;
  if (has_refs) ref = Rcpp::as<std::vector<std::string>>(refed.get());

  int32_t in_string_nums = string_nums;

  // pointer magic. even though these are extracted, they just point to the
  // memory in the data frame
  Rcpp::CharacterVector zz_row_r = Rcpp::as<Rcpp::CharacterVector>(zz["row_r"]);
  Rcpp::CharacterVector zz_c_cm  = Rcpp::as<Rcpp::CharacterVector>(zz["c_cm"]);
  Rcpp::CharacterVector zz_c_r   = Rcpp::as<Rcpp::CharacterVector>(zz["c_r"]);
  Rcpp::CharacterVector zz_v     = Rcpp::as<Rcpp::CharacterVector>(zz["v"]);
  Rcpp::CharacterVector zz_c_t   = Rcpp::as<Rcpp::CharacterVector>(zz["c_t"]);
  Rcpp::CharacterVector zz_is    = Rcpp::as<Rcpp::CharacterVector>(zz["is"]);
  Rcpp::CharacterVector zz_f     = Rcpp::as<Rcpp::CharacterVector>(zz["f"]);
  Rcpp::CharacterVector zz_f_t   = Rcpp::as<Rcpp::CharacterVector>(zz["f_t"]);
  Rcpp::CharacterVector zz_f_ref = Rcpp::as<Rcpp::CharacterVector>(zz["f_ref"]);
  Rcpp::CharacterVector zz_typ   = Rcpp::as<Rcpp::CharacterVector>(zz["typ"]);
  Rcpp::CharacterVector zz_r     = Rcpp::as<Rcpp::CharacterVector>(zz["r"]);

  // Convert na_strings only once outside the loop.
  na_strings = inline_strings ? txt_to_is(na_strings, 0, 1, 1) : txt_to_si(na_strings, 0, 1, 1);

  R_xlen_t idx = 0;

  SEXP blank_sexp      = Rf_mkChar("");
  SEXP array_sexp      = Rf_mkChar("array");
  SEXP inlineStr_sexp  = Rf_mkChar("inlineStr");
  SEXP bool_sexp       = Rf_mkChar("b");
  SEXP expr_sexp       = Rf_mkChar("e");
  SEXP sharedStr_sexp  = Rf_mkChar("s");
  SEXP string_sexp     = Rf_mkChar("str");

  SEXP na_sexp         = Rf_mkChar("#N/A");
  SEXP num_sexp        = Rf_mkChar("#NUM!");
  SEXP value_sexp      = Rf_mkChar("#VALUE!");
  SEXP na_strings_sexp = Rf_mkChar(na_strings.c_str());


  for (R_xlen_t i = 0; i < m; ++i) {

    Rcpp::CharacterVector cvec = Rcpp::as<Rcpp::CharacterVector>(z[i]);
    const std::string& col = scols[static_cast<size_t>(i)];
    int8_t vtyp_i = static_cast<int8_t>(vtyps[static_cast<size_t>(i)]);

    for (R_xlen_t j = 0; j < n; ++j, ++idx) {

      checkInterrupt(idx);

      // if colname is provided, the first row is always a character
      int8_t vtyp = (ColNames && j == 0) ? character : vtyp_i;
      SEXP vals_sexp = STRING_ELT(cvec, j);
      const char* vals = CHAR(vals_sexp);

      const std::string& row = srows[static_cast<size_t>(j)];

      R_xlen_t pos = (j * m) + i;

      // there should be no unicode character in ref_str
      std::string ref_str = "";
      if (vtyp == array_formula || vtyp == cm_formula) {
        ref_str = has_refs ? ref[static_cast<size_t>(i)] : col + row;
      }

      // factors can be numeric or string or both. tables require the
      // column name to be character and once we have overwritten for
      // a factor, we have to reset string_nums.
      if (!(ColNames && j == 0) && vtyp == factor)
        string_nums = 1;
      else
        string_nums = in_string_nums;

      switch(vtyp)
      {

      case currency:
      case short_date:
      case long_date:
      case accounting:
      case percentage:
      case scientific:
      case comma:
      case hms_time:
      case numeric:
        // v
        SET_STRING_ELT(zz_v, pos, vals_sexp);
        break;
      case logical:
        // v and c_t = "b"
        SET_STRING_ELT(zz_v,   pos, vals_sexp);
        SET_STRING_ELT(zz_c_t, pos, bool_sexp);
        break;
      case factor:
      case character:

        // test if string can be written as number
        if (string_nums && is_double(vals)) {
          // v
          SET_STRING_ELT(zz_v, pos, vals_sexp);
          vtyp     = (string_nums == 1) ? string_num : numeric;
        } else {
          // check if we write sst or inlineStr
          if (inline_strings) {
              // is and c_t = "inlineStr"
              SET_STRING_ELT(zz_c_t, pos, inlineStr_sexp);
              SET_STRING_ELT(zz_is,  pos, Rf_mkChar(txt_to_is(vals, 0, 1, 1).c_str()));
            } else {
              // v and c_t = "s"
              SET_STRING_ELT(zz_c_t, pos, sharedStr_sexp);
              SET_STRING_ELT(zz_v,   pos, Rf_mkChar(txt_to_si(vals, 0, 1, 1).c_str()));
          }
        }
        break;
      case hyperlink:
      case formula:
        // f and c_t = "str";
        SET_STRING_ELT(zz_c_t, pos, string_sexp);
        SET_STRING_ELT(zz_f,   pos, vals_sexp);
        break;
      case array_formula:
        // f, f_t = "array", and f_ref
        SET_STRING_ELT(zz_f,     pos, vals_sexp);
        SET_STRING_ELT(zz_f_t,   pos, array_sexp);
        SET_STRING_ELT(zz_f_ref, pos, Rf_mkChar(ref_str.c_str()));
        break;
      case cm_formula:
        // c_cm, f, f_t = "array", and f_ref
        SET_STRING_ELT(zz_c_cm,  pos, Rf_mkChar(c_cm.c_str()));
        SET_STRING_ELT(zz_f,     pos, vals_sexp);
        SET_STRING_ELT(zz_f_t,   pos, array_sexp);
        SET_STRING_ELT(zz_f_ref, pos, Rf_mkChar(ref_str.c_str()));
        break;
      }

      if (vals_sexp == NA_STRING || strcmp(vals, "_openxlsx_NA") == 0) {

        if (na_missing) {
          // v = "#N/A"
          SET_STRING_ELT(zz_v,   pos, na_sexp);
          SET_STRING_ELT(zz_c_t, pos, expr_sexp);
          // is = "" - required only if inline_strings
          // and vtyp = character || vtyp = factor
          SET_STRING_ELT(zz_is,  pos, blank_sexp);
        } else  {
          if (na_null) {
            // all three v, c_t, and is = "" - there should be nothing in this row
            SET_STRING_ELT(zz_v,   pos, blank_sexp);
            SET_STRING_ELT(zz_c_t, pos, blank_sexp);
            SET_STRING_ELT(zz_is,  pos, blank_sexp);
          } else {
            // c_t = "inlineStr" or "s"
            SET_STRING_ELT(zz_c_t, pos, inline_strings  ? inlineStr_sexp  : sharedStr_sexp);
            // for inlineStr is = na_strings else ""
            SET_STRING_ELT(zz_is,  pos, inline_strings  ? na_strings_sexp : blank_sexp);
            // otherwise v = na_strings else ""
            SET_STRING_ELT(zz_v,   pos, !inline_strings ? na_strings_sexp : blank_sexp);
          }
        }

      } else if (strcmp(vals, "NaN") == 0) {
        // v = "#VALUE!"
        SET_STRING_ELT(zz_v,   pos, value_sexp);
        SET_STRING_ELT(zz_c_t, pos, expr_sexp);
      } else if (strcmp(vals, "-Inf") == 0 || strcmp(vals, "Inf") == 0) {
        // v = "#NUM!"
        SET_STRING_ELT(zz_v,   pos, num_sexp);
        SET_STRING_ELT(zz_c_t, pos, expr_sexp);
      }

      // typ = std::to_string(vtyp)
      SET_STRING_ELT(zz_typ, pos, Rf_mkChar(std::to_string(vtyp).c_str()));

      std::string cell_r = has_dims ? dims[static_cast<size_t>(idx)] : col + row;
      SET_STRING_ELT(zz_r, pos, Rf_mkChar(cell_r.c_str()));

      if (has_dims) {
        // row_r = rm_colnum(r) and c_r = rm_rownum(r)
        SET_STRING_ELT(zz_row_r, pos, Rf_mkChar(rm_colnum(cell_r).c_str()));
        SET_STRING_ELT(zz_c_r, pos, Rf_mkChar(rm_rownum(cell_r).c_str()));
      } else {
        // row_r = row and c_r = col
        SET_STRING_ELT(zz_row_r, pos, Rf_mkChar(row.c_str()));
        SET_STRING_ELT(zz_c_r, pos, Rf_mkChar(col.c_str()));
      }
    } // n
  } // m
}

// simple helper function to create a data frame of type character
//' @param colnames a vector of the names of the data frame
//' @param n the length of the data frame
//' @noRd
// [[Rcpp::export]]
Rcpp::DataFrame create_char_dataframe(Rcpp::CharacterVector colnames, R_xlen_t n) {

  R_xlen_t kk = static_cast<R_xlen_t>(colnames.size());

  // 1. create the list
  Rcpp::List df(kk);
  for (R_xlen_t i = 0; i < kk; ++i) {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(n)));
  }

  Rcpp::IntegerVector rvec = Rcpp::IntegerVector::create(NA_INTEGER, n);

  // 3. Create a data.frame
  df.attr("row.names") = rvec;
  df.attr("names") = colnames;
  df.attr("class") = "data.frame";

  return df;
}

// TODO styles_xml.cpp should be converted to use these functions

// [[Rcpp::export]]
Rcpp::DataFrame read_xml2df(XPtrXML xml, std::string vec_name, std::vector<std::string> vec_attrs, std::vector<std::string> vec_chlds) {

  std::set<std::string> nam_attrs(vec_attrs.begin(), vec_attrs.end());
  std::set<std::string> nam_chlds(vec_chlds.begin(), vec_chlds.end());

  auto total_length = nam_attrs.size() + nam_chlds.size();
  std::vector<std::string> all_names(total_length);

  std::copy(nam_attrs.begin(), nam_attrs.end(), all_names.begin());
  std::copy(nam_chlds.begin(), nam_chlds.end(), all_names.begin() + static_cast<int32_t>(nam_attrs.size()));

  std::set<std::string> nams(std::make_move_iterator(all_names.begin()),
                             std::make_move_iterator(all_names.end()));


  R_xlen_t nn = std::distance(xml->begin(), xml->end());
  R_xlen_t kk = static_cast<R_xlen_t>(nams.size());
  uint32_t pugi_format_flags = pugi::format_raw | pugi::format_no_escapes;

  Rcpp::CharacterVector rvec(nn);

  // 1. create the list
  Rcpp::List df(kk);
  for (R_xlen_t i = 0; i < kk; ++i)
  {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
  }

  // 2. fill the list
  // <vec_name ...>
  auto itr = 0;
  for (auto xml_node : xml->children(vec_name.c_str())) {
    for (auto attrs : xml_node.attributes()) {

      std::string attr_name = attrs.name();
      std::string attr_value = attrs.value();
      auto find_res = nams.find(attr_name);

      // check if name is already known
      if (nams.count(attr_name) == 0) {
        Rcpp::warning("%s: not found in %s name table", attr_name, vec_name);
      } else {
        R_xlen_t mtc = std::distance(nams.begin(), find_res);
        Rcpp::as<Rcpp::CharacterVector>(df[mtc])[itr] = attr_value;
      }
    }

    for (auto cld : xml_node.children()) {

      std::string cld_name = cld.name();
      auto find_res = nams.find(cld_name);

      // check if name is already known
      if (nams.count(cld_name) == 0) {
        Rcpp::warning("%s: not found in %s name table", cld_name, vec_name);
      } else {
        std::ostringstream oss;
        cld.print(oss, " ", pugi_format_flags);
        std::string cld_value = oss.str();

        R_xlen_t mtc = std::distance(nams.begin(), find_res);
        Rcpp::as<Rcpp::CharacterVector>(df[mtc])[itr] = cld_value;
      }
    }

    rvec[itr] = std::to_string(itr);
    ++itr;

  }

  // 3. Create a data.frame
  df.attr("row.names") = rvec;
  df.attr("names") = nams;
  df.attr("class") = "data.frame";

  return df;
}


// [[Rcpp::export]]
Rcpp::CharacterVector write_df2xml(Rcpp::DataFrame df, std::string vec_name, std::vector<std::string> vec_attrs, std::vector<std::string> vec_chlds) {

  int64_t n = df.nrow();
  Rcpp::CharacterVector z(n);
  uint32_t pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_ws_pcdata | pugi::parse_eol;
  uint32_t pugi_format_flags = pugi::format_raw | pugi::format_no_escapes;

  // openxml 2.8.1
  std::vector<std::string>  attrnams = df.names();
  std::set<std::string> nam_attrs(vec_attrs.begin(), vec_attrs.end());
  std::set<std::string> nam_chlds(vec_chlds.begin(), vec_chlds.end());

  for (int64_t i = 0; i < n; ++i) {
    pugi::xml_document doc;
    pugi::xml_node xml_node = doc.append_child(vec_name.c_str());

    for (auto j = 0; j < df.ncol(); ++j) {

      std::string attr_j = attrnams[static_cast<size_t>(j)];

      // mimic which
      auto res1 = nam_attrs.find(attr_j);
      auto mtc1 = std::distance(nam_attrs.begin(), res1);

      std::vector<int32_t> idx1(static_cast<size_t>(mtc1) + 1);
      std::iota(idx1.begin(), idx1.end(), 0);

      auto res2 = nam_chlds.find(attr_j);
      auto mtc2 = std::distance(nam_chlds.begin(), res2);

      std::vector<int32_t> idx2(static_cast<size_t>(mtc2) + 1);
      std::iota(idx2.begin(), idx2.end(), 0);

      // check if name is already known
      if (nam_attrs.count(attr_j) != 0) {
        Rcpp::CharacterVector cv_s = "";
        cv_s = Rcpp::as<Rcpp::CharacterVector>(df[j])[i];

        // only write attributes where cv_s has a value
        if (cv_s[0] != "") {
          // Rf_PrintValue(cv_s);
          const std::string val_strl = Rcpp::as<std::string>(cv_s);
          xml_node.append_attribute(attr_j.c_str()) = val_strl.c_str();
        }
      }

      if (nam_chlds.count(attr_j) != 0) {
        Rcpp::CharacterVector cv_s = "";
        cv_s = Rcpp::as<Rcpp::CharacterVector>(df[j])[i];

        if (cv_s[0] != "") {

          std::string child_i = Rcpp::as<std::string>(cv_s[0]);

          pugi::xml_document xml_child;
          pugi::xml_parse_result result = xml_child.load_string(child_i.c_str(), pugi_parse_flags);
          if (!result) Rcpp::stop("loading %s child node fail: %s", vec_name, cv_s);

          xml_node.append_copy(xml_child.first_child());

        }
      }

      if (idx1.empty() && idx2.empty())
        Rcpp::warning("%s: not found in %s name table", attr_j, vec_name);
    }

    std::ostringstream oss;
    doc.print(oss, " ", pugi_format_flags);
    z[i] = oss.str();
  }

  return z;
}
