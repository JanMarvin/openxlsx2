#include "openxlsx2.h"
#include <algorithm>

#include <arrow/api.h>
#include <arrow/io/api.h>
#include <arrow/ipc/api.h>
#include <arrow/filesystem/api.h>
#include <arrow/table.h>
#include <parquet/arrow/writer.h>

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

// write xml_data structure into xml_data.parquet file
void flush_to_parquet(std::vector<xml_col> xml_data, std::string tmpfile) {

  // Step 2: Create Arrow Builders for each field in xml_col
  arrow::MemoryPool* pool = arrow::default_memory_pool();

  arrow::StringBuilder r_builder(pool), row_r_builder(pool), c_r_builder(pool),
  c_s_builder(pool), c_t_builder(pool), c_cm_builder(pool), c_ph_builder(pool),
  c_vm_builder(pool), v_builder(pool), f_builder(pool), f_t_builder(pool),
  f_ref_builder(pool), f_ca_builder(pool), f_si_builder(pool), is_builder(pool),
  typ_builder(pool);

  for (const auto& col : xml_data) {
    (void) r_builder.Append(col.r);
    (void) row_r_builder.Append(col.row_r);
    (void) c_r_builder.Append(col.c_r);
    (void) c_s_builder.Append(col.c_s);
    (void) c_t_builder.Append(col.c_t);
    (void) c_cm_builder.Append(col.c_cm);
    (void) c_ph_builder.Append(col.c_ph);
    (void) c_vm_builder.Append(col.c_vm);
    (void) v_builder.Append(col.v);
    (void) f_builder.Append(col.f);
    (void) f_t_builder.Append(col.f_t);
    (void) f_ref_builder.Append(col.f_ref);
    (void) f_ca_builder.Append(col.f_ca);
    (void) f_si_builder.Append(col.f_si);
    (void) is_builder.Append(col.is);
    (void) typ_builder.Append(col.typ);
  }

  // Step 3: Finalize the Arrow Arrays
  std::shared_ptr<arrow::Array> r_array, row_r_array, c_r_array, c_s_array,
  c_t_array, c_cm_array, c_ph_array, c_vm_array, v_array,
  f_array, f_t_array, f_ref_array, f_ca_array, f_si_array, is_array, typ_array;

  (void) r_builder.Finish(&r_array);
  (void) row_r_builder.Finish(&row_r_array);
  (void) c_r_builder.Finish(&c_r_array);
  (void) c_s_builder.Finish(&c_s_array);
  (void) c_t_builder.Finish(&c_t_array);
  (void) c_cm_builder.Finish(&c_cm_array);
  (void) c_ph_builder.Finish(&c_ph_array);
  (void) c_vm_builder.Finish(&c_vm_array);
  (void) v_builder.Finish(&v_array);
  (void) f_builder.Finish(&f_array);
  (void) f_t_builder.Finish(&f_t_array);
  (void) f_ref_builder.Finish(&f_ref_array);
  (void) f_ca_builder.Finish(&f_ca_array);
  (void) f_si_builder.Finish(&f_si_array);
  (void) is_builder.Finish(&is_array);
  (void) typ_builder.Finish(&typ_array);

  // Step 4: Create Schema and Table
  auto schema = arrow::schema({
    arrow::field("r", arrow::utf8()),
    arrow::field("row_r", arrow::utf8()),
    arrow::field("c_r", arrow::utf8()),
    arrow::field("c_s", arrow::utf8()),
    arrow::field("c_t", arrow::utf8()),
    arrow::field("c_cm", arrow::utf8()),
    arrow::field("c_ph", arrow::utf8()),
    arrow::field("c_vm", arrow::utf8()),
    arrow::field("v", arrow::utf8()),
    arrow::field("f", arrow::utf8()),
    arrow::field("f_t", arrow::utf8()),
    arrow::field("f_ref", arrow::utf8()),
    arrow::field("f_ca", arrow::utf8()),
    arrow::field("f_si", arrow::utf8()),
    arrow::field("is", arrow::utf8()),
    arrow::field("typ", arrow::utf8()),
  });

  auto table = arrow::Table::Make(schema, {r_array, row_r_array, c_r_array, c_s_array, c_t_array,
                                  c_cm_array, c_ph_array, c_vm_array, v_array, f_array, f_t_array,
                                  f_ref_array, f_ca_array, f_si_array, is_array, typ_array});

  // Step 5: Write to Parquet
  std::shared_ptr<arrow::io::FileOutputStream> outfile;
  PARQUET_ASSIGN_OR_THROW(
    outfile, arrow::io::FileOutputStream::Open(tmpfile)
  );

  PARQUET_THROW_NOT_OK(parquet::arrow::WriteTable(*table, arrow::default_memory_pool(), outfile, 1024));

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
    std::vector<std::string> dims,
    std::string tmpfile
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

  std::vector<xml_col> xml_cols(n * m);

  // Convert na_strings only once outside the loop.
  na_strings = inline_strings ? txt_to_is(na_strings, 0, 1, 1) : txt_to_si(na_strings, 0, 1, 1);

  R_xlen_t idx = 0;

  std::string blank_sexp      = "";
  std::string array_sexp      = "array";
  std::string inlineStr_sexp  = "inlineStr";
  std::string bool_sexp       = "b";
  std::string expr_sexp       = "e";
  std::string sharedStr_sexp  = "s";
  std::string string_sexp     = "str";

  std::string na_sexp         = "#N/A";
  std::string num_sexp        = "#NUM!";
  std::string value_sexp      = "#VALUE!";
  std::string na_strings_sexp = na_strings;


  for (R_xlen_t i = 0; i < m; ++i) {

    // FIXME somehow a logical can still appear as non character vector here
    Rcpp::CharacterVector cvec = Rcpp::as<Rcpp::CharacterVector>(z[i]);
    const std::string& col = scols[static_cast<size_t>(i)];
    int8_t vtyp_i = static_cast<int8_t>(vtyps[static_cast<size_t>(i)]);

    for (R_xlen_t j = 0; j < n; ++j, ++idx) {

      checkInterrupt(idx);

      // if colname is provided, the first row is always a character
      int8_t vtyp = (ColNames && j == 0) ? character : vtyp_i;
      std::string vals_sexp = Rcpp::as<std::string>(cvec[j]);

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
        xml_cols[pos].v = vals_sexp;
        break;
      case logical:
        // v and c_t = "b"
        xml_cols[pos].v =  vals_sexp;
        xml_cols[pos].c_t =  bool_sexp;
        break;
      case factor:
      case character:

        // test if string can be written as number
        if (string_nums && is_double(vals_sexp.c_str())) {
          // v
          xml_cols[pos].v = vals_sexp;
          vtyp     = (string_nums == 1) ? string_num : numeric;
        } else {
          // check if we write sst or inlineStr
          if (inline_strings) {
              // is and c_t = "inlineStr"
              xml_cols[pos].c_t = inlineStr_sexp;
              xml_cols[pos].is = txt_to_is(vals_sexp, 0, 1, 1);
            } else {
              // v and c_t = "s"
              xml_cols[pos].c_t = sharedStr_sexp;
              xml_cols[pos].v = txt_to_si(vals_sexp, 0, 1, 1);
          }
        }
        break;
      case hyperlink:
      case formula:
        // f and c_t = "str";
        xml_cols[pos].c_t = string_sexp;
        xml_cols[pos].f = vals_sexp;
        break;
      case array_formula:
        // f, f_t = "array", and f_ref
        xml_cols[pos].f = vals_sexp;
        xml_cols[pos].f_t = array_sexp;
        xml_cols[pos].f_ref = ref_str;
        break;
      case cm_formula:
        // c_cm, f, f_t = "array", and f_ref
        xml_cols[pos].c_cm  = c_cm;
        xml_cols[pos].f     = vals_sexp;
        xml_cols[pos].f_t   = array_sexp;
        xml_cols[pos].f_ref = ref_str;
        break;
      }

      if ((cvec[j] == NA_STRING) || vals_sexp == "_openxlsx_NA") {

        if (na_missing) {
          // v = "#N/A"
          xml_cols[pos].v = na_sexp;
          xml_cols[pos].c_t = expr_sexp;
          // is = "" - required only if inline_strings
          // and vtyp = character || vtyp = factor
          xml_cols[pos].is = blank_sexp;
        } else  {
          if (na_null) {
            // all three v, c_t, and is = "" - there should be nothing in this row
            xml_cols[pos].v = blank_sexp;
            xml_cols[pos].c_t = blank_sexp;
            xml_cols[pos].is = blank_sexp;
          } else {
            // c_t = "inlineStr" or "s"
            xml_cols[pos].c_t = inline_strings  ? inlineStr_sexp  : sharedStr_sexp;
            // for inlineStr is = na_strings else ""
            xml_cols[pos].is = inline_strings  ? na_strings_sexp : blank_sexp;
            // otherwise v = na_strings else ""
            xml_cols[pos].v = !inline_strings ? na_strings_sexp : blank_sexp;
          }
        }

      } else if (vals_sexp == "NaN") {
        // v = "#VALUE!"
        xml_cols[pos].v = value_sexp;
        xml_cols[pos].c_t = expr_sexp;
      } else if (vals_sexp == "-Inf" || vals_sexp == "Inf") {
        // v = "#NUM!"
        xml_cols[pos].v = num_sexp;
        xml_cols[pos].c_t = expr_sexp;
      }

      // typ = std::to_string(vtyp);
      xml_cols[pos].typ = std::to_string(vtyp);

      std::string cell_r = has_dims ? dims[static_cast<size_t>(idx)] : col + row;
      xml_cols[pos].r = cell_r;

      if (has_dims) {
        // row_r = rm_colnum(r) and c_r = rm_rownum(r)
        xml_cols[pos].row_r = rm_colnum(cell_r);
        xml_cols[pos].c_r = rm_rownum(cell_r);
      } else {
        // row_r = row and c_r = col
        xml_cols[pos].row_r = row;
        xml_cols[pos].c_r = col;
      }
    } // n
  } // m

  flush_to_parquet(xml_cols, tmpfile);
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
