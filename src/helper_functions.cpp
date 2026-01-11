#include <algorithm>
#include "openxlsx2.h"

// For R-devel 4.3 character length on Windows was modified. This caused an
// error when passing large character strings to file.exist() when called in
// read_xml(). We prevent bailing, by checking if the input is to long to be
// a path.
// Most likely this is only required until the dust has settled in R-devel, but
// CRAN checks must succeed on R-devel too.
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
  R_xlen_t ncol = Rf_xlength(x);

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
          type[i] = logical;  // logical
        } else {
          type[i] = factor;  // probably some custom class
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
            type[i] = numeric;  // numeric and integer
          } else {
            type[i] = factor;  // probably some custom class
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
  R_xlen_t n = Rf_xlength(x);
  std::unordered_map<std::string, int> col_map;
  Rcpp::IntegerVector colNums(n);

  for (R_xlen_t i = 0; i < n; ++i) {
    std::string a = Rcpp::as<std::string>(x[i]);

    if (a.empty()) Rcpp::stop("Empty string in conversion from column to integer");

    // check if the value is digit only, if yes, add it and continue the loop
    // at the top. This avoids slow:
    // suppressWarnings(isTRUE(as.character(as.numeric(x)) == x))
    if (std::all_of(a.begin(), a.end(), ::isdigit)) {
      colNums[i] = std::stoi(a);
      continue;
    }

    // Check if the column name is already in the map
    if (col_map.find(a) != col_map.end()) {
      colNums[i] = col_map[a];
    } else {
      // Compute the integer value and store it in the map
      int col_int = cell_to_colint(a);
      col_map[a] = col_int;
      colNums[i] = col_int;
    }
  }

  return colNums;
}

// [[Rcpp::export]]
Rcpp::CharacterVector ox_int_to_col(Rcpp::NumericVector x) {
  R_xlen_t n = Rf_xlength(x);
  Rcpp::CharacterVector colNames(n);
  std::unordered_map<uint32_t, std::string> cache;  // Updated key type

  for (R_xlen_t i = 0; i < n; ++i) {
    uint32_t num = static_cast<uint32_t>(x[i]);

    // Check if the column name is already in the cache
    if (cache.find(num) != cache.end()) {
      colNames[i] = cache[num];
    } else {
      // Compute the column name and store it in the cache
      std::string col_name = int_to_col(num);
      cache[num] = col_name;
      colNames[i] = col_name;
    }
  }

  return colNames;
}

// provide a basic rbindlist for lists of named characters
// [[Rcpp::export]]
SEXP rbindlist(Rcpp::List x) {
  R_xlen_t nn = Rf_xlength(x);

  // --- 1. Collect all unique names ---
  std::set<std::string> unique_names_set;
  for (R_xlen_t i = 0; i < nn; ++i) {
    if (Rf_isNull(x[i])) continue;

    // Check if the SEXP has names attribute before casting to CharacterVector
    if (!Rf_isNewList(x[i]) && !Rf_isVector(x[i])) continue;
    if (Rcpp::as<Rcpp::List>(x[i]).hasAttribute("names")) {
      Rcpp::CharacterVector names_i = Rcpp::as<Rcpp::List>(x[i]).attr("names");
      for (R_xlen_t j = 0; j < Rf_xlength(names_i); ++j) {
        unique_names_set.insert(Rcpp::as<std::string>(names_i[j]));
      }
    }
  }

  // --- 2. Map names to final column index ---
  std::vector<std::string> final_names(unique_names_set.begin(), unique_names_set.end());
  R_xlen_t kk = static_cast<R_xlen_t>(final_names.size());

  std::unordered_map<std::string, R_xlen_t> name_to_idx;
  name_to_idx.reserve(static_cast<uint64_t>(kk));
  for (R_xlen_t i = 0; i < kk; ++i) {
    name_to_idx[final_names[static_cast<uint64_t>(i)]] = i;
  }

  // --- 3. Allocate final DataFrame list ---
  Rcpp::List df(kk);
  for (R_xlen_t i = 0; i < kk; ++i) {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(nn));
  }

  // Pre-fetch references to the output vectors for fast assignment
  std::vector<Rcpp::CharacterVector> output_cols;
  output_cols.reserve(static_cast<size_t>(kk));
  for (R_xlen_t i = 0; i < kk; ++i) {
    output_cols.push_back(Rcpp::as<Rcpp::CharacterVector>(df[i]));
  }

  // --- 4. Fill the final columns row by row ---
  for (R_xlen_t i = 0; i < nn; ++i) {
    if (Rf_isNull(x[i])) continue;

    Rcpp::CharacterVector current_vec = x[i];

    // Check for names before accessing the attribute
    if (!current_vec.hasAttribute("names")) continue;

    Rcpp::CharacterVector current_names = current_vec.attr("names");

    for (R_xlen_t j = 0; j < Rf_xlength(current_vec); ++j) {
      // Get the column name and find its index (O(1) average lookup)
      std::string current_name = Rcpp::as<std::string>(current_names[j]);
      auto it = name_to_idx.find(current_name);

      if (it != name_to_idx.end()) {
        R_xlen_t col_idx = it->second;

        // Optimized assignment using pre-fetched vector reference
        output_cols[static_cast<size_t>(col_idx)][i] = current_vec[j];
      }
    }
  }

  // --- 5. Finalize DataFrame attributes ---
  df.attr("row.names") = Rcpp::IntegerVector::create(NA_INTEGER, -nn);
  df.attr("names") = Rcpp::CharacterVector(final_names.begin(), final_names.end());
  df.attr("class") = "data.frame";

  return df;
}

// [[Rcpp::export]]
SEXP copy(SEXP x) {
  return Rf_duplicate(x);
}

// [[Rcpp::export]]
std::vector<std::string> needed_cells(const std::string& range, bool all = true) {
  std::vector<std::string> cells;

  // --- 1. Parsing and Normalization ---

  std::string startCellStr, endCellStr;
  size_t colonPos = range.find(':');

  if (colonPos != std::string::npos) {
    startCellStr = range.substr(0, colonPos);
    endCellStr = range.substr(colonPos + 1);
  } else {
    startCellStr = range;
    endCellStr = range;
  }

  // NOTE: Assuming validate_dims, is_column_only, is_row_only are robust external functions.
  if (!validate_dims(startCellStr) || !validate_dims(endCellStr)) {
    if (is_column_only(startCellStr) && is_column_only(endCellStr)) {
      // Expand column range (e.g., "A:B") to max rows
      startCellStr += "1";
      endCellStr += "1048576";  // Max row count for modern XLSX
    } else if (is_row_only(startCellStr) && is_row_only(endCellStr)) {
      // Expand row range (e.g., "3:5") to max columns
      startCellStr = "A" + startCellStr;
      endCellStr = "XFD" + endCellStr;  // Max column name for modern XLSX
    } else {
      Rcpp::stop("Invalid input: dims must be something like A1 or A1:B2.");
    }
  }

  // --- 2. Quick Exit for Boundary Cells ---

  if (!all) {
    cells.reserve(2);
    cells.push_back(startCellStr);
    cells.push_back(endCellStr);
    return cells;
  }

  // --- 3. Cell Generation ---

  int32_t startRow = cell_to_rowint(startCellStr);
  int32_t endRow = cell_to_rowint(endCellStr);
  int32_t startCol = cell_to_colint(startCellStr);
  int32_t endCol = cell_to_colint(endCellStr);

  int32_t rowStep = (startRow <= endRow) ? 1 : -1;
  int32_t colStep = (startCol <= endCol) ? 1 : -1;

  // Estimate capacity to avoid reallocations
  size_t rowCount = static_cast<size_t>(std::abs(endRow - startRow)) + 1;
  size_t colCount = static_cast<size_t>(std::abs(endCol - startCol)) + 1;
  cells.reserve(rowCount * colCount);

  for (int32_t col = startCol;; col += colStep) {
    std::string colStr = int_to_col(col);

    for (int32_t row = startRow;; row += rowStep) {
      std::string cell = colStr;
      cell += std::to_string(row);
      cells.push_back(std::move(cell));

      if (row == endRow) break;
    }

    if (col == endCol) break;
  }

  return cells;
}

//' check if non consecutive dims is equal sized: "A1:A4,B1:B4"
//' @param dims dims
//' @param check check if all the same size
//' @keywords internal
//' @noRd
// [[Rcpp::export]]
SEXP get_dims(Rcpp::CharacterVector dims, bool check = false) {
  R_xlen_t n_dims = dims.size();
  if (n_dims == 0)
    return R_NilValue;

  // Output containers
  std::vector<Rcpp::IntegerVector> out_row_vectors;
  std::set<std::pair<int32_t, int32_t>> seen_rows;
  std::vector<int32_t> ordered_unique_cols;
  std::set<int32_t> seen_cols_globally;

  // Logic for 'check'
  int32_t first_min_r = -1, first_max_r = -1;
  bool check_failed = false;
  bool reference_set = false;

  for (R_xlen_t i = 0; i < n_dims; ++i) {
    std::string dim_str = Rcpp::as<std::string>(dims[i]);
    std::vector<std::string> cells = needed_cells(dim_str, false);

    if (cells.empty())
      continue;

    std::string left_cell = cells.front();
    std::string right_cell = cells.back();

    int32_t r1 = cell_to_rowint(left_cell);
    int32_t r2 = cell_to_rowint(right_cell);
    int32_t c1 = cell_to_colint(left_cell);
    int32_t c2 = cell_to_colint(right_cell);

    int32_t r_min = std::min(r1, r2);
    int32_t r_max = std::max(r1, r2);

    if (!reference_set) {
      first_min_r = r_min;
      first_max_r = r_max;
      reference_set = true;
    } else {
      if (r_min != first_min_r || r_max != first_max_r) {
        if (check) return Rcpp::wrap(false);
        check_failed = true;
      }
    }

    if (!check) {
      // Rows: Pair is much faster than Vector in a set
      if (seen_rows.find({r1, r2}) == seen_rows.end()) {
        out_row_vectors.push_back(Rcpp::IntegerVector::create(r1, r2));
        seen_rows.insert({r1, r2});
      }

      // Columns: Sequential collection
      if (c1 > 0 && c2 > 0) {
        int32_t step = (c1 <= c2) ? 1 : -1;
        for (int32_t curr_c = c1;; curr_c += step) {
          if (seen_cols_globally.find(curr_c) == seen_cols_globally.end()) {
            ordered_unique_cols.push_back(curr_c);
            seen_cols_globally.insert(curr_c);
          }
          if (curr_c == c2)
            break;
        }
      }
    }
  }

  if (check)
    return Rcpp::wrap(true);

  Rcpp::List res = Rcpp::List::create(Rcpp::Named("rows") = Rcpp::wrap(out_row_vectors),
                                      Rcpp::Named("cols") = Rcpp::wrap(ordered_unique_cols));

  // Attach the check result as an attribute so R can use it
  res.attr("is_equal_sized") = !check_failed;
  return res;
}

// [[Rcpp::export]]
SEXP dims_to_row_col_fill(Rcpp::CharacterVector dims, bool fills = false) {
  R_xlen_t n = Rf_xlength(dims);

  std::vector<int32_t> ordered_rows_final;
  std::vector<int32_t> ordered_cols_final;
  std::vector<std::string> fill_final;

  std::set<int32_t> seen_rows_globally;
  std::set<int32_t> seen_cols_globally;

  for (R_xlen_t i = 0; i < n; ++i) {
    std::string dim_str = Rcpp::as<std::string>(dims[i]);

    // If a single cell is given (e.g., "A1"), treat it as a range "A1:A1"
    if (dim_str.find(':') == std::string::npos) {
      dim_str += ":" + dim_str;
    }

    if (dim_str == "Inf:-Inf") {
      Rcpp::stop("dims are inf:-inf");
    }

    std::vector<std::string> cells_from_needed = needed_cells(dim_str, fills);

    if (cells_from_needed.empty()) {
      // Rcpp::Rcout << "Warning: needed_cells returned empty for dim: " << dim_str << std::endl;
      continue;  // Skip if no cells are derived
    }

    // If fills is true, append all cells returned by needed_cells to fill_final.
    // The order in fill_final depends on the order returned by needed_cells.
    if (fills) {
      fill_final.insert(fill_final.end(), cells_from_needed.begin(), cells_from_needed.end());
    }

    std::string l_cell_str = cells_from_needed.front();
    std::string r_cell_str = cells_from_needed.back();

    int32_t l_row_orig = cell_to_rowint(l_cell_str);
    int32_t r_row_orig = cell_to_rowint(r_cell_str);
    int32_t l_col_orig = cell_to_colint(l_cell_str);
    int32_t r_col_orig = cell_to_colint(r_cell_str);

    // Collect rows, preserving order and ensuring uniqueness
    if (l_row_orig <= r_row_orig) {  // Iterate forwards
      for (int32_t r = l_row_orig; r <= r_row_orig; ++r) {
        if (r > 0 && seen_rows_globally.find(r) == seen_rows_globally.end()) {
          ordered_rows_final.push_back(r);
          seen_rows_globally.insert(r);
        }
      }
    } else {  // Iterate backwards (l_row_orig > r_row_orig)
      for (int32_t r = l_row_orig; r >= r_row_orig; --r) {
        if (r > 0 && seen_rows_globally.find(r) == seen_rows_globally.end()) {
          ordered_rows_final.push_back(r);
          seen_rows_globally.insert(r);
        }
      }
    }

    // Collect columns, preserving order and ensuring uniqueness
    if (l_col_orig <= r_col_orig) {  // Iterate forwards
      for (int32_t c = l_col_orig; c <= r_col_orig; ++c) {
        if (c > 0 && seen_cols_globally.find(c) == seen_cols_globally.end()) {
          ordered_cols_final.push_back(c);
          seen_cols_globally.insert(c);
        }
      }
    } else {  // Iterate backwards (l_col_orig > r_col_orig)
      for (int32_t c = l_col_orig; c >= r_col_orig; --c) {
        if (c > 0 && seen_cols_globally.find(c) == seen_cols_globally.end()) {
          ordered_cols_final.push_back(c);
          seen_cols_globally.insert(c);
        }
      }
    }
  }  // End of loop over dims

  if (fills) {
    return Rcpp::List::create(
      Rcpp::Named("rows") = ordered_rows_final,
      Rcpp::Named("cols") = ordered_cols_final,
      Rcpp::Named("fill") = fill_final
    );
  } else {
    return Rcpp::List::create(
      Rcpp::Named("rows") = ordered_rows_final,
      Rcpp::Named("cols") = ordered_cols_final,
      Rcpp::Named("fill") = R_NilValue
    );
  }
}

// provide a basic rbindlist for lists of named characters
// [[Rcpp::export]]
SEXP dims_to_df(Rcpp::IntegerVector rows,
                Rcpp::CharacterVector cols,
                Rcpp::Nullable<Rcpp::CharacterVector> filled,
                bool fill,
                Rcpp::Nullable<Rcpp::IntegerVector> fcols) {
  R_xlen_t kk = Rf_xlength(cols);
  R_xlen_t nn = Rf_xlength(rows);

  bool has_fcols = fcols.isNotNull();
  bool has_filled = filled.isNotNull();

  std::vector<std::string> row_strs(static_cast<size_t>(nn));
  for (R_xlen_t j = 0; j < nn; ++j) {
    row_strs[static_cast<size_t>(j)] = std::to_string(rows[j]);
  }

  std::vector<bool> col_mask(static_cast<size_t>(kk), !has_fcols);
  if (has_fcols) {
    Rcpp::IntegerVector fcls(fcols.get());
    for (int idx : fcls) {
      if (idx >= 0 && idx < kk) {
        col_mask[static_cast<size_t>(idx)] = true;
      }
    }
  }

  std::unordered_set<std::string> filled_set;
  if (has_filled) {
    Rcpp::CharacterVector flld(filled.get());
    filled_set.reserve(static_cast<size_t>(flld.size()));
    for (R_xlen_t i = 0; i < flld.size(); ++i) {
      filled_set.insert(Rcpp::as<std::string>(flld[i]));
    }
  }

  Rcpp::List df(kk);
  SEXP default_val = fill ? Rf_mkChar("") : NA_STRING;

  for (R_xlen_t i = 0; i < kk; ++i) {
    Rcpp::CharacterVector cvec(nn, default_val);

    if (fill && col_mask[static_cast<size_t>(i)]) {
      const char* col_name = CHAR(STRING_ELT(cols, i));
      size_t col_len = strlen(col_name);

      for (R_xlen_t j = 0; j < nn; ++j) {
        const std::string& row_str = row_strs[static_cast<size_t>(j)];

        char buffer[32];
        memcpy(buffer, col_name, col_len);
        memcpy(buffer + col_len, row_str.c_str(), row_str.size() + 1);

        if (!has_filled || filled_set.count(buffer)) {
          cvec[j] = Rf_mkChar(buffer);
        }
      }
    }
    df[i] = cvec;
  }

  df.attr("row.names") = rows;
  df.attr("names") = cols;
  df.attr("class") = "data.frame";

  return df;
}

// similar to dcast converts cc dataframe to z dataframe
// [[Rcpp::export]]
void long_to_wide(Rcpp::DataFrame z, Rcpp::DataFrame tt, Rcpp::DataFrame zz) {
  R_xlen_t n = zz.nrow();
  Rcpp::IntegerVector cols = zz["cols"];
  Rcpp::IntegerVector rows = zz["rows"];
  Rcpp::CharacterVector vals = zz["val"];
  Rcpp::IntegerVector typs = zz["typ"];

  // Cache all column vectors to avoid repeated coercion
  std::vector<Rcpp::CharacterVector> z_cols(static_cast<size_t>(z.size()));
  std::vector<Rcpp::IntegerVector> tt_cols(static_cast<size_t>(tt.size()));

  for (R_xlen_t j = 0; j < z.size(); ++j) {
    z_cols[static_cast<size_t>(j)] = z[j];
    tt_cols[static_cast<size_t>(j)] = tt[j];
  }

  for (R_xlen_t i = 0; i < n; ++i) {
    int32_t col = cols[i];
    int32_t row = rows[i];

    if (row != NA_INTEGER && col != NA_INTEGER) {
      SET_STRING_ELT(z_cols[static_cast<size_t>(col)], row, STRING_ELT(vals, i));
      INTEGER(tt_cols[static_cast<size_t>(col)])[row] = INTEGER(typs)[i];
    }
  }
}

// function to apply on vector
// @param x a character vector as input
// [[Rcpp::export]]
Rcpp::LogicalVector is_charnum(Rcpp::CharacterVector x) {
  Rcpp::LogicalVector out(x.size());
  for (R_xlen_t i = 0; i < x.size(); ++i) {
    out[i] = is_double(x[i]);
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

  // --- 1. Pre-calculate Column/Row Strings ---
  std::vector<std::string> srows(static_cast<size_t>(n));
  for (size_t j = 0; j < static_cast<size_t>(n); ++j) {
    srows[j] = std::to_string(static_cast<size_t>(start_row) + j);
  }
  std::vector<std::string> scols(static_cast<size_t>(m));
  for (size_t i = 0; i < static_cast<size_t>(m); ++i) {
    scols[i] = int_to_col(static_cast<size_t>(start_col) + i);
  }

  // --- 2. Rcpp/C API Setup ---
  bool has_refs = refed.isNotNull();
  std::vector<std::string> ref;
  if (has_refs) ref = Rcpp::as<std::vector<std::string>>(refed.get());
  int32_t in_string_nums = string_nums;
  bool has_cm = zz.containsElementNamed("c_cm");
  bool has_typ = zz.containsElementNamed("typ");

  Rcpp::CharacterVector zz_row_r  = Rcpp::as<Rcpp::CharacterVector>(zz["row_r"]);
  Rcpp::CharacterVector zz_c_r    = Rcpp::as<Rcpp::CharacterVector>(zz["c_r"]);
  Rcpp::CharacterVector zz_v      = Rcpp::as<Rcpp::CharacterVector>(zz["v"]);
  Rcpp::CharacterVector zz_c_t    = Rcpp::as<Rcpp::CharacterVector>(zz["c_t"]);
  Rcpp::CharacterVector zz_is     = Rcpp::as<Rcpp::CharacterVector>(zz["is"]);
  Rcpp::CharacterVector zz_f      = Rcpp::as<Rcpp::CharacterVector>(zz["f"]);
  Rcpp::CharacterVector zz_f_attr = Rcpp::as<Rcpp::CharacterVector>(zz["f_attr"]);
  Rcpp::CharacterVector zz_r      = Rcpp::as<Rcpp::CharacterVector>(zz["r"]);
  Rcpp::CharacterVector zz_c_cm;
  if (has_cm) zz_c_cm = Rcpp::as<Rcpp::CharacterVector>(zz["c_cm"]);
  Rcpp::CharacterVector zz_typ;
  if (has_typ) zz_typ = Rcpp::as<Rcpp::CharacterVector>(zz["typ"]);

  // --- 3. Pre-convert ALL constant/pre-calculated strings to SEXP (Optimization) ---
  std::string final_na_string_content = inline_strings ?
    txt_to_is(na_strings, 0, 1, 1) :
    txt_to_si(na_strings, 0, 1, 1);

  SEXP blank_sexp      = Rf_mkChar("");
  SEXP inlineStr_sexp  = Rf_mkChar("inlineStr");
  SEXP bool_sexp       = Rf_mkChar("b");
  SEXP expr_sexp       = Rf_mkChar("e");
  SEXP sharedStr_sexp  = Rf_mkChar("s");
  SEXP string_sexp     = Rf_mkChar("str");
  SEXP na_sexp         = Rf_mkChar("#N/A");
  SEXP num_sexp        = Rf_mkChar("#NUM!");
  SEXP value_sexp      = Rf_mkChar("#VALUE!");
  SEXP na_strings_sexp = Rf_mkChar(final_na_string_content.c_str());
  SEXP c_cm_sexp_const = Rf_mkChar(c_cm.c_str());

  R_xlen_t iter_count = 0;

  // --- 4. Main Wide-to-Long Loop ---
  for (R_xlen_t i = 0; i < m; ++i) {
    Rcpp::CharacterVector cvec = Rcpp::as<Rcpp::CharacterVector>(z[i]);
    const std::string& col = scols[static_cast<size_t>(i)];
    int8_t vtyp_i = static_cast<int8_t>(vtyps[static_cast<size_t>(i)]);

    for (R_xlen_t j = 0; j < n; ++j) {
      // Calculate pos (Row-Major)
      R_xlen_t pos = (j * m) + i;
      checkInterrupt(iter_count);

      // Variables for this cell
      int8_t vtyp = (ColNames && j == 0) ? character : vtyp_i;
      SEXP vals_sexp = STRING_ELT(cvec, j);
      const char* vals = CHAR(vals_sexp);
      const std::string& row = srows[static_cast<size_t>(j)];

      // --- Coordinate Assignment ---
      if (has_dims) {
        R_xlen_t col_pos = (i * n) + j;
        const std::string& cell_r_str = dims[static_cast<size_t>(col_pos)];
        SET_STRING_ELT(zz_r, pos, Rf_mkChar(cell_r_str.c_str()));
        SET_STRING_ELT(zz_row_r, pos, Rf_mkChar(rm_colnum(cell_r_str).c_str()));
        SET_STRING_ELT(zz_c_r, pos, Rf_mkChar(rm_rownum(cell_r_str).c_str()));
      } else {
        const std::string& cell_r_str = col + row;
        SET_STRING_ELT(zz_r, pos, Rf_mkChar(cell_r_str.c_str()));
        SET_STRING_ELT(zz_row_r, pos, Rf_mkChar(row.c_str()));
        SET_STRING_ELT(zz_c_r, pos, Rf_mkChar(col.c_str()));
      }

      std::string ref_str = "";
      std::string f_attr = "";
      if (vtyp == array_formula || vtyp == cm_formula) {
        ref_str = has_refs ? ref[static_cast<size_t>(i)] : col + row;
        f_attr = "t=\"array\" ref=\"" + ref_str + "\"";
      }

      if (!(ColNames && j == 0) && vtyp == factor)
        string_nums = 1;
      else
        string_nums = in_string_nums;

      // --- Data Type Switch ---
      switch (vtyp) {
        // ... (numeric/logical cases) ...
        case currency:
        case short_date:
        case long_date:
        case accounting:
        case percentage:
        case scientific:
        case comma:
        case hms_time:
        case numeric:
          SET_STRING_ELT(zz_v, pos, vals_sexp);
          break;
        case logical:
          SET_STRING_ELT(zz_v, pos, vals_sexp);
          SET_STRING_ELT(zz_c_t, pos, bool_sexp);
          break;
        case factor:
        case character:
          if (string_nums && is_double(vals)) {
            SET_STRING_ELT(zz_v, pos, vals_sexp);
            vtyp = (string_nums == 1) ? string_num : numeric;
          } else {
            if (inline_strings) {
              SET_STRING_ELT(zz_c_t, pos, inlineStr_sexp);
              SET_STRING_ELT(zz_is, pos, Rf_mkChar(txt_to_is(vals, 0, 1, 1).c_str()));
            } else {
              SET_STRING_ELT(zz_c_t, pos, sharedStr_sexp);
              SET_STRING_ELT(zz_v, pos, Rf_mkChar(txt_to_si(vals, 0, 1, 1).c_str()));
            }
          }
          break;
        case hyperlink:
        case formula:
          SET_STRING_ELT(zz_c_t, pos, string_sexp);
          SET_STRING_ELT(zz_f, pos, vals_sexp);
          break;
        case array_formula:
          SET_STRING_ELT(zz_f, pos, vals_sexp);
          SET_STRING_ELT(zz_f_attr, pos, Rf_mkChar(f_attr.c_str()));
          break;
        case cm_formula:
          SET_STRING_ELT(zz_c_cm, pos, c_cm_sexp_const);
          SET_STRING_ELT(zz_f, pos, vals_sexp);
          SET_STRING_ELT(zz_f_attr, pos, Rf_mkChar(f_attr.c_str()));
          break;
      }

      // --- NA/Error Value Handling ---
      if (vals_sexp == NA_STRING || strcmp(vals, "_openxlsx_NA") == 0) {
        if (na_missing) {
          SET_STRING_ELT(zz_v, pos, na_sexp);
          SET_STRING_ELT(zz_c_t, pos, expr_sexp);
          SET_STRING_ELT(zz_is, pos, blank_sexp);
        } else {
          if (na_null) {
            SET_STRING_ELT(zz_v, pos, blank_sexp);
            SET_STRING_ELT(zz_c_t, pos, blank_sexp);
            SET_STRING_ELT(zz_is, pos, blank_sexp);
          } else {
            SET_STRING_ELT(zz_c_t, pos, inline_strings ? inlineStr_sexp : sharedStr_sexp);
            SET_STRING_ELT(zz_is, pos, inline_strings ? na_strings_sexp : blank_sexp);
            SET_STRING_ELT(zz_v, pos, !inline_strings ? na_strings_sexp : blank_sexp);
          }
        }
      } else if ((vtyp == numeric && strcmp(vals, "NaN") == 0) || (vtyp == factor && strcmp(vals, "_openxlsx_NaN") == 0)) {
        SET_STRING_ELT(zz_v, pos, value_sexp);
        SET_STRING_ELT(zz_c_t, pos, expr_sexp);
      } else if ((vtyp == numeric && (strcmp(vals, "-Inf") == 0 || strcmp(vals, "Inf") == 0)) ||
                 (vtyp == factor && (strcmp(vals, "_openxlsx_nInf") == 0 || strcmp(vals, "_openxlsx_Inf") == 0))) {
        SET_STRING_ELT(zz_v, pos, num_sexp);
        SET_STRING_ELT(zz_c_t, pos, expr_sexp);
      }

      if (has_typ) {
        SET_STRING_ELT(zz_typ, pos, Rf_mkChar(std::to_string(vtyp).c_str()));
      }
    }
  }
}

// simple helper function to create a data frame of type character
//' @param colnames a vector of the names of the data frame
//' @param n the length of the data frame
//' @noRd
// [[Rcpp::export]]
Rcpp::DataFrame create_char_dataframe(Rcpp::CharacterVector colnames, R_xlen_t n) {
  R_xlen_t kk = Rf_xlength(colnames);

  // 1. create the list
  Rcpp::List df(kk);
  for (R_xlen_t i = 0; i < kk; ++i) {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(n));
  }

  Rcpp::IntegerVector rvec = Rcpp::IntegerVector::create(NA_INTEGER, n);

  // 3. Create a data.frame
  df.attr("row.names") = rvec;
  df.attr("names") = colnames;
  df.attr("class") = "data.frame";

  return df;
}

// [[Rcpp::export]]
Rcpp::DataFrame create_int_dataframe(const Rcpp::DataFrame& char_df) {
  R_xlen_t n_rows = char_df.nrow();
  int32_t n_cols = static_cast<int32_t>(char_df.ncol());

  Rcpp::CharacterVector col_names = char_df.names();
  Rcpp::List int_list(n_cols);

  for (int32_t i = 0; i < n_cols; ++i) {
    SET_VECTOR_ELT(int_list, i, Rcpp::IntegerVector(n_rows, NA_INTEGER));
  }

  int_list.attr("row.names") = Rcpp::IntegerVector::create(NA_INTEGER, -n_rows);
  int_list.attr("names") = col_names;
  int_list.attr("class") = "data.frame";

  return int_list;
}

// TODO styles_xml.cpp should be converted to use these functions

// [[Rcpp::export]]
Rcpp::DataFrame read_xml2df(XPtrXML xml, std::string vec_name, std::vector<std::string> vec_attrs, std::vector<std::string> vec_chlds) {

  check_xptr_validity(xml);

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

  Rcpp::IntegerVector rvec(nn);
  std::iota(rvec.begin(), rvec.end(), 0);

  // 1. create the list
  Rcpp::List df(kk);
  for (R_xlen_t i = 0; i < kk; ++i) {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(nn));
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
  std::vector<std::string> attrnams = df.names();
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

// [[Rcpp::export]]
Rcpp::CharacterVector df_to_xml(std::string name, Rcpp::DataFrame df_col) {
  auto n = df_col.nrow();
  Rcpp::CharacterVector z(n);
  Rcpp::CharacterVector attrnams = df_col.names();

  for (auto i = 0; i < n; ++i) {
    pugi::xml_document doc;

    pugi::xml_node col = doc.append_child(name.c_str());

    for (auto j = 0; j < df_col.ncol(); ++j) {
      Rcpp::CharacterVector cv_s = "";
      cv_s = Rcpp::as<Rcpp::CharacterVector>(df_col[j])[i];

      // only write attributes where cv_s has a value
      if (cv_s[0] != "") {
        // Rf_PrintValue(cv_s);
        const std::string val_strl = Rcpp::as<std::string>(cv_s);
        col.append_attribute(attrnams[j]) = val_strl.c_str();
      }
    }

    std::ostringstream oss;
    doc.print(oss, " ", pugi::format_raw);

    z[i] = oss.str();
  }

  return z;
}
