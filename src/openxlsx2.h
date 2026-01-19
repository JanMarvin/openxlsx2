#ifndef OPENXLSX2_H
#define OPENXLSX2_H

#include "openxlsx2_types.h"

#define MAX_OOXML_COL_INT 16384
#define MAX_OOXML_ROW_INT 1048576


inline void check_xptr_validity(XPtrXML doc) {
  if (doc.get() == nullptr) {
    Rcpp::stop("The XML pointer is invalid.");
  }
}

Rcpp::IntegerVector col_to_int(Rcpp::CharacterVector x);

SEXP si_to_txt(XPtrXML doc);
SEXP is_to_txt(Rcpp::CharacterVector is_vec);

std::string txt_to_is(std::string txt, bool no_escapes, bool raw, bool skip_control);
std::string txt_to_si(std::string txt, bool no_escapes, bool raw, bool skip_control);

// helper function to access element from Rcpp::Character Vector as string
inline std::string to_string(Rcpp::Vector<16>::Proxy x) {
  return Rcpp::String(x);
}

inline void checkInterrupt(R_xlen_t& iteration, R_xlen_t frequency = 10000) {
  if (iteration % frequency == 0) {
    Rcpp::checkUserInterrupt();
  }
  ++iteration;
}

template <typename T>
static inline std::string int_to_col(T cell) {
  std::string col_name = "";

  if (cell <= 0 || cell > MAX_OOXML_COL_INT)
    Rcpp::stop("Column exceeds valid range");

  while (cell > 0) {
    auto modulo = (cell - 1) % 26;
    col_name = (char)('A' + modulo) + col_name;
    cell = (cell - modulo) / 26;
  }

  return col_name;
}

// similar to is.numeric(x)
// returns true if string can be written as numeric and is not Inf
// @param x a string input
static inline bool is_double(const char* x) {
  char* endp;
  double res;

  res = R_strtod(x, &endp);

  if (strlen(endp) == 0 && std::isfinite(res)) {
    return 1;
  }

  return 0;
}

static inline bool has_cell(const std::string& str, const std::unordered_set<std::string>& vec) {
  return vec.find(str) != vec.end();
}

// driver function for col_to_int
static inline uint32_t uint_col_to_int(std::string& a) {
  char A = 'A';
  int32_t aVal = (int)A - 1;
  uint32_t sum = 0;
  size_t k = a.length();

  if (!std::all_of(a.begin(), a.end(), ::isalpha)) {
    Rcpp::stop("found non alphabetic character in column to integer conversion");
  }

  for (size_t j = 0; j < k; ++j) {
    sum *= 26;
    sum += static_cast<uint32_t>(a[j] - aVal);
  }

  if (sum == 0 || sum > MAX_OOXML_COL_INT)
    Rcpp::stop("Column exceeds valid range");

  return sum;
}

static inline std::string rm_rownum(const std::string& str) {
  std::string result;
  for (char c : str) {
    if (!std::isdigit(c)) {
      result += c;
    }
  }
  return result;
}

static inline std::string rm_colnum(const std::string& str) {
  std::string result;
  for (char c : str) {
    if (std::isdigit(c)) {
      result += c;
    }
  }
  return result;
}

// Function to keep only digits in a string
inline int32_t cell_to_rowint(const std::string& str) {
  std::string result = rm_colnum(str);
  int32_t res = std::stoi(result);
  if (res < 1 || res > MAX_OOXML_ROW_INT)
    Rcpp::stop("Row exceeds valid range");

  return res;
}

static inline std::string str_toupper(std::string s) {
  std::transform(s.begin(), s.end(), s.begin(), [](unsigned char c) { return std::toupper(c); });
  return s;
}

// Function to remove digits from a string
static inline int32_t cell_to_colint(const std::string& str) {
  std::string result = rm_rownum(str);
  result = str_toupper(result);
  return static_cast<int32_t>(uint_col_to_int(result));
}

static inline bool is_column_only(const std::string& s) {
  for (char c : s) {
    if (!std::isalpha(c)) return false;
  }
  return !s.empty();
}

static inline bool is_row_only(const std::string& s) {
  for (char c : s) {
    if (!std::isdigit(c)) return false;
  }
  return !s.empty();
}

static inline bool validate_dims(const std::string& input) {
  bool has_col = false;
  bool has_row = false;

  for (char c : input) {
    if (std::isupper(c)) {
      has_col = true;
    } else if (std::isdigit(c)) {
      has_row = true;
    } else {
      return false;
    }
  }

  return has_col && has_row;
}

inline SEXP xml_cols_to_df(const std::vector<xml_col>& x, bool has_cm, bool has_ph, bool has_vm) {
  R_xlen_t n = static_cast<R_xlen_t>(x.size());

  // --- 1. Initialization ---

  // Base columns
  Rcpp::CharacterVector r      = Rcpp::CharacterVector(n);
  Rcpp::CharacterVector row_r  = Rcpp::CharacterVector(n);
  Rcpp::CharacterVector c_r    = Rcpp::CharacterVector(n);
  Rcpp::CharacterVector c_s    = Rcpp::CharacterVector(n);
  Rcpp::CharacterVector c_t    = Rcpp::CharacterVector(n);

  // Content columns
  Rcpp::CharacterVector v      = Rcpp::CharacterVector(n);
  Rcpp::CharacterVector f      = Rcpp::CharacterVector(n);
  Rcpp::CharacterVector f_attr = Rcpp::CharacterVector(n);
  Rcpp::CharacterVector is     = Rcpp::CharacterVector(n);

  // Conditional columns (only allocate if needed)
  Rcpp::CharacterVector c_cm, c_ph, c_vm;
  if (has_cm) c_cm             = Rcpp::CharacterVector(n);
  if (has_ph) c_ph             = Rcpp::CharacterVector(n);
  if (has_vm) c_vm             = Rcpp::CharacterVector(n);

  // --- 2. Fill Vectors ---

  for (R_xlen_t i = 0; i < n; ++i) {
    size_t ii = static_cast<size_t>(i);

    // Assignment. Use direct assignment for std::string source to Rcpp::CharacterVector elements.
    // Use Rcpp::String only where needed for correct UTF-8 handling (e.g., cell content).

    // Core attributes (usually simple ASCII, safe with direct assignment)
    if (!x[ii].r.empty())      r[i]      = x[ii].r;
    if (!x[ii].row_r.empty())  row_r[i]  = x[ii].row_r;
    if (!x[ii].c_r.empty())    c_r[i]    = x[ii].c_r;
    if (!x[ii].c_s.empty())    c_s[i]    = x[ii].c_s;
    if (!x[ii].c_t.empty())    c_t[i]    = x[ii].c_t;
    if (!x[ii].f_attr.empty()) f_attr[i] = x[ii].f_attr;

    // Conditional attributes
    if (has_cm && !x[ii].c_cm.empty()) c_cm[i] = x[ii].c_cm;
    if (has_vm && !x[ii].c_vm.empty()) c_vm[i] = x[ii].c_vm;

    // Textual attributes (use Rcpp::String for guaranteed UTF-8 handling)
    if (has_ph && !x[ii].c_ph.empty()) c_ph[i] = Rcpp::String(x[ii].c_ph);
    if (!x[ii].f.empty())              f[i]    = Rcpp::String(x[ii].f);
    if (!x[ii].is.empty())             is[i]   = Rcpp::String(x[ii].is);
    if (!x[ii].v.empty()) { // can only be utf8 if c_t = "str"
      if (x[ii].c_t.empty() && x[ii].f_attr.empty())
        v[i] = x[ii].v;
      else
        v[i] = Rcpp::String(x[ii].v);
    }
  }

  // --- 3. Create DataFrame (Simplified Structure) ---

  Rcpp::List df_list;
  Rcpp::CharacterVector df_names;

  // Base columns
  df_list.push_back(std::move(r));
  df_names.push_back("r");
  df_list.push_back(std::move(row_r));
  df_names.push_back("row_r");
  df_list.push_back(std::move(c_r));
  df_names.push_back("c_r");
  df_list.push_back(std::move(c_s));
  df_names.push_back("c_s");
  df_list.push_back(std::move(c_t));
  df_names.push_back("c_t");

  // Conditional columns
  if (has_cm) {
    df_list.push_back(std::move(c_cm));
    df_names.push_back("c_cm");
  }
  if (has_ph) {
    df_list.push_back(std::move(c_ph));
    df_names.push_back("c_ph");
  }
  if (has_vm) {
    df_list.push_back(std::move(c_vm));
    df_names.push_back("c_vm");
  }

  // Content columns
  df_list.push_back(std::move(v));
  df_names.push_back("v");
  df_list.push_back(std::move(f));
  df_names.push_back("f");
  df_list.push_back(std::move(f_attr));
  df_names.push_back("f_attr");
  df_list.push_back(std::move(is));
  df_names.push_back("is");

  df_list.attr("names") = df_names;
  df_list.attr("class") = "data.frame";
  df_list.attr("row.names") = Rcpp::IntegerVector::create(NA_INTEGER, -n);
  df_list.attr("stringsAsFactors") = false;

  return df_list;
}

#endif
