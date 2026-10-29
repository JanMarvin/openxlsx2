#ifndef OPENXLSX2_H
#define OPENXLSX2_H

#include <array>
#include <memory>

#include "openxlsx2_types.h"

#define MAX_OOXML_COL_INT 16384
#define MAX_OOXML_ROW_INT 1048576

struct xml_string_writer : pugi::xml_writer {
    std::string result;
    void write(const void* data, size_t size) override {
        result.append(static_cast<const char*>(data), size);
    }
};

inline void check_xptr_validity(XPtrXML doc) {
  if (doc.get() == nullptr) {
    Rcpp::stop("The XML pointer is invalid.");
  }
}

Rcpp::IntegerVector col_to_int(Rcpp::CharacterVector x);

SEXP si_to_txt(XPtrXML doc);
SEXP is_to_txt(Rcpp::CharacterVector is_vec);

std::string txt_to_xml(
    const std::string& text,
    bool no_escapes,
    bool raw,
    bool skip_control,
    const std::string& type
);
std::string txt_to_is(std::string txt, bool no_escapes, bool raw, bool skip_control);
std::string txt_to_si(std::string txt, bool no_escapes, bool raw, bool skip_control);

inline bool sexp_str_less(SEXP a, SEXP b) {
  return std::strcmp(CHAR(a), CHAR(b)) < 0;
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
static inline uint32_t uint_col_to_int(const std::string& a) {
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


// lookup table using 256 for all possible unsigned char values
static const std::array<char, 256> DIGIT_MASK = [] {
  std::array<char, 256> m = {{0}};
  m['0'] = 1; m['1'] = 1; m['2'] = 1; m['3'] = 1; m['4'] = 1;
  m['5'] = 1; m['6'] = 1; m['7'] = 1; m['8'] = 1; m['9'] = 1;
  return m;
}();

static inline std::string filter_digits(const char* s, bool keep_digits) {
    if (!s) return "";

    std::string buffer;
    while (*s) {
        bool is_dig = DIGIT_MASK[static_cast<unsigned char>(*s)];
        if (is_dig == keep_digits) {
            buffer.push_back(*s);
        }
        s++;
    }
    return buffer;
}

static inline std::string rm_rownum(const std::string& str) {
  return filter_digits(str.c_str(), false);
}

static inline std::string rm_colnum(const std::string& str) {
  return filter_digits(str.c_str(), true);
}

// Function to keep only digits in a string
inline int32_t cell_to_rowint(const std::string& str) {
  std::string result = rm_colnum(str);
  int32_t res = std::stoi(result);
  if (res < 1 || res > MAX_OOXML_ROW_INT)
    Rcpp::stop("Row exceeds valid range");

  return res;
}

// Function to remove digits from a string
static inline int32_t cell_to_colint(const std::string& str) {
  std::string result = rm_rownum(str);

  for (char &c : result) {
    if (c >= 'a' && c <= 'z') c -= 32;
  }

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
  Rcpp::NumericVector key      = Rcpp::NumericVector(n);

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

    key[i] = x[ii].key;
    if (!x[ii].r.empty())
      SET_STRING_ELT(r, i, Rf_mkCharLenCE(x[ii].r.data(), static_cast<int>(x[ii].r.size()), CE_UTF8));
    if (!x[ii].row_r.empty())
      SET_STRING_ELT(row_r, i, Rf_mkCharLenCE(x[ii].row_r.data(), static_cast<int>(x[ii].row_r.size()), CE_UTF8));
    if (!x[ii].c_r.empty())
      SET_STRING_ELT(c_r, i, Rf_mkCharLenCE(x[ii].c_r.data(), static_cast<int>(x[ii].c_r.size()), CE_UTF8));
    if (!x[ii].c_s.empty())
      SET_STRING_ELT(c_s, i, Rf_mkCharLenCE(x[ii].c_s.data(), static_cast<int>(x[ii].c_s.size()), CE_UTF8));
    if (!x[ii].c_t.empty())
      SET_STRING_ELT(c_t, i, Rf_mkCharLenCE(x[ii].c_t.data(), static_cast<int>(x[ii].c_t.size()), CE_UTF8));
    if (!x[ii].f_attr.empty())
      SET_STRING_ELT(f_attr, i, Rf_mkCharLenCE(x[ii].f_attr.data(), static_cast<int>(x[ii].f_attr.size()), CE_UTF8));
    if (!x[ii].v.empty())
      SET_STRING_ELT(v, i, Rf_mkCharLenCE(x[ii].v.data(), static_cast<int>(x[ii].v.size()), CE_UTF8));
    if (!x[ii].f.empty())
      SET_STRING_ELT(f, i, Rf_mkCharLenCE(x[ii].f.data(), static_cast<int>(x[ii].f.size()), CE_UTF8));
    if (!x[ii].is.empty())
      SET_STRING_ELT(is, i, Rf_mkCharLenCE(x[ii].is.data(), static_cast<int>(x[ii].is.size()), CE_UTF8));

    if (has_cm && !x[ii].c_cm.empty())
      SET_STRING_ELT(c_cm, i, Rf_mkCharLenCE(x[ii].c_cm.data(), static_cast<int>(x[ii].c_cm.size()), CE_UTF8));
    if (has_ph && !x[ii].c_ph.empty())
      SET_STRING_ELT(c_ph, i, Rf_mkCharLenCE(x[ii].c_ph.data(), static_cast<int>(x[ii].c_ph.size()), CE_UTF8));
    if (has_vm && !x[ii].c_vm.empty())
      SET_STRING_ELT(c_vm, i, Rf_mkCharLenCE(x[ii].c_vm.data(), static_cast<int>(x[ii].c_vm.size()), CE_UTF8));
  }

  // --- 3. Create DataFrame (Simplified Structure) ---

  Rcpp::List df_list;
  Rcpp::CharacterVector df_names;

  // Base columns
  df_list.push_back(std::move(key));
  df_names.push_back("key");
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
