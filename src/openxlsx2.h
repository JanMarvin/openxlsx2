#include "openxlsx2_types.h"

Rcpp::IntegerVector col_to_int(Rcpp::CharacterVector x);

SEXP si_to_txt(XPtrXML doc);
SEXP is_to_txt(Rcpp::CharacterVector is_vec);

std::string txt_to_is(std::string txt, bool no_escapes, bool raw, bool skip_control);
std::string txt_to_si(std::string txt, bool no_escapes, bool raw, bool skip_control);

// helper function to access element from Rcpp::Character Vector as string
inline std::string to_string(Rcpp::Vector<16>::Proxy x) {
  return Rcpp::String(x);
}

inline void checkInterrupt(R_xlen_t iteration, R_xlen_t frequency = 10000) {
  if (iteration % frequency == 0) {
    Rcpp::checkUserInterrupt();
  }
}

template <typename T>
static inline std::string int_to_col(T cell) {
  std::string col_name = "";

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
static inline bool is_double(std::string x) {
  char* endp;
  double res;

  res = R_strtod(x.c_str(), &endp);

  if (strlen(endp) == 0 && std::isfinite(res)) {
    return 1;
  }

  return 0;
}

static inline bool has_cell(const std::string& str, const std::unordered_set<std::string>& vec) {
  return vec.find(str) != vec.end();
}

// driver function for col_to_int
static inline uint32_t uint_col_to_int(std::string_view sv) {
  uint32_t col = 0;
  for (char c : sv) {
    col = col * 26 + (c - 'A' + 1);
  }
  return col;
}

static inline std::string rm_rownum(std::string_view sv) {
  std::string result;
  std::remove_copy_if(sv.begin(), sv.end(), std::back_inserter(result), ::isdigit);
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

static inline int32_t cell_to_rowint(std::string_view str) {
  std::string row_part;
  for (char c : str) {
    if (std::isdigit(c)) {
      row_part += c;
    }
  }
  return row_part.empty() ? 0 : std::stoi(row_part);
}

// Helper function to convert to uppercase
static inline std::string str_toupper(std::string_view sv) {
  std::string result;
  std::transform(sv.begin(), sv.end(), std::back_inserter(result), ::toupper);
  return result;
}
// Function to remove digits from a string and convert to column integer
static inline int32_t cell_to_colint(std::string_view sv) {
  std::string result_rm = rm_rownum(sv);
  std::string result_upper = str_toupper(result_rm);
  return static_cast<int32_t>(uint_col_to_int(result_upper));
}

static inline bool is_column_only(std::string_view s) {
  for (char c : s) {
    if (!std::isalpha(c)) return false;
  }
  return !s.empty();
}

static inline bool is_row_only(std::string_view s) {
  for (char c : s) {
    if (!std::isdigit(c)) return false;
  }
  return !s.empty();
}

static inline bool validate_dims(std::string_view input) {
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

  // Vector structure identical to xml_col from openxlsx2_types.h
  Rcpp::CharacterVector r(Rcpp::no_init(n));         // cell name: A1, A2 ...
  Rcpp::CharacterVector row_r(Rcpp::no_init(n));     // row name: 1, 2, ..., 9999

  Rcpp::CharacterVector c_r(Rcpp::no_init(n));       // col name: A, B, ..., ZZ
  Rcpp::CharacterVector c_s(Rcpp::no_init(n));       // cell style
  Rcpp::CharacterVector c_t(Rcpp::no_init(n));       // cell type
  Rcpp::CharacterVector c_cm, c_ph, c_vm;
  if (has_cm) c_cm = Rcpp::CharacterVector(Rcpp::no_init(n));
  if (has_ph) c_ph = Rcpp::CharacterVector(Rcpp::no_init(n));
  if (has_vm) c_vm = Rcpp::CharacterVector(Rcpp::no_init(n));

  Rcpp::CharacterVector v(Rcpp::no_init(n));         // <v> tag
  Rcpp::CharacterVector f(Rcpp::no_init(n));         // <f> tag
  Rcpp::CharacterVector f_attr(Rcpp::no_init(n));    // <f /> attributes
  Rcpp::CharacterVector is(Rcpp::no_init(n));        // <is> tag

  // struct to vector
  for (R_xlen_t i = 0; i < n; ++i) {
    size_t ii = static_cast<size_t>(i);
    if (!x[ii].r.empty())      r[i]      = std::string(x[ii].r);
    if (!x[ii].row_r.empty())  row_r[i]  = std::string(x[ii].row_r);
    if (!x[ii].c_r.empty())    c_r[i]    = std::string(x[ii].c_r);
    if (!x[ii].c_s.empty())    c_s[i]    = std::string(x[ii].c_s);
    if (!x[ii].c_t.empty())    c_t[i]    = std::string(x[ii].c_t);
    if (has_cm && !x[ii].c_cm.empty())   c_cm[i]   = Rcpp::String(std::string(x[ii].c_cm));
    if (has_ph && !x[ii].c_ph.empty())   c_ph[i]   = Rcpp::String(std::string(x[ii].c_ph));
    if (has_vm && !x[ii].c_vm.empty())   c_vm[i]   = Rcpp::String(std::string(x[ii].c_vm));
    if (!x[ii].v.empty()) {
      if (x[ii].c_t.empty() && x[ii].f_attr.empty())
        v[i] = Rcpp::String(std::string(x[ii].v));
      else
        v[i] = Rcpp::String(std::string(x[ii].v));
    }
    if (!x[ii].f_attr.empty()) f_attr[i] = std::string(x[ii].f_attr);
    if (!x[ii].f.empty())      f[i]      = Rcpp::String(std::string(x[ii].f));
    if (!x[ii].is.empty())     is[i]     = Rcpp::String(std::string(x[ii].is));
  }

  // Assign and return a dataframe
  Rcpp::List df_cols;
  df_cols["r"] = r;
  df_cols["row_r"] = row_r;
  df_cols["c_r"] = c_r;
  df_cols["c_s"] = c_s;
  df_cols["c_t"] = c_t;
  if (has_cm) df_cols["c_cm"] = c_cm;
  if (has_ph) df_cols["c_ph"] = c_ph;
  if (has_vm) df_cols["c_vm"] = c_vm;
  df_cols["v"] = v;
  df_cols["f"] = f;
  df_cols["f_attr"] = f_attr;
  df_cols["is"] = is;
  df_cols["stringsAsFactors"] = Rcpp::wrap(false);

  return Rcpp::DataFrame(df_cols);
}
