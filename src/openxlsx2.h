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
static inline uint32_t uint_col_to_int(std::string& a) {
  char A = 'A';
  int32_t aVal = (int)A - 1;
  uint32_t sum = 0;
  size_t k = a.length();

  for (size_t j = 0; j < k; ++j) {
    sum *= 26;
    sum += static_cast<uint32_t>(a[j] - aVal);
  }

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
inline uint32_t cell_to_rowint(const std::string& str) {
  std::string result = rm_colnum(str);
  return static_cast<uint32_t>(std::stoi(result));
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
  // We have to convert utf8 inputs via Rcpp::String for non unicode R sessions
  // Ideally there would be a function that calls Rcpp::String only if needed
  for (R_xlen_t i = 0; i < n; ++i) {
    size_t ii = static_cast<size_t>(i);
    if (!x[ii].r.empty())      r[i]      = std::string(x[ii].r);
    if (!x[ii].row_r.empty())  row_r[i]  = std::string(x[ii].row_r);
    if (!x[ii].c_r.empty())    c_r[i]    = std::string(x[ii].c_r);
    if (!x[ii].c_s.empty())    c_s[i]    = std::string(x[ii].c_s);
    if (!x[ii].c_t.empty())    c_t[i]    = std::string(x[ii].c_t);
    if (has_cm && !x[ii].c_cm.empty())   c_cm[i]   = std::string(x[ii].c_cm);
    if (has_ph && !x[ii].c_ph.empty())   c_ph[i]   = Rcpp::String(x[ii].c_ph);
    if (has_vm && !x[ii].c_vm.empty())   c_vm[i]   = std::string(x[ii].c_vm);
    if (!x[ii].v.empty()) { // can only be utf8 if c_t = "str"
      if (x[ii].c_t.empty() && x[ii].f_attr.empty())
        v[i] = std::string(x[ii].v);
      else
        v[i] = Rcpp::String(x[ii].v);
    }
    if (!x[ii].f.empty())      f[i]      = Rcpp::String(x[ii].f);
    if (!x[ii].f_attr.empty()) f_attr[i] = std::string(x[ii].f_attr);
    if (!x[ii].is.empty())     is[i]     = Rcpp::String(x[ii].is);
  }

  // Assign and return a dataframe
  if (has_cm && has_ph && has_vm) {
    return Rcpp::wrap(
      Rcpp::DataFrame::create(
        Rcpp::Named("r")      = r,
        Rcpp::Named("row_r")  = row_r,
        Rcpp::Named("c_r")    = c_r,
        Rcpp::Named("c_s")    = c_s,
        Rcpp::Named("c_t")    = c_t,
        Rcpp::Named("c_cm")   = c_cm,
        Rcpp::Named("c_ph")   = c_ph,
        Rcpp::Named("c_vm")   = c_vm,
        Rcpp::Named("v")      = v,
        Rcpp::Named("f")      = f,
        Rcpp::Named("f_attr") = f_attr,
        Rcpp::Named("is")     = is,
        Rcpp::Named("stringsAsFactors") = false
      )
    );
  } else if (has_cm && has_ph && !has_vm) {
    return Rcpp::wrap(
      Rcpp::DataFrame::create(
        Rcpp::Named("r")      = r,
        Rcpp::Named("row_r")  = row_r,
        Rcpp::Named("c_r")    = c_r,
        Rcpp::Named("c_s")    = c_s,
        Rcpp::Named("c_t")    = c_t,
        Rcpp::Named("c_cm")   = c_cm,
        Rcpp::Named("c_ph")   = c_ph,
        Rcpp::Named("v")      = v,
        Rcpp::Named("f")      = f,
        Rcpp::Named("f_attr") = f_attr,
        Rcpp::Named("is")     = is,
        Rcpp::Named("stringsAsFactors") = false
      )
    );
  } else if (has_cm && !has_ph && has_vm) {
    return Rcpp::wrap(
      Rcpp::DataFrame::create(
        Rcpp::Named("r")      = r,
        Rcpp::Named("row_r")  = row_r,
        Rcpp::Named("c_r")    = c_r,
        Rcpp::Named("c_s")    = c_s,
        Rcpp::Named("c_t")    = c_t,
        Rcpp::Named("c_cm")   = c_cm,
        Rcpp::Named("c_vm")   = c_vm,
        Rcpp::Named("v")      = v,
        Rcpp::Named("f")      = f,
        Rcpp::Named("f_attr") = f_attr,
        Rcpp::Named("is")     = is,
        Rcpp::Named("stringsAsFactors") = false
      )
    );
  } else if (!has_cm && has_ph && has_vm) {
    return Rcpp::wrap(
      Rcpp::DataFrame::create(
        Rcpp::Named("r")      = r,
        Rcpp::Named("row_r")  = row_r,
        Rcpp::Named("c_r")    = c_r,
        Rcpp::Named("c_s")    = c_s,
        Rcpp::Named("c_t")    = c_t,
        Rcpp::Named("c_ph")   = c_ph,
        Rcpp::Named("c_vm")   = c_vm,
        Rcpp::Named("v")      = v,
        Rcpp::Named("f")      = f,
        Rcpp::Named("f_attr") = f_attr,
        Rcpp::Named("is")     = is,
        Rcpp::Named("stringsAsFactors") = false
      )
    );
  } else if (has_cm && !has_ph && !has_vm) {
    return Rcpp::wrap(
      Rcpp::DataFrame::create(
        Rcpp::Named("r")      = r,
        Rcpp::Named("row_r")  = row_r,
        Rcpp::Named("c_r")    = c_r,
        Rcpp::Named("c_s")    = c_s,
        Rcpp::Named("c_t")    = c_t,
        Rcpp::Named("c_cm")   = c_cm,
        Rcpp::Named("v")      = v,
        Rcpp::Named("f")      = f,
        Rcpp::Named("f_attr") = f_attr,
        Rcpp::Named("is")     = is,
        Rcpp::Named("stringsAsFactors") = false
      )
    );
  } else if (!has_cm && has_ph && !has_vm) {
    return Rcpp::wrap(
      Rcpp::DataFrame::create(
        Rcpp::Named("r")      = r,
        Rcpp::Named("row_r")  = row_r,
        Rcpp::Named("c_r")    = c_r,
        Rcpp::Named("c_s")    = c_s,
        Rcpp::Named("c_t")    = c_t,
        Rcpp::Named("c_ph")   = c_ph,
        Rcpp::Named("v")      = v,
        Rcpp::Named("f")      = f,
        Rcpp::Named("f_attr") = f_attr,
        Rcpp::Named("is")     = is,
        Rcpp::Named("stringsAsFactors") = false
      )
    );
  } else if (!has_cm && !has_ph && has_vm) {
    return Rcpp::wrap(
      Rcpp::DataFrame::create(
        Rcpp::Named("r")      = r,
        Rcpp::Named("row_r")  = row_r,
        Rcpp::Named("c_r")    = c_r,
        Rcpp::Named("c_s")    = c_s,
        Rcpp::Named("c_t")    = c_t,
        Rcpp::Named("c_vm")   = c_vm,
        Rcpp::Named("v")      = v,
        Rcpp::Named("f")      = f,
        Rcpp::Named("f_attr") = f_attr,
        Rcpp::Named("is")     = is,
        Rcpp::Named("stringsAsFactors") = false
      )
    );
  } else {
    return Rcpp::wrap(
      Rcpp::DataFrame::create(
        Rcpp::Named("r")      = r,
        Rcpp::Named("row_r")  = row_r,
        Rcpp::Named("c_r")    = c_r,
        Rcpp::Named("c_s")    = c_s,
        Rcpp::Named("c_t")    = c_t,
        Rcpp::Named("v")      = v,
        Rcpp::Named("f")      = f,
        Rcpp::Named("f_attr") = f_attr,
        Rcpp::Named("is")     = is,
        Rcpp::Named("stringsAsFactors") = false
      )
    );
  }

}
