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

static inline std::string int_to_col(uint32_t cell) {
  std::string col_name = "";

  while (cell > 0)
  {
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

  char *endp;
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
  int aVal = (int)A - 1;
  int sum = 0;
  size_t k = a.length();

  for (size_t j = 0; j < k; ++j) {
    sum *= 26;
    sum += (a[j] - aVal);
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
  return std::stoi(result);
}

static inline std::string str_toupper(std::string s) {
  std::transform(s.begin(), s.end(), s.begin(),
                 [](unsigned char c){ return std::toupper(c); }
  );
  return s;
}

// Function to remove digits from a string
static inline uint32_t cell_to_colint(const std::string& str) {
  std::string result = rm_rownum(str);
  result = str_toupper(result);
  return uint_col_to_int(result);
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
