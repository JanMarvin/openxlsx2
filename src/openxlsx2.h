#include "openxlsx2_types.h"

std::string int_to_col(uint32_t cell);
uint32_t uint_col_to_int(std::string& a);
Rcpp::IntegerVector col_to_int(Rcpp::CharacterVector x);

SEXP si_to_txt(XPtrXML doc);
SEXP is_to_txt(Rcpp::CharacterVector is_vec);

std::string txt_to_is(std::string txt, bool no_escapes, bool raw, bool skip_control);

// helper function to access element from Rcpp::Character Vector as string
inline std::string to_string(Rcpp::Vector<16>::Proxy x) {
  return Rcpp::String(x);
}
