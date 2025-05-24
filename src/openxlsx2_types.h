#ifndef OPENXLSX2_TYPES_H_
#define OPENXLSX2_TYPES_H_

/******************************************************************************
 *                                                                            *
 * This file defines typedefs. R expects it to be called <pkgname>_types.h    *
 *                                                                            *
 ******************************************************************************/

/* create custom Rcpp::wrap function to be used with std::vector<xml_col> */
#include <RcppCommon.h>

typedef struct {
  std::string r;
  std::string row_r;
  std::string c_r;
  std::string_view c_s;
  std::string_view c_t;
  std::string_view c_cm;
  std::string_view c_ph;
  std::string_view c_vm;
  std::string_view v;
  std::string_view f;
  std::string f_attr;
  std::string_view is;
} xml_col;

typedef std::vector<std::string> vec_string;

// matches openxlsx2_celltype in openxlsx2.R
enum celltype {
  short_date     = 0,
  long_date      = 1,
  numeric        = 2,
  logical        = 3,
  character      = 4,
  formula        = 5,
  accounting     = 6,
  percentage     = 7,
  scientific     = 8,
  comma          = 9,
  hyperlink      = 10,
  array_formula  = 11,
  factor         = 12,
  string_num     = 13,
  cm_formula     = 14,
  hms_time       = 15,
  currency       = 16,
  list           = 17
};

// check for 1.0.8.0
#if RCPP_DEV_VERSION >= 1000800
#include <Rcpp/Lightest>
#else
#include <Rcpp.h>
#endif

#include <string_view>

// custom wrap function
// Converts the imported values from c++ std::vector<xml_col> to an R dataframe.
// Whenever new fields are spotted they have to be added here
namespace Rcpp {

template <>
inline SEXP wrap(const vec_string& x) {
  R_xlen_t n = static_cast<R_xlen_t>(x.size());

  Rcpp::CharacterVector z(no_init(n));

  for (R_xlen_t i = 0; i < n; ++i) {
    z[i] = Rcpp::String(x[static_cast<size_t>(i)]);
  }

  return Rcpp::wrap(z);
}


template <>
inline std::string_view as(SEXP s) {
  const char* ptr = Rcpp::String(s).get_cstring();
  return std::string_view(ptr);
}

}  // namespace Rcpp

// pugixml defines. This creates the xmlptr
#include "pugixml.hpp"

typedef pugi::xml_document xmldoc;
typedef Rcpp::XPtr<xmldoc> XPtrXML;

#endif // OPENXLSX2_TYPES_H_
