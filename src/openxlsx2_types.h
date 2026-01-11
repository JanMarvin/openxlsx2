#ifndef OPENXLSX2_TYPES_H
#define OPENXLSX2_TYPES_H

/******************************************************************************
 *                                                                            *
 * This file defines typedefs. R expects it to be called <pkgname>_types.h    *
 *                                                                            *
 ******************************************************************************/

/* create custom Rcpp::wrap function to be used with std::vector<xml_col> */
#include <RcppCommon.h>

struct xml_col {
  std::string r;
  std::string row_r;
  std::string c_r;     // CellReference
  std::string c_s;     // StyleIndex
  std::string c_t;     // DataType
  std::string c_cm;    // CellMetaIndex
  std::string c_ph;    // ShowPhonetic
  std::string c_vm;    // ValueMetaIndex
  std::string v;       // CellValue
  std::string f;       // CellFormula
  std::string f_attr;
  std::string is;      // inlineStr

  void clear() {
      r.clear(); row_r.clear(); c_r.clear(); c_s.clear(); c_t.clear();
      c_cm.clear(); c_ph.clear(); c_vm.clear();
      v.clear(); f.clear(); f_attr.clear(); is.clear();
    }
};

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

// custom wrap function
// Converts the imported values from c++ std::vector<xml_col> to an R dataframe.
// Whenever new fields are spotted they have to be added here
namespace Rcpp {

template <>
inline SEXP wrap(const vec_string& x) {
  R_xlen_t n = static_cast<R_xlen_t>(x.size());

  Rcpp::CharacterVector z(n);

  for (R_xlen_t i = 0; i < n; ++i) {
    z[i] = Rcpp::String(x[static_cast<size_t>(i)]);
  }

  return Rcpp::wrap(z);
}

}  // namespace Rcpp

// pugixml defines. This creates the xmlptr
#include "pugixml.hpp"

typedef pugi::xml_document xmldoc;
typedef Rcpp::XPtr<xmldoc> XPtrXML;

#endif
