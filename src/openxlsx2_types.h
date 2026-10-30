#ifndef OPENXLSX2_TYPES_H
#define OPENXLSX2_TYPES_H

/******************************************************************************
 *                                                                            *
 * This file defines typedefs. R expects it to be called <pkgname>_types.h    *
 *                                                                            *
 ******************************************************************************/

/* create custom Rcpp::wrap function to be used with std::vector<xml_col> */
#include <string>
#include <vector>
#include <RcppCommon.h>

struct xml_col {
  double key;
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

  void clear();
};

typedef std::vector<std::string> vec_string;

// Strings crossing between R and C++. Text is translated to UTF-8 when it
// comes from R and marked as UTF-8 when it goes to R. File and directory names
// are translated to the native encoding, because the file system expects that.
// Rcpp converts both types in the generated RcppExports.cpp.
struct utf8_string : std::string {
  utf8_string() = default;
  utf8_string(std::string x) : std::string(std::move(x)) {}
  utf8_string(const char* x) : std::string(x) {}
};

struct native_path : std::string {
  native_path() = default;
  native_path(std::string x) : std::string(std::move(x)) {}
};

namespace Rcpp {
  template <> utf8_string as(SEXP x);
  template <> std::vector<utf8_string> as(SEXP x);
  template <> native_path as(SEXP x);
  template <> std::vector<native_path> as(SEXP x);
  template <> SEXP wrap(const utf8_string& x);
  template <> SEXP wrap(const std::vector<utf8_string>& x);
}

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
  list           = 17,
  CELLTYPE_MAX  // always the last
};

#if RCPP_DEV_VERSION >= 1000800
#include <Rcpp/Lightest>
#else
#include <Rcpp.h>
#endif

inline SEXP utf8_charsxp(const std::string& x) {
  return Rf_mkCharLenCE(x.data(), static_cast<int>(x.size()), CE_UTF8);
}

namespace Rcpp {
  template <> inline utf8_string as(SEXP x) {
    if (TYPEOF(x) != STRSXP || Rf_xlength(x) != 1) Rcpp::stop("Expecting a single string value");
    return utf8_string(Rf_translateCharUTF8(STRING_ELT(x, 0)));
  }

  template <> inline std::vector<utf8_string> as(SEXP x) {
    // like Rcpp's as<std::vector<std::string>>(), other types are coerced
    Rcpp::Shield<SEXP> strs(Rcpp::r_cast<STRSXP>(x));
    std::vector<utf8_string> res;
    res.reserve(static_cast<size_t>(Rf_xlength(strs)));
    for (R_xlen_t i = 0; i < Rf_xlength(strs); ++i) {
      res.push_back(utf8_string(Rf_translateCharUTF8(STRING_ELT(strs, i))));
    }
    return res;
  }

  template <> inline native_path as(SEXP x) {
    if (TYPEOF(x) != STRSXP || Rf_xlength(x) != 1) Rcpp::stop("Expecting a single string value");
    return native_path(Rf_translateChar(STRING_ELT(x, 0)));
  }

  template <> inline std::vector<native_path> as(SEXP x) {
    Rcpp::Shield<SEXP> strs(Rcpp::r_cast<STRSXP>(x));
    std::vector<native_path> res;
    res.reserve(static_cast<size_t>(Rf_xlength(strs)));
    for (R_xlen_t i = 0; i < Rf_xlength(strs); ++i) {
      res.push_back(native_path(Rf_translateChar(STRING_ELT(strs, i))));
    }
    return res;
  }

  template <> inline SEXP wrap(const utf8_string& x) {
    return Rf_ScalarString(utf8_charsxp(x));
  }

  template <> inline SEXP wrap(const std::vector<utf8_string>& x) {
    Rcpp::Shield<SEXP> res(Rf_allocVector(STRSXP, static_cast<R_xlen_t>(x.size())));
    for (size_t i = 0; i < x.size(); ++i) {
      SET_STRING_ELT(res, static_cast<R_xlen_t>(i), utf8_charsxp(x[i]));
    }
    return res;
  }
}

inline void xml_col::clear() {
  key = NA_REAL;
  r.clear(); row_r.clear(); c_r.clear(); c_s.clear(); c_t.clear();
  c_cm.clear(); c_ph.clear(); c_vm.clear();
  v.clear(); f.clear(); f_attr.clear(); is.clear();
}

// pugixml defines. This creates the xmlptr
#include "pugixml.hpp"

typedef pugi::xml_document xmldoc;
typedef Rcpp::XPtr<xmldoc> XPtrXML;

#endif
