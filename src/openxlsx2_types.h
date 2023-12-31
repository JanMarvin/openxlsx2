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
  std::string c_r;   // CellReference
  std::string c_s;   // StyleIndex
  std::string c_t;   // DataType
  std::string c_cm;  // CellMetaIndex
  std::string c_ph;  // ShowPhonetic
  std::string c_vm;  // ValueMetaIndex
  std::string v;     // CellValue
  std::string f;     // CellFormula
  std::string f_t;
  std::string f_ref;
  std::string f_ca;
  std::string f_si;
  std::string is;    // inlineStr
} xml_col;


typedef struct {
  std::string v;
  std::string c_cm;
  std::string c_s;
  std::string c_t;
  std::string is;
  std::string f;
  std::string f_t;
  std::string f_ref;
  std::string typ;
  std::string r;
} celltyp;

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
  currency       = 16
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
inline SEXP wrap(const std::vector<xml_col> &x) {

  size_t n = x.size();

  // Vector structure identical to xml_col from openxlsx2_types.h
  Rcpp::CharacterVector r(no_init(n));         // cell name: A1, A2 ...
  Rcpp::CharacterVector row_r(no_init(n));     // row name: 1, 2, ..., 9999

  Rcpp::CharacterVector c_r(no_init(n));       // col name: A, B, ..., ZZ
  Rcpp::CharacterVector c_s(no_init(n));       // cell style
  Rcpp::CharacterVector c_t(no_init(n));       // cell type
  Rcpp::CharacterVector c_cm(no_init(n));
  Rcpp::CharacterVector c_ph(no_init(n));
  Rcpp::CharacterVector c_vm(no_init(n));

  Rcpp::CharacterVector v(no_init(n));         // <v> tag
  Rcpp::CharacterVector f(no_init(n));         // <f> tag
  Rcpp::CharacterVector f_t(no_init(n));       // <f t=""> attribute most likely shared
  Rcpp::CharacterVector f_ref(no_init(n));     // <f ref=""> attribute most likely reference
  Rcpp::CharacterVector f_ca(no_init(n));      // <f ca=""> attribute most likely conditional formatting
  Rcpp::CharacterVector f_si(no_init(n));      // <f si=""> attribute most likely sharedString
  Rcpp::CharacterVector is(no_init(n));        // <is> tag

  // struct to vector
  for (size_t i = 0; i < n; ++i) {
    if (!x[i].r.empty())     r[i]     = Rcpp::String(x[i].r);
    if (!x[i].row_r.empty()) row_r[i] = Rcpp::String(x[i].row_r);
    if (!x[i].c_r.empty())   c_r[i]   = Rcpp::String(x[i].c_r);
    if (!x[i].c_s.empty())   c_s[i]   = Rcpp::String(x[i].c_s);
    if (!x[i].c_t.empty())   c_t[i]   = Rcpp::String(x[i].c_t);
    if (!x[i].c_cm.empty())  c_cm[i]  = Rcpp::String(x[i].c_cm);
    if (!x[i].c_ph.empty())  c_ph[i]  = Rcpp::String(x[i].c_ph);
    if (!x[i].c_vm.empty())  c_vm[i]  = Rcpp::String(x[i].c_vm);
    if (!x[i].v.empty())     v[i]     = Rcpp::String(x[i].v);
    if (!x[i].f.empty())     f[i]     = Rcpp::String(x[i].f);
    if (!x[i].f_t.empty())   f_t[i]   = Rcpp::String(x[i].f_t);
    if (!x[i].f_ref.empty()) f_ref[i] = Rcpp::String(x[i].f_ref);
    if (!x[i].f_ca.empty())  f_ca[i]  = Rcpp::String(x[i].f_ca);
    if (!x[i].f_si.empty())  f_si[i]  = Rcpp::String(x[i].f_si);
    if (!x[i].is.empty())    is[i]    = Rcpp::String(x[i].is);
  }

  // Assign and return a dataframe
  return Rcpp::wrap(
    Rcpp::DataFrame::create(
      Rcpp::Named("r")     = r,
      Rcpp::Named("row_r") = row_r,
      Rcpp::Named("c_r")   = c_r,
      Rcpp::Named("c_s")   = c_s,
      Rcpp::Named("c_t")   = c_t,
      Rcpp::Named("c_cm")  = c_cm,
      Rcpp::Named("c_ph")  = c_ph,
      Rcpp::Named("c_vm")  = c_vm,
      Rcpp::Named("v")     = v,
      Rcpp::Named("f")     = f,
      Rcpp::Named("f_t")   = f_t,
      Rcpp::Named("f_ref") = f_ref,
      Rcpp::Named("f_ca")  = f_ca,
      Rcpp::Named("f_si")  = f_si,
      Rcpp::Named("is")    = is,
      Rcpp::Named("stringsAsFactors") = false
    )
  );
}

template <>
inline SEXP wrap(const vec_string &x) {

  size_t n = x.size();

  Rcpp::CharacterVector z(no_init(n));

  for (size_t i = 0; i < n; ++i) {
    z[i] = Rcpp::String(x[i]);
  }

  return Rcpp::wrap(z);
}

}

// pugixml defines. This creates the xmlptr
#include "pugixml.hpp"

typedef pugi::xml_document xmldoc;
typedef Rcpp::XPtr<xmldoc> XPtrXML;
