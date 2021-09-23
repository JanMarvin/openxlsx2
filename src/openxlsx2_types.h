/******************************************************************************
 *                                                                            *
 * This file defines typedefs. R expects it to be called <pkgname>_types.h    *
 *                                                                            *
 ******************************************************************************/

/* create custom Rcpp::wrap function to be used with std::vector<xml_col> */
#include <RcppCommon.h>

typedef struct {
  std::string row_r = "NA";
  std::string c_r   = "NA";
  std::string c_s   = "NA";
  std::string c_t   = "NA";
  std::string v     = "NA";
  std::string f     = "NA";
  std::string f_t   = "NA";
  std::string t     = "NA";
} xml_col;

#include <Rcpp.h>

// custom wrap function
// Converts the imported values from c++ std::vector<xml_col> to an R dataframe.
// Whenever new fields are spotted they have to be added here
namespace Rcpp {
template <>
inline SEXP wrap(const std::vector<xml_col> &x) {

  auto n = x.size();

  // Vector structure identical to xml_col from openxlsx2_types.h
  Rcpp::CharacterVector row_r(n);     // row name: 1, 2, ..., 9999

  Rcpp::CharacterVector c_r(n);       // col name: A, B, ..., ZZ
  Rcpp::CharacterVector c_s(n);       // cell style
  Rcpp::CharacterVector c_t(n);       // cell type

  Rcpp::CharacterVector v(n);         // <v> tag
  Rcpp::CharacterVector f(n);         // <f> tag
  Rcpp::CharacterVector f_t(n);       // <f t=""> attribute most likely shared
  Rcpp::CharacterVector t(n);         // <is><t> tag

  // struct to vector
  for (auto i = 0; i < n; ++i) {
    row_r[i] = x[i].row_r;
    c_r[i]   = x[i].c_r;
    c_s[i]   = x[i].c_s;
    c_t[i]   = x[i].c_t;
    v[i]     = x[i].v;
    f[i]     = x[i].f;
    f_t[i]   = x[i].f_t;
    t[i]     = x[i].t;
  }

  // Assign and return a dataframe
  return Rcpp::wrap(Rcpp::DataFrame::create(
      Rcpp::Named("row_r") = row_r,
      Rcpp::Named("c_r")   = c_r,
      Rcpp::Named("c_s")   = c_s,
      Rcpp::Named("c_t")   = c_t,
      Rcpp::Named("v")     = v,
      Rcpp::Named("f")     = f,
      Rcpp::Named("f_t")   = f_t,
      Rcpp::Named("t")     = t
  )
  );
};
}

// pugixml defines. This creates the xmlptr
#include "pugixml.hpp"

typedef pugi::xml_document xmldoc;
typedef Rcpp::XPtr<xmldoc> XPtrXML;
