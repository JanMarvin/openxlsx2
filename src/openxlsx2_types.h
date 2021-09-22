#include <RcppCommon.h>

typedef struct {
  std::string row_r = "NA";
  std::string c_r = "NA";
  std::string c_s = "NA";
  std::string c_t = "NA";
  std::string v = "NA";
  std::string f = "NA";
  std::string f_t = "NA";
  std::string t = "NA";
} xml_col;


#include <Rcpp.h>
#include "pugixml.hpp"

typedef pugi::xml_document xmldoc;
typedef Rcpp::XPtr<xmldoc> XPtrXML;
