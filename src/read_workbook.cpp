

#include "openxlsx2.h"



Rcpp::IntegerVector which_cpp(Rcpp::LogicalVector x) {
  Rcpp::IntegerVector v = Rcpp::seq(0, x.size() - 1);
  return v[x];
}
