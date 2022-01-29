#include "openxlsx2_types.h"

#include <iostream>
#include <fstream>
#include <sstream>

std::string int_to_col(uint32_t cell);
Rcpp::IntegerVector col_to_int(Rcpp::CharacterVector x);
std::string set_row(Rcpp::DataFrame row_attr, Rcpp::List cells, size_t row_idx);
