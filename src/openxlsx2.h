#include "openxlsx2_types.h"

#include <iostream>
#include <fstream>
#include <sstream>

// load workbook2
Rcpp::IntegerVector cell_ref_to_col(Rcpp::CharacterVector x );
Rcpp::CharacterVector int_2_cell_ref(Rcpp::IntegerVector);


std::vector<std::string> get_letters();


Rcpp::IntegerVector convert_from_excel_ref( Rcpp::CharacterVector x );

SEXP calc_column_widths(Rcpp::Environment sheet_data, std::vector<std::string> sharedStrings, Rcpp::IntegerVector autoColumns, Rcpp::NumericVector widths, float baseFontCharWidth, float minW, float maxW);

SEXP convert_to_excel_ref_expand(const std::vector<int>& cols, const std::vector<std::string>& LETTERS, const std::vector<std::string>& rows);

std::string set_row(Rcpp::DataFrame row_attr, Rcpp::List cells, size_t row_idx);
