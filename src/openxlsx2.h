#include "openxlsx2_types.h"

#include <iostream>
#include <fstream>
#include <sstream>

// load workbook2
Rcpp::IntegerVector cell_ref_to_col(Rcpp::CharacterVector x );
Rcpp::CharacterVector int_2_cell_ref(Rcpp::IntegerVector);


// write_data.cpp
Rcpp::CharacterVector map_cell_types_to_char(Rcpp::IntegerVector);
Rcpp::IntegerVector map_cell_types_to_integer(Rcpp::CharacterVector);



std::vector<std::string> get_letters();


Rcpp::IntegerVector convert_from_excel_ref( Rcpp::CharacterVector x );

SEXP calc_column_widths(Rcpp::Reference sheet_data, std::vector<std::string> sharedStrings, Rcpp::IntegerVector autoColumns, Rcpp::NumericVector widths, float baseFontCharWidth, float minW, float maxW);

SEXP getOpenClosedNode(std::string xml, std::string open_tag, std::string close_tag);

std::string cppReadFile(std::string xmlFile);

Rcpp::CharacterVector get_extLst_Major(std::string xml);
SEXP getAttr(Rcpp::CharacterVector x, std::string tag);

Rcpp::List buildCellList( Rcpp::CharacterVector r, Rcpp::CharacterVector t, Rcpp::CharacterVector v);
SEXP openxlsx_convert_to_excel_ref(Rcpp::IntegerVector cols, std::vector<std::string> LETTERS);

SEXP buildMatrixNumeric(Rcpp::CharacterVector v, Rcpp::IntegerVector rowInd, Rcpp::IntegerVector colInd,
                        Rcpp::CharacterVector colNames, int nRows, int nCols);

SEXP buildMatrixMixed(Rcpp::CharacterVector v,
                      Rcpp::IntegerVector rowInd,
                      Rcpp::IntegerVector colInd,
                      Rcpp::CharacterVector colNames,
                      int nRows,
                      int nCols,
                      Rcpp::IntegerVector charCols,
                      Rcpp::IntegerVector dateCols);

SEXP convert_to_excel_ref_expand(const std::vector<int>& cols, const std::vector<std::string>& LETTERS, const std::vector<std::string>& rows);

Rcpp::IntegerVector matrixRowInds(Rcpp::IntegerVector indices);
Rcpp::CharacterVector build_table_xml(std::string table, std::string ref, std::vector<std::string> colNames, bool showColNames, std::string tableStyle, bool withFilter);
int calc_number_rows(Rcpp::CharacterVector x, bool skipEmptyRows);
Rcpp::CharacterVector buildCellTypes(Rcpp::CharacterVector classes, int nRows);
Rcpp::LogicalVector isInternalHyperlink(Rcpp::CharacterVector x);


// helper functions
std::string itos(int i);
Rcpp::CharacterVector markUTF8(Rcpp::CharacterVector x, bool clone = false);

std::string set_row(Rcpp::DataFrame row_attr, Rcpp::List cells, size_t row_idx);
