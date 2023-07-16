#include <xlcpp/xlcpp.h>

// [[Rcpp::export]]
void read_encryption(std::string PATH, std::string OUT, std::string PASSWORD) {


  xlcpp::workbook wb(PATH, PASSWORD, OUT);

  // this creates xlcpp workbook, not the unzipped xlsx file
  // wb.save(OUT);
}
