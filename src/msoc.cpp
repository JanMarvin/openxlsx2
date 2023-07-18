#include "openxlsx2.h"
#include "msoc.h"

// [[Rcpp::export]]
int msoc(std::string mode, std::string inFile, std::string outFile, std::string pass)
{
  int err = 0;

  if (mode.compare("enc") == 0) {
    err = MSOC_encryptA(outFile.c_str(), inFile.c_str(), pass.c_str(), NULL);
  } else if (mode.compare("dec") == 0) {
    err = MSOC_decryptA(outFile.c_str(), inFile.c_str(), pass.c_str(), NULL);
  } else {
    Rcpp::stop("Input must be enc/dec. Was: %s\n", mode.c_str());
  }
  if (err != MSOC_NOERR) {
    Rcpp::stop("ERR %s\n", MSOC_getErrMessage(err));
    return 1;
  } else {
    // return err;
  }
  return 0;
}
