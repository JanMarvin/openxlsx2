

#include "openxlsx.h"



Rcpp::IntegerVector which_cpp(Rcpp::LogicalVector x) {
  Rcpp::IntegerVector v = Rcpp::seq(0, x.size() - 1);
  return v[x];
}


// [[Rcpp::export]]  
int calc_number_rows(Rcpp::CharacterVector x, bool skipEmptyRows){
  
  int n = x.size();
  if(n == 0)
    return(0);
  
  int nRows;
  
  if(skipEmptyRows){
    
    Rcpp::CharacterVector res(n);
    std::string r;
    for(int i = 0; i < n; i++){
      r = x[i];
      r.erase(std::remove_if(r.begin(), r.end(), ::isalpha), r.end());
      res[i] = r;
    }
    
    Rcpp::CharacterVector uRes = unique(res);
    nRows = uRes.size();
    
  }else{
    
    std::string fRef = Rcpp::as<std::string>(x[0]);
    std::string lRef = Rcpp::as<std::string>(x[n-1]);
    fRef.erase(std::remove_if(fRef.begin(), fRef.end(), ::isalpha), fRef.end());
    lRef.erase(std::remove_if(lRef.begin(), lRef.end(), ::isalpha), lRef.end());
    int firstRow = atoi(fRef.c_str());
    int lastRow = atoi(lRef.c_str());
    nRows = lastRow - firstRow + 1;
    
  }
  
  return(nRows);
  
}
