
#include "openxlsx2.h"



// [[Rcpp::export]]  
Rcpp::IntegerVector build_cell_types_integer(Rcpp::CharacterVector classes, int n_rows){
  
  // 0: "n"
  // 1: "s"
  // 2: "b"
  // 9: "h"
  // 4: TBC
  // 5: TBC
  
  size_t n_cols = classes.size();
  Rcpp::IntegerVector col_t(n_cols);
  
  for(size_t i = 0; i < n_cols; i++){
    
    if((classes[i] == "numeric") | (classes[i] == "integer") | (classes[i] == "raw") ){
      col_t[i] = 0; 
    }else if(classes[i] == "character"){
      col_t[i] = 1; 
    }else if(classes[i] == "logical"){
      col_t[i] = 2;
    }else if(classes[i] == "hyperlink"){
      col_t[i] = 9;
    }else if(classes[i] == "openxlsx_formula"){
      col_t[i] = NA_INTEGER;
    }else{
      col_t[i] = 1;
    }
    
  }
  
  Rcpp::IntegerVector cell_types = rep(col_t, n_rows); 
  
  return cell_types;
  
  
}

// [[Rcpp::export]]
Rcpp::List build_cell_merges(Rcpp::List comps){
  
  size_t nMerges = comps.size(); 
  Rcpp::List res(nMerges);
  
  for(size_t i =0; i < nMerges; i++){
    Rcpp::IntegerVector col = convert_from_excel_ref(comps[i]);  
    Rcpp::CharacterVector comp = comps[i];
    Rcpp::IntegerVector row(2);  
    
    for(size_t j = 0; j < 2; j++){
      std::string rt(comp[j]);      
      rt.erase(std::remove_if(rt.begin(), rt.end(), ::isalpha), rt.end());
      row[j] = atoi(rt.c_str());
    }
    
    size_t ca(col[0]);
    size_t ck = size_t(col[1]) - ca + 1;
    
    std::vector<int> v(ck) ;
    for(size_t j = 0; j < ck; j++)
      v[j] = j + ca;
    
    size_t ra(row[0]);
    
    size_t rk = int(row[1]) - ra + 1;
    std::vector<int> r(rk) ;
    for(size_t j = 0; j < rk; j++)
      r[j] = j + ra;
    
    Rcpp::CharacterVector M(ck*rk);
    int ind = 0;
    for(size_t j = 0; j < ck; j++){
      for(size_t k = 0; k < rk; k++){
        char name[30];
        sprintf(&(name[0]), "%d-%d", r[k], v[j]);
        M(ind) = name;
        ind++;
      }
    }
    
    res[i] = M;
  }
  
  return wrap(res) ;
  
}
