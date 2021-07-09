
#include "openxlsx.h"







// [[Rcpp::export]]   
Rcpp::IntegerVector map_cell_types_to_integer(Rcpp::CharacterVector t){
  
  // 0: "n"
  // 1: "s"
  // 2: "b"
  // 3: "str"
  // 4: "e"
  // 9: "h"
  
  size_t n = t.size();
  Rcpp::IntegerVector t_res(n);
  
  for(size_t i = 0; i < n; i++){
    
    if(Rcpp::CharacterVector::is_na(t[i])){
      t_res[i] = NA_INTEGER;
    }else if(t[i] == "n"){
      t_res[i] = 0;
    }else if(t[i] == "s"){
      t_res[i] = 1;
    }else if(t[i] == "b"){
      t_res[i] = 2;
    }else if(t[i] == "str"){
      t_res[i] = 3;
    }else if(t[i] == "e"){
      t_res[i] = 4;
    }
    
  }
  
  return t_res;
  
}




// [[Rcpp::export]]   
Rcpp::CharacterVector map_cell_types_to_char(Rcpp::IntegerVector t){
  
  // 0: "n"
  // 1: "s"
  // 2: "b"
  // 3: "str"
  // 4: "e"
  // 9: "h"
  
  size_t n = t.size();
  Rcpp::CharacterVector t_res(n);
  
  for(size_t i = 0; i < n; i++){
    
    if(Rcpp::IntegerVector::is_na(t[i])){
      t_res[i] = NA_STRING;
    }else if(t[i] == 0){
      t_res[i] = "n";
    }else if(t[i] == 1){
      t_res[i] = "s";
    }else if(t[i] == 2){
      t_res[i] = "b";
    }else if(t[i] == 3){
      t_res[i] = "str";
    }else if(t[i] == 4){
      t_res[i] = "e";
    }else{
      t_res[i] = "s";
    }
    
  }
  
  return t_res;
  
}



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
Rcpp::CharacterVector buildCellTypes(Rcpp::CharacterVector classes, int nRows){
  
  
  int nCols = classes.size();
  Rcpp::CharacterVector colLabels(nCols);
  for(int i=0; i < nCols; i++){
    
    if((classes[i] == "numeric") | (classes[i] == "integer") | (classes[i] == "raw") ){
      colLabels[i] = "n"; 
    }else if(classes[i] == "character"){
      colLabels[i] = "s"; 
    }else if(classes[i] == "logical"){
      colLabels[i] = "b";
    }else if(classes[i] == "hyperlink"){
      colLabels[i] = "h";
    }else if(classes[i] == "openxlsx_formula"){
      colLabels[i] = NA_STRING;
    }else{
      colLabels[i] = "s";
    }
    
  }
  
  Rcpp::CharacterVector cellTypes = rep(colLabels, nRows); 
  
  return wrap(cellTypes);
  
  
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




// [[Rcpp::export]]
Rcpp::List buildCellList( Rcpp::CharacterVector r, Rcpp::CharacterVector t, Rcpp::CharacterVector v) {
  
  //Valid combinations
  //  r t v	
  //  T	F	F	
  //  T	T	T
  //  F F	F	
  //  T F	T (must be a formula)	
  
  int n = r.size();
  Rcpp::List cells(n);
  Rcpp::LogicalVector hasV = !is_na(v);
  Rcpp::LogicalVector hasR = !is_na(r);
  Rcpp::LogicalVector hasT = !is_na(t);
  
  for(int i=0; i < n; i++){
    
    if(hasR[i]){
      
      if(hasV[i]){
        
        if(hasT[i]){
          
          //  r t v	
          //  T	T	T (2)
          cells[i] = Rcpp::CharacterVector::create(
            Rcpp::Named("r") = r[i],
                                Rcpp::Named("t") = t[i],
                                                    Rcpp::Named("v") = v[i],
                                                                        Rcpp::Named("f") = NA_STRING); 
          
        }else{
          
          //  r t f	
          //  T	T	T (4 - formula)
          cells[i] = Rcpp::CharacterVector::create(
            Rcpp::Named("r") = r[i],
                                Rcpp::Named("t") = "str",
                                Rcpp::Named("v") = NA_STRING,
                                Rcpp::Named("f") = "<f>" + v[i] + "</f>"); 
          
          
        }
        
      }else{
        
        //  r t v	
        //  T	F	F	(1)
        cells[i] = Rcpp::CharacterVector::create(
          Rcpp::Named("r") = r[i],
                              Rcpp::Named("t") = NA_STRING,
                              Rcpp::Named("v") = NA_STRING,
                              Rcpp::Named("f") = NA_STRING); 
      }
      
    }else{
      
      //  r t v	
      //  F F	F	(3)
      cells[i] = Rcpp::CharacterVector::create(
        Rcpp::Named("r") = NA_STRING,
        Rcpp::Named("t") = NA_STRING,
        Rcpp::Named("v") = NA_STRING,
        Rcpp::Named("f") = NA_STRING);  
    }
    
  } // end of for loop
  
  return Rcpp::wrap(cells) ;
}

