#include "openxlsx_types.h"

// [[Rcpp::export]]
Rcpp::DataFrame read_ws_d(std::string path,Rcpp::DataFrame shared_strings) {
  
  pugi::xml_document doc;
  
  doc.load_file(path.c_str());
  // Rcpp::List row;
  std::vector<std::string> col_r;
  std::vector<std::string> col_s;
  std::vector<std::string> col_t;
  std::vector<std::string> col_v;
  std::vector<int> col_rn;
  std::vector<int> col_cn;
  //pugi::xpath_query query_remote_tools("/worksheet/sheetData/row");
  
  std::vector<int> ssi = shared_strings["SharedStringIndex"];
  std::vector<std::string> ss = shared_strings["SharedString"];
  // std::vector<char> col_atr("r","s","t","v");
  

  
  // tag::code[]
  // Select nodes via compiled query
  //pugi::xpath_query query_remote_tools("/Profile/Tools/Tool[@AllowRemote='true']");
  
  // pugi::xpath_node_set tools = query_remote_tools.evaluate_node_set(doc);
  // std::cout << "Remote tool: ";
  // tools.print(std::cout);
  
  
  // Evaluate numbers via compiled query
  //pugi::xpath_query query_timeouts("sum(//Tool/@Timeout)");
  //std::cout << query_timeouts.evaluate_number(doc) << std::endl;
  
  // Evaluate strings via compiled query for different context nodes
  // pugi::xpath_query query_name_valid("string-length(substring-before(@Filename, '_')) > 0 and @OutputFileMasks");
  // pugi::xpath_query query_name("concat(substring-before(@Filename, '_'), ' produces ', @OutputFileMasks)");
  // 
  int row_ctr=1;
  for (pugi::xml_node rowx = doc.first_element_by_path("/worksheet/sheetData/row"); rowx; rowx = rowx.next_sibling())
  {
    
    
    
    int col_ctr=1;
    
    for (pugi::xml_node colx = rowx.first_child(); colx; colx = colx.next_sibling())
    {
      // for (pugi::xml_attribute attr = colx.first_attribute(); attr; attr = attr.next_attribute())
      // {
      //   std::cout << " " << attr.name() << "=" << attr.value();
      //   
      
      if(strcmp(colx.child_value("v")  ,"")==0){
        
      }else{
      
      
      if(rowx.attribute("r").as_int()==0){
        col_rn.push_back(row_ctr);
      } else{
      col_rn.push_back(rowx.attribute("r").as_int());
      }
      col_cn.push_back(col_ctr);
      col_r.push_back(colx.attribute("r").value());
      col_s.push_back(colx.attribute("s").value());
      col_t.push_back(colx.attribute("t").value());
      
      // std::cout << "Cell: "  << colx.attribute("r").value() << "\t";
      // std::cout << "s: "  << colx.attribute("s").value() << "\t";
      // std::cout << "t: "  << colx.attribute("t").value() << "\t";
      // std::cout << "v: "  << colx.child_value("v") << "\t";
      //x1[t=="s",ss,on=c("v==s"),':='(val=SharedString)]
      if(strcmp(colx.attribute("t").as_string(),"s")==0){
        col_v.push_back(ss[std::stoi( colx.child_value("v"))]);
        std::cout << "SS: "  << ss[std::stoi( colx.child_value("v"))] << "\n";
      } else{
        col_v.push_back(colx.child_value("v"));
        std::cout << "\n";
      }
      }
      
      
      col_ctr++;
      
      
      
      
      
         // col_r.push_back(attr.value());
        
      }
    
    row_ctr++;
      
    }
    
    
    // std::string s = query_name.evaluate_string(tool);
    // 
    // if (query_name_valid.evaluate_string(tool)) std::cout << s << std::endl;
    // rowx.print(std::cout);
    
    Rcpp::List res;
    
    res= Rcpp::List::create(Rcpp::Named("rn") = col_rn,
                            Rcpp::Named("cn") = col_cn,
                        Rcpp::Named("r") = col_r,
                       Rcpp::Named("s") = col_s,
                       Rcpp::Named("t") = col_t,
                       Rcpp::Named("v") = col_v);
  
    res.attr("class") = Rcpp::CharacterVector::create("data.table", "data.frame");
    return res;
    
  }


// [[Rcpp::export]]
Rcpp::DataFrame read_shared_strings(std::string path) {
 
  pugi::xml_document doc;
  
  doc.load_file(path.c_str());
  // Rcpp::List row;
  
  std::vector<std::string> shared_string;
  std::vector<int> shared_string_index;
  int index = 0;
  
  
  
  for (pugi::xml_node ss = doc.first_element_by_path("/sst/si"); ss; ss = ss.next_sibling())
  {
    shared_string.push_back(ss.child_value("t"));
    shared_string_index.push_back(index);
    index++;
  }
  
  
  
  Rcpp::List res_ss;
  res_ss["SharedStringIndex"]=shared_string_index;
  res_ss["SharedString"]=shared_string;
  
  
  res_ss.attr("class") = Rcpp::CharacterVector::create("data.table", "data.frame");
  
  return res_ss; 
  
}


