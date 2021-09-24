#include <Rcpp.h>
#include <sstream>
#include "pugixml.hpp"


// [[Rcpp::export]]
SEXP readXML(std::string path, bool isfile) {

  pugi::xml_document doc;
  pugi::xml_parse_result result;

  if (isfile) {
    result = doc.load_file(path.c_str(), pugi::parse_default | pugi::parse_escapes);
  } else {
    result = doc.load_string(path.c_str(), pugi::parse_default | pugi::parse_escapes);
  }

  if (!result) {
    Rcpp::stop("xml import unsuccessfull");
  }

  std::ostringstream oss;
  doc.print(oss, " ", pugi::format_raw);

  return  Rcpp::wrap(oss.str());
}

// [[Rcpp::export]]
SEXP getXML1(std::string str, std::string child) {

  pugi::xml_document doc;
  pugi::xml_parse_result result = doc.load_string(str.c_str(), pugi::parse_default | pugi::parse_escapes);
  if (!result) {
    Rcpp::stop("xml import unsuccessfull");
  }


  std::vector<std::string> res;

  for (pugi::xml_node worksheet = doc.child(child.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(child.c_str()))
  {
    std::ostringstream oss;
    worksheet.print(oss, " ", pugi::format_raw);
    res.push_back(oss.str());
  }

  return  Rcpp::wrap(res);
}

// [[Rcpp::export]]
SEXP getXML1val(std::string str, std::string child) {

  pugi::xml_document doc;
  pugi::xml_parse_result result = doc.load_string(str.c_str(), pugi::parse_default | pugi::parse_escapes);
  if (!result) {
    Rcpp::stop("xml import unsuccessfull");
  }


  std::vector<std::string> res;

  for (pugi::xml_node worksheet = doc.child(child.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(child.c_str()))
  {
    res.push_back(worksheet.child_value());
  }

  return  Rcpp::wrap(res);
}

// [[Rcpp::export]]
SEXP getXML2(std::string str, std::string level1, std::string child) {

  pugi::xml_document doc;
  pugi::xml_parse_result result = doc.load_string(str.c_str(), pugi::parse_default | pugi::parse_escapes);
  if (!result) {
    Rcpp::stop("xml import unsuccessfull");
  }


  std::vector<std::string> res;

  for (pugi::xml_node worksheet = doc.child(level1.c_str()).child(child.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(child.c_str()))
  {
    std::ostringstream oss;
    worksheet.print(oss, " ", pugi::format_raw);
    res.push_back(oss.str());
  }

  return  Rcpp::wrap(res);
}

// [[Rcpp::export]]
SEXP getXML2val(std::string str, std::string level1, std::string child) {

  pugi::xml_document doc;
  pugi::xml_parse_result result = doc.load_string(str.c_str(), pugi::parse_default | pugi::parse_escapes);
  if (!result) {
    Rcpp::stop("xml import unsuccessfull");
  }


  std::vector<std::string> res;

  for (pugi::xml_node worksheet = doc.child(level1.c_str()).child(child.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(child.c_str()))
  {
    res.push_back(worksheet.child_value());
  }

  return  Rcpp::wrap(res);
}

// [[Rcpp::export]]
SEXP getXML3(std::string str, std::string level1, std::string level2, std::string child) {

  pugi::xml_document doc;
  pugi::xml_parse_result result = doc.load_string(str.c_str(), pugi::parse_default | pugi::parse_escapes);
  if (!result) {
    Rcpp::stop("xml import unsuccessfull");
  }


  std::vector<std::string> res;

  for (pugi::xml_node worksheet = doc.child(level1.c_str()).child(level2.c_str()).child(child.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(child.c_str()))
  {
    std::ostringstream oss;
    worksheet.print(oss, " ", pugi::format_raw);
    res.push_back(oss.str());
  }

  return  Rcpp::wrap(res);
}

// [[Rcpp::export]]
SEXP getXML3val(std::string str, std::string level1, std::string level2, std::string child) {

  pugi::xml_document doc;
  pugi::xml_parse_result result = doc.load_string(str.c_str(), pugi::parse_default | pugi::parse_escapes);
  if (!result) {
    Rcpp::stop("xml import unsuccessfull");
  }


  std::vector<std::string> res;

  for (pugi::xml_node worksheet = doc.child(level1.c_str()).child(level2.c_str()).child(child.c_str());
       worksheet;
       worksheet = worksheet.next_sibling(child.c_str()))
  {
    res.push_back(worksheet.child_value());
  }

  return  Rcpp::wrap(res);
}
