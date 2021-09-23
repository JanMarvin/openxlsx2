#include "openxlsx2_types.h"

// [[Rcpp::export]]
Rcpp::CharacterVector set_sst(Rcpp::CharacterVector sharedStrings) {

  Rcpp::CharacterVector sst(sharedStrings.length());

    for (auto i = 0; i < sharedStrings.length(); ++i) {
      pugi::xml_document si;
      std::string sharedString = Rcpp::as<std::string>(sharedStrings[i]);

      si.append_child("si").append_child("t").append_child(pugi::node_pcdata).set_value(sharedString.c_str());

      std::ostringstream oss;
      si.print(oss, " ", pugi::format_raw);

      sst[i] = oss.str();
    }

    return sst;
}


// [[Rcpp::export]]
std::string list_to_attr(Rcpp::List attributes, std::string node) {

  pugi::xml_document doc;

  for (auto i = 0; i < attributes.length(); ++i) {
    pugi::xml_node nds = doc.append_child(node.c_str());

    Rcpp::List attrs = attributes[i];
    Rcpp::CharacterVector attrnams = attrs.names();
    for (auto j = 0; j < attrs.length(); ++j) {
      nds.append_attribute(attrnams[j]) = Rcpp::as<std::string>(attrs[j]).c_str();
    }
  }

  std::ostringstream oss;
  doc.print(oss, " ", pugi::format_raw);
  // doc.print(oss);

  return oss.str();
}


// [[Rcpp::export]]
std::string list_to_attr_full(Rcpp::List attributes, std::string node, std::string child) {

  pugi::xml_document doc;
  pugi::xml_node nds = doc.append_child(node.c_str());
  for (auto i = 0; i < attributes.length(); ++i) {
    nds.append_child(child.c_str());

    Rcpp::List attrs = attributes[i];
    Rcpp::CharacterVector attrnams = attrs.names();
    for (auto j = 0; j < attrs.length(); ++j) {
      nds.append_attribute(attrnams[j]) = Rcpp::as<std::string>(attrs[j]).c_str();
    }
  }

  std::ostringstream oss;
  doc.print(oss, " ", pugi::format_raw);
  // doc.print(oss);

  return oss.str();
}
