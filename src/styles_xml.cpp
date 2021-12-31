#include "openxlsx2.h"

// [[Rcpp::export]]
Rcpp::DataFrame read_xf(XPtrXML xml_doc_xf) {

  // reference
  // https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformat?view=openxml-2.8.1
  // openxml 2.8.1
  Rcpp::CharacterVector nams = {
    "applyAlignment",
    "applyBorder",
    "applyFill",
    "applyFont",
    "applyNumberFormat",
    "applyProtection",
    "borderId",
    "fillId",
    "fontId",
    "numFmtId",
    "pivotButton",
    "quotePrefix",
    "xfId",
    // child alignment
    "horizontal",
    "indent",
    "justifyLastLine",
    "readingOrder",
    "relativeIndent",
    "shrinkToFit",
    "textRotation",
    "vertical",
    "wrapText",
    // child extLst
    "extLst",
    // child protection
    "hidden",
    "locked"
  };


  auto nn = std::distance(xml_doc_xf->begin(), xml_doc_xf->end());
  auto kk = nams.length();

  Rcpp::IntegerVector rvec = Rcpp::seq_len(nn);

  // 1. create the list
  Rcpp::List df(kk);
  for (auto i = 0; i < kk; ++i)
  {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
  }

  // 2. fill the list
  // <xf ...>
  auto itr = 0;
  for (auto xml_xf : xml_doc_xf->children()) {

    std::string xf_name = xml_xf.name();
    if (xf_name != "xf")
      Rcpp::stop("xml_node is not xf");

    for (auto attrs : xml_xf.attributes()) {

      Rcpp::CharacterVector attr_name = attrs.name();
      Rcpp::CharacterVector attr_value = attrs.value();

      Rcpp::IntegerVector ii = Rcpp::match(nams, attr_name);

      // check if name is already known
      if (all(Rcpp::is_na(ii))) {
        Rcpp::Rcout << attr_name << ": not found in xf name table" << std::endl;
      } else {
        Rcpp::as<Rcpp::CharacterVector>(df[ii])[itr] = attr_value;
      }

    }

    // only handle known names
    // <alignment ...>
    // <extLst ...>
    // <protection ...>
    for (auto cld : xml_xf.children()) {

      std::string cld_name = cld.name();

      // check known names
      if (cld_name ==  "alignment" | cld_name == "extLst" | cld_name == "protection") {

        for (auto attrs : cld.attributes()) {

          Rcpp::CharacterVector attr_name = attrs.name();
          Rcpp::CharacterVector attr_value = attrs.value();

          Rcpp::IntegerVector ii = Rcpp::match(nams, attr_name);

          // check if name is already known
          if (all(Rcpp::is_na(ii))) {
            Rcpp::Rcout << attr_name << ": not found in xf name table" << std::endl;
          } else {
            Rcpp::as<Rcpp::CharacterVector>(df[ii])[itr] = attr_value;
          }

        }

      } else {
        Rcpp::Rcout << cld_name << ": not valid for xf in openxml 2.8.1" << std::endl;
      }

    } // end aligment, extLst, protection

    ++itr;
  }


  // 3. Create a data.frame
  df.attr("row.names") = rvec;
  df.attr("names") = nams;
  df.attr("class") = "data.frame";

  return df;
}
