#include "openxlsx2.h"
#include <cctype>

// converts sharedstrings xml tree to R-Character Vector
// [[Rcpp::export]]
SEXP si_to_txt(XPtrXML doc) {

  auto sst = doc->child("sst");
  auto n = std::distance(sst.begin(), sst.end());

  Rcpp::CharacterVector res(Rcpp::no_init(n));

  auto i = 0;
  for (auto si : doc->child("sst").children("si"))
  {
    // text to export
    std::string text = "";

    // has only t node
    for (auto t : si.children("t")) {
      text = t.child_value();
    }

    // has r node with t node
    // linebreaks and spaces are handled in the nodes
    for (auto r : si.children("r")) {
      for (auto t :r.children("t")) {
        text += t.child_value();
      }
    }

    // push everything back
    res[i] = Rcpp::String(text);
    ++i;
  }

  return res;
}


// [[Rcpp::export]]
std::string txt_to_si(Rcpp::CharacterVector txt,
                      bool no_escapes = false, bool raw = true, bool skip_control = true) {

  pugi::xml_document doc;

  unsigned int pugi_format_flags = pugi::format_indent;
  if (no_escapes) pugi_format_flags |= pugi::format_no_escapes;
  if (raw)  pugi_format_flags |= pugi::format_raw;
  if (skip_control) pugi_format_flags |= pugi::format_skip_control_chars;

  pugi::xml_node si_node = doc.append_child("si");

  auto i = 0;
  for (auto si = 0; si < txt.length(); ++si)
  {
    // text to export
    std::string text = Rcpp::as<std::string>(txt(si));
    pugi::xml_node t_node = si_node.append_child("t");

    if ((text.size() > 0) && (std::isspace(text.at(0)) || std::isspace(text.at(text.size()-1)))) {
      t_node.append_attribute("xml:space").set_value("preserve");
    }
    t_node.append_child(pugi::node_pcdata).set_value(text.c_str());

    ++i;
  }

  std::ostringstream oss;
  doc.print(oss, " ", pugi_format_flags);
  std::string xml_return = oss.str();

  return xml_return;
}


// converts inlineStr xml tree to R-Character Vector
// [[Rcpp::export]]
SEXP is_to_txt(Rcpp::CharacterVector is_vec) {

  auto n = is_vec.length();
  Rcpp::CharacterVector res(Rcpp::no_init(n));

  for (auto i = 0; i < n; ++i) {

    std::string tmp = Rcpp::as<std::string>(is_vec[i]);

    pugi::xml_document doc;
    pugi::xml_parse_result result = doc.load_string(tmp.c_str(), pugi::parse_default | pugi::parse_ws_pcdata | pugi::parse_escapes);

    if (!result) {
      Rcpp::stop("inlineStr xml import unsuccessfull");
    }

    for (auto is : doc.children("is"))
    {
      // text to export
      std::string text = "";

      // has only t node
      for (auto t : is.children("t")) {
        text = t.child_value();
      }

      // has r node with t node
      // phoneticPr (Phonetic Properties)
      // r (Rich Text Run)
      // rPr (Run Properties)
      // rPh (Phonetic Run)
      // t (Text)
      // linebreaks and spaces are handled in the nodes
      for (auto r : is.children("r")) {
        for (auto t :r.children("t")) {
          text += t.child_value();
        }
      }

      // push everything back
      res[i] = text;
    }

  }

  return res;
}


// [[Rcpp::export]]
std::string txt_to_is(std::string text,
                      bool no_escapes = false, bool raw = true, bool skip_control = true) {

  pugi::xml_document doc;

  unsigned int pugi_format_flags = pugi::format_indent;
  if (no_escapes) pugi_format_flags |= pugi::format_no_escapes;
  if (raw)  pugi_format_flags |= pugi::format_raw;
  if (skip_control) pugi_format_flags |= pugi::format_skip_control_chars;

  pugi::xml_node is_node = doc.append_child("is");

  // text to export
  pugi::xml_node t_node = is_node.append_child("t");;

  if ((text.size() > 0) && (std::isspace(text.at(0)) || std::isspace(text.at(text.size()-1)))) {
    t_node.append_attribute("xml:space").set_value("preserve");
  }

  t_node.append_child(pugi::node_pcdata).set_value(text.c_str());

  std::ostringstream oss;
  doc.print(oss, " ", pugi_format_flags);
  std::string xml_return = oss.str();

  return xml_return;
}
