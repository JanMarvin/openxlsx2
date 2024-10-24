#include "openxlsx2.h"
#include <cctype>

// converts sharedstrings xml tree to R-Character Vector
// [[Rcpp::export]]
SEXP xml_si_to_txt(XPtrXML doc) {

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
      text = t.text().get();
    }

    // has r node with t node
    // linebreaks and spaces are handled in the nodes
    for (auto r : si.children("r")) {
      for (auto t :r.children("t")) {
        text += t.text().get();
      }
    }

    // push everything back
    res[i] = Rcpp::String(text);
    ++i;
  }

  return res;
}

SEXP xml_to_txt(Rcpp::CharacterVector vec, std::string type) {

  auto n = vec.length();
  Rcpp::CharacterVector res(Rcpp::no_init(n));

  for (auto i = 0; i < n; ++i) {

    std::string tmp = Rcpp::as<std::string>(vec[i]);

    if (tmp.compare("") == 0) {
      res[i] = NA_STRING;
      continue;
    }

    pugi::xml_document doc;
    pugi::xml_parse_result result = doc.load_string(tmp.c_str(), pugi::parse_default | pugi::parse_ws_pcdata | pugi::parse_escapes);

    if (!result) {
      Rcpp::stop(type.c_str(), " xml import unsuccessful");
    }

    for (auto is : doc.children(type.c_str()))
    {
      // text to export
      std::string text = "";

      // has only t node
      for (auto t : is.children("t")) {
        text = t.text().get();
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
          text += t.text().get();
        }
      }

      // push everything back
      res[i] = Rcpp::String(text);
    }

  }

  return res;
}

// [[Rcpp::export]]
SEXP is_to_txt(Rcpp::CharacterVector is_vec) {
  return xml_to_txt(is_vec, "is");
}

// [[Rcpp::export]]
SEXP si_to_txt(Rcpp::CharacterVector si_vec) {
  return xml_to_txt(si_vec, "si");
}

std::string txt_to_xml(
    std::string text,
    bool no_escapes,
    bool raw,
    bool skip_control,
    std::string type
) {

  pugi::xml_document doc;

  unsigned int pugi_format_flags = pugi::format_indent;
  if (no_escapes) pugi_format_flags |= pugi::format_no_escapes;
  if (raw)  pugi_format_flags |= pugi::format_raw;
  if (skip_control) pugi_format_flags |= pugi::format_skip_control_chars;

  pugi::xml_node is_node = doc.append_child(type.c_str());

  // txt input beginning with "<r" is assumed to be a fmt_txt string
  if (text.rfind("<r>", 0) == 0 || text.rfind("<r/>", 0) == 0) {

    pugi::xml_document txt_node;
    pugi::xml_parse_result result = txt_node.load_string(text.c_str(), pugi::parse_default | pugi::parse_ws_pcdata | pugi::parse_escapes);
    if (!result) Rcpp::stop("Could not parse xml in txt_to_xml()");

    for (auto is_n : txt_node.children())
          is_node.append_copy(is_n);

  } else {
    // text to export
    pugi::xml_node t_node = is_node.append_child("t");

    if ((text.size() > 0) && (std::isspace(text.at(0)) || std::isspace(text.at(text.size()-1)))) {
      t_node.append_attribute("xml:space").set_value("preserve");
    }

    t_node.append_child(pugi::node_pcdata).set_value(text.c_str());
  }

  std::ostringstream oss;
  doc.print(oss, " ", pugi_format_flags);
  std::string xml_return = oss.str();

  return xml_return;
}

// [[Rcpp::export]]
std::string txt_to_is(std::string text,
                      bool no_escapes = false, bool raw = true, bool skip_control = true) {
  return txt_to_xml(text, no_escapes, raw, skip_control, "is");
}

// [[Rcpp::export]]
std::string txt_to_si(std::string text,
                      bool no_escapes = false, bool raw = true, bool skip_control = true) {
  return txt_to_xml(text, no_escapes, raw, skip_control, "si");
}
