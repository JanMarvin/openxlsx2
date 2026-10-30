#include <unordered_map>

#include "openxlsx2.h"

// helper: extract text from parsed XML
// Handles both simple text (<t>) and rich text runs (<r><t>)
static std::string extract_text(pugi::xml_document& doc, const char* type) {
  std::string text;

  for (auto is : doc.children(type)) {

    // has only t node
    if (auto t = is.child("t")) {
      text += t.text().get();
    }

    // has r node with t node
    // phoneticPr (Phonetic Properties)
    // r (Rich Text Run)
    // rPr (Run Properties)
    // rPh (Phonetic Run)
    // t (Text)
    // linebreaks and spaces are handled in the nodes
    for (auto r : is.children("r")) {
      for (auto t : r.children("t")) {
        text += t.text().get();
      }
    }
  }

  return text;
}

SEXP xml_to_txt(Rcpp::CharacterVector vec, const std::string& type) {
  const R_xlen_t n = vec.length();
  Rcpp::CharacterVector res(n);

  // CHARSXPs are cached by R, identical strings share one pointer. The values
  // are kept alive by res.
  std::unordered_map<SEXP, SEXP> cache;
  cache.reserve(static_cast<std::size_t>(n));

  const unsigned parse_flags =
    pugi::parse_default |
    pugi::parse_ws_pcdata |
    pugi::parse_escapes;

  const char* type_c = type.c_str();

  // strings translated from a non UTF-8 encoding are R_alloc'ed, release them per element
  const void* vmax = vmaxget();
  for (R_xlen_t i = 0; i < n; ++i) {
    vmaxset(vmax);
    SEXP key = STRING_ELT(vec, i);

    std::unordered_map<SEXP, SEXP>::const_iterator it = cache.find(key);
    if (it != cache.end()) {
      SET_STRING_ELT(res, i, it->second);
      continue;
    }

    const char* xml = Rf_translateCharUTF8(key);

    if (xml[0] == '\0') {
      cache.emplace(key, R_BlankString);
      SET_STRING_ELT(res, i, R_BlankString);
      continue;
    }

    pugi::xml_document doc;
    pugi::xml_parse_result result = doc.load_string(xml, parse_flags);

    if (!result) {
      Rcpp::stop(type + " xml import unsuccessful");
    }

    std::string text = extract_text(doc, type_c);

    SEXP txt = utf8_charsxp(text);
    SET_STRING_ELT(res, i, txt);
    cache.emplace(key, txt);
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
    const std::string& text,
    bool no_escapes,
    bool raw,
    bool skip_control,
    const std::string& type
) {
  pugi::xml_document doc;

  uint32_t pugi_format_flags = pugi::format_indent;
  if (no_escapes) pugi_format_flags |= pugi::format_no_escapes;
  if (raw)  pugi_format_flags |= pugi::format_raw;
  if (skip_control) pugi_format_flags |= pugi::format_skip_control_chars;

  pugi::xml_node is_node = doc.append_child(type.c_str());

  // txt input beginning with "<r" is assumed to be a fmt_txt string
  if (
      text.size() >= 3 &&
        (text.compare(0, 3, "<r>") == 0 ||
        text.compare(0, 4, "<r/>") == 0)
  ) {
    pugi::xml_document txt_node;
    pugi::xml_parse_result result =
      txt_node.load_string(
        text.c_str(),
        pugi::parse_default |
          pugi::parse_ws_pcdata |
          pugi::parse_escapes
      );

    if (!result) {
      Rcpp::stop("Could not parse xml in txt_to_xml()");
    }

    for (auto is_n : txt_node.children()) {
      is_node.append_copy(is_n);
    }

  } else {
    // text to export
    pugi::xml_node t_node = is_node.append_child("t");

    if (!text.empty() &&
        (std::isspace(static_cast<unsigned char>(text.front())) ||
        std::isspace(static_cast<unsigned char>(text.back())))) {
      t_node.append_attribute("xml:space").set_value("preserve");
    }

    t_node.append_child(pugi::node_pcdata).set_value(text.c_str());
  }

  xml_string_writer writer;
  doc.print(writer, " ", pugi_format_flags);
  return writer.result;
}

// [[Rcpp::export]]
utf8_string txt_to_is(
    utf8_string text,
    bool no_escapes = false,
    bool raw = true,
    bool skip_control = true
) {
  return utf8_string(txt_to_xml(text, no_escapes, raw, skip_control, "is"));
}

// [[Rcpp::export]]
utf8_string txt_to_si(
    utf8_string text,
    bool no_escapes = false,
    bool raw = true,
    bool skip_control = true
) {
  return utf8_string(txt_to_xml(text, no_escapes, raw, skip_control, "si"));
}
