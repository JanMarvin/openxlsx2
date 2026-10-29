#include "openxlsx2.h"

// [[Rcpp::export]]
SEXP readXML(Rcpp::CharacterVector path, bool isfile, bool escapes, bool declaration, bool whitespace, int comments, bool empty_tags, bool skip_control, bool pointer) {
  std::unique_ptr<xmldoc> doc(new xmldoc());
  pugi::xml_parse_result result;

  // pugi::parse_default without escapes flag
  uint32_t pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_eol;
  if (escapes) pugi_parse_flags |= pugi::parse_escapes;
  if (declaration) pugi_parse_flags |= pugi::parse_declaration;
  if (whitespace) pugi_parse_flags |= pugi::parse_ws_pcdata_single;
  if (!whitespace) pugi_parse_flags |= pugi::parse_trim_pcdata;
  if (comments) pugi_parse_flags |= pugi::parse_comments;

  // path is either a file name or xml content
  if (isfile) {
    result = doc->load_file(Rcpp::as<native_path>(path).c_str(), pugi_parse_flags, pugi::encoding_utf8);
  } else {
    result = doc->load_string(Rcpp::as<utf8>(path).c_str(), pugi_parse_flags);
  }

  if (!result) {
    Rcpp::stop("xml import unsuccessful");
  }

  if (comments == 2) {
    std::vector<utf8> comment_vec;
    std::vector<pugi::xml_node> stack;
    stack.push_back(*doc);

    while (!stack.empty()) {
      pugi::xml_node node = stack.back();
      stack.pop_back();

      if (node.type() == pugi::node_comment) {
        comment_vec.push_back(utf8(node.value()));
      }
      for (pugi::xml_node child = node.last_child(); child; child = child.previous_sibling()) {
        stack.push_back(child);
      }
    }
    return Rcpp::wrap(comment_vec);
  }

  uint32_t pugi_format_flags = pugi::format_raw;
  if (!escapes) pugi_format_flags |= pugi::format_no_escapes;
  if (empty_tags) pugi_format_flags |= pugi::format_no_empty_element_tags;
  if (skip_control) pugi_format_flags |= pugi::format_skip_control_chars;

  if (pointer) {
    XPtrXML ptr(doc.release(), true);
    ptr.attr("class") = Rcpp::CharacterVector::create("pugi_xml");
    ptr.attr("escapes") = escapes;
    ptr.attr("empty_tags") = empty_tags;
    ptr.attr("skip_control") = skip_control;
    return ptr;
  }

  xml_string_writer writer;
  doc->print(writer, " ", pugi_format_flags);
  return Rcpp::wrap(utf8(std::move(writer.result)));
}

// pugixml memory is invisible to R's garbage collector, so a document read
// with pointer = TRUE can stay alive long after its last use. Free it now.
// [[Rcpp::export]]
void xml_release(XPtrXML doc) {
  doc.release();
}

inline uint32_t pugi_format(XPtrXML doc) {

  check_xptr_validity(doc);

  bool escapes = Rcpp::as<bool>(doc.attr("escapes"));
  bool empty_tags = Rcpp::as<bool>(doc.attr("empty_tags"));
  bool skip_control = Rcpp::as<bool>(doc.attr("skip_control"));

  uint32_t pugi_format_flags = pugi::format_raw;
  if (!escapes) pugi_format_flags |= pugi::format_no_escapes;
  if (empty_tags) pugi_format_flags |= pugi::format_no_empty_element_tags;
  if (skip_control) pugi_format_flags |= pugi::format_skip_control_chars;

  return pugi_format_flags;
}

// [[Rcpp::export]]
Rcpp::LogicalVector is_xml(utf8 str) {
  pugi::xml_document doc;
  pugi::xml_parse_result result;

  result = doc.load_string(str.c_str());

  if (!result)
    return false;
  else
    return true;
}

// [[Rcpp::export]]
SEXP getXMLXPtrNamePath(XPtrXML doc, std::vector<utf8> path) {

  check_xptr_validity(doc);

  std::vector<pugi::xml_node> nodes = {*doc};  // Start from the root node
  std::vector<pugi::xml_node> next_nodes;

  for (const auto& tag : path) {
    next_nodes.clear();
    for (const auto& node : nodes) {
      if (tag == "*") {
        for (auto ch : node.children()) {
          next_nodes.push_back(ch);
        }
      } else {
        for (auto ch : node.children(tag.c_str())) {
          next_nodes.push_back(ch);
        }
      }
    }
    nodes = next_nodes;
  }

  std::vector<utf8> res;
  for (const auto& node : nodes) {
    for (auto ch : node.children()) {
      res.push_back(utf8(ch.name()));
    }
  }

  return Rcpp::wrap(res);
}

// helper for path-finding
std::vector<pugi::xml_node> find_nodes_by_path(pugi::xml_document* doc, const std::vector<utf8>& path) {
  std::vector<pugi::xml_node> current_nodes;
  if (path.empty()) return current_nodes;

  const std::string& first_tag = path[0];
  for (pugi::xml_node node : doc->children()) {
    if (first_tag == "*" || (node.type() == pugi::node_element && first_tag == node.name())) {
      current_nodes.push_back(node);
    }
  }

  for (size_t i = 1; i < path.size(); ++i) {
    if (current_nodes.empty()) break;

    const std::string& tag = path[i];
    if (tag.empty()) Rcpp::stop("Tag names cannot be empty strings");

    std::vector<pugi::xml_node> next_nodes;
    next_nodes.reserve(current_nodes.size());

    for (const auto& node : current_nodes) {
      for (pugi::xml_node child : node.children()) {
        if (tag == "*" || (child.type() == pugi::node_element && tag == child.name())) {
          next_nodes.push_back(child);
        }
      }
    }
    current_nodes = std::move(next_nodes);
  }

  return current_nodes;
}

// [[Rcpp::export]]
std::vector<utf8> getXMLXPtrPath(XPtrXML doc, std::vector<utf8> path) {
  check_xptr_validity(doc);
  uint32_t pugi_format_flags = pugi_format(doc);

  // If no path, return empty vector — or optionally, entire document text
  if (path.empty()) {
    xml_string_writer writer;
    doc->print(writer, " ", pugi_format_flags);
    return {utf8(std::move(writer.result))};
  }

  std::vector<pugi::xml_node> target_nodes = find_nodes_by_path(doc.get(), path);
  std::vector<utf8> res;
  res.reserve(target_nodes.size());

  for (const pugi::xml_node& node : target_nodes) {
    xml_string_writer writer;
    node.print(writer, " ", pugi_format_flags);
    res.push_back(utf8(std::move(writer.result)));
  }

  return res;
}

// [[Rcpp::export]]
std::vector<utf8> getXMLXPtrValPath(XPtrXML doc, std::vector<utf8> path) {
  check_xptr_validity(doc);

  if (path.empty()) return {};

  std::vector<pugi::xml_node> target_nodes = find_nodes_by_path(doc.get(), path);
  std::vector<utf8> res;
  res.reserve(target_nodes.size());

  for (const pugi::xml_node& node : target_nodes) {
    res.push_back(utf8(node.text().get()));
  }

  return res;
}

// [[Rcpp::export]]
SEXP getXMLXPtrAttrPath(XPtrXML doc, std::vector<utf8> path) {
  check_xptr_validity(doc);
  std::vector<pugi::xml_node> target_nodes;

  if (path.empty()) {
    for (pugi::xml_node node : doc->children()) {
      if (node.type() == pugi::node_element) {
        target_nodes.push_back(node);
      }
    }
  } else {
    target_nodes = find_nodes_by_path(doc.get(), path);
  }

  if (target_nodes.empty()) return Rcpp::List::create();

  Rcpp::List out(static_cast<R_xlen_t>(target_nodes.size()));
  for (R_xlen_t i = 0; i < out.size(); ++i) {
    const auto& node = target_nodes[static_cast<size_t>(i)];

    std::vector<utf8> attr_names, attr_vals;
    for (auto attr : node.attributes()) {
      attr_names.push_back(utf8(attr.name()));
      attr_vals.push_back(utf8(attr.value()));
    }

    Rcpp::CharacterVector vals = Rcpp::wrap(attr_vals);
    vals.attr("names") = Rcpp::wrap(attr_names);
    out[i] = vals;
  }

  return out;
}

// [[Rcpp::export]]
utf8 printXPtr(XPtrXML doc, utf8 indent, bool raw, bool attr_indent) {

  check_xptr_validity(doc);

  // pugi::parse_default without escapes flag
  uint32_t pugi_format_flags = pugi_format(doc);
  if (!raw) {
    // disable raw if not set. it is default in pugi_format()
    pugi_format_flags &= ~pugi::format_raw;
    pugi_format_flags |= pugi::format_indent;
  }
  if (attr_indent) pugi_format_flags |= pugi::format_indent_attributes;

  xml_string_writer writer;
  doc->print(writer, indent.c_str(), pugi_format_flags);
  return utf8(std::move(writer.result));
}

// [[Rcpp::export]]
XPtrXML write_xml_file(utf8 xml_content, bool escapes) {
  xmldoc* doc = new xmldoc;

  // pugi::parse_default without escapes flag
  uint32_t pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_ws_pcdata | pugi::parse_eol;
  if (escapes) pugi_parse_flags |= pugi::parse_escapes;

  // load and validate node
  if (xml_content != "") {
    pugi::xml_parse_result result;
    result = doc->load_string(xml_content.c_str(), pugi_parse_flags);
    if (!result) Rcpp::stop("Loading xml_content node failed: \n %s", xml_content);
  }

  // Needs to be added after the node has been loaded and validated
  pugi::xml_node decl = doc->prepend_child(pugi::node_declaration);
  decl.append_attribute("version") = "1.0";
  decl.append_attribute("encoding") = "UTF-8";
  decl.append_attribute("standalone") = "yes";

  XPtrXML ptr(doc, true);
  ptr.attr("class") = Rcpp::CharacterVector::create("pugi_xml");

  return ptr;
}

//' adds or updates attribute(s) in existing xml node
//'
//' @description Needs xml node and named character vector as input. Modifies
//' the arguments of each first child found in the xml node and adds or updates
//' the attribute vector.
//' @details If a named attribute in `xml_attributes` is "" remove the attribute
//' from the node.
//' If `xml_attributes` contains a named entry found in the xml node, it is
//' updated else it is added as attribute.
//'
//' @param xml_content some valid xml_node
//' @param xml_attributes R vector of named attributes
//' @param path optional path
//' @param escapes bool if escapes should be used
//' @param declaration bool if declaration should be imported
//' @param remove_empty_attr bool remove empty attributes or ignore them
//'
//' @examples
//'   # add single node
//'     xml_node <- "<a foo=\"bar\">openxlsx2</a><b />"
//'     xml_attr <- c(qux = "quux")
//'     # "<a foo=\"bar\" qux=\"quux\">openxlsx2</a><b qux=\"quux\"/>"
//'     xml_attr_mod(xml_node, xml_attr)
//'
//'   # update node and add node
//'     xml_node <- "<a foo=\"bar\">openxlsx2</a><b />"
//'     xml_attr <- c(foo = "baz", qux = "quux")
//'     # "<a foo=\"baz\" qux=\"quux\">openxlsx2</a><b foo=\"baz\" qux=\"quux\"/>"
//'     xml_attr_mod(xml_node, xml_attr)
//'
//'   # remove node and add node
//'     xml_node <- "<a foo=\"bar\">openxlsx2</a><b />"
//'     xml_attr <- c(foo = "", qux = "quux")
//'     # "<a qux=\"quux\">openxlsx2</a><b qux=\"quux\"/>"
//'     xml_attr_mod(xml_node, xml_attr)
//' @export
// [[Rcpp::export]]
utf8 xml_attr_mod(
  utf8 xml_content,
  Rcpp::CharacterVector xml_attributes,
  Rcpp::Nullable<Rcpp::CharacterVector> path = R_NilValue,
  bool escapes = false,
  bool declaration = false,
  bool remove_empty_attr = true
) {
  pugi::xml_document doc;
  pugi::xml_parse_result result;

  uint32_t pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_ws_pcdata | pugi::parse_eol;
  if (escapes) pugi_parse_flags |= pugi::parse_escapes;
  if (declaration) pugi_parse_flags |= pugi::parse_declaration;

  uint32_t pugi_format_flags = pugi::format_raw;
  if (!escapes) pugi_format_flags |= pugi::format_no_escapes;

  // load and validate node
  if (xml_content != "") {
    result = doc.load_string(xml_content.c_str(), pugi_parse_flags);
    if (!result) Rcpp::stop("Loading xml_content node failed: \n %s ", xml_content);
  }

  std::vector<pugi::xml_node> target_nodes;

  std::vector<utf8> p;
  if (path.isNotNull()) p = Rcpp::as<std::vector<utf8>>(path.get());

  if (!p.empty()) {
    target_nodes = find_nodes_by_path(&doc, p);
  } else {
    // Default behavior: targets immediate children of the document
    for (auto cld : doc.children()) {
      target_nodes.push_back(cld);
    }
  }

  if (!target_nodes.empty()) {
    std::vector<utf8> new_attr_nam = Rcpp::as<std::vector<utf8>>(Rf_getAttrib(xml_attributes, R_NamesSymbol));
    std::vector<utf8> new_attr_val = Rcpp::as<std::vector<utf8>>(xml_attributes);
    if (new_attr_nam.size() != new_attr_val.size()) Rcpp::stop("xml_attributes must be named");

    for (auto& node : target_nodes) {
      for (size_t i = 0; i < static_cast<size_t>(xml_attributes.length()); ++i) {
        if (new_attr_val[i].empty()) {
          if (remove_empty_attr)
            node.remove_attribute(new_attr_nam[i].c_str());
        } else {
          // Update attribute if found, else append
          pugi::xml_attribute attr = node.attribute(new_attr_nam[i].c_str());
          if (attr) {
            attr.set_value(new_attr_val[i].c_str());
          } else {
            node.append_attribute(new_attr_nam[i].c_str()) = new_attr_val[i].c_str();
          }
        }
      }
    }
  }

  xml_string_writer writer;
  doc.print(writer, " ", pugi_format_flags);
  return utf8(std::move(writer.result));
}

//' create xml_node from R objects
//' @description takes xml_name, xml_children and xml_attributes to create a new
//' xml_node.
//' @param xml_name the name of the new xml_node
//' @param xml_children character vector children attached to the xml_node
//' @param xml_attributes named character vector of attributes for the xml_node
//' @param escapes bool if escapes should be used
//' @param declaration bool if declaration should be imported
//' @details if xml_children or xml_attributes should be empty, use NULL
//'
//' @examples
//' xml_name <- "a"
//' # "<a/>"
//' xml_node_create(xml_name)
//'
//' xml_child <- "openxlsx"
//' # "<a>openxlsx</a>"
//' xml_node_create(xml_name, xml_children = xml_child)
//'
//' xml_attr <- c(foo = "baz", qux = "quux")
//' # "<a foo=\"baz\" qux=\"quux\"/>"
//' xml_node_create(xml_name, xml_attributes = xml_attr)
//'
//' # "<a foo=\"baz\" qux=\"quux\">openxlsx</a>"
//' xml_node_create(xml_name, xml_children = xml_child, xml_attributes = xml_attr)
//' @export
// [[Rcpp::export]]
utf8 xml_node_create(
  utf8 xml_name,
  Rcpp::Nullable<Rcpp::CharacterVector> xml_children = R_NilValue,
  Rcpp::Nullable<Rcpp::CharacterVector> xml_attributes = R_NilValue,
  bool escapes = false, bool declaration = false
) {
  pugi::xml_document doc;

  uint32_t pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_ws_pcdata | pugi::parse_eol;
  if (escapes) pugi_parse_flags |= pugi::parse_escapes;
  if (declaration) pugi_parse_flags |= pugi::parse_declaration;

  uint32_t pugi_format_flags = pugi::format_raw;
  if (!escapes) pugi_format_flags |= pugi::format_no_escapes;

  pugi::xml_node cld = doc.append_child(xml_name.c_str());

  // check if children are attached
  if (xml_children.isNotNull()) {
    std::vector<utf8> xml_child = Rcpp::as<std::vector<utf8>>(xml_children.get());

    for (const utf8& xml_cld : xml_child) {

      pugi::xml_document is_node;
      pugi::xml_parse_result result = is_node.load_string(xml_cld.c_str(), pugi_parse_flags);

      // check if result is a valid xml_node, else append as is
      if (result) {
        for (auto is_n : is_node.children())
          cld.append_copy(is_n);
      } else {
        cld.append_child(pugi::node_pcdata).set_value(xml_cld.c_str());
      }
    }
  }

  // check if attributes are attached
  if (xml_attributes.isNotNull()) {
    std::vector<utf8> new_attr_nam = Rcpp::as<std::vector<utf8>>(Rf_getAttrib(xml_attributes.get(), R_NamesSymbol));
    std::vector<utf8> new_attr_val = Rcpp::as<std::vector<utf8>>(xml_attributes.get());
    if (new_attr_nam.size() != new_attr_val.size()) Rcpp::stop("xml_attributes must be named");

    for (size_t i = 0; i < new_attr_val.size(); ++i) {
      if (!new_attr_val[i].empty())
        cld.append_attribute(new_attr_nam[i].c_str()) = new_attr_val[i].c_str();
    }
  }

  xml_string_writer writer;
  doc.print(writer, " ", pugi_format_flags);
  return utf8(std::move(writer.result));
}

// [[Rcpp::export]]
SEXP xml_append_child_path(XPtrXML node, XPtrXML child, std::vector<utf8> path, bool pointer) {

  check_xptr_validity(node);
  check_xptr_validity(child);

  uint32_t pugi_format_flags = pugi_format(node);

  // Start from the first child (consistent with your previous design)
  pugi::xml_node current = node->first_child();

  // Traverse path — allow first path element to match current node name
  for (size_t i = 0; i < path.size(); ++i) {
    const std::string& tag = path[i];

    if (i == 0 && std::string(current.name()) == tag) {
      // First path element matches current node — don't descend
      continue;
    }

    current = current.child(tag.c_str());
    if (!current) {
      Rcpp::stop("Invalid path: node <%s> not found", tag);
    }
  }

  // Append children
  for (pugi::xml_node cld : child->children()) {
    current.append_copy(cld);
  }

  // Return
  if (pointer) {
    return node;
  } else {
    xml_string_writer writer;
    node->print(writer, " ", pugi_format_flags);
    return Rcpp::wrap(utf8(std::move(writer.result)));
  }
}

// [[Rcpp::export]]
SEXP xml_remove_child_path(XPtrXML node, utf8 child, std::vector<utf8> path, int32_t which, bool pointer) {

  check_xptr_validity(node);

  uint32_t pugi_format_flags = pugi_format(node);

  pugi::xml_node current = node->first_child();

  // Allow first path element to match current node
  for (size_t i = 0; i < path.size(); ++i) {
    const std::string& tag = path[i];

    if (i == 0 && std::string(current.name()) == tag) {
      continue;
    }

    current = current.child(tag.c_str());
    if (!current) {
      Rcpp::stop("Invalid path: node <%s> not found", tag);
    }
  }

  // Iterate through children to remove
  int32_t ctr = 0;
  for (pugi::xml_node cld = current.child(child.c_str()); cld;) {
    pugi::xml_node next = cld.next_sibling(child.c_str());

    if (which < 0 || ctr == which) {
      current.remove_child(cld);
      if (which >= 0) break;  // only remove one if specified
    }

    cld = next;
    ++ctr;
  }

  if (pointer) {
    return node;
  } else {
    xml_string_writer writer;
    node->print(writer, " ", pugi_format_flags);
    return Rcpp::wrap(utf8(std::move(writer.result)));
  }
}
