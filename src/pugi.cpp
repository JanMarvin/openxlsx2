#include "openxlsx2.h"

// [[Rcpp::export]]
SEXP readXML(std::string path, bool isfile, bool escapes, bool declaration, bool whitespace, bool empty_tags, bool skip_control, bool pointer) {
  xmldoc* doc = new xmldoc;
  pugi::xml_parse_result result;

  // pugi::parse_default without escapes flag
  uint32_t pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_eol;
  if (escapes) pugi_parse_flags |= pugi::parse_escapes;
  if (declaration) pugi_parse_flags |= pugi::parse_declaration;
  if (whitespace) pugi_parse_flags |= pugi::parse_ws_pcdata_single;
  if (!whitespace) pugi_parse_flags |= pugi::parse_trim_pcdata;

  if (isfile) {
    result = doc->load_file(path.c_str(), pugi_parse_flags, pugi::encoding_utf8);
  } else {
    result = doc->load_string(path.c_str(), pugi_parse_flags);
  }

  if (!result) {
    Rcpp::stop("xml import unsuccessful");
  }

  uint32_t pugi_format_flags = pugi::format_raw;
  if (!escapes) pugi_format_flags |= pugi::format_no_escapes;
  if (empty_tags) pugi_format_flags |= pugi::format_no_empty_element_tags;
  if (skip_control) pugi_format_flags |= pugi::format_skip_control_chars;

  if (pointer) {
    XPtrXML ptr(doc, true);
    ptr.attr("class") = Rcpp::CharacterVector::create("pugi_xml");
    ptr.attr("escapes") = escapes;
    ptr.attr("empty_tags") = empty_tags;
    ptr.attr("skip_control") = skip_control;
    return ptr;
  }

  std::ostringstream oss;
  doc->print(oss, " ", pugi_format_flags);
  return Rcpp::wrap(Rcpp::String(oss.str()));
}

inline uint32_t pugi_format(XPtrXML doc) {
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
Rcpp::LogicalVector is_xml(std::string str) {
  pugi::xml_document doc;
  pugi::xml_parse_result result;

  result = doc.load_string(str.c_str());

  if (!result)
    return false;
  else
    return true;
}

// [[Rcpp::export]]
SEXP getXMLXPtrNamePath(XPtrXML doc, std::vector<std::string> path) {
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

  vec_string res;
  for (const auto& node : nodes) {
    for (auto ch : node.children()) {
      res.push_back(ch.name());
    }
  }

  return Rcpp::wrap(res);
}

// [[Rcpp::export]]
SEXP getXMLXPtrPath(XPtrXML doc, std::vector<std::string> path) {
  vec_string res;
  uint32_t pugi_format_flags = pugi_format(doc);

  // Validate tag names: no empty strings allowed
  for (const std::string& tag : path) {
    if (tag.empty()) {
      Rcpp::stop("Tag names cannot be empty strings");
    }
  }

  // Return whole document if no path is specified
  if (path.empty()) {
    std::ostringstream oss;
    doc->print(oss, " ", pugi_format_flags);
    res.push_back(Rcpp::String(oss.str()));
    return Rcpp::wrap(res);
  }

  // Step 1: start from top-level nodes that match path[0] or wildcard
  std::vector<pugi::xml_node> current_nodes;
  if (path[0] == "*") {
    for (pugi::xml_node node : doc->children()) {
      current_nodes.push_back(node);
    }
  } else {
    for (pugi::xml_node node : doc->children(path[0].c_str())) {
      current_nodes.push_back(node);
    }
  }

  // Step 2: traverse remaining path
  for (size_t i = 1; i < path.size(); ++i) {
    const std::string& tag = path[i];
    std::vector<pugi::xml_node> next_nodes;

    for (const pugi::xml_node& node : current_nodes) {
      if (tag == "*") {
        for (pugi::xml_node child : node.children()) {
          next_nodes.push_back(child);
        }
      } else {
        for (pugi::xml_node child : node.children(tag.c_str())) {
          next_nodes.push_back(child);
        }
      }
    }

    current_nodes = std::move(next_nodes);
  }

  // Step 3: serialize final result
  for (const pugi::xml_node& node : current_nodes) {
    std::ostringstream oss;
    node.print(oss, " ", pugi_format_flags);
    res.push_back(Rcpp::String(oss.str()));
  }

  return Rcpp::wrap(res);
}

// [[Rcpp::export]]
SEXP getXMLXPtrValPath(XPtrXML doc, std::vector<std::string> path) {
  std::vector<std::string> values;

  // Validate: no empty tag names
  for (const std::string& tag : path) {
    if (tag.empty()) {
      Rcpp::stop("Tag names cannot be empty strings");
    }
  }

  // If no path, return empty vector — or optionally, entire document text
  if (path.empty()) {
    return Rcpp::wrap(values);
  }

  // Step 1: start from top-level nodes matching path[0] or wildcard
  std::vector<pugi::xml_node> current_nodes;
  if (path[0] == "*") {
    for (pugi::xml_node node : doc->children()) {
      current_nodes.push_back(node);
    }
  } else {
    for (pugi::xml_node node : doc->children(path[0].c_str())) {
      current_nodes.push_back(node);
    }
  }

  // Step 2: traverse the path
  for (size_t i = 1; i < path.size(); ++i) {
    const std::string& tag = path[i];
    std::vector<pugi::xml_node> next_nodes;

    for (const pugi::xml_node& node : current_nodes) {
      if (tag == "*") {
        for (pugi::xml_node child : node.children()) {
          next_nodes.push_back(child);
        }
      } else {
        for (pugi::xml_node child : node.children(tag.c_str())) {
          next_nodes.push_back(child);
        }
      }
    }

    current_nodes = std::move(next_nodes);
  }

  // Step 3: extract .text().get() values
  for (const pugi::xml_node& node : current_nodes) {
    const char* text = node.text().get();
    values.push_back(std::string(text));
  }

  return Rcpp::wrap(values);
}

// [[Rcpp::export]]
SEXP getXMLXPtrAttrPath(XPtrXML doc, std::vector<std::string> path) {
  std::vector<pugi::xml_node> current_nodes;

  // Validate: no empty tag names
  for (const std::string& tag : path) {
    if (tag.empty()) {
      Rcpp::stop("Tag names cannot be empty strings");
    }
  }

  // Handle root case (like getXMLXPtr1attr with no children)
  if (path.empty()) {
    for (pugi::xml_node node : doc->children()) {
      current_nodes.push_back(node);
    }
  } else {
    // First level
    if (path[0] == "*") {
      for (pugi::xml_node node : doc->children()) {
        current_nodes.push_back(node);
      }
    } else {
      for (pugi::xml_node node : doc->children(path[0].c_str())) {
        current_nodes.push_back(node);
      }
    }

    // Traverse further
    for (size_t i = 1; i < path.size(); ++i) {
      const std::string& tag = path[i];
      std::vector<pugi::xml_node> next_nodes;

      for (const auto& node : current_nodes) {
        if (tag == "*") {
          for (pugi::xml_node child : node.children()) {
            next_nodes.push_back(child);
          }
        } else {
          for (pugi::xml_node child : node.children(tag.c_str())) {
            next_nodes.push_back(child);
          }
        }
      }

      current_nodes = std::move(next_nodes);
    }
  }

  // Prepare result list
  R_xlen_t n = std::max(static_cast<R_xlen_t>(1), static_cast<R_xlen_t>(current_nodes.size()));
  Rcpp::List out(n);
  R_xlen_t i = 0;

  for (const auto& node : current_nodes) {
    R_xlen_t attr_count = std::distance(node.attributes_begin(), node.attributes_end());

    Rcpp::CharacterVector attr_names(attr_count);
    Rcpp::CharacterVector attr_values(attr_count);

    R_xlen_t j = 0;
    for (auto attr = node.attributes_begin(); attr != node.attributes_end(); ++attr, ++j) {
      attr_names[j] = attr->name();
      attr_values[j] = attr->value();
    }

    attr_values.attr("names") = attr_names;
    out[i++] = attr_values;
  }

  // If no nodes matched, return empty list
  if (current_nodes.empty()) {
    return Rcpp::List();
  }

  return out;
}

// [[Rcpp::export]]
SEXP printXPtr(XPtrXML doc, std::string indent, bool raw, bool attr_indent) {
  // pugi::parse_default without escapes flag
  uint32_t pugi_format_flags = pugi_format(doc);
  if (!raw) {
    // disable raw if not set. it is default in pugi_format()
    pugi_format_flags &= ~pugi::format_raw;
    pugi_format_flags |= pugi::format_indent;
  }
  if (attr_indent) pugi_format_flags |= pugi::format_indent_attributes;

  std::ostringstream oss;
  doc->print(oss, indent.c_str(), pugi_format_flags);

  return Rcpp::wrap(Rcpp::String(oss.str()));
}

// [[Rcpp::export]]
XPtrXML write_xml_file(std::string xml_content, bool escapes) {
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
Rcpp::CharacterVector xml_attr_mod(std::string xml_content,
                                   Rcpp::CharacterVector xml_attributes,
                                   bool escapes = false,
                                   bool declaration = false,
                                   bool remove_empty_attr = true) {
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

  std::vector<std::string> new_attr_nam = xml_attributes.names();
  std::vector<std::string> new_attr_val = Rcpp::as<std::vector<std::string>>(xml_attributes);

  for (auto cld : doc.children()) {
    for (size_t i = 0; i < static_cast<size_t>(xml_attributes.length()); ++i) {
      // check if attribute_val is empty. if yes, remove the attribute.
      // otherwise add or update the attribute
      if (new_attr_val[i].empty()) {
        if (remove_empty_attr)
          cld.remove_attribute(new_attr_nam[i].c_str());
      } else {
        // update attribute if found else add attribute
        if (cld.attribute(new_attr_nam[i].c_str())) {
          cld.attribute(new_attr_nam[i].c_str()).set_value(new_attr_val[i].c_str());
        } else {
          cld.append_attribute(new_attr_nam[i].c_str()) = new_attr_val[i].c_str();
        }
      }
    }
  }

  std::ostringstream oss;
  doc.print(oss, " ", pugi_format_flags);

  return Rcpp::wrap(Rcpp::String(oss.str()));
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
Rcpp::CharacterVector xml_node_create(
    std::string xml_name,
    Rcpp::Nullable<Rcpp::CharacterVector> xml_children = R_NilValue,
    Rcpp::Nullable<Rcpp::CharacterVector> xml_attributes = R_NilValue,
    bool escapes = false, bool declaration = false) {
  pugi::xml_document doc;

  uint32_t pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_ws_pcdata | pugi::parse_eol;
  if (escapes) pugi_parse_flags |= pugi::parse_escapes;
  if (declaration) pugi_parse_flags |= pugi::parse_declaration;

  uint32_t pugi_format_flags = pugi::format_raw;
  if (!escapes) pugi_format_flags |= pugi::format_no_escapes;

  pugi::xml_node cld = doc.append_child(xml_name.c_str());

  // check if children are attached
  if (xml_children.isNotNull()) {
    Rcpp::CharacterVector xml_child(xml_children.get());

    for (auto i = 0; i < xml_child.size(); ++i) {
      std::string xml_cld = std::string(xml_child[i]);

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
    Rcpp::CharacterVector xml_attr(xml_attributes.get());

    std::vector<std::string> new_attr_nam = xml_attr.names();
    std::vector<std::string> new_attr_val = Rcpp::as<std::vector<std::string>>(xml_attr);

    for (size_t i = 0; i < static_cast<size_t>(xml_attr.length()); ++i) {
      if (!new_attr_val[i].empty())
        cld.append_attribute(new_attr_nam[i].c_str()) = new_attr_val[i].c_str();
    }
  }

  std::ostringstream oss;
  doc.print(oss, " ", pugi_format_flags);

  return Rcpp::wrap(Rcpp::String(oss.str()));
}

// [[Rcpp::export]]
SEXP xml_append_child_path(XPtrXML node, XPtrXML child, std::vector<std::string> path, bool pointer) {
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
    std::ostringstream oss;
    node->print(oss, " ", pugi_format_flags);
    return Rcpp::wrap(Rcpp::String(oss.str()));
  }
}

// [[Rcpp::export]]
SEXP xml_remove_child_path(XPtrXML node, std::string child, std::vector<std::string> path, int32_t which, bool pointer) {
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
    std::ostringstream oss;
    node->print(oss, " ", pugi_format_flags);
    return Rcpp::wrap(Rcpp::String(oss.str()));
  }
}
