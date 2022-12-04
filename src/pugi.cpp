#include "openxlsx2.h"

// [[Rcpp::export]]
SEXP readXMLPtr(std::string path, bool isfile, bool escapes, bool declaration, bool whitespace, bool empty_tags, bool skip_control) {

  xmldoc *doc = new xmldoc;
  pugi::xml_parse_result result;

  // pugi::parse_default without escapes flag
  unsigned int pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_eol;
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
    Rcpp::stop("xml import unsuccessfull");
  }

  XPtrXML ptr(doc, true);
  ptr.attr("class") = Rcpp::CharacterVector::create("pugi_xml");
  ptr.attr("escapes") = escapes;
  ptr.attr("empty_tags") = empty_tags;
  ptr.attr("skip_control") = skip_control;
  return ptr;
}

// [[Rcpp::export]]
SEXP readXML(std::string path, bool isfile, bool escapes, bool declaration, bool whitespace, bool empty_tags, bool skip_control) {

  pugi::xml_document doc;
  pugi::xml_parse_result result;

  // pugi::parse_default without escapes flag
  unsigned int pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_eol;
  if (escapes) pugi_parse_flags |= pugi::parse_escapes;
  if (declaration) pugi_parse_flags |= pugi::parse_declaration;
  if (whitespace) pugi_parse_flags |= pugi::parse_ws_pcdata_single;
  if (!whitespace) pugi_parse_flags |= pugi::parse_trim_pcdata;

  unsigned int pugi_format_flags = pugi::format_raw;
  if (!escapes) pugi_format_flags |= pugi::format_no_escapes;
  if (empty_tags) pugi_format_flags |= pugi::format_no_empty_element_tags;
  if (skip_control) pugi_format_flags |= pugi::format_skip_control_chars;

  if (isfile) {
    result = doc.load_file(path.c_str(), pugi_parse_flags, pugi::encoding_utf8);
  } else {
    result = doc.load_string(path.c_str(), pugi_parse_flags);
  }

  if (!result) {
    Rcpp::stop("xml import unsuccessfull");
  }

  std::ostringstream oss;
  doc.print(oss, " ", pugi_format_flags);
  return  Rcpp::wrap(Rcpp::String(oss.str()));
}

inline unsigned int pugi_format(XPtrXML doc){
  bool escapes = Rcpp::as<bool>(doc.attr("escapes"));
  bool empty_tags = Rcpp::as<bool>(doc.attr("empty_tags"));
  bool skip_control = Rcpp::as<bool>(doc.attr("skip_control"));

  unsigned int pugi_format_flags = pugi::format_raw;
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
SEXP getXMLXPtrName1(XPtrXML doc) {

  vec_string res;

  for (auto lvl0 : doc->children())
  {
    res.push_back(lvl0.name());
  }

  return  Rcpp::wrap(res);
}

// [[Rcpp::export]]
SEXP getXMLXPtrName2(XPtrXML doc, std::string level1) {

  vec_string res;

  for (auto lvl0 : doc->children(level1.c_str()))
  {
    for (auto lvl1 : lvl0.children())
    {
      res.push_back(lvl1.name());
    }
  }

  return  Rcpp::wrap(res);
}

// [[Rcpp::export]]
SEXP getXMLXPtrName3(XPtrXML doc, std::string level1, std::string level2) {

  vec_string res;

  for (auto lvl0 : doc->children(level1.c_str()))
  {
    for (auto lvl1 : lvl0.children())
    {
      for (auto lvl2 : lvl1.children())
      {
        res.push_back(lvl2.name());
      }
    }
  }

  return  Rcpp::wrap(res);
}

// [[Rcpp::export]]
SEXP getXMLXPtr0(XPtrXML doc) {

  vec_string res;
  unsigned int  pugi_format_flags = pugi_format(doc);

  for (auto worksheet : doc->children())
  {
    std::ostringstream oss;
    worksheet.print(oss, " ", pugi_format_flags);
    res.push_back(Rcpp::String(oss.str()));
  }

  return  Rcpp::wrap(res);
}

// [[Rcpp::export]]
SEXP getXMLXPtr1(XPtrXML doc, std::string child) {

  vec_string res;
  unsigned int  pugi_format_flags = pugi_format(doc);

  for (auto cld : doc->children(child.c_str()))
  {
    std::ostringstream oss;
    cld.print(oss, " ", pugi_format_flags);
    res.push_back(Rcpp::String(oss.str()));
  }

  return  Rcpp::wrap(res);
}


// [[Rcpp::export]]
SEXP getXMLXPtr2(XPtrXML doc, std::string level1, std::string child) {

  vec_string res;
  unsigned int  pugi_format_flags = pugi_format(doc);

  for (auto lvl1 : doc->children(level1.c_str())) {
    for (auto cld : lvl1.children(child.c_str())) {
      std::ostringstream oss;
      cld.print(oss, " ", pugi_format_flags);
      res.push_back(Rcpp::String(oss.str()));
    }
  }

  return  Rcpp::wrap(res);
}

// [[Rcpp::export]]
SEXP getXMLXPtr3(XPtrXML doc, std::string level1, std::string level2, std::string child) {

  vec_string res;
  unsigned int  pugi_format_flags = pugi_format(doc);

  for (auto lvl1 : doc->children(level1.c_str())) {
    for (auto lvl2 : lvl1.children(level2.c_str())) {
      for (auto cld : lvl2.children(child.c_str())) {
        std::ostringstream oss;
        cld.print(oss, " ", pugi_format_flags);
        res.push_back(Rcpp::String(oss.str()));
      }
    }
  }

  return  Rcpp::wrap(res);
}

// level2 is wildcard. (for border only color nodes are imported.
// Do not know why :'( )
// [[Rcpp::export]]
SEXP unkgetXMLXPtr3(XPtrXML doc, std::string level1, std::string child) {

  vec_string res;
  unsigned int  pugi_format_flags = pugi_format(doc);

  for (auto lvl1 : doc->children(level1.c_str())) {
    for (auto lvl2 : lvl1.children()) {
      for (auto cld : lvl2.children(child.c_str())) {
        std::ostringstream oss;
        cld.print(oss, " ", pugi_format_flags);
        res.push_back(Rcpp::String(oss.str()));
      }
    }
  }

  return  Rcpp::wrap(res);
}

// [[Rcpp::export]]
SEXP getXMLPtr1con(XPtrXML doc) {

  vec_string res;
  unsigned int  pugi_format_flags = pugi_format(doc);

  for (auto node : doc->children())
  {
    for (auto cld : node.children()) {
      std::ostringstream oss;
      cld.print(oss, " ", pugi_format_flags);
      res.push_back(Rcpp::String(oss.str()));
    }
  }

  return  Rcpp::wrap(res);
}

// [[Rcpp::export]]
SEXP getXMLXPtr1val(XPtrXML doc, std::string child) {

  // returns a single vector, not a list of vectors!
  std::vector<std::string> x;

  for (auto worksheet : doc->children(child.c_str()))
  {
    x.push_back(Rcpp::String(worksheet.child_value()));
  }

  return  Rcpp::wrap(x);
}

// [[Rcpp::export]]
SEXP getXMLXPtr2val(XPtrXML doc, std::string level1, std::string child) {

  // returns a single vector, not a list of vectors!
  std::vector<std::string> x;

  for (auto worksheet : doc->children(level1.c_str())) {

    for (auto col : worksheet.children(child.c_str())) {
      x.push_back(Rcpp::String(col.child_value()));
    }
  }

  return  Rcpp::wrap(x);
}

// [[Rcpp::export]]
SEXP getXMLXPtr3val(XPtrXML doc, std::string level1, std::string level2, std::string child) {

  // returns a single vector, not a list of vectors!
  std::vector<std::string> x;

  for (auto worksheet : doc->child(level1.c_str()).children(level2.c_str()))
  {

    for (auto col : worksheet.children(child.c_str()))
      x.push_back(Rcpp::String(col.child_value()));
  }

  return  Rcpp::wrap(x);
}


// [[Rcpp::export]]
SEXP getXMLXPtr1attr(XPtrXML doc, std::string child) {

  auto children = doc->children(child.c_str());

  size_t n = std::distance(children.begin(),
                           children.end());

  // for a childless single line node the distance might be zero
  if (n == 0) n++;
  Rcpp::List z(n);

  auto itr = 0;
  for (auto child : children) {

    Rcpp::CharacterVector res;
    std::vector<std::string> nam;

    for (auto attrs : child.attributes())
    {
      nam.push_back(Rcpp::String(attrs.name()));
      res.push_back(Rcpp::String(attrs.value()));
    }

    // assign names
    res.attr("names") = nam;

    z[itr] = res;
    ++itr;
  }

  return Rcpp::wrap(z);
}

// [[Rcpp::export]]
Rcpp::List getXMLXPtr2attr(XPtrXML doc, std::string level1, std::string child) {

  auto worksheet = doc->child(level1.c_str()).children(child.c_str());
  auto n = std::distance(worksheet.begin() , worksheet.end());
  Rcpp::List z(n);

  auto itr = 0;
  for (auto ws : worksheet)
  {

    auto n = std::distance(ws.attributes_begin(), ws.attributes_end());

    Rcpp::CharacterVector res(n);
    Rcpp::CharacterVector nam(n);

    auto attr_itr = 0;
    for (auto attr : ws.attributes())
    {
      nam[attr_itr] = Rcpp::String(attr.name());
      res[attr_itr] = Rcpp::String(attr.value());
      ++attr_itr;
    }

    // assign names
    res.attr("names") = nam;

    z[itr] = res;
    ++itr;
  }

  return z;
}

// [[Rcpp::export]]
SEXP getXMLXPtr3attr(XPtrXML doc, std::string level1, std::string level2, std::string child) {

  auto worksheet = doc->child(level1.c_str()).child(level2.c_str()).children(child.c_str());
  auto n = std::distance(worksheet.begin(), worksheet.end());
  Rcpp::List z(n);

  auto itr = 0;
  for (auto ws : worksheet)
  {

    auto n = std::distance(ws.attributes_begin(), ws.attributes_end());

    Rcpp::CharacterVector res(n);
    Rcpp::CharacterVector nam(n);

    auto attr_itr = 0;
    for (auto attr : ws.attributes())
    {
      nam[attr_itr] = Rcpp::String(attr.name());
      res[attr_itr] = Rcpp::String(attr.value());
      ++attr_itr;
    }

    // assign names
    res.attr("names") = nam;

    z[itr] = res;
    ++itr;
  }

  return z;
}


// [[Rcpp::export]]
SEXP printXPtr(XPtrXML doc, std::string indent, bool raw, bool attr_indent) {

  // pugi::parse_default without escapes flag
  unsigned int pugi_format_flags = pugi_format(doc);
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

  xmldoc *doc = new xmldoc;

  // pugi::parse_default without escapes flag
  unsigned int pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_ws_pcdata | pugi::parse_eol;
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
Rcpp::CharacterVector xml_attr_mod(std::string xml_content, Rcpp::CharacterVector xml_attributes,
                                   bool escapes = false, bool declaration = false) {

  pugi::xml_document doc;
  pugi::xml_parse_result result;

  unsigned int pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_ws_pcdata | pugi::parse_eol;
  if (escapes) pugi_parse_flags |= pugi::parse_escapes;
  if (declaration) pugi_parse_flags |= pugi::parse_declaration;

  unsigned int pugi_format_flags = pugi::format_raw;
  if (!escapes) pugi_format_flags |= pugi::format_no_escapes;

  // load and validate node
  if (xml_content != "") {
    result = doc.load_string(xml_content.c_str(), pugi_parse_flags);
    if (!result) Rcpp::stop("Loading xml_content node failed: \n %s ", xml_content);
  }

  std::vector<std::string> new_attr_nam = xml_attributes.names();
  std::vector<std::string> new_attr_val = Rcpp::as<std::vector<std::string>>(xml_attributes);

  for (auto cld : doc.children()) {
    for (auto i = 0; i < xml_attributes.length(); ++i){

      // check if attribute_val is empty. if yes, remove the attribute.
      // otherwise add or update the attribute
      if (new_attr_val[i] == "") {
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
  pugi::xml_parse_result result;

  unsigned int pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_ws_pcdata | pugi::parse_eol;
  if (escapes) pugi_parse_flags |= pugi::parse_escapes;
  if (declaration) pugi_parse_flags |= pugi::parse_declaration;

  unsigned int pugi_format_flags = pugi::format_raw;
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

    for (auto i = 0; i < xml_attr.length(); ++i){
      cld.append_attribute(new_attr_nam[i].c_str()) = new_attr_val[i].c_str();
    }
  }

  std::ostringstream oss;
  doc.print(oss, " ", pugi_format_flags);

  return Rcpp::wrap(Rcpp::String(oss.str()));
}

// xml_append_child1
// @param node xml_node a child is appended to
// @param child the xml_node appended to node
// @param pointer bool if pointer should be returned
// @export
// [[Rcpp::export]]
SEXP xml_append_child1(XPtrXML node, XPtrXML child, bool pointer) {

  unsigned int pugi_format_flags = pugi_format(node);

  for (auto cld: child->children()) {
    node->first_child().append_copy(cld);
  }

  if (pointer) {
    return (node);
  } else {
    std::ostringstream oss;
    node->print(oss, " ", pugi_format_flags);
    return Rcpp::wrap(Rcpp::String(oss.str()));
  }
}

// xml_append_child2
// @param node xml_node a child is appended to
// @param child the xml_node appended to node
// @param level1 level the child will be added to
// @param pointer bool if pointer should be returned
// @export
// [[Rcpp::export]]
SEXP xml_append_child2(XPtrXML node, XPtrXML child, std::string level1, bool pointer) {

  unsigned int pugi_format_flags = pugi_format(node);

  for (auto cld: child->children()) {
    node->first_child().child(level1.c_str()).append_copy(cld);
  }

  if (pointer) {
    return (node);
  } else {
    std::ostringstream oss;
    node->print(oss, " ", pugi_format_flags);
    return Rcpp::wrap(Rcpp::String(oss.str()));
  }
}

// xml_append_child3
// @param node xml_node a child is appended to
// @param child the xml_node appended to node
// @param level1 level the child will be added to
// @param level2 level the child will be added to
// @param pointer bool if pointer should be returned
// @export
// [[Rcpp::export]]
SEXP xml_append_child3(XPtrXML node, XPtrXML child, std::string level1, std::string level2, bool pointer) {

  unsigned int pugi_format_flags = pugi_format(node);

  for (auto cld: child->children()) {
    node->first_child().child(level1.c_str()).child(level2.c_str()).append_copy(cld);
  }

  if (pointer) {
    return (node);
  } else {
    std::ostringstream oss;
    node->print(oss, " ", pugi_format_flags);
    return Rcpp::wrap(Rcpp::String(oss.str()));
  }
}

// xml_remove_child1
// @param node xml_node a child is removed from
// @param child the xml_node to remove
// @param pointer bool if pointer should be returned
// @param escapes bool if escapes should be used
// @export
// [[Rcpp::export]]
SEXP xml_remove_child1(XPtrXML node, std::string child, int which, bool pointer) {

  unsigned int pugi_format_flags = pugi_format(node);

  auto ctr = 0;
  for (pugi::xml_node cld = node->first_child().child(child.c_str()); cld; ) {
    auto next = cld.next_sibling();
    if (ctr == which || which < 0) cld.parent().remove_child(cld);
     cld = next;
    ++ctr;
  }

  if (pointer) {
    return (node);
  } else {
    std::ostringstream oss;
    node->print(oss, " ", pugi_format_flags);
    return Rcpp::wrap(Rcpp::String(oss.str()));
  }
}

// xml_remove_child2
// @param node xml_node a child is removed from
// @param child the xml_node to remove
// @param level1 level the child will be removed from
// @param pointer bool if pointer should be returned
// @param escapes bool if escapes should be used
// @export
// [[Rcpp::export]]
SEXP xml_remove_child2(XPtrXML node, std::string child, std::string level1, int which,  bool pointer) {

  unsigned int pugi_format_flags = pugi_format(node);

  auto ctr = 0;
  for (pugi::xml_node cld = node->first_child().child(level1.c_str()).child(child.c_str()); cld; ) {
    auto next = cld.next_sibling();
    if (ctr == which || which < 0)  cld.parent().remove_child(cld);
    cld = next;
    ++ctr;
  }

  if (pointer) {
    return (node);
  } else {
    std::ostringstream oss;
    node->print(oss, " ", pugi_format_flags);
    return Rcpp::wrap(Rcpp::String(oss.str()));
  }
}

// xml_remove_child3
// @param node xml_node a child is removed from
// @param child the xml_node to remove
// @param level1 level the child will be removed from
// @param level2 level the child will be removed from
// @param pointer bool if pointer should be returned
// @param escapes bool if escapes should be used
// @export
// [[Rcpp::export]]
SEXP xml_remove_child3(XPtrXML node, std::string child, std::string level1, std::string level2, int which, bool pointer) {

  unsigned int pugi_format_flags = pugi_format(node);

  auto ctr = 0;
  for (pugi::xml_node cld = node->first_child().child(level1.c_str()).child(level2.c_str()).child(child.c_str()); cld; ) {
    auto next = cld.next_sibling();
    if (ctr == which || which < 0) cld.parent().remove_child(cld);
    cld = next;
    ++ctr;
  }

  if (pointer) {
    return (node);
  } else {
    std::ostringstream oss;
    node->print(oss, " ", pugi_format_flags);
    return Rcpp::wrap(Rcpp::String(oss.str()));
  }
}
