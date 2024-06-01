#include "openxlsx2.h"

// helper function to check if row contains any of the expected types
bool has_it(Rcpp::DataFrame df_xf, std::set<std::string> xf_nams, size_t row) {

  bool has_it = false;

  // names as vector and set. because set is ordered, we have to change the
  // order of the data frame columns as well.
  std::vector<std::string> df_nms = df_xf.names();
  std::set<std::string> df_names(df_nms.begin(), df_nms.end());
  std::vector<std::string> xf_names(xf_nams.begin(), xf_nams.end());

  Rcpp::CharacterVector sel_chr;
  Rcpp::IntegerVector sel_int;
  Rcpp::DataFrame df_tmp;

  sel_chr = Rcpp::wrap(df_names);
  df_tmp = df_xf[sel_chr];

  // get the position of the xf_nams in the sorted df_xf data frame
  std::vector<R_xlen_t> idx;
  for (size_t i = 0; i < xf_names.size(); ++i) {
    std::string xf_name = xf_names[i];
    if (df_names.count(xf_name) > 0) {
      auto res = df_names.find(xf_name);
      R_xlen_t mtc = std::distance(df_names.begin(), res);
      idx.push_back(mtc);
    }
  }

  // get a subset of the original data frame
  sel_int = Rcpp::wrap(idx);
  df_tmp = df_tmp[sel_int];

  // check every column of the selected row
  for (auto ii = 0; ii < df_tmp.ncol(); ++ii) {
    std::string cv_tmp;
    cv_tmp = Rcpp::as<Rcpp::CharacterVector>(df_tmp[ii])[row];
    if (!cv_tmp.empty()) {
      has_it = true;
    }
  }

  return has_it;
}

// we modify CellStyleFormats. Not CellStyles. The latter are pre defined sets,
// and we create the set from scratch. Though we might think about adding pre-
// defined style sets as well.

// [[Rcpp::export]]
Rcpp::DataFrame read_xf(XPtrXML xml_doc_xf) {

  // https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformat?view=openxml-2.8.1

  // openxml 2.8.1
  std::set<std::string> nams {
    "numFmtId",
    "fontId",
    "fillId",
    "borderId",
    "xfId",
    "applyNumberFormat",
    "applyFont",
    "applyFill",
    "applyBorder",
    "applyAlignment",
    "applyProtection",
    "pivotButton",
    "quotePrefix",
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

  size_t nn = std::distance(xml_doc_xf->begin(), xml_doc_xf->end());
  size_t kk = nams.size();

  Rcpp::CharacterVector rvec(nn);

  // 1. create the list
  Rcpp::List df(kk);
  for (size_t i = 0; i < kk; ++i)
  {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
  }

  // 2. fill the list
  // <xf ...>
  auto itr = 0;
  for (auto xml_xf : xml_doc_xf->children("xf")) {
    for (auto attrs : xml_xf.attributes()) {

      std::string attr_name = attrs.name();
      std::string attr_value = attrs.value();
      auto find_res = nams.find(attr_name);

      // check if name is already known
      if (nams.count(attr_name) == 0) {
        Rcpp::warning("%s: not found in xf name table", attr_name);
      } else {
        R_xlen_t mtc = std::distance(nams.begin(), find_res);
        Rcpp::as<Rcpp::CharacterVector>(df[mtc])[itr] = attr_value;
      }
    }

    // only handle known names
    // <alignment ...>
    // <extLst ...> // FIXME should be imported as single node. most likely broken
    // <protection ...>
    for (auto cld : xml_xf.children()) {

      std::string cld_name = cld.name();

      // check known names
      if (cld_name ==  "alignment" || cld_name == "extLst" || cld_name == "protection") {

        for (auto attrs : cld.attributes()) {
          std::string attr_name = attrs.name();
          std::string attr_value = attrs.value();
          auto find_res = nams.find(attr_name);

          // check if name is already known
          if (nams.count(attr_name) == 0) {
            Rcpp::warning("%s: not found in xf name table", attr_name);
          } else {
            R_xlen_t mtc = std::distance(nams.begin(), find_res);
            Rcpp::as<Rcpp::CharacterVector>(df[mtc])[itr] = attr_value;
          }
        }
      } else {
        Rcpp::warning("%s: not valid for xf in openxml 2.8.1", cld_name);
      }

    } // end aligment, extLst, protection

    // rownames as character vectors matching to <c s= ...>
    rvec[itr] = std::to_string(itr);

    ++itr;
  }

  // 3. Create a data.frame
  df.attr("row.names") = rvec;
  df.attr("names") = nams;
  df.attr("class") = "data.frame";

  return df;
}


// [[Rcpp::export]]
Rcpp::CharacterVector write_xf(Rcpp::DataFrame df_xf) {

  auto n = df_xf.nrow();
  Rcpp::CharacterVector z(n);

  unsigned int pugi_format_flags = pugi::format_raw | pugi::format_no_escapes;

  std::vector<std::string> attrnams = df_xf.names();

  std::set<std::string> xf_nams = {
    "numFmtId",
    "fontId",
    "fillId",
    "borderId",
    "xfId",
    "applyFont",
    "applyFill",
    "applyBorder",
    "applyAlignment",
    "applyNumberFormat",
    "applyProtection",
    "pivotButton",
    "quotePrefix"
  };

  std::set<std::string> xf_nams_alignment = {
    "horizontal",
    "indent",
    "justifyLastLine",
    "readingOrder",
    "relativeIndent",
    "shrinkToFit",
    "textRotation",
    "vertical",
    "wrapText"
  };

  std::set<std::string> xf_nams_extLst = {
    "extLst"
  };

  std::set<std::string> xf_nams_protection = {
    "hidden",
    "locked"
  };

  for (auto i = 0; i < n; ++i) {
    pugi::xml_document doc;

    pugi::xml_node xf = doc.append_child("xf");

    // check if alignment node is required
    bool has_alignment = false;
    has_alignment = has_it(df_xf, xf_nams_alignment, i);

    pugi::xml_node xf_alignment;
    if (has_alignment) {
      xf_alignment = xf.append_child("alignment");
    }

    // check if extLst node is required
    bool has_extLst = false;
    has_extLst = has_it(df_xf, xf_nams_extLst, i);

    pugi::xml_node xf_extLst;
    if (has_extLst) {
      xf_extLst = xf.append_child("extLst");
    }

    // check if protection node is required
    bool has_protection = false;
    has_protection = has_it(df_xf, xf_nams_protection, i);

    pugi::xml_node xf_protection;
    if (has_protection) {
      xf_protection = xf.append_child("protection");
    }

    for (auto j = 0; j < df_xf.ncol(); ++j) {

      std::string attrnam = attrnams[j];

      // not all missing in match: ergo they are
      bool is_xf = xf_nams.count(attrnam) > 0;
      bool is_alignment = xf_nams_alignment.count(attrnam) > 0;
      bool is_extLst = xf_nams_extLst.count(attrnam) > 0;
      bool is_protection = xf_nams_protection.count(attrnam) > 0;

      // <xf ...>
      if (is_xf) {
        Rcpp::CharacterVector cv_s = "";
        cv_s = Rcpp::as<Rcpp::CharacterVector>(df_xf[j])[i];

        // only write attributes where cv_s has a value
        if (cv_s[0] != "") {
          // Rf_PrintValue(cv_s);
          const std::string val_strl = Rcpp::as<std::string>(cv_s);
          xf.append_attribute(attrnams[j].c_str()) = val_strl.c_str();
        }
      }

      if (has_alignment && is_alignment) {
        Rcpp::CharacterVector cv_s = "";
        cv_s = Rcpp::as<Rcpp::CharacterVector>(df_xf[j])[i];

        if (cv_s[0] != "") {
          const std::string val_strl = Rcpp::as<std::string>(cv_s);
          xf_alignment.append_attribute(attrnams[j].c_str()) = val_strl.c_str();
        }
      }

      // FIXME should be written as single node. most likely broken
      if (has_extLst && is_extLst) {
        Rcpp::CharacterVector cv_s = "";
        cv_s = Rcpp::as<Rcpp::CharacterVector>(df_xf[j])[i];

        if (cv_s[0] != "") {
          const std::string val_strl = Rcpp::as<std::string>(cv_s);
          xf_extLst.append_attribute(attrnams[j].c_str()) = val_strl.c_str();
        }
      }

      if (has_protection && is_protection) {
        Rcpp::CharacterVector cv_s = "";
        cv_s = Rcpp::as<Rcpp::CharacterVector>(df_xf[j])[i];

        if (cv_s[0] != "") {
          const std::string val_strl = Rcpp::as<std::string>(cv_s);
          xf_protection.append_attribute(attrnams[j].c_str()) = val_strl.c_str();
        }
      }

    }

    std::ostringstream oss;
    doc.print(oss, " ", pugi_format_flags);

    z[i] = oss.str();
  }

  return z;
}


// [[Rcpp::export]]
Rcpp::DataFrame read_font(XPtrXML xml_doc_font) {

  // https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.font?view=openxml-2.8.1

  // openxml 2.8.1
  std::set<std::string> nams {
    "b",
    "charset",
    "color",
    "condense",
    "extend",
    "family",
    "i",
    "name",
    "outline",
    // TODO might contain child <localName ...>
    "scheme",
    "shadow",
    "strike",
    "sz",
    "u",
    "vertAlign"
  };

  size_t nn = std::distance(xml_doc_font->begin(), xml_doc_font->end());
  size_t kk = nams.size();
  unsigned int pugi_format_flags = pugi::format_raw | pugi::format_no_escapes;

  Rcpp::CharacterVector rvec(nn);

  // 1. create the list
  Rcpp::List df(kk);
  for (size_t i = 0; i < kk; ++i)
  {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
  }

  // 2. fill the list
  // <xf ...>
  auto itr = 0;
  for (auto xml_font : xml_doc_font->children("font")) {

    for (auto cld : xml_font.children()) {

      std::string name = cld.name();
      std::string value = cld.value();
      auto find_res = nams.find(name);

      // check if name is already known
      if (nams.count(name) == 0) {
        Rcpp::warning("%s: not found in font name table", name);
      } else {
        // TODO why is this needed here?
        std::ostringstream oss;
        cld.print(oss, " ", pugi_format_flags);

        R_xlen_t mtc = std::distance(nams.begin(), find_res);
        Rcpp::as<Rcpp::CharacterVector>(df[mtc])[itr] = oss.str();
      }

    } // end aligment, extLst, protection

    // rownames as character vectors matching to <c s= ...>
    rvec[itr] = std::to_string(itr);

    ++itr;
  }

  // 3. Create a data.frame
  df.attr("row.names") = rvec;
  df.attr("names") = nams;
  df.attr("class") = "data.frame";

  return df;
}


// [[Rcpp::export]]
Rcpp::CharacterVector write_font(Rcpp::DataFrame df_font) {

  auto n = df_font.nrow();
  Rcpp::CharacterVector z(n);
  unsigned int pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_ws_pcdata | pugi::parse_eol;
  unsigned int pugi_format_flags = pugi::format_raw | pugi::format_no_escapes;

  for (auto i = 0; i < n; ++i) {
    pugi::xml_document doc;

    pugi::xml_node font = doc.append_child("font");


    for (auto j = 0; j < df_font.ncol(); ++j) {

      Rcpp::CharacterVector cv_s = "";
      cv_s = Rcpp::as<Rcpp::CharacterVector>(df_font[j])[i];

      if (cv_s[0] != "") {

        std::string font_i = Rcpp::as<std::string>(cv_s[0]);

        pugi::xml_document font_node;
        pugi::xml_parse_result result = font_node.load_string(font_i.c_str(), pugi_parse_flags);
        if (!result) Rcpp::stop("loading font node fail: %s", cv_s);

        font.append_copy(font_node.first_child());

      }

    }

    std::ostringstream oss;
    doc.print(oss, " ", pugi_format_flags);

    z[i] = oss.str();
  }

  return z;
}


// [[Rcpp::export]]
Rcpp::DataFrame read_numfmt(XPtrXML xml_doc_numfmt) {

  // https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-2.8.1

  // openxml 2.8.1
  std::set<std::string> nams {
    "formatCode",
    "numFmtId"
  };

  size_t nn = std::distance(xml_doc_numfmt->begin(), xml_doc_numfmt->end());
  size_t kk = nams.size();

  Rcpp::CharacterVector rvec(nn);

  // 1. create the list
  Rcpp::List df(kk);
  for (size_t i = 0; i < kk; ++i)
  {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
  }

  // 2. fill the list
  // <numFmt ...>
  auto itr = 0;
  for (auto xml_numfmt : xml_doc_numfmt->children("numFmt")) {
    for (auto attrs : xml_numfmt.attributes()) {

      std::string attr_name = attrs.name();
      std::string attr_value = attrs.value();
      auto find_res = nams.find(attr_name);

      // check if name is already known
      if (nams.count(attr_name) == 0) {
        Rcpp::warning("%s: not found in numfmt name table", attr_name);
      } else {
        R_xlen_t mtc = std::distance(nams.begin(), find_res);
        Rcpp::as<Rcpp::CharacterVector>(df[mtc])[itr] = attr_value;
      }
    }

    // rownames as character vectors matching to <c s= ...>
    rvec[itr] = std::to_string(itr);

    ++itr;
  }

  // 3. Create a data.frame
  df.attr("row.names") = rvec;
  df.attr("names") = nams;
  df.attr("class") = "data.frame";

  return df;
}


// [[Rcpp::export]]
Rcpp::CharacterVector write_numfmt(Rcpp::DataFrame df_numfmt) {

  auto n = df_numfmt.nrow();
  Rcpp::CharacterVector z(n);
  unsigned int pugi_format_flags = pugi::format_raw | pugi::format_no_escapes;

  for (auto i = 0; i < n; ++i) {
    pugi::xml_document doc;
    Rcpp::CharacterVector attrnams = df_numfmt.names();

    pugi::xml_node numFmt = doc.append_child("numFmt");

    for (auto j = 0; j < df_numfmt.ncol(); ++j) {
      Rcpp::CharacterVector cv_s = "";
      cv_s = Rcpp::as<Rcpp::CharacterVector>(df_numfmt[j])[i];

      // only write attributes where cv_s has a value
      if (cv_s[0] != "") {
        // Rf_PrintValue(cv_s);
        const std::string val_strl = Rcpp::as<std::string>(cv_s);
        numFmt.append_attribute(attrnams[j]) = val_strl.c_str();
      }
    }

    std::ostringstream oss;
    doc.print(oss, " ", pugi_format_flags);

    z[i] = oss.str();
  }

  return z;
}

// [[Rcpp::export]]
Rcpp::DataFrame read_border(XPtrXML xml_doc_border) {

  // https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.border?view=openxml-2.8.1

  // openxml 2.8.1
  std::set<std::string> nam_attrs {"diagonalDown", "diagonalUp", "outline"};

  std::set<std::string> nam_chlds {"start", "end",
                                   "left", "right", "top", "bottom", "diagonal",
                                   "vertical",  "horizontal"};

  auto total_length = nam_attrs.size() + nam_chlds.size();
  std::vector<std::string> all_names(total_length);

  std::copy(nam_attrs.begin(), nam_attrs.end(), all_names.begin());
  std::copy(nam_chlds.begin(), nam_chlds.end(), all_names.begin() + nam_attrs.size());

  std::set<std::string> nams(std::make_move_iterator(all_names.begin()),
                             std::make_move_iterator(all_names.end()));


  size_t nn = std::distance(xml_doc_border->begin(), xml_doc_border->end());
  size_t kk = nams.size();
  unsigned int pugi_format_flags = pugi::format_raw | pugi::format_no_escapes;

  Rcpp::CharacterVector rvec(nn);

  // 1. create the list
  Rcpp::List df(kk);
  for (size_t i = 0; i < kk; ++i)
  {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
  }

  // 2. fill the list
  // <numFmt ...>
  auto itr = 0;
  for (auto xml_border : xml_doc_border->children("border")) {
    for (auto attrs : xml_border.attributes()) {

      std::string attr_name = attrs.name();
      std::string attr_value = attrs.value();
      auto find_res = nams.find(attr_name);

      // check if name is already known
      if (nams.count(attr_name) == 0) {
        Rcpp::warning("%s: not found in border name table", attr_name);
      } else {
        R_xlen_t mtc = std::distance(nams.begin(), find_res);
        Rcpp::as<Rcpp::CharacterVector>(df[mtc])[itr] = attr_value;
      }
    }

    for (auto cld : xml_border.children()) {

      std::string cld_name = cld.name();
      auto find_res = nams.find(cld_name);

      // check if name is already known
      if (nams.count(cld_name) == 0) {
        Rcpp::warning("%s: not found in border name table", cld_name);
      } else {
        std::ostringstream oss;
        cld.print(oss, " ", pugi_format_flags);
        std::string cld_value = oss.str();

        R_xlen_t mtc = std::distance(nams.begin(), find_res);
        Rcpp::as<Rcpp::CharacterVector>(df[mtc])[itr] = cld_value;
      }
    }

    // rownames as character vectors matching to <c s= ...>
    rvec[itr] = std::to_string(itr);
    ++itr;

  }

  // 3. Create a data.frame
  df.attr("row.names") = rvec;
  df.attr("names") = nams;
  df.attr("class") = "data.frame";

  return df;
}


// [[Rcpp::export]]
Rcpp::CharacterVector write_border(Rcpp::DataFrame df_border) {

  auto n = df_border.nrow();
  Rcpp::CharacterVector z(n);
  unsigned int pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_ws_pcdata | pugi::parse_eol;
  unsigned int pugi_format_flags = pugi::format_raw | pugi::format_no_escapes;


  // openxml 2.8.1
  std::vector<std::string> attrnams = df_border.names();
  std::set<std::string> nam_attrs = {"diagonalDown", "diagonalUp", "outline"};
  std::set<std::string> nam_chlds = {"bottom", "diagonal", "end", "horizontal",
                                     "left", "right", "start", "top",
                                     "vertical"};

  for (auto i = 0; i < n; ++i) {
    pugi::xml_document doc;
    pugi::xml_node border = doc.append_child("border");

    for (auto j = 0; j < df_border.ncol(); ++j) {

      std::string attr_j = attrnams[j];

      // mimic which
      auto res1 = nam_attrs.find(attr_j);
      R_xlen_t mtc1 = std::distance(nam_attrs.begin(), res1);

      std::vector<int> idx1(mtc1 + 1);
      std::iota(idx1.begin(), idx1.end(), 0);

      auto res2 = nam_chlds.find(attr_j);
      R_xlen_t mtc2 = std::distance(nam_chlds.begin(), res2);

      std::vector<int> idx2(mtc2 + 1);
      std::iota(idx2.begin(), idx2.end(), 0);

      // check if name is already known
      if (nam_attrs.count(attr_j) != 0) {
        Rcpp::CharacterVector cv_s = "";
        cv_s = Rcpp::as<Rcpp::CharacterVector>(df_border[j])[i];

        // only write attributes where cv_s has a value
        if (cv_s[0] != "") {
          // Rf_PrintValue(cv_s);
          const std::string val_strl = Rcpp::as<std::string>(cv_s);
          border.append_attribute(attrnams[j].c_str()) = val_strl.c_str();
        }
      }

      if (nam_chlds.count(attr_j) != 0) {
        Rcpp::CharacterVector cv_s = "";
        cv_s = Rcpp::as<Rcpp::CharacterVector>(df_border[j])[i];

        if (cv_s[0] != "") {

          std::string font_i = Rcpp::as<std::string>(cv_s[0]);

          pugi::xml_document border_node;
          pugi::xml_parse_result result = border_node.load_string(font_i.c_str(), pugi_parse_flags);
          if (!result) Rcpp::stop("loading border node fail: %s", cv_s);

          border.append_copy(border_node.first_child());

        }
      }

      if (idx1.empty() && idx2.empty())
        Rcpp::warning("%s: not found in border name table", attr_j);
    }

    std::ostringstream oss;
    doc.print(oss, " ", pugi_format_flags);
    z[i] = oss.str();
  }

  return z;
}


// [[Rcpp::export]]
Rcpp::DataFrame read_fill(XPtrXML xml_doc_fill) {

  // https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.font?view=openxml-2.8.1

  // openxml 2.8.1
  std::set<std::string> nams {
    "gradientFill",
    "patternFill"
  };

  size_t nn = std::distance(xml_doc_fill->begin(), xml_doc_fill->end());
  size_t kk = nams.size();
  unsigned int pugi_format_flags = pugi::format_raw | pugi::format_no_escapes;

  Rcpp::CharacterVector rvec(nn);

  // 1. create the list
  Rcpp::List df(kk);
  for (size_t i = 0; i < kk; ++i)
  {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
  }

  // 2. fill the list
  // <xf ...>
  auto itr = 0;
  for (auto xml_fill : xml_doc_fill->children("fill")) {

    for (auto cld : xml_fill.children()) {

      std::string name = cld.name();
      auto find_res = nams.find(name);

      // check if name is already known
      if (nams.count(name) == 0) {
        Rcpp::warning("%s: not found in fill name table", name);
      } else {
        // TODO why is this needed here?
        std::ostringstream oss;
        cld.print(oss, " ", pugi_format_flags);

        R_xlen_t mtc = std::distance(nams.begin(), find_res);
        Rcpp::as<Rcpp::CharacterVector>(df[mtc])[itr] = oss.str();
      }

    } // end aligment, extLst, protection

    // rownames as character vectors matching to <c s= ...>
    rvec[itr] = std::to_string(itr);

    ++itr;
  }

  // 3. Create a data.frame
  df.attr("row.names") = rvec;
  df.attr("names") = nams;
  df.attr("class") = "data.frame";

  return df;
}


// [[Rcpp::export]]
Rcpp::CharacterVector write_fill(Rcpp::DataFrame df_fill) {

  auto n = df_fill.nrow();
  Rcpp::CharacterVector z(n);
  unsigned int pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_ws_pcdata | pugi::parse_eol;
  unsigned int pugi_format_flags = pugi::format_raw | pugi::format_no_escapes;

  for (auto i = 0; i < n; ++i) {
    pugi::xml_document doc;

    pugi::xml_node fill = doc.append_child("fill");


    for (auto j = 0; j < df_fill.ncol(); ++j) {

      Rcpp::CharacterVector cv_s = "";
      cv_s = Rcpp::as<Rcpp::CharacterVector>(df_fill[j])[i];

      if (cv_s[0] != "") {

        std::string font_i = Rcpp::as<std::string>(cv_s[0]);

        pugi::xml_document font_node;
        pugi::xml_parse_result result = font_node.load_string(font_i.c_str(), pugi_parse_flags);
        if (!result) Rcpp::stop("loading fill node fail: %s", cv_s);

        fill.append_copy(font_node.first_child());

      }

    }

    std::ostringstream oss;
    doc.print(oss, " ", pugi_format_flags);

    z[i] = oss.str();
  }

  return z;
}


// [[Rcpp::export]]
Rcpp::DataFrame read_cellStyle(XPtrXML xml_doc_cellStyle) {

  // https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.border?view=openxml-2.8.1

  // openxml 2.8.1
  std::set<std::string> nam_attrs {"builtinId", "customBuiltin", "hidden",
                                   "iLevel", "name", "xfId", "xr:uid"};

  std::set<std::string> nam_chlds {"extLst"};

  auto total_length = nam_attrs.size() + nam_chlds.size();
  std::vector<std::string> all_names(total_length);

  std::copy(nam_attrs.begin(), nam_attrs.end(), all_names.begin());
  std::copy(nam_chlds.begin(), nam_chlds.end(), all_names.begin() + nam_attrs.size());

  std::set<std::string> nams(std::make_move_iterator(all_names.begin()),
                             std::make_move_iterator(all_names.end()));


  size_t nn = std::distance(xml_doc_cellStyle->begin(), xml_doc_cellStyle->end());
  size_t kk = nams.size();
  unsigned int pugi_format_flags = pugi::format_raw | pugi::format_no_escapes;

  Rcpp::CharacterVector rvec(nn);

  // 1. create the list
  Rcpp::List df(kk);
  for (size_t i = 0; i < kk; ++i)
  {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
  }

  // 2. fill the list
  // <numFmt ...>
  auto itr = 0;
  for (auto xml_cellStyle : xml_doc_cellStyle->children("cellStyle")) {
    for (auto attrs : xml_cellStyle.attributes()) {

      std::string attr_name = attrs.name();
      std::string attr_value = attrs.value();
      auto find_res = nams.find(attr_name);

      // check if name is already known
      if (nams.count(attr_name) == 0) {
        Rcpp::warning("%s: not found in cellstyle name table", attr_name);
      } else {
        R_xlen_t mtc = std::distance(nams.begin(), find_res);
        Rcpp::as<Rcpp::CharacterVector>(df[mtc])[itr] = attr_value;
      }
    }

    for (auto cld : xml_cellStyle.children()) {

      std::string cld_name = cld.name();
      auto find_res = nams.find(cld_name);

      // check if name is already known
      if (nams.count(cld_name) == 0) {
        Rcpp::warning("%s: not found in cellstyle name table", cld_name);
      } else {
        std::ostringstream oss;
        cld.print(oss, " ", pugi_format_flags);
        std::string cld_value = oss.str();

        R_xlen_t mtc = std::distance(nams.begin(), find_res);
        Rcpp::as<Rcpp::CharacterVector>(df[mtc])[itr] = cld_value;
      }
    }

    // rownames as character vectors matching to <c s= ...>
    rvec[itr] = std::to_string(itr);
    ++itr;

  }

  // 3. Create a data.frame
  df.attr("row.names") = rvec;
  df.attr("names") = nams;
  df.attr("class") = "data.frame";

  return df;
}


// [[Rcpp::export]]
Rcpp::CharacterVector write_cellStyle(Rcpp::DataFrame df_cellstyle) {

  auto n = df_cellstyle.nrow();
  Rcpp::CharacterVector z(n);
  unsigned int pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_ws_pcdata | pugi::parse_eol;
  unsigned int pugi_format_flags = pugi::format_raw | pugi::format_no_escapes;


  // openxml 2.8.1
  std::vector<std::string>  attrnams = df_cellstyle.names();
  std::set<std::string> nam_attrs = {"builtinId", "customBuiltin", "hidden", "iLevel", "name", "xfId"};
  std::set<std::string> nam_chlds = {"extLst"};


  Rcpp::IntegerVector mtc1, mtc2, idx1, idx2;

  for (auto i = 0; i < n; ++i) {
    pugi::xml_document doc;
    pugi::xml_node cellstyle = doc.append_child("cellStyle");

    for (auto j = 0; j < df_cellstyle.ncol(); ++j) {

      std::string attr_j = attrnams[j];

      // mimic which
      auto res1 = nam_attrs.find(attr_j);
      auto mtc1 = std::distance(nam_attrs.begin(), res1);

      std::vector<int> idx1(mtc1 + 1);
      std::iota(idx1.begin(), idx1.end(), 0);

      auto res2 = nam_chlds.find(attr_j);
      auto mtc2 = std::distance(nam_chlds.begin(), res2);

      std::vector<int> idx2(mtc2 + 1);
      std::iota(idx2.begin(), idx2.end(), 0);

      // check if name is already known
      if (nam_attrs.count(attr_j) != 0) {
        Rcpp::CharacterVector cv_s = "";
        cv_s = Rcpp::as<Rcpp::CharacterVector>(df_cellstyle[j])[i];

        // only write attributes where cv_s has a value
        if (cv_s[0] != "") {
          // Rf_PrintValue(cv_s);
          const std::string val_strl = Rcpp::as<std::string>(cv_s);
          cellstyle.append_attribute(attrnams[j].c_str()) = val_strl.c_str();
        }
      }

      if (nam_chlds.count(attr_j) != 0) {
        Rcpp::CharacterVector cv_s = "";
        cv_s = Rcpp::as<Rcpp::CharacterVector>(df_cellstyle[j])[i];

        if (cv_s[0] != "") {

          std::string font_i = Rcpp::as<std::string>(cv_s[0]);

          pugi::xml_document border_node;
          pugi::xml_parse_result result = border_node.load_string(font_i.c_str(), pugi_parse_flags);
          if (!result) Rcpp::stop("loading cellStyle node fail: %s", cv_s);

          cellstyle.append_copy(border_node.first_child());

        }
      }

      if (idx1.empty() && idx2.empty())
        Rcpp::warning("%s: not found in cellStyle name table", attr_j);
    }

    std::ostringstream oss;
    doc.print(oss, " ", pugi_format_flags);
    z[i] = oss.str();
  }

  return z;
}


// [[Rcpp::export]]
Rcpp::DataFrame read_tableStyle(XPtrXML xml_doc_tableStyle) {

  // https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.border?view=openxml-2.8.1

  // openxml 2.8.1
  std::set<std::string> nam_attrs {"count", "name", "pivot", "table", "xr9:uid"};

  std::set<std::string> nam_chlds {"tableStyleElement"};

  auto total_length = nam_attrs.size() + nam_chlds.size();
  std::vector<std::string> all_names(total_length);

  std::copy(nam_attrs.begin(), nam_attrs.end(), all_names.begin());
  std::copy(nam_chlds.begin(), nam_chlds.end(), all_names.begin() + nam_attrs.size());

  std::set<std::string> nams(std::make_move_iterator(all_names.begin()),
                             std::make_move_iterator(all_names.end()));


  size_t nn = std::distance(xml_doc_tableStyle->begin(), xml_doc_tableStyle->end());
  size_t kk = nams.size();
  unsigned int pugi_format_flags = pugi::format_raw | pugi::format_no_escapes;

  Rcpp::CharacterVector rvec(nn);

  // 1. create the list
  Rcpp::List df(kk);
  for (size_t i = 0; i < kk; ++i)
  {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
  }

  // 2. fill the list
  // <numFmt ...>
  auto itr = 0;
  for (auto xml_tableStyle : xml_doc_tableStyle->children("tableStyle")) {
    for (auto attrs : xml_tableStyle.attributes()) {

      std::string attr_name = attrs.name();
      std::string attr_value = attrs.value();
      auto find_res = nams.find(attr_name);

      // check if name is already known
      if (nams.count(attr_name) == 0) {
        Rcpp::warning("%s: not found in tablestyle name table", attr_name);
      } else {
        R_xlen_t mtc = std::distance(nams.begin(), find_res);
        Rcpp::as<Rcpp::CharacterVector>(df[mtc])[itr] = attr_value;
      }
    }

    std::string cld_value;
    for (auto cld : xml_tableStyle.children()) {

      std::string cld_name = cld.name();
      auto find_res = nams.find(cld_name);

      // check if name is already known
      if (nams.count(cld_name) == 0) {
        Rcpp::warning("%s: not found in tablestyle name table", cld_name);
      } else {
        std::ostringstream oss;
        cld.print(oss, " ", pugi_format_flags);
        cld_value += oss.str();

        R_xlen_t mtc = std::distance(nams.begin(), find_res);
        Rcpp::as<Rcpp::CharacterVector>(df[mtc])[itr] = cld_value;
      }
    }

    // rownames as character vectors matching to <c s= ...>
    rvec[itr] = std::to_string(itr);
    ++itr;

  }

  // 3. Create a data.frame
  df.attr("row.names") = rvec;
  df.attr("names") = nams;
  df.attr("class") = "data.frame";

  return df;
}


// [[Rcpp::export]]
Rcpp::CharacterVector write_tableStyle(Rcpp::DataFrame df_tablestyle) {

  auto n = df_tablestyle.nrow();
  Rcpp::CharacterVector z(n);
  unsigned int pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_ws_pcdata | pugi::parse_eol;
  unsigned int pugi_format_flags = pugi::format_raw | pugi::format_no_escapes;


  // openxml 2.8.1
  std::vector<std::string> attrnams = df_tablestyle.names();

  std::set<std::string> nam_attrs = {"count", "name", "pivot", "table", "xr9:uid"};
  std::set<std::string> nam_chlds = {"tableStyleElement"};


  for (auto i = 0; i < n; ++i) {
    pugi::xml_document doc;
    pugi::xml_node tablestyle = doc.append_child("tableStyle");

    for (auto j = 0; j < df_tablestyle.ncol(); ++j) {

      std::string attr_j = attrnams[j];

      // mimic which
      auto res1 = nam_attrs.find(attr_j);
      auto mtc1 = std::distance(nam_attrs.begin(), res1);

      std::vector<int> idx1(mtc1 + 1);
      std::iota(idx1.begin(), idx1.end(), 0);

      auto res2 = nam_chlds.find(attr_j);
      auto mtc2 = std::distance(nam_chlds.begin(), res2);

      std::vector<int> idx2(mtc2 + 1);
      std::iota(idx2.begin(), idx2.end(), 0);

      // check if name is already known
      if (nam_attrs.count(attr_j) != 0) {
        Rcpp::CharacterVector cv_s = "";
        cv_s = Rcpp::as<Rcpp::CharacterVector>(df_tablestyle[j])[i];

        // only write attributes where cv_s has a value
        if (cv_s[0] != "") {
          // Rf_PrintValue(cv_s);
          const std::string val_strl = Rcpp::as<std::string>(cv_s);
          tablestyle.append_attribute(attrnams[j].c_str()) = val_strl.c_str();
        }
      }

      if (nam_chlds.count(attr_j) != 0) {
        Rcpp::CharacterVector cv_s = "";
        cv_s = Rcpp::as<Rcpp::CharacterVector>(df_tablestyle[j])[i];

        if (cv_s[0] != "") {

          std::string font_i = Rcpp::as<std::string>(cv_s[0]);

          pugi::xml_document tableStyleElement;
          pugi::xml_parse_result result = tableStyleElement.load_string(font_i.c_str(), pugi_parse_flags);
          if (!result) Rcpp::stop("loading df_tablestyle node fail: %s", cv_s);

          for (auto chld: tableStyleElement.children())
            tablestyle.append_copy(chld);

        }
      }

      if (idx1.empty() && idx2.empty())
        Rcpp::warning("%s: not found in df_tablestyle name table", attr_j);
    }

    std::ostringstream oss;
    doc.print(oss, " ", pugi_format_flags);
    z[i] = oss.str();
  }

  return z;
}



// [[Rcpp::export]]
Rcpp::DataFrame read_dxf(XPtrXML xml_doc_dxf) {

  // https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.font?view=openxml-2.8.1

  // openxml 2.8.1
  std::set<std::string> nams {
    "alignment",
    "border",
    "extLst",
    "fill",
    "font",
    "numFmt",
    "protection"
  };

  size_t nn = std::distance(xml_doc_dxf->begin(), xml_doc_dxf->end());
  size_t kk = nams.size();
  unsigned int pugi_format_flags = pugi::format_raw | pugi::format_no_escapes;

  Rcpp::CharacterVector rvec(nn);

  // 1. create the list
  Rcpp::List df(kk);
  for (size_t i = 0; i < kk; ++i)
  {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
  }

  // 2. fill the list
  // <xf ...>
  auto itr = 0;
  for (auto xml_dxf : xml_doc_dxf->children("dxf")) {

    for (auto cld : xml_dxf.children()) {

      std::string name = cld.name();
      auto find_res = nams.find(name);

      // check if name is already known
      if (nams.count(name) == 0) {
        Rcpp::warning("%s: not found in dxf name table", name);
      } else {
        std::ostringstream oss;
        cld.print(oss, " ", pugi_format_flags);
        std::string value = oss.str();

        R_xlen_t mtc = std::distance(nams.begin(), find_res);
        Rcpp::as<Rcpp::CharacterVector>(df[mtc])[itr] = value;
      }

    } // end aligment, extLst, protection

    // rownames as character vectors matching to <c s= ...>
    rvec[itr] = std::to_string(itr);

    ++itr;
  }

  // 3. Create a data.frame
  df.attr("row.names") = rvec;
  df.attr("names") = nams;
  df.attr("class") = "data.frame";

  return df;
}


// [[Rcpp::export]]
Rcpp::CharacterVector write_dxf(Rcpp::DataFrame df_dxf) {

  auto n = df_dxf.nrow();
  Rcpp::CharacterVector z(n);
  unsigned int pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_ws_pcdata | pugi::parse_eol;
  unsigned int pugi_format_flags = pugi::format_raw | pugi::format_no_escapes;

  for (auto i = 0; i < n; ++i) {
    pugi::xml_document doc;

    pugi::xml_node dxf = doc.append_child("dxf");


    for (auto j = 0; j < df_dxf.ncol(); ++j) {

      Rcpp::CharacterVector cv_s = "";
      cv_s = Rcpp::as<Rcpp::CharacterVector>(df_dxf[j])[i];

      if (cv_s[0] != "") {

        std::string font_i = Rcpp::as<std::string>(cv_s[0]);

        pugi::xml_document font_node;
        pugi::xml_parse_result result = font_node.load_string(font_i.c_str(), pugi_parse_flags);
        if (!result) Rcpp::stop("loading dxf node fail: %s", cv_s);

        dxf.append_copy(font_node.first_child());

      }

    }

    std::ostringstream oss;
    doc.print(oss, " ", pugi_format_flags);

    z[i] = oss.str();
  }

  return z;
}



// [[Rcpp::export]]
Rcpp::DataFrame read_colors(XPtrXML xml_doc_colors) {

  // https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.colors?view=openxml-2.8.1

  // openxml 2.8.1
  std::set<std::string> nams {
    "indexedColors",
    "mruColors"
  };

  size_t nn = std::distance(xml_doc_colors->begin(), xml_doc_colors->end());
  size_t kk = nams.size();
  unsigned int pugi_format_flags = pugi::format_raw | pugi::format_no_escapes;

  Rcpp::CharacterVector rvec(nn);

  // 1. create the list
  Rcpp::List df(kk);
  for (size_t i = 0; i < kk; ++i)
  {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
  }

  // 2. fill the list
  // <xf ...>
  auto itr = 0;
  for (auto xml_color : xml_doc_colors->children("colors")) {

    for (auto cld : xml_color.children()) {

      std::string name = cld.name();
      auto find_res = nams.find(name);

      // check if name is already known
      if (nams.count(name) == 0) {
        Rcpp::warning("%s: not found in color name table", name);
      } else {
        std::ostringstream oss;
        cld.print(oss, " ", pugi_format_flags);
        std::string cld_value = oss.str();

        R_xlen_t mtc = std::distance(nams.begin(), find_res);
        Rcpp::as<Rcpp::CharacterVector>(df[mtc])[itr] = cld_value;
      }

    } // end aligment, extLst, protection

    // rownames as character vectors matching to <c s= ...>
    rvec[itr] = std::to_string(itr);

    ++itr;
  }

  // 3. Create a data.frame
  df.attr("row.names") = rvec;
  df.attr("names") = nams;
  df.attr("class") = "data.frame";

  return df;
}


// [[Rcpp::export]]
Rcpp::CharacterVector write_colors(Rcpp::DataFrame df_colors) {

  auto n = df_colors.nrow();
  Rcpp::CharacterVector z(n);
  unsigned int pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_ws_pcdata | pugi::parse_eol;
  unsigned int pugi_format_flags = pugi::format_raw | pugi::format_no_escapes;

  for (auto i = 0; i < n; ++i) {
    pugi::xml_document doc;

    pugi::xml_node color = doc.append_child("colors");


    for (auto j = 0; j < df_colors.ncol(); ++j) {

      Rcpp::CharacterVector cv_s = "";
      cv_s = Rcpp::as<Rcpp::CharacterVector>(df_colors[j])[i];

      if (cv_s[0] != "") {

        std::string font_i = Rcpp::as<std::string>(cv_s[0]);

        pugi::xml_document font_node;
        pugi::xml_parse_result result = font_node.load_string(font_i.c_str(), pugi_parse_flags);
        if (!result) Rcpp::stop("loading color node fail: %s", cv_s);

        color.append_copy(font_node.first_child());

      }

    }

    std::ostringstream oss;
    doc.print(oss, " ", pugi_format_flags);

    z[i] = oss.str();
  }

  return z;
}
