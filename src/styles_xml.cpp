#include "openxlsx2.h"

// [[Rcpp::export]]
Rcpp::DataFrame read_xf(XPtrXML xml_doc_xf) {

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

  Rcpp::CharacterVector rvec(nn);

  // 1. create the list
  Rcpp::List df(kk);
  for (auto i = 0; i < kk; ++i)
  {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
  }

  // 2. fill the list
  // <xf ...>
  auto itr = 0;
  for (auto xml_xf : xml_doc_xf->children("xf")) {
    for (auto attrs : xml_xf.attributes()) {

      Rcpp::CharacterVector attr_name = attrs.name();
      std::string attr_value = attrs.value();

      // mimic which
      Rcpp::IntegerVector mtc = Rcpp::match(nams, attr_name);
      Rcpp::IntegerVector idx = Rcpp::seq(0, mtc.length()-1);

      // check if name is already known
      if (all(Rcpp::is_na(mtc))) {
        Rcpp::Rcout << attr_name << ": not found in xf name table" << std::endl;
      } else {
        size_t ii = Rcpp::as<size_t>(idx[!Rcpp::is_na(mtc)]);
        Rcpp::as<Rcpp::CharacterVector>(df[ii])[itr] = attr_value;
      }
    }

    // only handle known names
    // <alignment ...>
    // <extLst ...> // FIXME should be imported as single node. most likely broken
    // <protection ...>
    for (auto cld : xml_xf.children()) {

      std::string cld_name = cld.name();

      // check known names
      if (cld_name ==  "alignment" | cld_name == "extLst" | cld_name == "protection") {

        for (auto attrs : cld.attributes()) {
          Rcpp::CharacterVector attr_name = attrs.name();
          std::string attr_value = attrs.value();

          // mimic which
          Rcpp::IntegerVector mtc = Rcpp::match(nams, attr_name);
          Rcpp::IntegerVector idx = Rcpp::seq(0, mtc.length()-1);

          // check if name is already known
          if (all(Rcpp::is_na(mtc))) {
            Rcpp::Rcout << attr_name << ": not found in xf name table" << std::endl;
          } else {
            size_t ii = Rcpp::as<size_t>(idx[!Rcpp::is_na(mtc)]);
            Rcpp::as<Rcpp::CharacterVector>(df[ii])[itr] = attr_value;
          }
        }
      } else {
        Rcpp::Rcout << cld_name << ": not valid for xf in openxml 2.8.1" << std::endl;
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




// helper function to check if row contains any of the expected types
bool has_it(Rcpp::DataFrame df_xf, Rcpp::CharacterVector xf_nams, size_t row) {

  bool has_it = false;
  Rcpp::CharacterVector attrnams = df_xf.names();

  Rcpp::DataFrame df_tmp;
  Rcpp::CharacterVector cv_tmp;

  Rcpp::IntegerVector mtc = Rcpp::match(attrnams, xf_nams);
  Rcpp::IntegerVector idx = Rcpp::seq(0, mtc.length()-1);

  idx = idx[!Rcpp::is_na(mtc)];
  df_tmp = df_xf[idx];

  for (auto ii = 0; ii < df_tmp.ncol(); ++ii) {
    cv_tmp = Rcpp::as<Rcpp::CharacterVector>(df_tmp[ii])[row];
    if (cv_tmp[0] != "") has_it = true;
  }

  return has_it;
}


// [[Rcpp::export]]
Rcpp::CharacterVector write_xf(Rcpp::DataFrame df_xf) {

  auto n = df_xf.nrow();
  Rcpp::CharacterVector z(n);

  Rcpp::CharacterVector xf_nams = {
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
    "xfId"
  };

  Rcpp::CharacterVector xf_nams_alignment = {
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

  Rcpp::CharacterVector xf_nams_extLst = {
    "extLst"
  };

  Rcpp::CharacterVector xf_nams_protection = {
    "hidden",
    "locked"
  };

  for (auto i = 0; i < n; ++i) {
    pugi::xml_document doc;

    pugi::xml_node xf = doc.append_child("xf");
    Rcpp::CharacterVector attrnams = df_xf.names();

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

      Rcpp::CharacterVector attrnam = Rcpp::as<Rcpp::CharacterVector>(attrnams[j]);

      // not all missing in match: ergo they are
      bool is_xf = !Rcpp::as<bool>(Rcpp::all(Rcpp::is_na(Rcpp::match(xf_nams, attrnam))));
      bool is_alignment = !Rcpp::as<bool>(Rcpp::all(Rcpp::is_na(Rcpp::match(xf_nams_alignment, attrnam))));
      bool is_extLst = !Rcpp::as<bool>(Rcpp::all(Rcpp::is_na(Rcpp::match(xf_nams_extLst, attrnam))));
      bool is_protection = !Rcpp::as<bool>(Rcpp::all(Rcpp::is_na(Rcpp::match(xf_nams_protection, attrnam))));

      // <xf ...>
      if (is_xf) {
        Rcpp::CharacterVector cv_s = "";
        cv_s = Rcpp::as<Rcpp::CharacterVector>(df_xf[j])[i];

        // only write attributes where cv_s has a value
        if (cv_s[0] != "") {
          // Rf_PrintValue(cv_s);
          const std::string val_strl = Rcpp::as<std::string>(cv_s);
          xf.append_attribute(attrnams[j]) = val_strl.c_str();
        }
      }

      if (has_alignment && is_alignment) {
        Rcpp::CharacterVector cv_s = "";
        cv_s = Rcpp::as<Rcpp::CharacterVector>(df_xf[j])[i];

        if (cv_s[0] != "") {
          const std::string val_strl = Rcpp::as<std::string>(cv_s);
          xf_alignment.append_attribute(attrnams[j]) = val_strl.c_str();
        }
      }

      // FIXME should be written as single node. most likely broken
      if (has_extLst && is_extLst) {
        Rcpp::CharacterVector cv_s = "";
        cv_s = Rcpp::as<Rcpp::CharacterVector>(df_xf[j])[i];

        if (cv_s[0] != "") {
          const std::string val_strl = Rcpp::as<std::string>(cv_s);
          xf_extLst.append_attribute(attrnams[j]) = val_strl.c_str();
        }
      }

      if (has_protection && is_protection) {
        Rcpp::CharacterVector cv_s = "";
        cv_s = Rcpp::as<Rcpp::CharacterVector>(df_xf[j])[i];

        if (cv_s[0] != "") {
          const std::string val_strl = Rcpp::as<std::string>(cv_s);
          xf_protection.append_attribute(attrnams[j]) = val_strl.c_str();
        }
      }

    }

    std::ostringstream oss;
    doc.print(oss, " ", pugi::format_raw);

    z[i] = oss.str();
  }

  return z;
}


// [[Rcpp::export]]
Rcpp::DataFrame read_font(XPtrXML xml_doc_font) {

  // https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.font?view=openxml-2.8.1

  // openxml 2.8.1
  Rcpp::CharacterVector nams = {
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

  auto nn = std::distance(xml_doc_font->begin(), xml_doc_font->end());
  auto kk = nams.length();

  Rcpp::CharacterVector rvec(nn);

  // 1. create the list
  Rcpp::List df(kk);
  for (auto i = 0; i < kk; ++i)
  {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
  }

  // 2. fill the list
  // <xf ...>
  auto itr = 0;
  for (auto xml_font : xml_doc_font->children("font")) {

    for (auto cld : xml_font.children()) {

      Rcpp::CharacterVector name = cld.name();
      std::string value = cld.value();

      // mimic which
      Rcpp::IntegerVector mtc = Rcpp::match(nams, name);
      Rcpp::IntegerVector idx = Rcpp::seq(0, mtc.length()-1);

      // check if name is already known
      if (all(Rcpp::is_na(mtc))) {
        Rcpp::Rcout << name << ": not found in font name table" << std::endl;
      } else {
        std::ostringstream oss;
        cld.print(oss, " ", pugi::format_raw);

        size_t ii = Rcpp::as<size_t>(idx[!Rcpp::is_na(mtc)]);
        Rcpp::as<Rcpp::CharacterVector>(df[ii])[itr] = oss.str();
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

  for (auto i = 0; i < n; ++i) {
    pugi::xml_document doc;

    pugi::xml_node font = doc.append_child("font");


    for (auto j = 0; j < df_font.ncol(); ++j) {

      Rcpp::CharacterVector cv_s = "";
      cv_s = Rcpp::as<Rcpp::CharacterVector>(df_font[j])[i];

      if (cv_s[0] != "") {

        std::string font_i = Rcpp::as<std::string>(cv_s[0]);

        pugi::xml_document font_node;
        pugi::xml_parse_result result = font_node.load_string(font_i.c_str(), pugi::parse_default | pugi::parse_escapes);
        if (!result) Rcpp::stop("loading font node fail: %s", cv_s);

        font.append_copy(font_node.first_child());

      }

    }

    std::ostringstream oss;
    doc.print(oss, " ", pugi::format_raw);

    z[i] = oss.str();
  }

  return z;
}


// [[Rcpp::export]]
Rcpp::DataFrame read_numfmt(XPtrXML xml_doc_font) {

  // https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-2.8.1

  // openxml 2.8.1
  Rcpp::CharacterVector nams = {
    "formatCode",
    "numFmtId"
  };

  auto nn = std::distance(xml_doc_font->begin(), xml_doc_font->end());
  auto kk = nams.length();

  Rcpp::CharacterVector rvec(nn);

  // 1. create the list
  Rcpp::List df(kk);
  for (auto i = 0; i < kk; ++i)
  {
    SET_VECTOR_ELT(df, i, Rcpp::CharacterVector(Rcpp::no_init(nn)));
  }

  // 2. fill the list
  // <numFmt ...>
  auto itr = 0;
  for (auto xml_numfmt : xml_doc_font->children("numFmt")) {
    for (auto attrs : xml_numfmt.attributes()) {

      Rcpp::CharacterVector attr_name = attrs.name();
      std::string attr_value = attrs.value();

      // mimic which
      Rcpp::IntegerVector mtc = Rcpp::match(nams, attr_name);
      Rcpp::IntegerVector idx = Rcpp::seq(0, mtc.length()-1);

      // check if name is already known
      if (all(Rcpp::is_na(mtc))) {
        Rcpp::Rcout << attr_name << ": not found in xf name table" << std::endl;
      } else {
        size_t ii = Rcpp::as<size_t>(idx[!Rcpp::is_na(mtc)]);
        Rcpp::as<Rcpp::CharacterVector>(df[ii])[itr] = attr_value;
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
    doc.print(oss, " ", pugi::format_raw);

    z[i] = oss.str();
  }

  return z;
}
