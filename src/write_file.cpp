#include "openxlsx2.h"
#include <fstream>

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

R_xlen_t select_rows(std::vector<std::string> x, std::string row, std::string col) {
  std::string y = col + row;
  for (auto i = 0; i < x.size(); ++i) {
    if (x[i] == y)
      return(i);
  }
  // else
  return(0);
}

// @param cc cc
// @param x x
// @param data_class data_class
// @param rows rows
// @param cols cols
// @param colNames colNames
// @param removeCellStyle removeCellStyle
// @param cell cell
// @param hyperlinkstyle hyperlinkstyle
// @param no_na_strings no_na_strings
// @param na_strings_ na_strings
// [[Rcpp::export]]
void update_cell_loop(
    Rcpp::DataFrame cc,
    Rcpp::DataFrame x,
    Rcpp::CharacterVector data_class,
    std::vector<std::string> rows,
    std::vector<std::string> cols,
    bool colNames,
    bool removeCellStyle,
    std::string cell,
    std::string hyperlinkstyle,
    bool no_na_strings,
    Rcpp::Nullable<Rcpp::String> na_strings_ = R_NilValue
) {

  auto i = 0, m = 0;
    for (auto &col : cols) {

    auto n = 0;

    for (auto &row : rows) {

      // check if is data frame or matrix
      Rcpp::String value = "";
      value = Rcpp::wrap(Rcpp::as<Rcpp::CharacterVector>(x[m])[n]);

      R_xlen_t sel = select_rows(cc["r"], row, col);

      if (removeCellStyle) Rcpp::as<Rcpp::CharacterVector>(cc["c_s"])[sel] = "";
      Rcpp::as<Rcpp::CharacterVector>(cc["c_t"])[sel] = "";
      Rcpp::as<Rcpp::CharacterVector>(cc["c_cm"])[sel] = "";
      Rcpp::as<Rcpp::CharacterVector>(cc["c_ph"])[sel] = "";
      Rcpp::as<Rcpp::CharacterVector>(cc["c_vm"])[sel] = "";
      Rcpp::as<Rcpp::CharacterVector>(cc["v"])[sel] = "";
      Rcpp::as<Rcpp::CharacterVector>(cc["f"])[sel] = "";
      Rcpp::as<Rcpp::CharacterVector>(cc["f_t"])[sel] = "";
      Rcpp::as<Rcpp::CharacterVector>(cc["f_ref"])[sel] = "";
      Rcpp::as<Rcpp::CharacterVector>(cc["f_ca"])[sel] = "";
      Rcpp::as<Rcpp::CharacterVector>(cc["f_si"])[sel] = "";
      Rcpp::as<Rcpp::CharacterVector>(cc["is"])[sel] = "";

      // for now convert all R-characters to inlineStr (e.g. names() of a data frame)
      if ((data_class[m] == character) || ((colNames) && (n == 0))) {
        if (value == NA_STRING) {
          Rcpp::as<Rcpp::CharacterVector>(cc["v"])[sel] = "#N/A";
          Rcpp::as<Rcpp::CharacterVector>(cc["c_t"])[sel] = "e";
        } else {
          Rcpp::as<Rcpp::CharacterVector>(cc["c_t"])[sel] = "inlineStr";
          Rcpp::as<Rcpp::CharacterVector>(cc["is"])[sel] = txt_to_is(value.get_cstring(), 0, 1).c_str();
        }
      } else if (data_class[m] == formula) {
        Rcpp::as<Rcpp::CharacterVector>(cc["c_t"])[sel] = "str";
        Rcpp::as<Rcpp::CharacterVector>(cc["f"])[sel] = value.get_cstring();
      } else if (data_class[m] == array_formula) {
        Rcpp::as<Rcpp::CharacterVector>(cc["f"])[sel] = value.get_cstring();
        Rcpp::as<Rcpp::CharacterVector>(cc["f_t"])[sel] = "array";
        Rcpp::as<Rcpp::CharacterVector>(cc["f_ref"])[sel] = cell.c_str();
      } else if (data_class[m] == hyperlink) {
        Rcpp::as<Rcpp::CharacterVector>(cc["f"])[sel] = value.get_cstring();
        /* ** not yet implemented **
        //FIXME assign the hyperlinkstyle if no style found. This might not be
        // desired. We should provide an option to prevent this.
        if (cc[sel, "c_s"] == "" || is.na(cc[sel, "c_s"]))
          cc[sel, "c_s"] <- hyperlinkstyle.c_str()
        */
      } else {
        if (value == NA_STRING) {
          if (no_na_strings) {
            Rcpp::as<Rcpp::CharacterVector>(cc["v"])[sel]   = "#N/A";
            Rcpp::as<Rcpp::CharacterVector>(cc["c_t"])[sel] = "e";
          } else {
            if (na_strings_.isNull()) {
              // do not add any value: <c/>
            } else {
              Rcpp::String na_strings(na_strings_);

              Rcpp::as<Rcpp::CharacterVector>(cc["c_t"])[sel] = "inlineStr";
              Rcpp::as<Rcpp::CharacterVector>(cc["is"])[sel]  = txt_to_is(na_strings, 0, 1).c_str();
            }
          }
        } else if (value == "NaN") {
          Rcpp::as<Rcpp::CharacterVector>(cc["v"])[sel]   = "#VALUE!";
          Rcpp::as<Rcpp::CharacterVector>(cc["c_t"])[sel] = "e";
        } else if (value == "-Inf" || value == "Inf") {
          Rcpp::as<Rcpp::CharacterVector>(cc["v"])[sel]   = "#NUM!";
          Rcpp::as<Rcpp::CharacterVector>(cc["c_t"])[sel] = "e";
        } else {
          Rcpp::as<Rcpp::CharacterVector>(cc["v"])[sel] = value.get_cstring();
        }
      }
      ++i;
      ++n;

      // Rcpp::Rcout << "col: " << m << std::endl;
    }
    ++m;

    // Rcpp::Rcout << "row: " << n << std::endl;
  }
}


// helper function to access element from Rcpp::Character Vector as string
std::string to_string(Rcpp::Vector<16>::Proxy x) {
  return Rcpp::String(x);
}


// creates an xml row
// data in xml is ordered row wise. therefore we need the row attributes and
// the column data used in this row. This function uses both to create a single
// row and passes it to write_worksheet_xml_2 which will create the entire
// sheet_data part for this worksheet
pugi::xml_document xml_sheet_data(Rcpp::DataFrame row_attr, Rcpp::DataFrame cc) {

  auto lastrow = 0; // integer value of the last row with column data
  auto thisrow = 0; // integer value of the current row with column data
  auto row_idx = 0; // the index of the row_attr file. this is != rowid
  auto rowid   = 0; // integer value of the r field in row_attr

  pugi::xml_document doc;
  pugi::xml_node row;

  std::string xml_preserver = " ";

  // non optional: treat input as valid at this stage
  unsigned int pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_ws_pcdata | pugi::parse_eol;
  // unsigned int pugi_format_flags = pugi::format_raw | pugi::format_no_escapes;

  // we cannot access rows directly in the dataframe.
  // Have to extract the columns and use these
  Rcpp::CharacterVector cc_row_r = cc["row_r"]; // 1
  Rcpp::CharacterVector cc_r     = cc["r"];     // A1
  Rcpp::CharacterVector cc_v     = cc["v"];
  Rcpp::CharacterVector cc_c_t   = cc["c_t"];
  Rcpp::CharacterVector cc_c_s   = cc["c_s"];
  Rcpp::CharacterVector cc_c_cm  = cc["c_cm"];
  Rcpp::CharacterVector cc_c_ph  = cc["c_ph"];
  Rcpp::CharacterVector cc_c_vm  = cc["c_vm"];
  Rcpp::CharacterVector cc_f     = cc["f"];
  Rcpp::CharacterVector cc_f_t   = cc["f_t"];
  Rcpp::CharacterVector cc_f_ref = cc["f_ref"];
  Rcpp::CharacterVector cc_f_ca  = cc["f_ca"];
  Rcpp::CharacterVector cc_f_si  = cc["f_si"];
  Rcpp::CharacterVector cc_is    = cc["is"];

  Rcpp::CharacterVector row_r    = row_attr["r"];

  for (auto i = 0; i < cc.nrow(); ++i) {

    thisrow = std::stoi(Rcpp::as<std::string>(cc_row_r[i]));

    if (lastrow < thisrow) {

      // there might be entirely empty rows in between. this is the case for
      // loadExample. We check the rowid and write the line and skip until we
      // have every row and only then continue writing the column
      while (rowid < thisrow) {

        rowid = std::stoi(Rcpp::as<std::string>(
          row_r[row_idx]
        ));

        row = doc.append_child("row");
        Rcpp::CharacterVector attrnams = row_attr.names();

        for (auto j = 0; j < row_attr.ncol(); ++j) {

          Rcpp::CharacterVector cv_s = "";
          cv_s = Rcpp::as<Rcpp::CharacterVector>(row_attr[j])[row_idx];

          if (cv_s[0] != "") {
            const std::string val_strl = Rcpp::as<std::string>(cv_s);
            row.append_attribute(attrnams[j]) = val_strl.c_str();
          }
        }

        // read the next row_idx when visiting again
        ++row_idx;
      }
    }

    // create node <c>
    pugi::xml_node cell = row.append_child("c");

    // Every cell consists of a typ and a val list. Certain functions have an
    // additional attr list.

    // append attributes <c r="A1" ...>
    cell.append_attribute("r") = to_string(cc_r[i]).c_str();

    if (!to_string(cc_c_s[i]).empty())
      cell.append_attribute("s") = to_string(cc_c_s[i]).c_str();

    // assign type if not <v> aka numeric
    if (!to_string(cc_c_t[i]).empty())
      cell.append_attribute("t") = to_string(cc_c_t[i]).c_str();

    // CellMetaIndex: suppress curly brackets in spreadsheet software
    if (!to_string(cc_c_cm[i]).empty())
      cell.append_attribute("cm") = to_string(cc_c_cm[i]).c_str();

    // phonetics spelling
    if (!to_string(cc_c_ph[i]).empty())
      cell.append_attribute("ph") = to_string(cc_c_ph[i]).c_str();

    // suppress curly brackets in spreadsheet software
    if (!to_string(cc_c_vm[i]).empty())
      cell.append_attribute("vm") = to_string(cc_c_vm[i]).c_str();

    // append nodes <c r="A1" ...><v>...</v></c>

    bool f_si = false;

    // <f> ... </f>
    // f node: formula to be evaluated
    if (!to_string(cc_f[i]).empty() || !to_string(cc_f_t[i]).empty() || !to_string(cc_f_si[i]).empty()) {
      pugi::xml_node f = cell.append_child("f");
      if (!to_string(cc_f_t[i]).empty()) {
        f.append_attribute("t") = to_string(cc_f_t[i]).c_str();
      }
      if (!to_string(cc_f_ref[i]).empty()) {
        f.append_attribute("ref") = to_string(cc_f_ref[i]).c_str();
      }
      if (!to_string(cc_f_ca[i]).empty()) {
        f.append_attribute("ca") = to_string(cc_f_ca[i]).c_str();
      }
      if (!to_string(cc_f_si[i]).empty()) {
        f.append_attribute("si") = to_string(cc_f_si[i]).c_str();
        f_si = true;
      }

      f.append_child(pugi::node_pcdata).set_value(to_string(cc_f[i]).c_str());
    }

    // v node: value stored from evaluated formula
    if (!to_string(cc_v[i]).empty()) {
      if (!f_si & (to_string(cc_v[i]).compare(xml_preserver.c_str()) == 0)) {
        cell.append_child("v").append_attribute("xml:space").set_value("preserve");
        cell.child("v").append_child(pugi::node_pcdata).set_value(" ");
      } else {
        cell.append_child("v").append_child(pugi::node_pcdata).set_value(to_string(cc_v[i]).c_str());
      }
    }

    // <is><t> ... </t></is>
    if (to_string(cc_c_t[i]).compare("inlineStr") == 0) {
      if (!to_string(cc_is[i]).empty()) {

        pugi::xml_document is_node;
        pugi::xml_parse_result result = is_node.load_string(to_string(cc_is[i]).c_str(), pugi_parse_flags);
        if (!result) Rcpp::stop("loading inlineStr node while writing failed");

        cell.append_copy(is_node.first_child());
      }
    }

    // update lastrow
    lastrow = thisrow;
  }

  return doc;
}


// TODO: convert to pugi
// function that creates the xml worksheet
// uses preparated data and writes it. It passes data to set_row() which will
// create single xml rows of sheet_data.
//
// [[Rcpp::export]]
XPtrXML write_worksheet(
    std::string prior,
    std::string post,
    Rcpp::Environment sheet_data
) {

  unsigned int pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_ws_pcdata | pugi::parse_eol;


  // sheet_data will be in order, just need to check for row_heights
  // CharacterVector cell_col = int_to_col(sheet_data.field("cols"));
  Rcpp::DataFrame row_attr = Rcpp::as<Rcpp::DataFrame>(sheet_data["row_attr"]);
  Rcpp::DataFrame cc = Rcpp::as<Rcpp::DataFrame>(sheet_data["cc_out"]);


  xmldoc *doc = new xmldoc;
  pugi::xml_parse_result result;

  pugi::xml_document xml_pr;
  result = xml_pr.load_string(prior.c_str(), pugi_parse_flags);
  if (!result) Rcpp::stop("loading prior while writing failed");
  pugi::xml_node worksheet = doc->append_copy(xml_pr.child("worksheet"));

  pugi::xml_node sheetData = worksheet.append_child("sheetData");

  if (cc.size() > 0) {
    pugi::xml_document xml_sd;
    xml_sd = xml_sheet_data(row_attr, cc);
    for (auto sd : xml_sd.children())
      sheetData.append_copy(sd);
  }

  if (!post.empty()) {
    pugi::xml_document xml_po;
    result = xml_po.load_string(post.c_str(), pugi_parse_flags);
    if (!result) Rcpp::stop("loading post while writing failed");
    for (auto po : xml_po.children())
      worksheet.append_copy(po);
  }


  // doc->load_string(post.c_str());

  pugi::xml_node decl = doc->prepend_child(pugi::node_declaration);
  decl.append_attribute("version") = "1.0";
  decl.append_attribute("encoding") = "UTF-8";
  decl.append_attribute("standalone") = "yes";

  XPtrXML ptr(doc, true);
  ptr.attr("class") = Rcpp::CharacterVector::create("pugi_xml");
  // ptr.attr("escapes") = escapes;
  // ptr.attr("is_utf8") = utf8;

  return ptr;
}

// [[Rcpp::export]]
void write_xmlPtr(
    XPtrXML doc,
    std::string fl
) {
  unsigned int pugi_format_flags = pugi::format_raw | pugi::format_no_escapes;
  doc->save_file(fl.c_str(), "", pugi_format_flags, pugi::encoding_utf8);
}
