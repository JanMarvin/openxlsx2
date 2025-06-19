#include <fstream>
#include "openxlsx2.h"

// write xml by streaming to files. this takes whatever input we provide and
// dumps it into the file. no xml checking, no unicode checking
void xml_sheet_data_slim(
    Rcpp::DataFrame row_attr,
    Rcpp::DataFrame cc,
    std::string prior,
    std::string post,
    std::string fl
) {

  bool has_cm = cc.containsElementNamed("c_cm");
  bool has_ph = cc.containsElementNamed("c_ph");
  bool has_vm = cc.containsElementNamed("c_vm");

  std::ofstream file(fl);
  if (!file.is_open()) Rcpp::stop("could not open output file");

  auto lastrow = 0;  // integer value of the last row with column data
  auto thisrow = 0;  // integer value of the current row with column data
  auto row_idx = 0;  // the index of the row_attr file. this is != rowid
  auto rowid   = 0;  // integer value of the r field in row_attr

  std::string xml_preserver = " ";

  file << "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
  file << prior;

  Rcpp::CharacterVector cc_c_cm, cc_c_ph, cc_c_vm;

  if (cc.nrow() && cc.ncol()) {
    // we cannot access rows directly in the dataframe.
    // Have to extract the columns and use these
    Rcpp::CharacterVector cc_row_r  = cc["row_r"];  // 1
    Rcpp::CharacterVector cc_r      = cc["r"];      // A1
    Rcpp::CharacterVector cc_v      = cc["v"];
    Rcpp::CharacterVector cc_c_t    = cc["c_t"];
    Rcpp::CharacterVector cc_c_s    = cc["c_s"];
    if (has_cm) cc_c_cm  = cc["c_cm"];
    if (has_ph) cc_c_ph  = cc["c_ph"];
    if (has_vm) cc_c_vm  = cc["c_vm"];
    Rcpp::CharacterVector cc_f      = cc["f"];
    Rcpp::CharacterVector cc_f_attr = cc["f_attr"];
    Rcpp::CharacterVector cc_is     = cc["is"];

    Rcpp::CharacterVector row_r     = row_attr["r"];
    Rcpp::CharacterVector attrnams  = row_attr.names();

    file << "<sheetData>";
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

          if (row_idx) file << "</row>";
          file << "<row";

          for (auto j = 0; j < row_attr.ncol(); ++j) {
            Rcpp::CharacterVector cv_s = "";
            cv_s = Rcpp::as<Rcpp::CharacterVector>(row_attr[j])[row_idx];

            if (cv_s[0] != "") {
              const std::string val_strl = Rcpp::as<std::string>(cv_s);
              file << " " << attrnams[j] << "=\"" << val_strl.c_str() << "\"";
            }
          }
          file << ">";  // end <r ...>

          // read the next row_idx when visiting again
          ++row_idx;
        }
      }

      if (
        cc_c_s[i].empty() &&
        cc_c_t[i].empty() &&
        (!has_cm || (has_cm && cc_c_cm[i].empty())) &&
        (!has_ph || (has_ph && cc_c_ph[i].empty())) &&
        (!has_vm || (has_vm && cc_c_vm[i].empty())) &&
        cc_v[i].empty() &&
        cc_f[i].empty() &&
        cc_f_attr[i].empty() &&
        cc_is[i].empty()
      ) {
        continue;
      }

      // create node <c>
      file << "<c";

      // Every cell consists of a typ and a val list. Certain functions have an
      // additional attr list.

      // append attributes <c r="A1" ...>
      file << " r" << "=\"" << std::string(cc_r[i]).c_str() << "\"";

      if (!cc_c_s[i].empty())
        file << " s" << "=\"" << std::string(cc_c_s[i]).c_str() << "\"";

      // assign type if not <v> aka numeric
      if (!cc_c_t[i].empty())
        file << " t" << "=\"" << std::string(cc_c_t[i]).c_str() << "\"";

      // CellMetaIndex: suppress curly brackets in spreadsheet software
      if (has_cm && !cc_c_cm[i].empty())
        file << " cm" << "=\"" << std::string(cc_c_cm[i]).c_str() << "\"";

      // phonetics spelling
      if (has_ph && !cc_c_ph[i].empty())
        file << " ph" << "=\"" << std::string(cc_c_ph[i]).c_str() << "\"";

      // suppress curly brackets in spreadsheet software
      if (has_vm && !cc_c_vm[i].empty())
        file << " vm" << "=\"" << std::string(cc_c_vm[i]).c_str() << "\"";

      file << ">";  // end <c ...>

      bool f_si = false;

      // <f> ... </f>
      // f node: formula to be evaluated
      if (!cc_f[i].empty() || !cc_f_attr[i].empty()) {
        file << "<f";
        if (!cc_f_attr[i].empty()) {
          file << " " << std::string(cc_f_attr[i]).c_str();
        }
        file << ">";

        file << to_string(cc_f[i]).c_str();
        if (!f_si && std::string(cc_f[i]).find("\"si\"=") != std::string::npos) f_si = true;

        file << "</f>";
      }

      // v node: value stored from evaluated formula
      if (!cc_v[i].empty()) {
        if (!f_si & (std::string(cc_v[i]).compare(xml_preserver.c_str()) == 0)) {
          // this looks strange
          file << "<v xml:space=\"preserve\">";
          file << " ";
          file << "</v>";
        } else {
          if (cc_c_t[i].empty() && cc_f_attr[i].empty())
            file << "<v>" << to_string(cc_v[i]).c_str() << "</v>";
          else
            file << "<v>" << std::string(cc_v[i]).c_str() << "</v>";
        }
      }

      // <is><t> ... </t></is>
      if (std::string(cc_c_t[i]).compare("inlineStr") == 0) {
        if (!cc_is[i].empty()) {
          file << to_string(cc_is[i]).c_str();
        }
      }

      file << "</c>";

      // update lastrow
      lastrow = thisrow;
    }

    file << "</row>";
    file << "</sheetData>";
  } else {
    file << "<sheetData/>";
  }

  file << post;
  file << "</worksheet>";

  file.close();
}

// export worksheet without pugixml
// this should be way quicker, uses far less memory, but also skips all of the checks pugi does
//
// [[Rcpp::export]]
void write_worksheet_slim(
    Rcpp::Environment sheet_data,
    std::string prior,
    std::string post,
    std::string fl
){
  // sheet_data will be in order, just need to check for row_heights
  // CharacterVector cell_col = int_to_col(sheet_data.field("cols"));
  Rcpp::DataFrame row_attr = Rcpp::as<Rcpp::DataFrame>(sheet_data["row_attr"]);
  Rcpp::DataFrame cc = Rcpp::as<Rcpp::DataFrame>(sheet_data["cc"]);

  xml_sheet_data_slim(row_attr, cc, prior, post, fl);
}

// creates an xml row
// data in xml is ordered row wise. therefore we need the row attributes and
// the column data used in this row. This function uses both to create a single
// row and passes it to write_worksheet_xml_2 which will create the entire
// sheet_data part for this worksheet

void xml_sheet_data(pugi::xml_node &doc, Rcpp::DataFrame &row_attr, Rcpp::DataFrame &cc) {
  auto lastrow = 0;  // integer value of the last row with column data
  auto thisrow = 0;  // integer value of the current row with column data
  auto row_idx = 0;  // the index of the row_attr file. this is != rowid
  auto rowid   = 0;  // integer value of the r field in row_attr

  pugi::xml_node row;

  std::string xml_preserver = " ";

  // non optional: treat input as valid at this stage
  uint32_t pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_ws_pcdata | pugi::parse_eol;
  // uint32_t pugi_format_flags = pugi::format_raw | pugi::format_no_escapes;

  bool has_cm = cc.containsElementNamed("c_cm");
  bool has_ph = cc.containsElementNamed("c_ph");
  bool has_vm = cc.containsElementNamed("c_vm");

  Rcpp::CharacterVector cc_c_cm, cc_c_ph, cc_c_vm;

  // we cannot access rows directly in the dataframe.
  // Have to extract the columns and use these
  Rcpp::CharacterVector cc_row_r  = cc["row_r"];  // 1
  Rcpp::CharacterVector cc_r      = cc["r"];      // A1
  Rcpp::CharacterVector cc_v      = cc["v"];      // can be utf8
  Rcpp::CharacterVector cc_c_t    = cc["c_t"];
  Rcpp::CharacterVector cc_c_s    = cc["c_s"];
  if (has_cm) cc_c_cm   = cc["c_cm"];
  if (has_ph) cc_c_ph   = cc["c_ph"];
  if (has_vm) cc_c_vm   = cc["c_vm"];
  Rcpp::CharacterVector cc_f      = cc["f"];      // can be utf8
  Rcpp::CharacterVector cc_f_attr = cc["f_attr"];
  Rcpp::CharacterVector cc_is     = cc["is"];     // can be utf8

  Rcpp::CharacterVector row_r    = row_attr["r"];
  Rcpp::CharacterVector attrnams = row_attr.names();

  for (auto i = 0; i < cc.nrow(); ++i) {
    thisrow = std::stoi(Rcpp::as<std::string>(cc_row_r[i]));

    if (lastrow < thisrow) {
      // there might be entirely empty rows in between. this is the case for
      // loadExample. We check the rowid and write the line and skip until we
      // have every row and only then continue writing the column
      while (rowid < thisrow) {
        rowid = std::stoi(Rcpp::as<std::string>(row_r[row_idx]));

        row = doc.append_child("row");

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

    // update lastrow
    lastrow = thisrow;

    if (
      cc_c_s[i].empty() &&
      cc_c_t[i].empty() &&
      (!has_cm || (has_cm && cc_c_cm[i].empty())) &&
      (!has_ph || (has_ph && cc_c_ph[i].empty())) &&
      (!has_vm || (has_vm && cc_c_vm[i].empty())) &&
      cc_v[i].empty() &&
      cc_f[i].empty() &&
      cc_f_attr[i].empty() &&
      cc_is[i].empty()
    ) {
      continue;
    }

    // create node <c>
    pugi::xml_node cell = row.append_child("c");

    // Every cell consists of a typ and a val list. Certain functions have an
    // additional attr list.

    // append attributes <c r="A1" ...>
    cell.append_attribute("r") = std::string(cc_r[i]).c_str();

    if (!cc_c_s[i].empty())
      cell.append_attribute("s") = std::string(cc_c_s[i]).c_str();

    // assign type if not <v> aka numeric
    if (!cc_c_t[i].empty())
      cell.append_attribute("t") = std::string(cc_c_t[i]).c_str();

    // CellMetaIndex: suppress curly brackets in spreadsheet software
    if (has_cm && !cc_c_cm[i].empty())
      cell.append_attribute("cm") = std::string(cc_c_cm[i]).c_str();

    // phonetics spelling
    if (has_ph && !cc_c_ph[i].empty())
      cell.append_attribute("ph") = std::string(cc_c_ph[i]).c_str();

    // suppress curly brackets in spreadsheet software
    if (has_vm && !cc_c_vm[i].empty())
      cell.append_attribute("vm") = std::string(cc_c_vm[i]).c_str();

    // append nodes <c r="A1" ...><v>...</v></c>

    bool f_si = false;

    // <f> ... </f>
    // f node: formula to be evaluated
    if (!cc_f[i].empty() || !cc_f_attr[i].empty()) {
      // Fix Most Vexing Parse
      std::istringstream attr_stream((std::string(cc_f_attr[i])));
      pugi::xml_node f = cell.append_child("f");

      // Parse attributes from f_attr
      std::string key_value;
      while (std::getline(attr_stream, key_value, ' ')) {
          auto pos = key_value.find('=');
          if (pos != std::string::npos) {
              std::string key = key_value.substr(0, pos);
              std::string value = key_value.substr(pos + 1);

              // Remove quotes from value
              if (!value.empty() && value.front() == '\"' && value.back() == '\"') {
                  value = value.substr(1, value.size() - 2);
              }

              // Append attribute to <f>
              f.append_attribute(key.c_str()) = value.c_str();
              if (key == "si") f_si = true;
          }
      }

      // Add the content of <f>
      if (!cc_f[i].empty()) {
          f.append_child(pugi::node_pcdata).set_value(to_string(cc_f[i]).c_str());
      }
    }

    // v node: value stored from evaluated formula
    if (!cc_v[i].empty()) {
      if (!f_si & (std::string(cc_v[i]).compare(xml_preserver.c_str()) == 0)) {
        cell.append_child("v").append_attribute("xml:space").set_value("preserve");
        cell.child("v").append_child(pugi::node_pcdata).set_value(" ");
      } else {
        if (cc_c_t[i].empty() && cc_f_attr[i].empty())
          cell.append_child("v").append_child(pugi::node_pcdata).set_value(std::string(cc_v[i]).c_str());
        else
          cell.append_child("v").append_child(pugi::node_pcdata).set_value(to_string(cc_v[i]).c_str());
      }
    }

    // <is><t> ... </t></is>
    if (std::string(cc_c_t[i]).compare("inlineStr") == 0) {
      if (!cc_is[i].empty()) {
        pugi::xml_document is_node;
        pugi::xml_parse_result result = is_node.load_string(to_string(cc_is[i]).c_str(), pugi_parse_flags);
        if (!result) Rcpp::stop("loading inlineStr node while writing failed");

        cell.append_copy(is_node.first_child());
      }
    }
  }
}

// TODO: convert to pugi
// function that creates the xml worksheet
// uses preparated data and writes it. It passes data to set_row() which will
// create single xml rows of sheet_data.
//
// [[Rcpp::export]]
XPtrXML write_worksheet(std::string prior, std::string post, Rcpp::Environment sheet_data) {
  uint32_t pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_ws_pcdata | pugi::parse_eol;

  // sheet_data will be in order, just need to check for row_heights
  // CharacterVector cell_col = int_to_col(sheet_data.field("cols"));
  Rcpp::DataFrame row_attr = Rcpp::as<Rcpp::DataFrame>(sheet_data["row_attr"]);
  Rcpp::DataFrame cc = Rcpp::as<Rcpp::DataFrame>(sheet_data["cc"]);

  xmldoc* doc = new xmldoc;
  pugi::xml_parse_result result;

  pugi::xml_document xml_pr;
  result = xml_pr.load_string(prior.c_str(), pugi_parse_flags);
  if (!result) Rcpp::stop("loading prior while writing failed");
  pugi::xml_node worksheet = doc->append_copy(xml_pr.child("worksheet"));

  pugi::xml_node sheetData = worksheet.append_child("sheetData");

  if (cc.size() > 0) {
    xml_sheet_data(sheetData, row_attr, cc);
  }

  if (!post.empty()) {
    pugi::xml_document xml_po;
    result = xml_po.load_string(post.c_str(), pugi_parse_flags);
    if (!result) Rcpp::stop("loading post while writing failed");
    for (auto po : xml_po.children())
      worksheet.append_copy(po);
  }

  pugi::xml_node decl = doc->prepend_child(pugi::node_declaration);
  decl.append_attribute("version") = "1.0";
  decl.append_attribute("encoding") = "UTF-8";
  decl.append_attribute("standalone") = "yes";

  XPtrXML ptr(doc, true);
  ptr.attr("class") = Rcpp::CharacterVector::create("pugi_xml");

  return ptr;
}

// [[Rcpp::export]]
void write_xmlPtr(XPtrXML doc, std::string fl) {
  uint32_t pugi_format_flags = pugi::format_raw | pugi::format_no_escapes;
  const bool success = doc->save_file(fl.c_str(), "", pugi_format_flags, pugi::encoding_utf8);
  if (!success) Rcpp::stop("could not save file");
}
