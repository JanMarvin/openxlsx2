#include "openxlsx2.h"


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


// creates an xml row
// data in xml is ordered row wise. therefore we need the row attributes and
// the column data used in this row. This function uses both to create a single
// row and passes it to write_worksheet_xml_2 which will create the entire
// sheet_data part for this worksheet
std::string set_row(Rcpp::DataFrame row_attr, std::vector<xml_col> cells, size_t row_idx) {

  pugi::xml_document doc;

  // non optional: treat input as valid at this stage
  unsigned int pugi_parse_flags = pugi::parse_cdata | pugi::parse_wconv_attribute | pugi::parse_eol;
  unsigned int pugi_format_flags = pugi::format_raw | pugi::format_no_escapes;

  pugi::xml_node row = doc.append_child("row");
  Rcpp::CharacterVector attrnams = row_attr.names();

  for (auto j = 0; j < row_attr.ncol(); ++j) {

    Rcpp::CharacterVector cv_s = "";
    cv_s = Rcpp::as<Rcpp::CharacterVector>(row_attr[j])[row_idx];

    if (cv_s[0] != "") {
      const std::string val_strl = Rcpp::as<std::string>(cv_s);
      row.append_attribute(attrnams[j]) = val_strl.c_str();
    }
  }

  std::string rnastring = "_openxlsx_NA_";
  std::string xml_preserver = "";

  auto cells_size = cells.size();
  for (auto i = 0; i < cells_size; ++i) {

    // create node <c>
    pugi::xml_node cell = row.append_child("c");

    xml_col cll = cells[i];

    std::vector<std::string> cell_val_names;


    // Every cell consists of a typ and a val list. Certain functions have an
    // additional attr list.
    // std::string c_typ = cells.c_t;
    // std::string c_sty = cells.c_s;
    // std::string c_rnm = cells.row_r;
    // std::string c_val = cells.v.;

    // append attributes <c r="A1" ...>
    cell.append_attribute("r") = cll.row_r.c_str();

    if (cll.c_s.compare(rnastring.c_str()) != 0)
      cell.append_attribute("s") = cll.c_s.c_str();

    // assign type if not <v> aka numeric
    if (cll.c_t.compare(rnastring.c_str()) != 0)
      cell.append_attribute("t") = cll.c_t.c_str();

    // append nodes <c r="A1" ...><v>...</v></c>

    // Rcpp::Rcout << c_val << std::endl;
    bool f_si = false;

    // <f> ... </f>

    // std::string fml = Rcpp::as<std::string>(cll["f"]);
    // std::string fml_type = Rcpp::as<std::string>(cll["f_t"]);
    // std::string fml_si = Rcpp::as<std::string>(cll["f_si"]);
    // std::string fml_ref = Rcpp::as<std::string>(cll["f_ref"]);

    // f node: formula to be evaluated
    if (cll.f.compare(rnastring.c_str()) != 0 ||
        cll.f_t.compare(rnastring.c_str()) != 0 ||
        cll.f_si.compare(rnastring.c_str()) != 0) {
      pugi::xml_node f = cell.append_child("f");
      if (cll.f_t.compare(rnastring.c_str()) != 0) {
        f.append_attribute("t") = cll.f_t.c_str();
      }
      if (cll.f_ref.compare(rnastring.c_str()) != 0) {
        f.append_attribute("ref") = cll.f_ref.c_str();
      }
      if (cll.f_si.compare(rnastring.c_str()) != 0) {
        f.append_attribute("si") = cll.f_si.c_str();
        f_si = true;
      }

      f.append_child(pugi::node_pcdata).set_value(cll.f.c_str());
    }

    // v node: value stored from evaluated formula
    if (cll.v.compare(rnastring.c_str()) != 0) {
      if (!f_si & (cll.v.compare(xml_preserver.c_str()) == 0)) {
        cell.append_child("v").append_attribute("xml:space").set_value("preserve");
        cell.child("v").append_child(pugi::node_pcdata).set_value(" ");
      } else {
        cell.append_child("v").append_child(pugi::node_pcdata).set_value(cll.v.c_str());
      }
    }


    // <is><t> ... </t></is>
    if(cll.c_t.compare("inlineStr") == 0) {
      if (cll.is.compare(rnastring.c_str()) != 0) {

        pugi::xml_document is_node;
        pugi::xml_parse_result result = is_node.load_string(cll.is.c_str(), pugi_parse_flags);
        if (!result) Rcpp::stop("loading inlineStr node while writing failed");

        cell.append_copy(is_node.first_child());
      }
    }

  }

  std::ostringstream oss;
  doc.print(oss, " ", pugi::format_raw | pugi::format_no_escapes);
  // doc.print(oss);

  return oss.str();
}

// TODO: convert to pugi
// function that creates the xml worksheet
// uses preparated data and writes it. It passes data to set_row() which will
// create single xml rows of sheet_data.
//
// [[Rcpp::export]]
SEXP write_worksheet_xml_2( std::string prior,
                            std::string post,
                            Rcpp::Environment sheet_data,
                            Rcpp::CharacterVector cols_attr, // currently unused
                            std::string R_fileName = "output"){


  // open file and write header XML
  const char * s = R_fileName.c_str();
  std::ofstream xmlFile;
  xmlFile.open (s);
  xmlFile << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
  xmlFile << prior;


  // sheet_data will be in order, just need to check for row_heights
  // CharacterVector cell_col = int_to_col(sheet_data.field("cols"));
  Rcpp::DataFrame row_attr = sheet_data["row_attr"];
  Rcpp::DataFrame cc = sheet_data["cc_out"];

  // TODO prev. this was Rf_isNull() no we have a zero col, zero row dataframe?
  if ((row_attr.nrow() == 0) || (cc.nrow() == 0)) {
    xmlFile << "<sheetData />";
  } else {

    xmlFile << "<sheetData>";

    // we cannot access rows directly. Have to extract the columns and use these
    Rcpp::CharacterVector cc_row_r = cc["row_r"]; // 1
    Rcpp::CharacterVector cc_r     = cc["r"];     // A1
    Rcpp::CharacterVector cc_v     = cc["v"];
    Rcpp::CharacterVector cc_c_t   = cc["c_t"];
    Rcpp::CharacterVector cc_c_s   = cc["c_s"];
    Rcpp::CharacterVector cc_f     = cc["f"];
    Rcpp::CharacterVector cc_f_t   = cc["f_t"];
    Rcpp::CharacterVector cc_f_ref = cc["f_ref"];
    Rcpp::CharacterVector cc_f_si  = cc["f_si"];
    Rcpp::CharacterVector cc_is    = cc["is"];

    for (int i = 0; i < row_attr.nrow(); ++i) {

      Rcpp::CharacterVector ii = std::to_string(i + 1);
      // which
      Rcpp::IntegerVector sel_int = match(cc_row_r, ii);
      Rcpp::IntegerVector sel = Rcpp::seq(0, sel_int.size() - 1);
      sel = sel[!Rcpp::is_na(sel_int)];
      // Rcpp::LogicalVector sel = !Rcpp::is_na(sel_int);

      Rcpp::CharacterVector cc_i_row_r = cc_row_r[sel]; // 1
      Rcpp::CharacterVector cc_i_r     = cc_r[sel];     // A1
      Rcpp::CharacterVector cc_i_v     = cc_v[sel];
      Rcpp::CharacterVector cc_i_c_t   = cc_c_t[sel];
      Rcpp::CharacterVector cc_i_c_s   = cc_c_s[sel];
      Rcpp::CharacterVector cc_i_f     = cc_f[sel];
      Rcpp::CharacterVector cc_i_r_t   = cc_f_t[sel];
      Rcpp::CharacterVector cc_i_f_ref = cc_f_ref[sel];
      Rcpp::CharacterVector cc_i_f_si  = cc_f_si[sel];
      Rcpp::CharacterVector cc_i_is    = cc_is[sel];

      // Rf_PrintValue(sel);
      // Rf_PrintValue(cc_i_row_r);

      std::vector<xml_col> cell(cc_i_r.size());
      for (auto j = 0; j < cc_i_r.size(); ++ j) {
        // Rcpp::Rcout << j << std::endl;

        cell[j].row_r = cc_i_r[j];
        cell[j].v     = cc_i_v[j];
        cell[j].c_t   = cc_i_c_t[j];
        cell[j].c_s   = cc_i_c_s[j];
        cell[j].f     = cc_i_f[j];
        cell[j].f_t   = cc_i_r_t[j];
        cell[j].f_ref = cc_i_f_ref[j];
        cell[j].f_si  = cc_i_f_si[j];
        cell[j].is    = cc_i_is[j];
      }

      xmlFile << set_row(row_attr, cell, i);

    }

    // write closing tag and XML post data
    xmlFile << "</sheetData>";

  }
  xmlFile << post;

  //close file
  xmlFile.close();

  return Rcpp::wrap(0);

}
