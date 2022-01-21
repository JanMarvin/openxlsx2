#include "openxlsx2.h"


// creates an xml row
// data in xml is ordered row wise. therefore we need the row attributes and
// the colum data used in this row. This function uses both to create a single
// row and passes it to write_worksheet_xml_2 which will create the entire
// sheet_data part for this worksheet
//
// [[Rcpp::export]]
std::string set_row(Rcpp::DataFrame row_attr, Rcpp::List cells, size_t row_idx) {

  pugi::xml_document doc;

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

  // Rf_PrintValue(attrnams);

  for (auto i = 0; i < cells.length(); ++i) {

    // create node <c>
    pugi::xml_node cell = row.append_child("c");

    Rcpp::List cll = cells[i];
    // Rf_PrintValue(cll);

    Rcpp::List cell_atr, cell_val, cell_isval;
    std::vector<std::string> cell_val_names;


    // Every cell consists of a typ and a val list. Certain functions have an
    // additional attr list.
    std::string c_typ        = Rcpp::as<std::string>(cll["c_t"]);
    std::string c_sty        = Rcpp::as<std::string>(cll["c_s"]);
    std::string c_rnm        = Rcpp::as<std::string>(cll["r"]);

    std::string c_val        = Rcpp::as<std::string>(cll["v"]);

    // Rf_PrintValue(cell_atr);
    // Rf_PrintValue(cell_val);
    // Rf_PrintValue(attr_val);

    // append attributes <c r="A1" ...>
    cell.append_attribute("r") = c_rnm.c_str();

    if (c_sty.compare(rnastring.c_str()) != 0)
      cell.append_attribute("s") = c_sty.c_str();

    // assign type if not <v> aka numeric
    if (c_typ.compare(rnastring.c_str()) != 0)
      cell.append_attribute("t") = c_typ.c_str();

    // append nodes <c r="A1" ...><v>...</v></c>

    // Rcpp::Rcout << c_val << std::endl;

    // <f> ... </f>
    if(c_typ.compare("e") == 0 || c_typ.compare("str") == 0) {

      std::string fml = Rcpp::as<std::string>(cll["f"]);
      std::string fml_type = Rcpp::as<std::string>(cll["f_t"]);
      std::string fml_si = Rcpp::as<std::string>(cll["f_si"]);
      std::string fml_ref = Rcpp::as<std::string>(cll["f_ref"]);

      // f node: formula to be evaluated
      if (fml.compare(rnastring.c_str()) != 0) {
        pugi::xml_node f = cell.append_child("f");
        if (fml_type.compare(rnastring.c_str()) != 0) {
          f.append_attribute("t") = fml_type.c_str();
        }
        if (fml_si.compare(rnastring.c_str()) != 0) {
          f.append_attribute("ref") = fml_ref.c_str();
        }
        if (fml_si.compare(rnastring.c_str()) != 0) {
          f.append_attribute("si") = fml_si.c_str();
        }

        f.append_child(pugi::node_pcdata).set_value(fml.c_str());
      }

      // v node: value stored from evaluated formula
      if (c_val.compare(rnastring.c_str()) != 0)
        cell.append_child("v").append_child(pugi::node_pcdata).set_value(c_val.c_str());

    }

    // <is><t> ... </t></is>
    if(c_typ.compare("inlineStr") == 0) {
      std::string c_ist = Rcpp::as<std::string>(cll["is"]);
      if (c_ist.compare(rnastring.c_str()) != 0) {

        pugi::xml_document is_node;
        pugi::xml_parse_result result = is_node.load_string(c_ist.c_str(), pugi::parse_default | pugi::parse_escapes);
        if (!result) Rcpp::stop("loading inlineStr node while writing failed");

        cell.append_copy(is_node.first_child());
      }
    }

    // <v> ... </v>
    if(c_typ.compare("b") == 0) {
      cell.append_child("v").append_child(pugi::node_pcdata).set_value(c_val.c_str());
    }

    // <v> ... </v>
    if(c_typ.compare("s") == 0) {
      cell.append_child("v").append_child(pugi::node_pcdata).set_value(c_val.c_str());
    }

    // <v> ... </v>
    if(c_typ.compare("n") == 0) { // random numeric type, shall we simply treat all non strings as numeric?
      cell.append_child("v").append_child(pugi::node_pcdata).set_value(c_val.c_str());
    }

    // <v> ... </v>
    if(c_typ.compare(rnastring.c_str()) == 0) {
      if (c_val.compare(rnastring.c_str()) != 0) // dont write defined missings (NA might be to generic for )
        cell.append_child("v").append_child(pugi::node_pcdata).set_value(c_val.c_str());
    }

  }

  std::ostringstream oss;
  doc.print(oss, " ", pugi::format_raw);
  // doc.print(oss);

  return oss.str();
}

// function that creates the xml worksheet
// uses preparated data and writes it. It passes data to set_row() which will
// create single xml rows of sheet_data.
//
// [[Rcpp::export]]
SEXP write_worksheet_xml_2( std::string prior,
                            std::string post,
                            Rcpp::Reference sheet_data,
                            Rcpp::CharacterVector cols_attr, // currently unused
                            std::string R_fileName = "output"){


  // open file and write header XML
  const char * s = R_fileName.c_str();
  std::ofstream xmlFile;
  xmlFile.open (s);
  xmlFile << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
  xmlFile << prior;


  // sheet_data will be in order, just need to check for row_heights
  // CharacterVector cell_col = int_2_cell_ref(sheet_data.field("cols"));
  Rcpp::DataFrame row_attr = Rcpp::as<Rcpp::DataFrame>(sheet_data.field("row_attr"));
  Rcpp::List cc = sheet_data.field("cc_out");

  if (Rf_isNull(row_attr)) {
    xmlFile << "<sheetData />";
  } else {

    xmlFile << "<sheetData>";

    for (int i = 0; i < row_attr.nrow(); ++i) {

      xmlFile << set_row(row_attr, cc[i], i);

    }

    // write closing tag and XML post data
    xmlFile << "</sheetData>";

  }
  xmlFile << post;

  //close file
  xmlFile.close();

  return Rcpp::wrap(0);

}
