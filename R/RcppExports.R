# Generated by using Rcpp::compileAttributes() -> do not edit by hand
# Generator token: 10BE3573-1514-4C36-9D1C-5A225CD40393

date_to_unix <- function(x, origin = "1900-01-01", datetime = TRUE) {
    .Call(`_openxlsx2_date_to_unix`, x, origin, datetime)
}

#' Check if path is to long to be an R file path
#' @param path the file path used in file.exists()
#' @noRd
to_long <- function(path) {
    .Call(`_openxlsx2_to_long`, path)
}

openxlsx2_type <- function(x) {
    .Call(`_openxlsx2_openxlsx2_type`, x)
}

col_to_int <- function(x) {
    .Call(`_openxlsx2_col_to_int`, x)
}

ox_int_to_col <- function(x) {
    .Call(`_openxlsx2_ox_int_to_col`, x)
}

rbindlist <- function(x) {
    .Call(`_openxlsx2_rbindlist`, x)
}

copy <- function(x) {
    .Call(`_openxlsx2_copy`, x)
}

needed_cells <- function(range, all = TRUE) {
    .Call(`_openxlsx2_needed_cells`, range, all)
}

#' check if non consecutive dims is equal sized: "A1:A4,B1:B4"
#' @param dims dims
#' @param check check if all the same size
#' @keywords internal
#' @noRd
get_dims <- function(dims, check = FALSE) {
    .Call(`_openxlsx2_get_dims`, dims, check)
}

dims_to_row_col_fill <- function(dims, fills = FALSE) {
    .Call(`_openxlsx2_dims_to_row_col_fill`, dims, fills)
}

dims_to_df <- function(rows, cols, filled, fill, fcols) {
    .Call(`_openxlsx2_dims_to_df`, rows, cols, filled, fill, fcols)
}

long_to_wide <- function(z, tt, zz) {
    invisible(.Call(`_openxlsx2_long_to_wide`, z, tt, zz))
}

is_charnum <- function(x) {
    .Call(`_openxlsx2_is_charnum`, x)
}

wide_to_long <- function(z, vtyps, zz, ColNames, start_col, start_row, refed, string_nums, na_null, na_missing, na_strings, inline_strings, c_cm, dims) {
    invisible(.Call(`_openxlsx2_wide_to_long`, z, vtyps, zz, ColNames, start_col, start_row, refed, string_nums, na_null, na_missing, na_strings, inline_strings, c_cm, dims))
}

#' @param colnames a vector of the names of the data frame
#' @param n the length of the data frame
#' @noRd
create_char_dataframe <- function(colnames, n) {
    .Call(`_openxlsx2_create_char_dataframe`, colnames, n)
}

read_xml2df <- function(xml, vec_name, vec_attrs, vec_chlds) {
    .Call(`_openxlsx2_read_xml2df`, xml, vec_name, vec_attrs, vec_chlds)
}

write_df2xml <- function(df, vec_name, vec_attrs, vec_chlds) {
    .Call(`_openxlsx2_write_df2xml`, df, vec_name, vec_attrs, vec_chlds)
}

col_to_df <- function(doc) {
    .Call(`_openxlsx2_col_to_df`, doc)
}

df_to_xml <- function(name, df_col) {
    .Call(`_openxlsx2_df_to_xml`, name, df_col)
}

loadvals <- function(sheet_data, doc) {
    invisible(.Call(`_openxlsx2_loadvals`, sheet_data, doc))
}

readXML <- function(path, isfile, escapes, declaration, whitespace, empty_tags, skip_control, pointer) {
    .Call(`_openxlsx2_readXML`, path, isfile, escapes, declaration, whitespace, empty_tags, skip_control, pointer)
}

is_xml <- function(str) {
    .Call(`_openxlsx2_is_xml`, str)
}

getXMLXPtrNamePath <- function(doc, path) {
    .Call(`_openxlsx2_getXMLXPtrNamePath`, doc, path)
}

getXMLXPtrPath <- function(doc, path) {
    .Call(`_openxlsx2_getXMLXPtrPath`, doc, path)
}

getXMLXPtrValPath <- function(doc, path) {
    .Call(`_openxlsx2_getXMLXPtrValPath`, doc, path)
}

getXMLXPtrAttrPath <- function(doc, path) {
    .Call(`_openxlsx2_getXMLXPtrAttrPath`, doc, path)
}

printXPtr <- function(doc, indent, raw, attr_indent) {
    .Call(`_openxlsx2_printXPtr`, doc, indent, raw, attr_indent)
}

write_xml_file <- function(xml_content, escapes) {
    .Call(`_openxlsx2_write_xml_file`, xml_content, escapes)
}

#' adds or updates attribute(s) in existing xml node
#'
#' @description Needs xml node and named character vector as input. Modifies
#' the arguments of each first child found in the xml node and adds or updates
#' the attribute vector.
#' @details If a named attribute in `xml_attributes` is "" remove the attribute
#' from the node.
#' If `xml_attributes` contains a named entry found in the xml node, it is
#' updated else it is added as attribute.
#'
#' @param xml_content some valid xml_node
#' @param xml_attributes R vector of named attributes
#' @param escapes bool if escapes should be used
#' @param declaration bool if declaration should be imported
#' @param remove_empty_attr bool remove empty attributes or ignore them
#'
#' @examples
#'   # add single node
#'     xml_node <- "<a foo=\"bar\">openxlsx2</a><b />"
#'     xml_attr <- c(qux = "quux")
#'     # "<a foo=\"bar\" qux=\"quux\">openxlsx2</a><b qux=\"quux\"/>"
#'     xml_attr_mod(xml_node, xml_attr)
#'
#'   # update node and add node
#'     xml_node <- "<a foo=\"bar\">openxlsx2</a><b />"
#'     xml_attr <- c(foo = "baz", qux = "quux")
#'     # "<a foo=\"baz\" qux=\"quux\">openxlsx2</a><b foo=\"baz\" qux=\"quux\"/>"
#'     xml_attr_mod(xml_node, xml_attr)
#'
#'   # remove node and add node
#'     xml_node <- "<a foo=\"bar\">openxlsx2</a><b />"
#'     xml_attr <- c(foo = "", qux = "quux")
#'     # "<a qux=\"quux\">openxlsx2</a><b qux=\"quux\"/>"
#'     xml_attr_mod(xml_node, xml_attr)
#' @export
xml_attr_mod <- function(xml_content, xml_attributes, escapes = FALSE, declaration = FALSE, remove_empty_attr = TRUE) {
    .Call(`_openxlsx2_xml_attr_mod`, xml_content, xml_attributes, escapes, declaration, remove_empty_attr)
}

#' create xml_node from R objects
#' @description takes xml_name, xml_children and xml_attributes to create a new
#' xml_node.
#' @param xml_name the name of the new xml_node
#' @param xml_children character vector children attached to the xml_node
#' @param xml_attributes named character vector of attributes for the xml_node
#' @param escapes bool if escapes should be used
#' @param declaration bool if declaration should be imported
#' @details if xml_children or xml_attributes should be empty, use NULL
#'
#' @examples
#' xml_name <- "a"
#' # "<a/>"
#' xml_node_create(xml_name)
#'
#' xml_child <- "openxlsx"
#' # "<a>openxlsx</a>"
#' xml_node_create(xml_name, xml_children = xml_child)
#'
#' xml_attr <- c(foo = "baz", qux = "quux")
#' # "<a foo=\"baz\" qux=\"quux\"/>"
#' xml_node_create(xml_name, xml_attributes = xml_attr)
#'
#' # "<a foo=\"baz\" qux=\"quux\">openxlsx</a>"
#' xml_node_create(xml_name, xml_children = xml_child, xml_attributes = xml_attr)
#' @export
xml_node_create <- function(xml_name, xml_children = NULL, xml_attributes = NULL, escapes = FALSE, declaration = FALSE) {
    .Call(`_openxlsx2_xml_node_create`, xml_name, xml_children, xml_attributes, escapes, declaration)
}

xml_append_child_path <- function(node, child, path, pointer) {
    .Call(`_openxlsx2_xml_append_child_path`, node, child, path, pointer)
}

xml_remove_child_path <- function(node, child, path, which, pointer) {
    .Call(`_openxlsx2_xml_remove_child_path`, node, child, path, which, pointer)
}

is_to_txt <- function(is_vec) {
    .Call(`_openxlsx2_is_to_txt`, is_vec)
}

si_to_txt <- function(si_vec) {
    .Call(`_openxlsx2_si_to_txt`, si_vec)
}

txt_to_is <- function(text, no_escapes = FALSE, raw = TRUE, skip_control = TRUE) {
    .Call(`_openxlsx2_txt_to_is`, text, no_escapes, raw, skip_control)
}

txt_to_si <- function(text, no_escapes = FALSE, raw = TRUE, skip_control = TRUE) {
    .Call(`_openxlsx2_txt_to_si`, text, no_escapes, raw, skip_control)
}

read_xf <- function(xml_doc_xf) {
    .Call(`_openxlsx2_read_xf`, xml_doc_xf)
}

write_xf <- function(df_xf) {
    .Call(`_openxlsx2_write_xf`, df_xf)
}

read_font <- function(xml_doc_font) {
    .Call(`_openxlsx2_read_font`, xml_doc_font)
}

write_font <- function(df_font) {
    .Call(`_openxlsx2_write_font`, df_font)
}

read_numfmt <- function(xml_doc_numfmt) {
    .Call(`_openxlsx2_read_numfmt`, xml_doc_numfmt)
}

write_numfmt <- function(df_numfmt) {
    .Call(`_openxlsx2_write_numfmt`, df_numfmt)
}

read_border <- function(xml_doc_border) {
    .Call(`_openxlsx2_read_border`, xml_doc_border)
}

write_border <- function(df_border) {
    .Call(`_openxlsx2_write_border`, df_border)
}

read_fill <- function(xml_doc_fill) {
    .Call(`_openxlsx2_read_fill`, xml_doc_fill)
}

write_fill <- function(df_fill) {
    .Call(`_openxlsx2_write_fill`, df_fill)
}

read_cellStyle <- function(xml_doc_cellStyle) {
    .Call(`_openxlsx2_read_cellStyle`, xml_doc_cellStyle)
}

write_cellStyle <- function(df_cellstyle) {
    .Call(`_openxlsx2_write_cellStyle`, df_cellstyle)
}

read_tableStyle <- function(xml_doc_tableStyle) {
    .Call(`_openxlsx2_read_tableStyle`, xml_doc_tableStyle)
}

write_tableStyle <- function(df_tablestyle) {
    .Call(`_openxlsx2_write_tableStyle`, df_tablestyle)
}

read_dxf <- function(xml_doc_dxf) {
    .Call(`_openxlsx2_read_dxf`, xml_doc_dxf)
}

write_dxf <- function(df_dxf) {
    .Call(`_openxlsx2_write_dxf`, df_dxf)
}

read_colors <- function(xml_doc_colors) {
    .Call(`_openxlsx2_read_colors`, xml_doc_colors)
}

write_colors <- function(df_colors) {
    .Call(`_openxlsx2_write_colors`, df_colors)
}

write_worksheet_slim <- function(sheet_data, prior, post, fl) {
    invisible(.Call(`_openxlsx2_write_worksheet_slim`, sheet_data, prior, post, fl))
}

write_worksheet <- function(prior, post, sheet_data) {
    .Call(`_openxlsx2_write_worksheet`, prior, post, sheet_data)
}

write_xmlPtr <- function(doc, fl) {
    invisible(.Call(`_openxlsx2_write_xmlPtr`, doc, fl))
}

styles_bin <- function(filePath, outPath, debug) {
    .Call(`_openxlsx2_styles_bin`, filePath, outPath, debug)
}

table_bin <- function(filePath, outPath, debug) {
    .Call(`_openxlsx2_table_bin`, filePath, outPath, debug)
}

comments_bin <- function(filePath, outPath, debug) {
    .Call(`_openxlsx2_comments_bin`, filePath, outPath, debug)
}

externalreferences_bin <- function(filePath, outPath, debug) {
    .Call(`_openxlsx2_externalreferences_bin`, filePath, outPath, debug)
}

sharedstrings_bin <- function(filePath, outPath, debug) {
    .Call(`_openxlsx2_sharedstrings_bin`, filePath, outPath, debug)
}

workbook_bin <- function(filePath, outPath, debug) {
    .Call(`_openxlsx2_workbook_bin`, filePath, outPath, debug)
}

worksheet_bin <- function(filePath, chartsheet, outPath, debug) {
    .Call(`_openxlsx2_worksheet_bin`, filePath, chartsheet, outPath, debug)
}

