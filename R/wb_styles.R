
# functions interacting with styles


#' clone sheets style
#'
#' @param wb workbook
#' @param from_sheet sheet we select the style from
#' @param to_sheet sheet we apply the style from
#' @export
cloneSheetStyle <- function(wb, from_sheet, to_sheet) {

  # check if sheets exist in wb
  id_org <- wb$validateSheet(from_sheet)
  id_new <- wb$validateSheet(to_sheet)

  org_style <- wb$worksheets[[id_org]]$sheet_data$cc
  new_style <- wb$worksheets[[id_new]]$sheet_data$cc

  # remove all values
  org_style <- org_style[c("row_r", "c_r", "c_s")]

  merged_style <- merge(org_style, new_style, all = TRUE)
  merged_style[is.na(merged_style)] <- "_openxlsx_NA_"

  # TODO order this as well?
  wb$worksheets[[id_new]]$sheet_data$cc <- merged_style

  # copy entire attributes from original sheet to new sheet
  org_rows <- wb$worksheets[[id_org]]$sheet_data$row_attr
  new_rows <- wb$worksheets[[id_new]]$sheet_data$row_attr

  merged_rows <- merge(org_rows[c("r")], new_rows, all = TRUE)
  merged_rows[is.na(merged_rows)] <- ""
  ordr <- ordered(order(as.integer(merged_rows$r)))
  merged_rows <- merged_rows[ordr,]

  wb$worksheets[[id_new]]$sheet_data$row_attr <- merged_rows

  wb$worksheets[[id_new]]$cols_attr <-
    wb$worksheets[[id_org]]$cols_attr

  wb$worksheets[[id_new]]$dimension <-
    wb$worksheets[[id_org]]$dimension

  wb$worksheets[[id_new]]$mergeCells <-
    wb$worksheets[[id_org]]$mergeCells

}


# internal function used in loadWorkbook
# @param x character string containing styles.xml
import_styles <- function(x) {

  sxml <- openxlsx2::read_xml(x)

  z <- NULL

  # Borders borderId
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.border?view=openxml-2.8.1
  z$borders <- xml_node(sxml, "styleSheet", "borders", "border")

  # Cell Styles (no clue)
  # links
  # + xfId
  # + builtinId
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellstyle?view=openxml-2.8.1
  z$cellStyles <- xml_node(sxml, "styleSheet", "cellStyles", "cellStyle")

  # Formatting Records
  # links
  # + numFmtId
  # + fontId
  # + fillId
  # + borderId
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellstyleformats?view=openxml-2.8.1
  z$cellStyleXfs <- xml_node(sxml, "styleSheet", "cellStyleXfs", "xf")

  # Cell Formats
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformat?view=openxml-2.8.1
  #
  # Position is 0 index s="value" used in worksheet <c ...>
  # links
  # + numFmtId
  # + fontId
  # + fillId
  # + borderId
  # + xfId
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformats?view=openxml-2.8.1
  z$cellXfs <- xml_node(sxml, "styleSheet", "cellXfs", "xf")

  # Colors
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.colors?view=openxml-2.8.1
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.indexedcolors?view=openxml-2.8.1
  z$colors <- xml_node(sxml, "styleSheet", "colors")

  # Formats (No clue?)
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.differentialformat?view=openxml-2.8.1
  z$dxfs <- xml_node(sxml, "styleSheet", "dxfs", "dxf")

  # Future extensions No clue, some special styles
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.extensionlist?view=openxml-2.8.1
  z$extLst <- xml_node(sxml, "styleSheet", "extLst")

  # Fills fillId
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fill?view=openxml-2.8.1
  z$fills <- xml_node(sxml, "styleSheet", "fills", "fill")

  # Fonts fontId
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.font?view=openxml-2.8.1
  z$fonts <- xml_node(sxml, "styleSheet", "fonts", "font")

  # Number Formats numFmtId
  # overrides for
  # + numFmtId
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-2.8.1
  z$numFmts <- xml_node(sxml, "styleSheet", "numFmts", "numFmt")

  # Table Styles Maybe position Id?
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tablestyle?view=openxml-2.8.1
  z$tableStyles <- xml_node(sxml, "styleSheet", "tableStyles")

  z
}


#' get all styles on a sheet
#'
#' @param wb workbook
#' @param sheet worksheet
#'
#' @export
styles_on_sheet <- function(wb, sheet) {

  sheet_id <- wb$validateSheet(sheet)

  z <- unique(wb$worksheets[[sheet_id]]$sheet_data$cc$c_s)

  as.numeric(z)

}


# TODO can be further generalized with additional xf attributes and children
#' create_builtin_cell_style
#' @param x a numFmt ID for a builin style
#' @param horizontal alignment can be "", "center", "right"
#' @param vertical alignment can be "", "center", "right"
#'
#' @details
#'  | "ID" | "numFmt"                    |
#'  |------|-----------------------------|
#'  | "0"  | "General"                   |
#'  | "1"  | "0"                         |
#'  | "2"  | "0.00"                      |
#'  | "3"  | "#,##0"                     |
#'  | "4"  | "#,##0.00"                  |
#'  | "9"  | "0%"                        |
#'  | "10" | "0.00%"                     |
#'  | "11" | "0.00E+00"                  |
#'  | "12" | "# ?/?"                     |
#'  | "13" | "# ??/??"                   |
#'  | "14" | "mm-dd-yy"                  |
#'  | "15" | "d-mmm-yy"                  |
#'  | "16" | "d-mmm"                     |
#'  | "17" | "mmm-yy"                    |
#'  | "18" | "h:mm AM/PM"                |
#'  | "19" | "h:mm:ss AM/PM"             |
#'  | "20" | "h:mm"                      |
#'  | "21" | "h:mm:ss"                   |
#'  | "22" | "m/d/yy h:mm"               |
#'  | "37" | "#,##0 ;(#,##0)"            |
#'  | "38" | "#,##0 ;\[Red\](#,##0)"     |
#'  | "39" | "#,##0.00;(#,##0.00)"       |
#'  | "40" | "#,##0.00;\[Red\](#,##0.00)"|
#'  | "45" | "mm:ss"                     |
#'  | "46" | "\[h\]:mm:ss"               |
#'  | "47" | "mmss.0"                    |
#'  | "48" | "##0.0E+0"                  |
#'  | "49" | "@"                         |
#' @export
create_builtin_cell_style <- function(x, horizontal = "", vertical = "") {
  n <- length(x)
  x <- as.character(x)

  applyNumberFormat <- ""
  if (any(x != "0")) applyNumberFormat <- "1"

  applyAlignment <- ""
  if (horizontal != "" | vertical != "") applyAlignment <- "1"

  df_cellXfs <- data.frame(
    numFmtId = as.character(x),
    fontId = rep("0", n),
    fillId = rep("0", n),
    borderId = rep("0", n),
    xfId = rep("0", n),
    applyNumberFormat = rep(applyNumberFormat, n),
    applyAlignment = rep(applyAlignment, n),
    horizontal = rep(horizontal, n),
    vertical = rep(vertical, n),
    stringsAsFactors = FALSE
  )

  cellXfs <- write_xf(df_cellXfs)
  cellXfs
}


#' helper get_cell_style
#' @param wb a workbook
#' @param sheet a worksheet
#' @param cell a cell
#' @export
get_cell_style <- function(wb, sheet, cell) {

  sheet <- wb$validateSheet(sheet)

  cell <- as.character(unlist(dims_to_dataframe(cell, fill = TRUE)))
  cc <- wb$worksheets[[sheet]]$sheet_data$cc

  cc$cell <- paste0(cc$c_r, cc$row_r)

  sel <- cc$cell %in% cell
  cc$c_s[sel]
}


#' helper set_cell_style
#' @param wb a workbook
#' @param sheet a worksheet
#' @param cell a cell
#' @param value a value to assign
#' @examples
#' wb <- createWorkbook()
#' addWorksheet(wb, "test")
#'
#' mat <- matrix(rnorm(28*28, mean = 44444, sd = 555), ncol = 28)
#' writeData(wb, "test", mat, colNames = FALSE)
#'
#'
#' x <- c("0", "1", "2", "3", "4", "9", "10", "11", "12", "13", "14", "15", "16",
#'        "17", "18", "19", "20", "21", "22", "37", "38", "39", "40", "45", "46",
#'        "47", "48", "49")
#'
#' # combine original and new styles
#' wb$styles$cellXfs <- c(
#'   wb$styles$cellXfs,
#'   create_builtin_cell_style(x, horizontal = "center")
#' )
#'
#' # new styles are 1:28, because s in a 0-index
#' for (i in seq_len(28)) {
#'   cell <- sprintf("%s1:%s28", int2col(i), int2col(i))
#'   set_cell_style(wb, "test", cell, as.character(i))
#' }
#'
#' wb_save(wb, "test.xlsx")
#' @export
set_cell_style <- function(wb, sheet, cell, value) {

  sheet <- wb$validateSheet(sheet)

  cell <- as.character(unlist(dims_to_dataframe(cell, fill = TRUE)))
  cc <- wb$worksheets[[sheet]]$sheet_data$cc

  cc$cell <- paste0(cc$c_r, cc$row_r)

  sel <- cc$cell %in% cell
  cc$c_s[sel] <- value
  cc$cell <- NULL

  wb$worksheets[[sheet]]$sheet_data$cc <- cc
}
