# `read_xlsx()` -----------------------------------------------------------------
#' Read from an Excel file or Workbook object
#'
#' Read data from an Excel file or Workbook object into a data.frame
#'
#' @details
#' Formulae written using write_formula to a Workbook object will not get picked up by read_xlsx().
#' This is because only the formula is written and left to be evaluated when the file is opened in Excel.
#' Opening, saving and closing the file with Excel will resolve this.
#' @seealso [wb_get_named_regions()] [wb_to_df()]
#'
#' @param xlsx_file An xlsx file, Workbook object or URL to xlsx file.
#' @param sheet The name or index of the sheet to read data from.
#' @param start_row first row to begin looking for data.
#' @param start_col first column to begin looking for data.
#' @param col_names If `TRUE`, the first row of data will be used as column names.
#' @param skip_empty_rows If `TRUE`, empty rows are skipped else empty rows after the first row containing data
#' will return a row of NAs.
#' @param row_names If `TRUE`, first column of data will be used as row names.
#' @param detect_dates If `TRUE`, attempt to recognize dates and perform conversion.
#' @param cols A numeric vector specifying which columns in the Excel file to read.
#' If NULL, all columns are read.
#' @param rows A numeric vector specifying which rows in the Excel file to read.
#' If NULL, all rows are read.
#' @param check.names logical. If TRUE then the names of the variables in the data frame
#' are checked to ensure that they are syntactically valid variable names
#' @param sep.names (unimplemented) One character which substitutes blanks in column names. By default, "."
#' @param named_region A named region in the Workbook. If not NULL startRow, rows and cols parameters are ignored.
#' @param na.strings A character vector of strings which are to be interpreted as NA. Blank cells will be returned as NA.
#' @param na.numbers A numeric vector of digits which are to be interpreted as NA. Blank cells will be returned as NA.
#' @param fill_merged_cells If TRUE, the value in a merged cell is given to all cells within the merge.
#' @param skip_empty_cols If `TRUE`, empty columns are skipped.
#' @param ... additional arguments passed to `wb_to_df()`
#' @return A data.frame
#'
#' @examples
#' xlsxFile <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
#' read_xlsx(xlsxFile = xlsxFile)
#' @export
read_xlsx <- function(
  xlsx_file,
  sheet,
  start_row         = 1,
  start_col         = NULL,
  row_names         = FALSE,
  col_names         = TRUE,
  skip_empty_rows   = FALSE,
  skip_empty_cols   = FALSE,
  rows              = NULL,
  cols              = NULL,
  detect_dates      = TRUE,
  named_region,
  na.strings        = "#N/A",
  na.numbers        = NA,
  check.names       = FALSE,
  sep.names         = ".",
  fill_merged_cells = FALSE,
  ...
) {

  # keep sheet missing // read_xlsx is the function to replace.
  # dont mess with wb_to_df
  if (missing(sheet))
    sheet <- substitute()

  wb_to_df(
    xlsx_file,
    sheet             = sheet,
    start_row         = start_row,
    start_col         = start_col,
    row_names         = row_names,
    col_names         = col_names,
    skip_empty_rows   = skip_empty_rows,
    skip_empty_cols   = skip_empty_cols,
    rows              = rows,
    cols              = cols,
    detect_dates      = detect_dates,
    named_region      = named_region,
    na.strings        = na.strings,
    na.numbers        = na.numbers,
    fill_merged_cells = fill_merged_cells,
    ...
  )
}

# `wb_read()` ------------------------------------------------------------------
#' Read from an Excel file or Workbook object
#'
#' Read data from an Excel file or Workbook object into a data.frame
#'
#' @details
#' Creates a data.frame of all data in worksheet.
#' @seealso [wb_get_named_regions()] [wb_to_df()] [read_xlsx()]
#'
#' @inheritParams read_xlsx
#' @inherit read_xlsx return
#' @examples
#' xlsxFile <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
#' df1 <- wb_read(xlsxFile = xlsxFile, sheet = 1)
#'
#' xlsxFile <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
#' df1 <- wb_read(xlsxFile = xlsxFile, sheet = 1, rows = c(1, 3, 5), cols = 1:3)
#' @export
wb_read <- function(
  xlsx_file,
  sheet           = 1,
  start_row       = 1,
  start_col       = NULL,
  row_names       = FALSE,
  col_names       = TRUE,
  skip_empty_rows = FALSE,
  skip_empty_cols = FALSE,
  rows            = NULL,
  cols            = NULL,
  detect_dates    = TRUE,
  named_region,
  na.strings      = "NA",
  na.numbers      = NA,
  ...
) {

  # keep sheet missing // read_xlsx is the function to replace.
  # dont mess with wb_to_df
  if (missing(sheet))
    sheet <- substitute()

  wb_to_df(
    xlsx_file       = xlsx_file,
    sheet           = sheet,
    start_row       = start_row,
    start_col       = start_col,
    row_names       = row_names,
    col_names       = col_names,
    skip_empty_rows = skip_empty_rows,
    skip_empty_cols = skip_empty_cols,
    rows            = rows,
    cols            = cols,
    detect_dates    = detect_dates,
    named_region    = named_region,
    na.strings      = na.strings,
    na.numbers      = na.numbers,
    ...
  )

}
# `read_sheet_names()` ----------------------------------------------
#' Get names of worksheets
#'
#' Returns the worksheet names within an xlsx file
#'
#' @param file An xlsx or xlsm file.
#' @return Character vector of worksheet names.
#' @examples
#' wb_load(system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2"))$get_sheet_names()
#' @keywords internal
#' @export
read_sheet_names <- function(file) {
  # TODO Move `read_sheet_names()` to a file named R/openxlsx2-deprecated.R
  if (!inherits(file, "wbWorkbook")) {
    wb <- wb_load(file)
  }
  .Deprecated(old = "read_sheet_names", new = "wb_get_sheet_names")

  unname(wb$get_sheet_names())
}
