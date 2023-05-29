#' @name read_xlsx
#' @title  Read from an Excel file or Workbook object
#' @description Read data from an Excel file or Workbook object into a data.frame
#' @param xlsxFile An xlsx file, Workbook object or URL to xlsx file.
#' @param sheet The name or index of the sheet to read data from.
#' @param startRow first row to begin looking for data.
#' @param startCol first column to begin looking for data.
#' @param colNames If `TRUE`, the first row of data will be used as column names.
#' @param skipEmptyRows If `TRUE`, empty rows are skipped else empty rows after the first row containing data
#' will return a row of NAs.
#' @param rowNames If `TRUE`, first column of data will be used as row names.
#' @param detectDates If `TRUE`, attempt to recognize dates and perform conversion.
#' @param cols A numeric vector specifying which columns in the Excel file to read.
#' If NULL, all columns are read.
#' @param rows A numeric vector specifying which rows in the Excel file to read.
#' If NULL, all rows are read.
#' @param check.names logical. If TRUE then the names of the variables in the data frame
#' are checked to ensure that they are syntactically valid variable names
#' @param sep.names (unimplemented) One character which substitutes blanks in column names. By default, "."
#' @param namedRegion A named region in the Workbook. If not NULL startRow, rows and cols parameters are ignored.
#' @param na.strings A character vector of strings which are to be interpreted as NA. Blank cells will be returned as NA.
#' @param na.numbers A numeric vector of digits which are to be interpreted as NA. Blank cells will be returned as NA.
#' @param fillMergedCells If TRUE, the value in a merged cell is given to all cells within the merge.
#' @param skipEmptyCols If `TRUE`, empty columns are skipped.
#' @param ... additional arguments passed to `wb_to_df()`
#' @seealso [wb_get_named_regions()] [wb_to_df()]
#' @details Formulae written using write_formula to a Workbook object will not get picked up by read_xlsx().
#' This is because only the formula is written and left to be evaluated when the file is opened in Excel.
#' Opening, saving and closing the file with Excel will resolve this.
#' @return data.frame
#' @export
#' @examples
#' xlsxFile <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
#' read_xlsx(xlsxFile = xlsxFile)
#' @export
read_xlsx <- function(
  xlsxFile,
  sheet,
  startRow        = 1,
  startCol        = NULL,
  rowNames        = FALSE,
  colNames        = TRUE,
  skipEmptyRows   = FALSE,
  skipEmptyCols   = FALSE,
  rows            = NULL,
  cols            = NULL,
  detectDates     = TRUE,
  namedRegion,
  na.strings      = "#N/A",
  na.numbers      = NA,
  check.names     = FALSE,
  sep.names       = ".",
  fillMergedCells = FALSE,
  ...
) {

  # keep sheet missing // read_xlsx is the function to replace.
  # dont mess with wb_to_df
  if (missing(sheet))
    sheet <- substitute()

  wb_to_df(
    xlsxFile,
    sheet           = sheet,
    startRow        = startRow,
    startCol        = startCol,
    rowNames        = rowNames,
    colNames        = colNames,
    skipEmptyRows   = skipEmptyRows,
    skipEmptyCols   = skipEmptyCols,
    rows            = rows,
    cols            = cols,
    detectDates     = detectDates,
    named_region    = namedRegion,
    na.strings      = na.strings,
    na.numbers      = na.numbers,
    fillMergedCells = fillMergedCells,
    ...
  )
}


#' @name wb_read
#' @title  Read from an Excel file or Workbook object
#' @description Read data from an Excel file or Workbook object into a data.frame
#' @inheritParams read_xlsx
#' @details Creates a data.frame of all data in worksheet.
#' @param ... additional arguments passed to `wb_to_df()`
#' @seealso [wb_get_named_regions()] [wb_to_df()] [read_xlsx()]
#' @return data.frame
#' @export
#' @examples
#' xlsxFile <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
#' df1 <- wb_read(xlsxFile = xlsxFile, sheet = 1)
#'
#' xlsxFile <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
#' df1 <- wb_read(xlsxFile = xlsxFile, sheet = 1, rows = c(1, 3, 5), cols = 1:3)
wb_read <- function(
  xlsxFile,
  sheet         = 1,
  startRow      = 1,
  startCol      = NULL,
  rowNames      = FALSE,
  colNames      = TRUE,
  skipEmptyRows = FALSE,
  skipEmptyCols = FALSE,
  rows          = NULL,
  cols          = NULL,
  detectDates   = TRUE,
  namedRegion,
  na.strings    = "NA",
  na.numbers    = NA,
  ...
) {

  # keep sheet missing // read_xlsx is the function to replace.
  # dont mess with wb_to_df
  if (missing(sheet))
    sheet <- substitute()

  wb_to_df(
    xlsxFile      = xlsxFile,
    sheet         = sheet,
    startRow      = startRow,
    startCol      = startCol,
    rowNames      = rowNames,
    colNames      = colNames,
    skipEmptyRows = skipEmptyRows,
    skipEmptyCols = skipEmptyCols,
    rows          = rows,
    cols          = cols,
    detectDates   = detectDates,
    named_region  = namedRegion,
    na.strings    = na.strings,
    na.numbers    = na.numbers,
    ...
  )

}

#' @name read_sheet_names
#' @title Get names of worksheets
#' @description Returns the worksheet names within an xlsx file
#' @param file An xlsx or xlsm file.
#' @return Character vector of worksheet names.
#' @examples
#' wb_load(system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2"))$get_sheet_names()
#' @export
read_sheet_names <- function(file) {
  if (!inherits(file, "wbWorkbook")) {
    wb <- wb_load(file)
  }
  .Deprecated(old = "read_sheet_names", new = "wb_get_sheet_names")

  unname(wb$get_sheet_names())
}
