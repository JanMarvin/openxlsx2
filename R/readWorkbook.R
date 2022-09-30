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
#' @seealso [get_named_regions()]
#' @details Formulae written using write_formula to a Workbook object will not get picked up by read_xlsx().
#' This is because only the formula is written and left to be evaluated when the file is opened in Excel.
#' Opening, saving and closing the file with Excel will resolve this.
#' @return data.frame
#' @export
#' @examples
#'
#' xlsxFile <- system.file("extdata", "readTest.xlsx", package = "openxlsx2")
#' df1 <- read_xlsx(xlsxFile = xlsxFile, sheet = 1, skipEmptyRows = FALSE)
#' sapply(df1, class)
#'
#' df2 <- read_xlsx(xlsxFile = xlsxFile, sheet = 3, skipEmptyRows = TRUE)
#' df2$Date <- convert_date(df2$Date)
#' sapply(df2, class)
#' head(df2)
#'
#' df2 <- read_xlsx(
#'   xlsxFile = xlsxFile, sheet = 3, skipEmptyRows = TRUE,
#'   detectDates = TRUE
#' )
#' sapply(df2, class)
#' head(df2)
#'
#' wb <- wb_load(system.file("extdata", "readTest.xlsx", package = "openxlsx2"))
#' df3 <- read_xlsx(wb, sheet = 2, skipEmptyRows = FALSE, colNames = TRUE)
#' df4 <- read_xlsx(xlsxFile, sheet = 2, skipEmptyRows = FALSE, colNames = TRUE)
#' all.equal(df3, df4)
#'
#' wb <- wb_load(system.file("extdata", "readTest.xlsx", package = "openxlsx2"))
#' df3 <- read_xlsx(wb,
#'   sheet = 2, skipEmptyRows = FALSE,
#'   cols = c(1, 4), rows = c(1, 3, 4)
#' )
#'
#' ## URL
#' ##
#' xlsxFile <- "https://github.com/JanMarvin/openxlsx2/raw/main/inst/extdata/readTest.xlsx"
#' head(read_xlsx(xlsxFile))
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
  fillMergedCells = FALSE
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
    fillMergedCells = fillMergedCells
  )
}


#' @name wb_read
#' @title  Read from an Excel file or Workbook object
#' @description Read data from an Excel file or Workbook object into a data.frame
#' @inheritParams read_xlsx
#' @details Creates a data.frame of all data in worksheet.
#' @return data.frame
#' @seealso [get_named_regions()]
#' @seealso [read_xlsx()]
#' @export
#' @examples
#' xlsxFile <- system.file("extdata", "readTest.xlsx", package = "openxlsx2")
#' df1 <- wb_read(xlsxFile = xlsxFile, sheet = 1)
#'
#' xlsxFile <- system.file("extdata", "readTest.xlsx", package = "openxlsx2")
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
  na.numbers    = NA
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
    na.numbers    = na.numbers
  )

}

#' @name read_sheet_names
#' @title Get names of worksheets
#' @description Returns the worksheet names within an xlsx file
#' @param file An xlsx or xlsm file.
#' @return Character vector of worksheet names.
#' @examples
#' read_sheet_names(system.file("extdata", "readTest.xlsx", package = "openxlsx2"))
#' @export
read_sheet_names <- function(file) {
  if (!file.exists(file)) {
    stop("file does not exist.")
  }

  if (grepl("\\.xls$|\\.xlm$", file)) {
    stop("openxlsx can not read .xls or .xlm files!")
  }

  ## create temp dir and unzip
  xmlDir <- temp_dir("_excelXMLRead")
  on.exit(unlink(xmlDir, recursive = TRUE), add = TRUE)
  xmlFiles <- unzip(file, exdir = xmlDir)

  workbook <- grep("workbook.xml$", xmlFiles, perl = TRUE, value = TRUE)
  workbook <- read_xml(workbook)
  sheets <- xml_node(workbook, "workbook", "sheets", "sheet")
  sheets <- xml_attr(sheets, "sheet")
  sheets <- rbindlist(sheets)

  ## Some veryHidden sheets do not have a sheet content and their rId is empty.
  ## Such sheets need to be filtered out because otherwise their sheet names
  ## occur in the list of all sheet names, leading to a wrong association
  ## of sheet names with sheet indeces.
  sheets <- sheets$name[sheets$`r:id` != ""]
  sheets <- replaceXMLEntities(sheets)

  return(sheets)
}
