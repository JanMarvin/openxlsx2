#' @name write_data
#' @title Write an object to a worksheet
#' @description Write an object to worksheet with optional styling.
#' @param wb A Workbook object containing a worksheet.
#' @param sheet The worksheet to write to. Can be the worksheet index or name.
#' @param x Object to be written. For classes supported look at the examples.
#' @param startCol A vector specifying the starting column to write to.
#' @param startRow A vector specifying the starting row to write to.
#' @param array A bool if the function written is of type array
#' @param xy An alternative to specifying `startCol` and
#' `startRow` individually.  A vector of the form
#' `c(startCol, startRow)`.
#' @param colNames If `TRUE`, column names of x are written.
#' @param rowNames If `TRUE`, data.frame row names of x are written.
#' @param withFilter If `TRUE`, add filters to the column name row. NOTE can only have one filter per worksheet.
#' @param sep Only applies to list columns. The separator used to collapse list columns to a character vector e.g. sapply(x$list_column, paste, collapse = sep).
#' @param name If not NULL, a named region is defined.
#' @param removeCellStyle keep the cell style?
#' @seealso [write_datatable()]
#' @export write_data
#' @details Formulae written using write_formula to a Workbook object will not get picked up by read_xlsx().
#' This is because only the formula is written and left to Excel to evaluate the formula when the file is opened in Excel.
#' The string `"_openxlsx_NA"` is reserved for `openxlsx2`. If the data frame contains this string, the output will be broken.
#' @rdname write_data
#' @return invisible(0)
#' @examples
#'
#' ## See formatting vignette for further examples.
#'
#' ## Options for default styling (These are the defaults)
#' options("openxlsx2.dateFormat" = "mm/dd/yyyy")
#' options("openxlsx2.datetimeFormat" = "yyyy-mm-dd hh:mm:ss")
#' options("openxlsx2.numFmt" = NULL)
#'
#'
#' #####################################################################################
#' ## Create Workbook object and add worksheets
#' wb <- wb_workbook()
#'
#' ## Add worksheets
#' wb$add_worksheet("Cars")
#' wb$add_worksheet("Formula")
#'
#'
#' x <- mtcars[1:6, ]
#' wb$add_data("Cars", x, startCol = 2, startRow = 3, rowNames = TRUE)
#'
#'
#' #####################################################################################
#' ## Hyperlinks
#' ## - vectors/columns with class 'hyperlink' are written as hyperlinks'
#'
#' v <- rep("https://CRAN.R-project.org/", 4)
#' names(v) <- paste0("Hyperlink", 1:4) # Optional: names will be used as display text
#' class(v) <- "hyperlink"
#' wb$add_data("Cars", x = v, xy = c("B", 32))
#'
#'
#' #####################################################################################
#' ## Formulas
#' ## - vectors/columns with class 'formula' are written as formulas'
#'
#' df <- data.frame(
#'   x = 1:3, y = 1:3,
#'   z = paste(paste0("A", 1:3 + 1L), paste0("B", 1:3 + 1L), sep = "+"),
#'   stringsAsFactors = FALSE
#' )
#'
#' class(df$z) <- c(class(df$z), "formula")
#'
#' wb$add_data(sheet = "Formula", x = df)
#'
#' ###########################################################################
#' # update cell range and add mtcars
#' xlsxFile <- system.file("extdata", "inline_str.xlsx", package = "openxlsx2")
#' wb2 <- wb_load(xlsxFile)
#'
#' # read dataset with inlinestr
#' wb_to_df(wb2)
#' # read_xlsx(wb2)
#' write_data(wb2, 1, mtcars, startCol = 4, startRow = 4)
#' wb_to_df(wb2)
#' \dontrun{
#' file <- tempfile(fileext = ".xlsx")
#' wb_save(wb2, file, overwrite = TRUE)
#' file.remove(file)
#' }
#'
#' #####################################################################################
#' ## Save workbook
#' ## Open in excel without saving file: xl_open(wb)
#' \dontrun{
#' wb_save(wb, "write_dataExample.xlsx", overwrite = TRUE)
#' }
write_data <- function(
    wb,
    sheet,
    x,
    startCol = 1,
    startRow = 1,
    array = FALSE,
    xy = NULL,
    colNames = TRUE,
    rowNames = FALSE,
    withFilter = FALSE,
    sep = ", ",
    name = NULL,
    removeCellStyle = TRUE
) {
  write_data_table(
    wb = wb,
    sheet = sheet,
    x = x,
    startCol = startCol,
    startRow = startRow,
    array = array,
    xy = xy,
    colNames = colNames,
    rowNames = rowNames,
    tableStyle = NULL,
    tableName = NULL,
    withFilter = withFilter,
    sep = sep,
    stack = FALSE,
    firstColumn = FALSE,
    lastColumn = FALSE,
    bandedRows = FALSE,
    bandedCols = FALSE,
    name = name,
    removeCellStyle = removeCellStyle,
    data_table = FALSE
  )
}



#' @name write_formula
#' @title Write a character vector as an Excel Formula
#' @description Write a a character vector containing Excel formula to a worksheet.
#' @details Currently only the English version of functions are supported. Please don't use the local translation.
#' The examples below show a small list of possible formulas:
#' \itemize{
#'     \item{SUM(B2:B4)}
#'     \item{AVERAGE(B2:B4)}
#'     \item{MIN(B2:B4)}
#'     \item{MAX(B2:B4)}
#'     \item{...}
#'
#' }
#' @param wb A Workbook object containing a worksheet.
#' @param sheet The worksheet to write to. Can be the worksheet index or name.
#' @param x A character vector.
#' @param startCol A vector specifying the starting column to write to.
#' @param startRow A vector specifying the starting row to write to.
#' @param array A bool if the function written is of type array
#' @param xy An alternative to specifying `startCol` and
#' `startRow` individually.  A vector of the form
#' `c(startCol, startRow)`.
#' @seealso [write_data()]
#' @export write_formula
#' @rdname write_formula
#' @examples
#'
#' ## There are 3 ways to write a formula
#'
#' wb <- wb_workbook()
#' wb$add_worksheet("Sheet 1")
#' wb$add_data("Sheet 1", x = iris)
#'
#' ## SEE int2col() to convert int to Excel column label
#'
#' ## 1. -  As a character vector using write_formula
#'
#' v <- c("SUM(A2:A151)", "AVERAGE(B2:B151)") ## skip header row
#' write_formula(wb, sheet = 1, x = v, startCol = 10, startRow = 2)
#' write_formula(wb, 1, x = "A2 + B2", startCol = 10, startRow = 10)
#'
#'
#' ## 2. - As a data.frame column with class "formula" using write_data
#'
#' df <- data.frame(
#'   x = 1:3,
#'   y = 1:3,
#'   z = paste(paste0("A", 1:3 + 1L), paste0("B", 1:3 + 1L), sep = " + "),
#'   z2 = sprintf("ADDRESS(1,%s)", 1:3),
#'   stringsAsFactors = FALSE
#' )
#'
#' class(df$z) <- c(class(df$z), "formula")
#' class(df$z2) <- c(class(df$z2), "formula")
#'
#' wb$add_worksheet("Sheet 2")
#' wb$add_data(sheet = 2, x = df)
#'
#'
#'
#' ## 3. - As a vector with class "formula" using write_data
#'
#' v2 <- c("SUM(A2:A4)", "AVERAGE(B2:B4)", "MEDIAN(C2:C4)")
#' class(v2) <- c(class(v2), "formula")
#'
#' wb$add_data(sheet = 2, x = v2, startCol = 10, startRow = 2)
#'
#' ## Save workbook
#' \dontrun{
#' wb_save(wb, "write_formulaExample.xlsx", overwrite = TRUE)
#' }
#'
#'
#' ## 4. - Writing internal hyperlinks
#'
#' wb <- wb_workbook()
#' wb$add_worksheet("Sheet1")
#' wb$add_worksheet("Sheet2")
#' write_formula(wb, "Sheet1", x = '=HYPERLINK("#Sheet2!B3", "Text to Display - Link to Sheet2")')
#'
#' ## Save workbook
#' \dontrun{
#' wb_save(wb, "write_formulaHyperlinkExample.xlsx", overwrite = TRUE)
#' }
#'
#' ## 5. - Writing array formulas
#' set.seed(123)
#' df <- data.frame(C = rnorm(10), D = rnorm(10))
#'
#' wb <- wb_workbook()
#' wb <- wb_add_worksheet(wb, "df")
#'
#' wb$add_data("df", df, startCol = "C")
#'
#' write_formula(wb, "df", startCol = "E", startRow = "2",
#'              x = "SUM(C2:C11*D2:D11)",
#'              array = TRUE)
#'
write_formula <- function(wb,
  sheet,
  x,
  startCol = 1,
  startRow = 1,
  array = FALSE,
  xy = NULL) {
  assert_class(x, "character")
  dfx <- data.frame("X" = x, stringsAsFactors = FALSE)
  class(dfx$X) <- c("character", if (array) "array_formula" else "formula")

  if (any(grepl("^(=|)HYPERLINK\\(", x, ignore.case = TRUE))) {
    class(dfx$X) <- c("character", "formula", "hyperlink")
  }

  write_data(
    wb = wb,
    sheet = sheet,
    x = dfx,
    startCol = startCol,
    startRow = startRow,
    array = array,
    xy = xy,
    colNames = FALSE,
    rowNames = FALSE
  )


  invisible(0)
}
