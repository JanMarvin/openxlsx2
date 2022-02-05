#' @name writeData
#' @title Write an object to a worksheet
#' @import stringi
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
#' @param headerStyle Custom style to apply to column names.
#' @param borders Either "`none`" (default), "`surrounding`",
#' "`columns`", "`rows`" or *respective abbreviations*.  If
#' "`surrounding`", a border is drawn around the data.  If "`rows`",
#' a surrounding border is drawn with a border around each row. If
#' "`columns`", a surrounding border is drawn with a border between
#' each column. If "`all`" all cell borders are drawn.
#' @param borderColour Colour of cell border.  A valid colour (belonging to `colours()` or a hex colour code, eg see [here](https://www.webfx.com/web-design/color-picker/)).
#' @param borderStyle Border line style
#' \itemize{
#'    \item{**none**}{ no border}
#'    \item{**thin**}{ thin border}
#'    \item{**medium**}{ medium border}
#'    \item{**dashed**}{ dashed border}
#'    \item{**dotted**}{ dotted border}
#'    \item{**thick**}{ thick border}
#'    \item{**double**}{ double line border}
#'    \item{**hair**}{ hairline border}
#'    \item{**mediumDashed**}{ medium weight dashed border}
#'    \item{**dashDot**}{ dash-dot border}
#'    \item{**mediumDashDot**}{ medium weight dash-dot border}
#'    \item{**dashDotDot**}{ dash-dot-dot border}
#'    \item{**mediumDashDotDot**}{ medium weight dash-dot-dot border}
#'    \item{**slantDashDot**}{ slanted dash-dot border}
#'   }
#' @param withFilter If `TRUE`, add filters to the column name row. NOTE can only have one filter per worksheet.
#' @param keepNA If `TRUE`, NA values are converted to #N/A (or `na.string`, if not NULL) in Excel, else NA cells will be empty.
#' @param na.string If not NULL, and if `keepNA` is `TRUE`, NA values are converted to this string in Excel.
#' @param name If not NULL, a named region is defined.
#' @param sep Only applies to list columns. The separator used to collapse list columns to a character vector e.g. sapply(x$list_column, paste, collapse = sep).
#' @param removeCellStyle keep the cell style?
#' @seealso [writeDataTable()]
#' @export writeData
#' @details Formulae written using writeFormula to a Workbook object will not get picked up by read.xlsx().
#' This is because only the formula is written and left to Excel to evaluate the formula when the file is opened in Excel.
#' @rdname writeData
#' @return invisible(0)
#' @examples
#'
#' ## See formatting vignette for further examples.
#'
#' ## Options for default styling (These are the defaults)
#' options("openxlsx.borderColour" = "black")
#' options("openxlsx.borderStyle" = "thin")
#' options("openxlsx.dateFormat" = "mm/dd/yyyy")
#' options("openxlsx.datetimeFormat" = "yyyy-mm-dd hh:mm:ss")
#' options("openxlsx.numFmt" = NULL)
#'
#' ## Change the default border colour to #4F81BD
#' options("openxlsx.borderColour" = "#4F81BD")
#'
#'
#' #####################################################################################
#' ## Create Workbook object and add worksheets
#' wb <- createWorkbook()
#'
#' ## Add worksheets
#' addWorksheet(wb, "Cars")
#' addWorksheet(wb, "Formula")
#'
#'
#' x <- mtcars[1:6, ]
#' writeData(wb, "Cars", x, startCol = 2, startRow = 3, rowNames = TRUE)
#'
#' #####################################################################################
#' ## Bordering
#'
#' writeData(wb, "Cars", x,
#'   rowNames = TRUE, startCol = "O", startRow = 3,
#'   borders = "surrounding", borderColour = "black"
#' ) ## black border
#'
#' writeData(wb, "Cars", x,
#'   rowNames = TRUE,
#'   startCol = 2, startRow = 12, borders = "columns"
#' )
#'
#' writeData(wb, "Cars", x,
#'   rowNames = TRUE,
#'   startCol = "O", startRow = 12, borders = "rows"
#' )
#'
#'
#' #####################################################################################
#' ## Header Styles
#'
#' hs1 <- createStyle(
#'   fgFill = "#DCE6F1", halign = "center", textDecoration = "italic",
#'   border = "bottom"
#' )
#'
#' writeData(wb, "Cars", x,
#'   colNames = TRUE, rowNames = TRUE, startCol = "B",
#'   startRow = 23, borders = "rows", headerStyle = hs1, borderStyle = "dashed"
#' )
#'
#'
#' hs2 <- createStyle(
#'   fontColour = "#ffffff", fgFill = "#4F80BD",
#'   halign = "center", valign = "center", textDecoration = "bold",
#'   border = "all"
#' )
#'
#' writeData(wb, "Cars", x,
#'   colNames = TRUE, rowNames = TRUE,
#'   startCol = "O", startRow = 23, borders = "columns", headerStyle = hs2
#' )
#'
#'
#'
#'
#' #####################################################################################
#' ## Hyperlinks
#' ## - vectors/columns with class 'hyperlink' are written as hyperlinks'
#'
#' v <- rep("https://CRAN.R-project.org/", 4)
#' names(v) <- paste0("Hyperlink", 1:4) # Optional: names will be used as display text
#' class(v) <- "hyperlink"
#' writeData(wb, "Cars", x = v, xy = c("B", 32))
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
#' writeData(wb, sheet = "Formula", x = df)
#'
#' ###########################################################################
#' # update cell range and add mtcars
#' xlsxFile <- system.file("extdata", "inline_str.xlsx", package = "openxlsx2")
#' wb2 <- loadWorkbook(xlsxFile)
#'
#' # read dataset with inlinestr
#' wb_to_df(wb2)
#' # read.xlsx(wb2)
#' writeData(wb2, 1, mtcars, startCol = 4, startRow = 4)
#' wb_to_df(wb2)
#' \dontrun{
#' file <- tempfile(fileext = ".xlsx")
#' saveWorkbook(wb2, file, overwrite = TRUE)
#' file.remove(file)
#' }
#'
#' #####################################################################################
#' ## Save workbook
#' ## Open in excel without saving file: openXL(wb)
#' \dontrun{
#' saveWorkbook(wb, "writeDataExample.xlsx", overwrite = TRUE)
#' }
writeData <- function(wb,
  sheet,
  x,
  startCol = 1,
  startRow = 1,
  array = FALSE,
  xy = NULL,
  colNames = TRUE,
  rowNames = FALSE,
  headerStyle = NULL,
  borders = c("none", "surrounding", "rows", "columns", "all"),
  borderColour = getOption("openxlsx.borderColour", "black"),
  borderStyle = getOption("openxlsx.borderStyle", "thin"),
  withFilter = FALSE,
  keepNA = FALSE,
  na.string = NULL,
  name = NULL,
  sep = ", ",
  removeCellStyle = TRUE) {

  ## increase scipen to avoid writing in scientific
  exSciPen <- getOption("scipen")
  od <- getOption("OutDec")
  exDigits <- getOption("digits")

  options("scipen" = 200)
  options("OutDec" = ".")
  options("digits" = 22)

  on.exit(options("scipen" = exSciPen), add = TRUE)
  on.exit(expr = options("OutDec" = od), add = TRUE)
  on.exit(options("digits" = exDigits), add = TRUE)



  if (is.null(x)) {
    return(invisible(0))
  }

  ## All input conversions/validations
  if (!is.null(xy)) {
    if (length(xy) != 2) {
      stop("xy parameter must have length 2")
    }
    startCol <- xy[[1]]
    startRow <- xy[[2]]
  }

  ## convert startRow and startCol
  if (!is.numeric(startCol)) {
    startCol <- col2int(startCol)
  }
  startRow <- as.integer(startRow)

  assert_workbook(wb)
  if (!is.null(headerStyle)) assert_style(headerStyle)
  assert_class(colNames, "logical")
  assert_class(rowNames, "logical")

  if ((!is.character(sep)) | (length(sep) != 1)) stop("sep must be a character vector of length 1")

  borders <- match.arg(borders)
  if (length(borders) != 1) stop("borders argument must be length 1.")

  ## borderColours validation
  borderColour <- validateColour(borderColour, "Invalid border colour")
  borderStyle <- validate_border_style(borderStyle)[[1]]

  ## special case - vector of hyperlinks
  hlinkNames <- NULL
  if (inherits(x, "hyperlink")) {
    # consider wbHyperlink?
    hlinkNames <- names(x)
    colNames <- FALSE
  }

  ## special case - formula
  if (inherits(x, "formula")) {
    x <- data.frame("X" = x, stringsAsFactors = FALSE)
    class(x[[1]]) <- ifelse(array, "array_formula", "formula")
    colNames <- FALSE
  }

  ## named region
  if (!is.null(name)) { ## validate name
    ex_names <- regmatches(wb$workbook$definedNames, regexpr('(?<=name=")[^"]+', wb$workbook$definedNames, perl = TRUE))
    ex_names <- replaceXMLEntities(ex_names)

    if (name %in% ex_names) {
      stop(sprintf("Named region with name '%s' already exists!", name))
    } else if (grepl("^[A-Z]{1,3}[0-9]+$", name)) {
      stop("name cannot look like a cell reference.")
    }
  }

  if (is.vector(x) | is.factor(x) | inherits(x, "Date")) {
    colNames <- FALSE
  } ## this will go to coerce.default and rowNames will be ignored

  ## Coerce to data.frame
  x <- openxlsxCoerce(x = x, rowNames = rowNames)

  nCol <- ncol(x)
  nRow <- nrow(x)

  ## If no rows and not writing column names return as nothing to write
  if (nRow == 0 & !colNames) {
    return(invisible(0))
  }

  ## If no columns and not writing row names return as nothing to write
  if (nCol == 0 & !rowNames) {
    return(invisible(0))
  }

  colClasses <- lapply(x, function(x) tolower(class(x)))
  colClasss2 <- colClasses
  colClasss2[sapply(colClasses, function(x) "formula" %in% x) & sapply(colClasses, function(x) "hyperlink" %in% x)] <- "formula"

  if (is.numeric(sheet)) {
    sheetX <- wb_validate_sheet(wb, sheet)
  } else {
    sheetX <- wb_validate_sheet(wb, replaceXMLEntities(sheet))
    sheet <- replaceXMLEntities(sheet)
  }

  if (wb$isChartSheet[[sheetX]]) {
    stop("Cannot write to chart sheet.")
    return(NULL)
  }



  ## Check not overwriting existing table headers
  wb_check_overwrite_tables(
    wb,
    sheet = sheet,
    new_rows = c(startRow, startRow + nRow - 1L + colNames),
    new_cols = c(startCol, startCol + nCol - 1L),
    check_table_header_only = TRUE,
    error_msg =
      "Cannot overwrite table headers. Avoid writing over the header row or see getTables() & removeTables() to remove the table object."
  )



  ## write autoFilter, can only have a single filter per worksheet
  if (withFilter) {
    coords <- data.frame("x" = c(startRow, startRow + nRow + colNames - 1L), "y" = c(startCol, startCol + nCol - 1L))
    ref <- stri_join(getCellRefs(coords), collapse = ":")

    wb$worksheets[[sheetX]]$autoFilter <- sprintf('<autoFilter ref="%s"/>', ref)

    l <- int2col(unlist(coords[, 2]))
    dfn <- sprintf("'%s'!%s", names(wb)[sheetX], stri_join("$", l, "$", coords[, 1], collapse = ":"))

    dn <- sprintf('<definedName name="_xlnm._FilterDatabase" localSheetId="%s" hidden="1">%s</definedName>', sheetX - 1L, dfn)

    if (length(wb$workbook$definedNames) > 0) {
      ind <- grepl('name="_xlnm._FilterDatabase"', wb$workbook$definedNames)
      if (length(ind) > 0) {
        wb$workbook$definedNames[ind] <- dn
      }
    } else {
      wb$workbook$definedNames <- dn
    }
  }

  if (is.data.frame(x))
    rownames(x) <- seq_len(NROW(x))

  # actual driver, the rest should not create data used for writing
  wb <- writeData2(wb = wb, sheet = sheet, data = x, colNames = colNames, rowNames = FALSE, startRow = startRow, startCol = startCol, removeCellStyle = removeCellStyle)

  invisible(0)
}



#' @name writeFormula
#' @title Write a character vector as an Excel Formula
#' @description Write a a character vector containing Excel formula to a worksheet.
#' @details Currently only the english version of functions are supported. Please don't use the local translation.
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
#' @seealso [writeData()]
#' @export writeFormula
#' @rdname writeFormula
#' @examples
#'
#' ## There are 3 ways to write a formula
#'
#' wb <- createWorkbook()
#' addWorksheet(wb, "Sheet 1")
#' writeData(wb, "Sheet 1", x = iris)
#'
#' ## SEE int2col() to convert int to Excel column label
#'
#' ## 1. -  As a character vector using writeFormula
#'
#' v <- c("SUM(A2:A151)", "AVERAGE(B2:B151)") ## skip header row
#' writeFormula(wb, sheet = 1, x = v, startCol = 10, startRow = 2)
#' writeFormula(wb, 1, x = "A2 + B2", startCol = 10, startRow = 10)
#'
#'
#' ## 2. - As a data.frame column with class "formula" using writeData
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
#' addWorksheet(wb, "Sheet 2")
#' writeData(wb, sheet = 2, x = df)
#'
#'
#'
#' ## 3. - As a vector with class "formula" using writeData
#'
#' v2 <- c("SUM(A2:A4)", "AVERAGE(B2:B4)", "MEDIAN(C2:C4)")
#' class(v2) <- c(class(v2), "formula")
#'
#' writeData(wb, sheet = 2, x = v2, startCol = 10, startRow = 2)
#'
#' ## Save workbook
#' \dontrun{
#' saveWorkbook(wb, "writeFormulaExample.xlsx", overwrite = TRUE)
#' }
#'
#'
#' ## 4. - Writing internal hyperlinks
#'
#' wb <- createWorkbook()
#' addWorksheet(wb, "Sheet1")
#' addWorksheet(wb, "Sheet2")
#' writeFormula(wb, "Sheet1", x = '=HYPERLINK("#Sheet2!B3", "Text to Display - Link to Sheet2")')
#'
#' ## Save workbook
#' \dontrun{
#' saveWorkbook(wb, "writeFormulaHyperlinkExample.xlsx", overwrite = TRUE)
#' }
#'
writeFormula <- function(wb,
  sheet,
  x,
  startCol = 1,
  startRow = 1,
  array = FALSE,
  xy = NULL) {
  assert_class(x, "character")
  dfx <- data.frame("X" = x, stringsAsFactors = FALSE)
  class(dfx$X) <- c("character", ifelse(array, "array_formula", "formula"))

  if (any(grepl("^(=|)HYPERLINK\\(", x, ignore.case = TRUE))) {
    class(dfx$X) <- c("character", "formula", "hyperlink")
  }



  writeData(
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
