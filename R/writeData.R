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
#' @param name If not NULL, a named region is defined.
#' @param sep Only applies to list columns. The separator used to collapse list columns to a character vector e.g. sapply(x$list_column, paste, collapse = sep).
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
write_data <- function(wb,
  sheet,
  x,
  startCol = 1,
  startRow = 1,
  array = FALSE,
  xy = NULL,
  colNames = TRUE,
  rowNames = FALSE,
  withFilter = FALSE,
  name = NULL,
  sep = ", ",
  removeCellStyle = TRUE) {
  op <- openxlsx_options()
  on.exit(options(op), add = TRUE)

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
  assert_class(colNames, "logical")
  assert_class(rowNames, "logical")

  if ((!is.character(sep)) || (length(sep) != 1)) stop("sep must be a character vector of length 1")

  ## special case - vector of hyperlinks
  # # hlinkNames not used?
  # hlinkNames <- NULL

  is_hyperlink <- FALSE
  if (is.null(dim(x))) {
    is_hyperlink <- inherits(x, "hyperlink")
  } else {
    is_hyperlink <- vapply(x, inherits, what = "hyperlink", FALSE)
  }

  if (any(is_hyperlink)) {
    # consider wbHyperlink?
    # hlinkNames <- names(x)
    if (is.null(dim(x))) {
      colNames <- FALSE
      x[is_hyperlink] <- create_hyperlink(text = x[is_hyperlink])
      class(x[is_hyperlink]) <- c("character", "hyperlink")
    } else {
      # check should be in create_hyperlink and that apply should not be required either
      if (!any(grepl("^(=|)HYPERLINK\\(", x[is_hyperlink], ignore.case = TRUE))) {
        x[is_hyperlink] <- apply(x[is_hyperlink], 1, FUN=function(str) create_hyperlink(text = str))
      }
      class(x[,is_hyperlink]) <- c("character", "hyperlink")
    }
  }

  ## special case - formula
  if (inherits(x, "formula")) {
    x <- data.frame("X" = x, stringsAsFactors = FALSE)
    class(x[[1]]) <- if (array) "array_formula" else "formula"
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

  if (is.vector(x) || is.factor(x) || inherits(x, "Date")) {
    colNames <- FALSE
  } ## this will go to coerce.default and rowNames will be ignored

  ## Coerce to data.frame
  if (inherits(x, "hyperlink")) { 
    ## vector of hyperlinks
    class(x) <- c("character", "hyperlink")
    x <- as.data.frame(x, stringsAsFactors = FALSE)
  } else if (!inherits(x, "data.frame")) {
    x <- as.data.frame(x, stringsAsFactors = FALSE)
  } else {
    ## cbind rownames to x
    if (rowNames) {
      x <- cbind(
        data.frame("row names" = rownames(x), stringsAsFactors = FALSE),
        as.data.frame(x, stringsAsFactors = FALSE)
      )
      names(x)[[1]] <- ""
    }
  }

  nCol <- ncol(x)
  nRow <- nrow(x)

  ## If no rows and not writing column names return as nothing to write
  if (nRow == 0 && !colNames) {
    return(invisible(0))
  }

  ## If no columns and not writing row names return as nothing to write
  if (nCol == 0 && !rowNames) {
    return(invisible(0))
  }

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
      "Cannot overwrite table headers. Avoid writing over the header row or see wb_get_tables() & wb_remove_tabless() to remove the table object."
  )



  ## write autoFilter, can only have a single filter per worksheet
  if (withFilter) {
    coords <- data.frame("x" = c(startRow, startRow + nRow + colNames - 1L), "y" = c(startCol, startCol + nCol - 1L))
    ref <- stri_join(get_cell_refs(coords), collapse = ":")

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

  # actual driver, the rest should not create data used for writing
  wb <- write_data2(
    wb = wb,
    sheet = sheet,
    data = x,
    name = name,
    colNames = colNames,
    rowNames = FALSE,
    startRow = startRow,
    startCol = startCol,
    removeCellStyle = removeCellStyle
  )

  invisible(0)
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
