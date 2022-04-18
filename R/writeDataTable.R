
#' @name writeDataTable
#' @title Write to a worksheet as an Excel table
#' @description Write to a worksheet and format as an Excel table
#' @param wb A Workbook object containing a
#' worksheet.
#' @param sheet The worksheet to write to. Can be the worksheet index or name.
#' @param x A dataframe.
#' @param startCol A vector specifying the starting column to write df
#' @param startRow A vector specifying the starting row to write df
#' @param xy An alternative to specifying startCol and startRow individually.
#' A vector of the form c(startCol, startRow)
#' @param colNames If `TRUE`, column names of x are written.
#' @param rowNames If `TRUE`, row names of x are written.
#' @param tableStyle Any excel table style name or "none" (see "formatting" vignette).
#' @param tableName name of table in workbook. The table name must be unique.
#' @param withFilter If `TRUE`, columns with have filters in the first row.
#' @param sep Only applies to list columns. The separator used to collapse list columns to a character vector e.g. sapply(x$list_column, paste, collapse = sep).
#' @param stack If `TRUE` the new style is merged with any existing cell styles.  If FALSE, any
#' existing style is replaced by the new style.
#' \cr\cr
#' \cr**The below options correspond to Excel table options:**
#' \cr
#' \if{html}{\figure{tableoptions.png}{options: width="40\%" alt="Figure: table_options.png"}}
#' \if{latex}{\figure{tableoptions.pdf}{options: width=7cm}}
#'
#' @param firstColumn logical. If TRUE, the first column is bold
#' @param lastColumn logical. If TRUE, the last column is bold
#' @param bandedRows logical. If TRUE, rows are colour banded
#' @param bandedCols logical. If TRUE, the columns are colour banded
#' @details columns of x with class Date/POSIXt, currency, accounting,
#' hyperlink, percentage are automatically styled as dates, currency, accounting,
#' hyperlinks, percentages respectively.
#' The string `"_openxlsx_NA"` is reserved for `openxlsx2`. If the data frame
#' contains this string, the output will be broken.
#' @seealso [wb_add_worksheet()]
#' @seealso [writeData()]
#' @seealso [removeTable()]
#' @seealso [getTables()]
#' @export
#' @examples
#' ## see package vignettes for further examples.
#'
#' #####################################################################################
#' ## Create Workbook object and add worksheets
#' wb <- wb_workbook()
#' wb$addWorksheet("S1")
#' wb$addWorksheet("S2")
#' wb$addWorksheet("S3")
#'
#'
#' #####################################################################################
#' ## -- write data.frame as an Excel table with column filters
#' ## -- default table style is "TableStyleMedium2"
#'
#' writeDataTable(wb, "S1", x = iris)
#'
#' writeDataTable(wb, "S2",
#'   x = mtcars, xy = c("B", 3), rowNames = TRUE,
#'   tableStyle = "TableStyleLight9"
#' )
#'
#' df <- data.frame(
#'   "Date" = Sys.Date() - 0:19,
#'   "T" = TRUE, "F" = FALSE,
#'   "Time" = Sys.time() - 0:19 * 60 * 60,
#'   "Cash" = paste("$", 1:20), "Cash2" = 31:50,
#'   "hLink" = "https://CRAN.R-project.org/",
#'   "Percentage" = seq(0, 1, length.out = 20),
#'   "TinyNumbers" = runif(20) / 1E9, stringsAsFactors = FALSE
#' )
#'
#' ## openxlsx will apply default Excel styling for these classes
#' class(df$Cash) <- c(class(df$Cash), "currency")
#' class(df$Cash2) <- c(class(df$Cash2), "accounting")
#' class(df$hLink) <- "hyperlink"
#' class(df$Percentage) <- c(class(df$Percentage), "percentage")
#' class(df$TinyNumbers) <- c(class(df$TinyNumbers), "scientific")
#'
#' writeDataTable(wb, "S3", x = df, startRow = 4, rowNames = TRUE, tableStyle = "TableStyleMedium9")
#'
#' #####################################################################################
#' ## Additional Header Styling and remove column filters
#'
#' writeDataTable(wb,
#'   # todo header styling not implemented
#'   sheet = 1, x = iris, startCol = 7,
#'   withFilter = FALSE
#' )
#'
#'
#' #####################################################################################
#' ## Save workbook
#' ## Open in excel without saving file: openXL(wb)
#' \dontrun{
#' wb_save(wb, "writeDataTableExample.xlsx", overwrite = TRUE)
#' }
#'
#'
#'
#'
#'
#' #####################################################################################
#' ## Pre-defined table styles gallery
#'
#' wb <- wb_workbook(paste0("tableStylesGallery.xlsx"))
#' wb$addWorksheet("Style Samples")
#' for (i in 1:21) {
#'   style <- paste0("TableStyleLight", i)
#'   writeDataTable(wb,
#'     x = data.frame(style), sheet = 1,
#'     tableStyle = style, startRow = 1, startCol = i * 3 - 2
#'   )
#' }
#'
#' for (i in 1:28) {
#'   style <- paste0("TableStyleMedium", i)
#'   writeDataTable(wb,
#'     x = data.frame(style), sheet = 1,
#'     tableStyle = style, startRow = 4, startCol = i * 3 - 2
#'   )
#' }
#'
#' for (i in 1:11) {
#'   style <- paste0("TableStyleDark", i)
#'   writeDataTable(wb,
#'     x = data.frame(style), sheet = 1,
#'     tableStyle = style, startRow = 7, startCol = i * 3 - 2
#'   )
#' }
#'
#' ## openXL(wb)
#' \dontrun{
#' wb_save(wb, path = "tableStylesGallery.xlsx", overwrite = TRUE)
#' }
#'
writeDataTable <- function(wb, sheet, x,
  startCol = 1,
  startRow = 1,
  xy = NULL,
  colNames = TRUE,
  rowNames = FALSE,
  tableStyle = "TableStyleLight9",
  tableName = NULL,
  withFilter = TRUE,
  sep = ", ",
  stack = FALSE,
  firstColumn = FALSE,
  lastColumn = FALSE,
  bandedRows = TRUE,
  bandedCols = FALSE) {
  if (!is.null(xy)) {
    if (length(xy) != 2) {
      stop("xy parameter must have length 2")
    }
    startCol <- xy[[1]]
    startRow <- xy[[2]]
  }

  ## Input validating
  assert_workbook(wb)
  assert_class(x, "data.frame")

  # TODO sipmlify these --
  if (!is.logical(colNames)) stop("colNames must be a logical.")
  if (!is.logical(rowNames)) stop("rowNames must be a logical.")

  if (!is.logical(withFilter)) stop("withFilter must be a logical.")
  if ((!is.character(sep)) || (length(sep) != 1)) stop("sep must be a character vector of length 1")

  if (!is.logical(firstColumn)) stop("firstColumn must be a logical.")
  if (!is.logical(lastColumn)) stop("lastColumn must be a logical.")
  if (!is.logical(bandedRows)) stop("bandedRows must be a logical.")
  if (!is.logical(bandedCols)) stop("bandedCols must be a logical.")

  if (is.null(tableName)) {
    tableName <- paste0("Table", as.character(length(wb$tables) + 1L))
  } else {
    tableName <- wb_validate_table_name(wb, tableName)
  }
  op <- openxlsx_options()
  on.exit(options(op), add = TRUE)

  ## convert startRow and startCol
  if (!is.numeric(startCol)) {
    startCol <- col2int(startCol)
  }
  startRow <- as.integer(startRow)


  showColNames <- colNames

  if (colNames) {
    col_names <- colnames(x)
    if (any(duplicated(tolower(col_names)))) {
      stop("Column names of x must be case-insensitive unique.")
    }

    ## zero char names are invalid
    char0 <- nchar(col_names) == 0
    if (any(char0)) {
      col_names[char0] <- colnames(x)[char0] <- paste0("Column", which(char0))
    }
  } else {
    col_names <- paste0("Column", seq_along(x))
    names(x) <- col_names
  }

  if (rowNames) {
    nam <- names(x)
    x <- cbind(rownames(x), x)
    names(x) <- c("rownames", nam)
  }

  ## If 0 rows append a blank row
  validNames <- c("none", paste0("TableStyleLight", 1:21), paste0("TableStyleMedium", 1:28), paste0("TableStyleDark", 1:11))
  if (!tolower(tableStyle) %in% tolower(validNames)) {
    stop("Invalid table style.")
  } else {
    tableStyle <- grep(paste0("^", tableStyle, "$"), validNames, ignore.case = TRUE, value = TRUE)
  }

  tableStyle <- na.omit(tableStyle)
  if (length(tableStyle) == 0) {
    stop("Unknown table style.")
  }

  ## If zero rows, append an empty row (prevent XML from corrupting)
  if (nrow(x) == 0) {
    x <- rbind(as.data.frame(x), matrix("", nrow = 1, ncol = ncol(x), dimnames = list(character(), colnames(x))))
    names(x) <- colNames
  }

  ref1 <- paste0(int2col(startCol), startRow)
  ref2 <- paste0(int2col(startCol + ncol(x) - 1), startRow + nrow(x))
  ref <- paste(ref1, ref2, sep = ":")

  ## check not overwriting another table
  wb_check_overwrite_tables(
    wb,
    sheet = sheet,
    # header
    new_rows = c(startRow, startRow + nrow(x) - 1L + 1L),
    new_cols = c(startCol, startCol + ncol(x) - 1L)
  )

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
      x[is_hyperlink] <- makeHyperlinkString(text = x[is_hyperlink])
      class(x[is_hyperlink]) <- c("character", "hyperlink")
    } else {
      # check should be in makeHyperlinkString and that apply should not be required either
      if (!any(grepl("^(=|)HYPERLINK\\(", x[is_hyperlink], ignore.case = TRUE))) {
        x[is_hyperlink] <- apply(x[is_hyperlink], 1, FUN=function(str) makeHyperlinkString(text = str))
      }
      class(x[,is_hyperlink]) <- c("character", "hyperlink")
    }
  }


  ## write data to worksheet
  # TODO writeData2 should be wb$writeData
  wb <- writeData2(
    wb =  wb,
    sheet = sheet,
    data = x,
    colNames = TRUE,
    startRow = startRow,
    startCol = startCol
  )

  ## replace invalid XML characters
  col_names <- replaceIllegalCharacters(colnames(x))

  ## create table.xml and assign an id to worksheet tables
  wb$buildTable(
    sheet = sheet,
    colNames = col_names,
    ref = ref,
    showColNames = showColNames,
    tableStyle = tableStyle,
    tableName = tableName,
    withFilter = withFilter,
    totalsRowCount = 0L,
    showFirstColumn = firstColumn,
    showLastColumn = lastColumn,
    showRowStripes = bandedRows,
    showColumnStripes = bandedCols
  )
}
