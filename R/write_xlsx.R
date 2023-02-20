#' @name write_xlsx
#' @title write data to an xlsx file
#' @description write a data.frame or list of data.frames to an xlsx file
#' @param x object or a list of objects that can be handled by [write_data()] to write to file
#' @param file xlsx file name
#' @param asTable write using write_datatable as opposed to write_data
#' @param ... optional parameters to pass to functions:
#' \itemize{
#'   \item{[wb_workbook()]}
#'   \item{[wb_add_worksheet()]}
#'   \item{[wb_add_data()]}
#'   \item{[wb_freeze_pane]}
#'   \item{[wb_save()]}
#' }
#'
#' see details.
#' @details Optional parameters are:
#'
#' **wb_workbook Parameters**
#' \itemize{
#'   \item{**creator**}{ A string specifying the workbook author}
#' }
#'
#' **wb_add_worksheet() Parameters**
#' \itemize{
#'   \item{**sheetName**}{ Name of the worksheet}
#'   \item{**gridLines**}{ A logical. If `FALSE`, the worksheet grid lines will be hidden.}
#'   \item{**tabColor**}{ Color of the worksheet tab. A valid color (belonging to colors())
#'   or a valid hex color beginning with "#".}
#'   \item{**zoom**}{ A numeric between 10 and 400. Worksheet zoom level as a percentage.}
#' }
#'
#' **write_data/write_datatable Parameters**
#' \itemize{
#'   \item{**startCol**}{ A vector specifying the starting column(s) to write df}
#'   \item{**startRow**}{ A vector specifying the starting row(s) to write df}
#'   \item{**xy**}{ An alternative to specifying startCol and startRow individually.
#'  A vector of the form c(startCol, startRow)}
#'   \item{**colNames or col.names**}{ If `TRUE`, column names of x are written.}
#'   \item{**rowNames or row.names**}{ If `TRUE`, row names of x are written.}
#'   \item{**na.string**} {If not NULL NA values are converted to this string in Excel. Defaults to NULL.}
#' }
#'
#' **freezePane Parameters**
#' \itemize{
#'   \item{**firstActiveRow**} {Top row of active region to freeze pane.}
#'   \item{**firstActiveCol**} {Furthest left column of active region to freeze pane.}
#'   \item{**firstRow**} {If `TRUE`, freezes the first row (equivalent to firstActiveRow = 2)}
#'   \item{**firstCol**} {If `TRUE`, freezes the first column (equivalent to firstActiveCol = 2)}
#' }
#'
#' **colWidths Parameters**
#' \itemize{
#'   \item{**colWidths**} {May be a single value for all columns (or "auto"), or a list of vectors that will be recycled for each sheet (see examples)}
#' }
#'
#'
#' **wb_save Parameters**
#' \itemize{
#'   \item{**overwrite**}{ Overwrite existing file (Defaults to TRUE as with write.table)}
#' }
#'
#'
#' columns of x with class Date or POSIXt are automatically
#' styled as dates and datetimes respectively.
#' @seealso [wb_add_worksheet()], [write_data()]
#' @return A workbook object
#' @examples
#'
#' ## write to working directory
#' write_xlsx(iris, file = temp_xlsx(), colNames = TRUE)
#'
#' write_xlsx(iris,
#'   file = temp_xlsx(),
#'   colNames = TRUE
#' )
#'
#' ## Lists elements are written to individual worksheets, using list names as sheet names if available
#' l <- list("IRIS" = iris, "MTCATS" = mtcars, matrix(runif(1000), ncol = 5))
#' write_xlsx(l, temp_xlsx(), colWidths = c(NA, "auto", "auto"))
#'
#' ## different sheets can be given different parameters
#' write_xlsx(l, temp_xlsx(),
#'   startCol = c(1, 2, 3), startRow = 2,
#'   asTable = c(TRUE, TRUE, FALSE), withFilter = c(TRUE, FALSE, FALSE)
#' )
#'
#' # specify column widths for multiple sheets
#' write_xlsx(l, temp_xlsx(), colWidths = 20)
#' write_xlsx(l, temp_xlsx(), colWidths = list(100, 200, 300))
#' write_xlsx(l, temp_xlsx(), colWidths = list(rep(10, 5), rep(8, 11), rep(5, 5)))
#'
#' @export
write_xlsx <- function(x, file, asTable = FALSE, ...) {


  ## set scientific notation penalty

  params <- list(...)

  ## Possible parameters

  #---wb_workbook---#
  ## creator

  #---add_worksheet---#
  ## sheetName
  ## gridLines
  ## tabColor = NULL
  ## zoom = 100
  ## header = NULL
  ## footer = NULL
  ## evenHeader = NULL
  ## evenFooter = NULL
  ## firstHeader = NULL
  ## firstFooter = NULL

  #---write_data---#
  ## startCol = 1,
  ## startRow = 1,
  ## xy = NULL,
  ## colNames = TRUE,
  ## rowNames = FALSE,
  ## na.strings = NULL

  #----write_datatable---#
  ## startCol = 1
  ## startRow = 1
  ## xy = NULL
  ## colNames = TRUE
  ## rowNames = FALSE
  ## tableStyle = "TableStyleLight9"
  ## tableName = NULL
  ## withFilter = TRUE

  #---freezePane---#
  ## firstActiveRow = NULL
  ## firstActiveCol = NULL
  ## firstRow = FALSE
  ## firstCol = FALSE


  #---wb_save---#
  #   overwrite = TRUE

  if (!is.logical(asTable)) {
    stop("asTable must be a logical.")
  }

  creator <- if ("creator" %in% names(params)) params$creator else ""
  title <- params$title ### will return NULL of not exist
  subject <- params$subject ### will return NULL of not exist
  category <- params$category ### will return NULL of not exist


  sheetName <- "Sheet 1"
  if ("sheetName" %in% names(params)) {
    if (any(nchar(params$sheetName) > 31)) {
      stop("sheetName too long! Max length is 31 characters.")
    }

    sheetName <- as.character(params$sheetName)

    if (inherits(x, "list") && (length(sheetName) == length(x))) {
      names(x) <- sheetName
    }
  }

  tabColor <- NULL
  if ("tabColor" %in% names(params)) {
    tabColor <- params$tabColor
  } else if ("tabColor" %in% names(params)) {
    tabColor <- params$tabColor
  }

  zoom <- 100
  if ("zoom" %in% names(params)) {
    if (is.numeric(params$zoom)) {
      zoom <- params$zoom
    } else {
      stop("zoom must be numeric")
    }
  }

  ## wb_add_worksheet()
  gridLines <- TRUE
  if ("gridLines" %in% names(params)) {
    if (all(is.logical(params$gridLines))) {
      gridLines <- params$gridLines
    } else {
      stop("Argument gridLines must be TRUE or FALSE")
    }
  }

  overwrite <- TRUE
  if ("overwrite" %in% names(params)) {
    if (is.logical(params$overwrite)) {
      overwrite <- params$overwrite
    } else {
      stop("Argument overwrite must be TRUE or FALSE")
    }
  }


  withFilter <- TRUE
  if ("withFilter" %in% names(params)) {
    if (is.logical(params$withFilter)) {
      withFilter <- params$withFilter
    } else {
      stop("Argument withFilter must be TRUE or FALSE")
    }
  }

  startRow <- 1
  if ("startRow" %in% names(params)) {
    if (all(params$startRow > 0)) {
      startRow <- params$startRow
    } else {
      stop("startRow must be a positive integer")
    }
  }

  startCol <- 1
  if ("startCol" %in% names(params)) {
    if (all(col2int(params$startCol) > 0)) {
      startCol <- params$startCol
    } else {
      stop("startCol must be a positive integer")
    }
  }

  colNames <- TRUE
  if ("colNames" %in% names(params)) {
    if (is.logical(params$colNames)) {
      colNames <- params$colNames
    } else {
      stop("Argument colNames must be TRUE or FALSE")
    }
  }

  ## to be consistent with write.csv
  if ("col.names" %in% names(params)) {
    if (is.logical(params$col.names)) {
      colNames <- params$col.names
    } else {
      stop("Argument col.names must be TRUE or FALSE")
    }
  }


  rowNames <- FALSE
  if ("rowNames" %in% names(params)) {
    if (is.logical(params$rowNames)) {
      rowNames <- params$rowNames
    } else {
      stop("Argument colNames must be TRUE or FALSE")
    }
  }

  ## to be consistent with write.csv
  if ("row.names" %in% names(params)) {
    if (is.logical(params$row.names)) {
      rowNames <- params$row.names
    } else {
      stop("Argument row.names must be TRUE or FALSE")
    }
  }

  xy <- NULL
  if ("xy" %in% names(params)) {
    if (length(params$xy) != 2) {
      stop("xy parameter must have length 2")
    }
    xy <- params$xy
  }

  colWidths <- NULL
  if ("colWidths" %in% names(params)) {
    colWidths <- params$colWidths
    if (any(is.na(colWidths))) colWidths[is.na(colWidths)] <- 8.43
  }


  tableStyle <- "TableStyleLight9"
  if ("tableStyle" %in% names(params)) {
    tableStyle <- params$tableStyle
  }

  na.strings <- params$na.strings %||% na_strings()


  ## create new Workbook object
  wb <- wb_workbook(creator = creator, title = title, subject = subject, category = category)


  ## If a list is supplied write to individual worksheets using names if available
  if (!inherits(x, "list"))
    x <- list(x)

  nms <- names(x)
  nSheets <- length(x)

  if (is.null(nms)) {
    nms <- paste("Sheet", seq_len(nSheets))
  } else if (any("" %in% nms)) {
    nms[nms == ""] <- paste("Sheet", (seq_len(nSheets))[nms %in% ""])
  } else {
    nms <- make.unique(nms)
  }

  if (any(nchar(nms) > 31)) {
    warning("Truncating list names to 31 characters.")
    nms <- substr(nms, 1, 31)
  }

  ## make all inputs as long as the list
  if (!is.null(tabColor)) {
    if (length(tabColor) != nSheets) {
      tabColor <- rep_len(tabColor, length.out = nSheets)
    }
  }

  if (length(zoom) != nSheets) {
    zoom <- rep_len(zoom, length.out = nSheets)
  }

  if (length(gridLines) != nSheets) {
    gridLines <- rep_len(gridLines, length.out = nSheets)
  }

  if (length(withFilter) != nSheets) {
    withFilter <- rep_len(withFilter, length.out = nSheets)
  }

  if (length(colNames) != nSheets) {
    colNames <- rep_len(colNames, length.out = nSheets)
  }

  if (length(rowNames) != nSheets) {
    rowNames <- rep_len(rowNames, length.out = nSheets)
  }

  if (length(startRow) != nSheets) {
    startRow <- rep_len(startRow, length.out = nSheets)
  }

  if (length(startCol) != nSheets) {
    startCol <- rep_len(startCol, length.out = nSheets)
  }

  if (length(asTable) != nSheets) {
    asTable <- rep_len(asTable, length.out = nSheets)
  }

  if (length(tableStyle) != nSheets) {
    tableStyle <- rep_len(tableStyle, length.out = nSheets)
  }

  if (length(colWidths) != nSheets) {
    if (!is.null(colWidths))
      colWidths <- rep_len(colWidths, length.out = nSheets)
  }

  for (i in seq_len(nSheets)) {
    wb$add_worksheet(nms[[i]], gridLines = gridLines[i], tabColor = tabColor[i], zoom = zoom[i])

    if (asTable[i]) {
      write_datatable(
        wb = wb,
        sheet = i,
        x = x[[i]],
        startCol = startCol[[i]],
        startRow = startRow[[i]],
        xy = xy,
        colNames = colNames[[i]],
        rowNames = rowNames[[i]],
        tableStyle = tableStyle[[i]],
        tableName = NULL,
        withFilter = withFilter[[i]],
        na.strings = na.strings
      )
    } else {
      write_data(
        wb = wb,
        sheet = i,
        x = x[[i]],
        startCol = startCol[[i]],
        startRow = startRow[[i]],
        xy = xy,
        colNames = colNames[[i]],
        rowNames = rowNames[[i]],
        na.strings = na.strings
      )
    }

    # colWidth is not required for the output
    if (!is.null(colWidths)) {
      cols <- seq_len(NCOL(x[[i]])) + startCol[[i]] - 1L
      if (identical(colWidths[[i]], "auto")) {
        wb$set_col_widths(sheet = i, cols = cols, widths = "auto")
      } else if (!identical(colWidths[[i]], "")) {
        wb$set_col_widths(sheet = i, cols = cols, widths = colWidths[[i]])
      }
    }
  }

  ### --Freeze Panes---###
  ## firstActiveRow = NULL
  ## firstActiveCol = NULL
  ## firstRow = FALSE
  ## firstCol = FALSE

  freeze_pane <- FALSE
  firstActiveRow <- rep_len(1L, length.out = nSheets)
  if ("firstActiveRow" %in% names(params)) {
    firstActiveRow <- params$firstActiveRow
    freeze_pane <- TRUE
    if (length(firstActiveRow) != nSheets) {
      firstActiveRow <- rep_len(firstActiveRow, length.out = nSheets)
    }
  }

  firstActiveCol <- rep_len(1L, length.out = nSheets)
  if ("firstActiveCol" %in% names(params)) {
    firstActiveCol <- params$firstActiveCol
    freeze_pane <- TRUE
    if (length(firstActiveCol) != nSheets) {
      firstActiveCol <- rep_len(firstActiveCol, length.out = nSheets)
    }
  }

  firstRow <- rep_len(FALSE, length.out = nSheets)
  if ("firstRow" %in% names(params)) {
    firstRow <- params$firstRow
    freeze_pane <- TRUE
    if (inherits(x, "list") && (length(firstRow) != nSheets)) {
      firstRow <- rep_len(firstRow, length.out = nSheets)
    }
  }

  firstCol <- rep_len(FALSE, length.out = nSheets)
  if ("firstCol" %in% names(params)) {
    firstCol <- params$firstCol
    freeze_pane <- TRUE
    if (inherits(x, "list") && (length(firstCol) != nSheets)) {
      firstCol <- rep_len(firstCol, length.out = nSheets)
    }
  }

  if (freeze_pane) {
    for (i in seq_len(nSheets)) {
      wb <- wb_freeze_pane(
        wb = wb,
        sheet = i,
        firstActiveRow = firstActiveRow[i],
        firstActiveCol = firstActiveCol[i],
        firstRow = firstRow[i],
        firstCol = firstCol[i]
      )
    }
  }

  wb_save(wb, path = file, overwrite = overwrite)
  invisible(wb)
}
