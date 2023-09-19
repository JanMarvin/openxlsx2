# `write_xlsx()` ---------------------------------------------------------------
#' Write data to an xlsx file
#'
#' Write a data frame or list of data frames to an xlsx file.
#'
#' columns of `x` with class `Date` or `POSIXt` are automatically
#' styled as dates and datetimes respectively.
#'
#' @param x An object or a list of objects that can be handled by [wb_add_data()] to write to file
#' @param file An xlsx file name
#' @param as_table If `TRUE`, will write as a data table, instead of data.
#' @inheritDotParams wb_workbook creator
#' @inheritDotParams wb_add_worksheet sheet grid_lines tab_color zoom
#' @inheritDotParams wb_add_data_table start_col start_row col_names row_names na.strings
#' @inheritDotParams wb_add_data start_col start_row col_names row_names na.strings
#' @inheritDotParams wb_freeze_pane first_active_row first_active_col first_row first_col
#' @inheritDotParams wb_set_col_widths widths
#' @inheritDotParams wb_save overwrite
#' @return A workbook object
#' @examples
#' ## write to working directory
#' write_xlsx(iris, file = temp_xlsx(), col_names = TRUE)
#'
#' write_xlsx(iris,
#'   file = temp_xlsx(),
#'   col_names = TRUE
#' )
#'
#' ## Lists elements are written to individual worksheets, using list names as sheet names if available
#' l <- list("IRIS" = iris, "MTCARS" = mtcars, matrix(runif(1000), ncol = 5))
#' write_xlsx(l, temp_xlsx(), col_widths = c(NA, "auto", "auto"))
#'
#' ## different sheets can be given different parameters
#' write_xlsx(l, temp_xlsx(),
#'   start_col = c(1, 2, 3), start_row = 2,
#'   as_table = c(TRUE, TRUE, FALSE), with_filter = c(TRUE, FALSE, FALSE)
#' )
#'
#' # specify column widths for multiple sheets
#' write_xlsx(l, temp_xlsx(), col_widths = 20)
#' write_xlsx(l, temp_xlsx(), col_widths = list(100, 200, 300))
#' write_xlsx(l, temp_xlsx(), col_widths = list(rep(10, 5), rep(8, 11), rep(5, 5)))
#' @export
write_xlsx <- function(x, file, as_table = FALSE, ...) {


  ## set scientific notation penalty

  arguments <- c(ls(), "creator",  "sheet_name", "grid_lines",
    "tab_color", "tab_colour",
    "zoom", "header", "footer", "even_header", "even_footer", "first_header",
    "first_footer", "start_col", "start_row",
    "col.names", "row.names", "col_names", "row_names", "table_style",
    "table_name", "with_filter", "first_active_row", "first_active_col",
    "first_row", "first_col", "col_widths", "na.strings",
    "overwrite", "title", "subject", "category"
  )

  params <- list(...)

  # we need them in params
  params <- standardize_case_names(params, arguments = arguments, return = TRUE)

  # and in global env for `asTable`
  standardize_case_names(..., arguments = arguments)

  ## Possible parameters

  #---wb_workbook---#
  ## creator
  creator <- NULL
  if ("creator" %in% names(params)) {
    creator <- params$creator
  }

  title <- NULL
  if ("subject" %in% names(params)) {
    title <- params$title
  }

  subject <- NULL
  if ("subject" %in% names(params)) {
    subject <- params$subject
  }

  category <- NULL
  if ("category" %in% names(params)) {
    category <- params$category
  }

  creator <- creator %||%
    getOption("openxlsx2.creator") %||%
    # USERNAME may only be present for windows
    Sys.getenv("USERNAME", Sys.getenv("USER"))

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
  ## colNames = TRUE,
  ## rowNames = FALSE,
  ## na.strings = NULL

  #----write_datatable---#
  ## startCol = 1
  ## startRow = 1
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

  if (!is.logical(as_table)) {
    stop("as_table must be a logical.")
  }


  sheetName <- "Sheet 1"
  if ("sheet_name" %in% names(params)) {
    if (any(nchar(params$sheet_name) > 31)) {
      stop("sheet_name too long! Max length is 31 characters.")
    }

    sheetName <- as.character(params$sheet_name)

    if (inherits(x, "list") && (length(sheetName) == length(x))) {
      names(x) <- sheetName
    }
  }

  tabColor <- NULL
  if ("tab_color" %in% names(params)) {
    tabColor <- params$tab_color
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
  if ("grid_lines" %in% names(params)) {
    if (all(is.logical(params$grid_lines))) {
      gridLines <- params$grid_lines
    } else {
      stop("Argument grid_lines must be TRUE or FALSE")
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
  if ("with_filter" %in% names(params)) {
    if (is.logical(params$with_filter)) {
      withFilter <- params$with_filter
    } else {
      stop("Argument with_filter must be TRUE or FALSE")
    }
  }

  startRow <- 1
  if ("start_row" %in% names(params)) {
    if (all(params$start_row > 0)) {
      startRow <- params$start_row
    } else {
      stop("start_row must be a positive integer")
    }
  }

  startCol <- 1
  if ("start_col" %in% names(params)) {
    if (all(col2int(params$start_col) > 0)) {
      startCol <- params$start_col
    } else {
      stop("start_col must be a positive integer")
    }
  }

  colNames <- TRUE
  if ("col_names" %in% names(params)) {
    if (is.logical(params$col_names)) {
      colNames <- params$col_names
    } else {
      stop("Argument col_names must be TRUE or FALSE")
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
  if ("row_names" %in% names(params)) {
    if (is.logical(params$row_names)) {
      rowNames <- params$row_names
    } else {
      stop("Argument row_names must be TRUE or FALSE")
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

  colWidths <- NULL
  if ("col_widths" %in% names(params)) {
    colWidths <- params$col_widths
    if (any(is.na(colWidths))) colWidths[is.na(colWidths)] <- 8.43
  }


  tableStyle <- "TableStyleLight9"
  if ("table_style" %in% names(params)) {
    tableStyle <- params$table_style
  }

  na.strings <-
    if ("na.strings" %in% names(params)) {
      params$na.strings
    } else {
      na_strings()
    }

  ## create new Workbook object
  wb <- wb_workbook(creator = creator, title = title, subject = subject, category = category)


  ## If a list is supplied write to individual worksheets using names if available
  if (!inherits(x, "list")) {
    x <- list(x)
    names(x) <- sheetName
  }

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

  if (length(as_table) != nSheets) {
    as_table <- rep_len(as_table, length.out = nSheets)
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

    if (as_table[i]) {
      write_datatable(
        wb          = wb,
        sheet       = i,
        x           = x[[i]],
        start_col   = startCol[[i]],
        start_row   = startRow[[i]],
        col_names   = colNames[[i]],
        row_names   = rowNames[[i]],
        table_style = tableStyle[[i]],
        table_name  = NULL,
        with_filter = withFilter[[i]],
        na.strings  = na.strings
      )
    } else {
      write_data(
        wb = wb,
        sheet = i,
        x = x[[i]],
        start_col  = startCol[[i]],
        start_row  = startRow[[i]],
        col_names  = colNames[[i]],
        row_names  = rowNames[[i]],
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
      # TODO replace with snake case arguments
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

  wb_save(wb, file = file, overwrite = overwrite)
  invisible(wb)
}
