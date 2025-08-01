# `write_xlsx()` ---------------------------------------------------------------
#' Write data to an xlsx file
#'
#' Write a data frame or list of data frames to an xlsx file.
#'
#' columns of `x` with class `Date` or `POSIXt` are automatically
#' styled as dates and datetimes respectively.
#'
#' @param x An object or a list of objects that can be handled by [wb_add_data()] to write to file.
#' @param file An optional xlsx file name. If no file is passed, the object is not written to disk and only a workbook object is returned.
#' @param as_table If `TRUE`, will write as a data table, instead of data.
#' @inheritDotParams wb_workbook creator
#' @inheritDotParams wb_add_worksheet sheet grid_lines tab_color zoom
#' @inheritDotParams wb_add_data_table start_col start_row col_names row_names na.strings total_row
#' @inheritDotParams wb_add_data start_col start_row col_names row_names na.strings
#' @inheritDotParams wb_freeze_pane first_active_row first_active_col first_row first_col
#' @inheritDotParams wb_set_col_widths widths
#' @inheritDotParams wb_save overwrite
#' @inheritDotParams wb_set_base_font font_size font_color font_name
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
#'
#' # set base font color to automatic so LibreOffice dark mode works as expected
#' write_xlsx(l, temp_xlsx(), font_color = wb_color(auto = TRUE))
#' @export
write_xlsx <- function(x, file, as_table = FALSE, ...) {


  ## set scientific notation penalty

  arguments <- c(ls(), "creator", "sheet", "sheet_name", "grid_lines",
    "tab_color", "tab_colour",
    "zoom", "header", "footer", "even_header", "even_footer", "first_header",
    "first_footer", "start_col", "start_row", "total_row",
    "col.names", "row.names", "col_names", "row_names", "table_style",
    "table_name", "with_filter", "first_active_row", "first_active_col",
    "first_row", "first_col", "col_widths", "na.strings",
    "overwrite", "title", "subject", "category",
    "font_size", "font_color", "font_name",
    "flush", "widths"
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
    getOption("openxlsx2.creator",
              default = Sys.getenv("USERNAME", unset = Sys.getenv("USER")))
    # USERNAME should be present for Windows and Linux (USER on Mac)

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

  #---wb_set_base_font---#
  ## font_size = 11,
  ## font_color = wb_color(theme = "1"),
  ## font_name = "Aptos Narrow"

  if (!is.logical(as_table)) {
    stop("as_table must be a logical.")
  }


  sheetName <- "Sheet 1"
  if (any(c("sheet_name", "sheet") %in% names(params))) {

    sheetName <- as.character(params$sheet_name %||% params$sheet)

    if (any(nchar(params$sheet_name) > 31)) {
      stop("sheet_name too long! Max length is 31 characters.")
    }

    if (inherits(x, "list") && (length(sheetName) == length(x))) {
      names(x) <- sheetName
    }
  }

  tabColor <- NULL
  if ("tab_color" %in% names(params) || "tab_colour" %in% names(params)) {
    tabColor <- params$tab_color %||% params$tab_colour
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

  # unable to add col_widths to the roxygen2 array in ..., therefore
  # accept widths as well. #1361
  if ("widths" %in% names(params)) {
    params$col_widths <- params$widths
  }

  colWidths <- NULL
  if ("col_widths" %in% names(params)) {
    colWidths <- params$col_widths
    if (anyNA(colWidths)) colWidths[is.na(colWidths)] <- 8.43
  }

  table_name <- NULL
  if ("table_name" %in% names(params)) {
    table_name <- params$table_name
  }

  tableStyle <- "TableStyleLight9"
  if ("table_style" %in% names(params)) {
    tableStyle <- params$table_style
  }

  totalRow <- FALSE
  if ("total_row" %in% names(params)) {
    totalRow <- params$total_row
  }

  na.strings <-
    if ("na.strings" %in% names(params)) {
      params$na.strings
    } else if (!is.null(getOption("openxlsx2.na.strings"))) {
      getOption("openxlsx2.na.strings")
    } else {
      na_strings()
    }

  # Get base font parameters if provided
  font_args <- list()
  if ("font_size" %in% names(params)) {
    font_args$font_size <- params$font_size
  }
  if ("font_color" %in% names(params)) {
    font_args$font_color <- params$font_color
  }
  if ("font_name" %in% names(params)) {
    font_args$font_name <- params$font_name
  }

  # Flush stream file to disk
  flush <- FALSE
  if ("flush" %in% names(params)) {
    flush <- params$flush
  }


  ## create new Workbook object
  wb <- wb_workbook(creator = creator, title = title, subject = subject, category = category)

  # Set base font if any parameters were provided
  if (length(font_args)) {
    font_args$wb <- wb
    wb <- do.call(wb_set_base_font, font_args)
  }


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

  if (!is.null(table_name) && length(table_name) != nSheets) {
    table_name <- c(table_name, paste0(table_name[1], seq_len(nSheets)[-nSheets]))
  }

  if (length(tableStyle) != nSheets) {
    tableStyle <- rep_len(tableStyle, length.out = nSheets)
  }

  if (length(colWidths) != nSheets) {
    if (!is.null(colWidths))
      colWidths <- rep_len(colWidths, length.out = nSheets)
  }

  for (i in seq_len(nSheets)) {
    wb$add_worksheet(nms[[i]], grid_lines = gridLines[i], tab_color = tabColor[i], zoom = zoom[i])

    if (as_table[i]) {
      # add data table??
      do_write_datatable(
        wb          = wb,
        sheet       = i,
        x           = x[[i]],
        start_col   = startCol[[i]],
        start_row   = startRow[[i]],
        col_names   = colNames[[i]],
        row_names   = rowNames[[i]],
        table_style = tableStyle[[i]],
        table_name  = table_name[[i]],
        with_filter = withFilter[[i]],
        na.strings  = na.strings,
        total_row   = totalRow
      )
    } else {
      # TODO add_data()?
      do_write_data(
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
  if ("first_active_row" %in% names(params)) {
    firstActiveRow <- params$first_active_row
    freeze_pane <- TRUE
    if (length(firstActiveRow) != nSheets) {
      firstActiveRow <- rep_len(firstActiveRow, length.out = nSheets)
    }
  }

  firstActiveCol <- rep_len(1L, length.out = nSheets)
  if ("first_active_col" %in% names(params)) {
    firstActiveCol <- params$first_active_col
    freeze_pane <- TRUE
    if (length(firstActiveCol) != nSheets) {
      firstActiveCol <- rep_len(firstActiveCol, length.out = nSheets)
    }
  }

  firstRow <- rep_len(FALSE, length.out = nSheets)
  if ("first_row" %in% names(params)) {
    firstRow <- params$first_row
    freeze_pane <- TRUE
    if (inherits(x, "list") && (length(firstRow) != nSheets)) {
      firstRow <- rep_len(firstRow, length.out = nSheets)
    }
  }

  firstCol <- rep_len(FALSE, length.out = nSheets)
  if ("first_col" %in% names(params)) {
    firstCol <- params$first_col
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
        first_active_row = firstActiveRow[i],
        first_active_col = firstActiveCol[i],
        first_row = firstRow[i],
        first_col = firstCol[i]
      )
    }
  }

  if (!missing(file))
    wb_save(wb, file = file, overwrite = overwrite, flush = flush)

  invisible(wb)
}
