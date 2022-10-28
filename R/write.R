#' Replace data cell(s)
#'
#' Minimal invasive update of cell(s) inside of imported workbooks.
#'
#' @param x cc dataframe of the updated cells
#' @param wb the workbook you want to update
#' @param sheet the sheet you want to update
#' @param cell the cell you want to update in Excel connotation e.g. "A1"
#' @param colNames if TRUE colNames are passed down
#' @param removeCellStyle keep the cell style?
#' @param na.strings optional na.strings argument. if missing #N/A is used. If NULL no cell value is written, if character or numeric this is written (even if NA is part of numeric data)
#'
#' @keywords internal
#' @noRd
update_cell <- function(x, wb, sheet, cell, colNames = FALSE,
                        removeCellStyle = FALSE, na.strings) {

  sheet_id <- wb$validate_sheet(sheet)

  dims <- dims_to_dataframe(cell, fill = TRUE)
  rows <- rownames(dims)

  cells_needed <- unname(unlist(dims))


  # 1) pull sheet to modify from workbook; 2) modify it; 3) push it back
  cc  <- wb$worksheets[[sheet_id]]$sheet_data$cc
  row_attr <- wb$worksheets[[sheet_id]]$sheet_data$row_attr

  # workbooks contain only entries for values currently present.
  # if A1 is filled, B1 is not filled and C1 is filled the sheet will only
  # contain fields A1 and C1.
  cells_in_wb <- cc$r
  rows_in_wb <- row_attr$r

  # check if there are rows not available
  if (!all(rows %in% rows_in_wb)) {
    # message("row(s) not in workbook")

    missing_rows <- rows[!rows %in% rows_in_wb]

    # new row_attr
    row_attr_missing <- empty_row_attr(n = length(missing_rows))
    row_attr_missing$r <- missing_rows

    row_attr <- rbind(row_attr, row_attr_missing)

    # order
    row_attr <- row_attr[order(as.numeric(row_attr$r)), ]

    wb$worksheets[[sheet_id]]$sheet_data$row_attr <- row_attr
    # provide output
    rows_in_wb <- row_attr$r

  }

  if (!all(cells_needed %in% cells_in_wb)) {
    # message("cell(s) not in workbook")

    missing_cells <- cells_needed[!cells_needed %in% cells_in_wb]

    # create missing cells
    cc_missing <- create_char_dataframe(names(cc), length(missing_cells))
    cc_missing$r     <- missing_cells
    cc_missing$row_r <- gsub("[[:upper:]]", "", cc_missing$r)
    cc_missing$c_r   <- gsub("[[:digit:]]", "", cc_missing$r)

    # assign to cc
    cc <- rbind(cc, cc_missing)

    # order cc (not really necessary, will be done when saving)
    cc <- cc[order(as.integer(cc[, "row_r"]), col2int(cc[, "c_r"])), ]

    # update dimensions (only required if new cols and rows are added) ------
    all_rows <- as.numeric(unique(cc$row_r))
    all_cols <- col2int(unique(cc$c_r))

    min_cell <- trimws(paste0(int2col(min(all_cols, na.rm = TRUE)), min(all_rows, na.rm = TRUE)))
    max_cell <- trimws(paste0(int2col(max(all_cols, na.rm = TRUE)), max(all_rows, na.rm = TRUE)))

    # i know, i know, i'm lazy
    wb$worksheets[[sheet_id]]$dimension <- paste0("<dimension ref=\"", min_cell, ":", max_cell, "\"/>")
  }

  if (missing(na.strings)) {
    na.strings <- NULL
  }

  if (removeCellStyle) {
    cell_style <- "c_s"
  } else {
    cell_style <- NULL
  }

  replacement <- c("r", cell_style, "c_t", "c_cm", "c_ph", "c_vm", "v",
                   "f", "f_t", "f_ref", "f_ca", "f_si", "is")

  sel <- match(x$r, cc$r)
  cc[sel, replacement] <- x[replacement]

  # avoid missings in cc
  if (any(is.na(cc)))
    cc[is.na(cc)] <- ""

  # push everything back to workbook
  wb$worksheets[[sheet_id]]$sheet_data$cc  <- cc

  wb
}


nmfmt_df <- function(x) {
  data.frame(
    numFmtId = as.character(x),
    fontId = "0",
    fillId = "0",
    borderId = "0",
    xfId = "0",
    applyNumberFormat = "1",
    stringsAsFactors = FALSE
  )
}


#' dummy function to write data
#' @param wb workbook
#' @param sheet sheet
#' @param data data to export
#' @param name If not NULL, a named region is defined.
#' @param colNames include colnames?
#' @param rowNames include rownames?
#' @param startRow row to place it
#' @param startCol col to place it
#' @param applyCellStyle apply styles when writing on the sheet
#' @param removeCellStyle keep the cell style?
#' @param na.strings optional na.strings argument. if missing #N/A is used. If NULL no cell value is written, if character or numeric this is written (even if NA is part of numeric data)
#' @param data_table logical. if `TRUE` and `rowNames = TRUE`, do not write the cell containing  `"_rowNames_"`
#' @details
#' The string `"_openxlsx_NA"` is reserved for `openxlsx2`. If the data frame
#' contains this string, the output will be broken.
#'
#' @examples
#' # create a workbook and add some sheets
#' wb <- wb_workbook()
#'
#' wb$add_worksheet("sheet1")
#' write_data2(wb, "sheet1", mtcars, colNames = TRUE, rowNames = TRUE)
#'
#' wb$add_worksheet("sheet2")
#' write_data2(wb, "sheet2", cars, colNames = FALSE)
#'
#' wb$add_worksheet("sheet3")
#' write_data2(wb, "sheet3", letters)
#'
#' wb$add_worksheet("sheet4")
#' write_data2(wb, "sheet4", as.data.frame(Titanic), startRow = 2, startCol = 2)
#'
#' @export
write_data2 <- function(
    wb,
    sheet,
    data,
    name = NULL,
    colNames = TRUE,
    rowNames = FALSE,
    startRow = 1,
    startCol = 1,
    applyCellStyle = TRUE,
    removeCellStyle = FALSE,
    na.strings,
    data_table = FALSE
  ) {

  if (missing(na.strings)) na.strings <- substitute()

  is_data_frame <- FALSE
  #### prepare the correct data formats for openxml
  dc <- openxlsx2_type(data)

  # convert factor to character
  if (any(dc == openxlsx2_celltype[["factor"]])) {
    is_factor <- dc == openxlsx2_celltype[["factor"]]
    fcts <- names(dc[is_factor])
    data[fcts] <- lapply(data[fcts], as.character)
    dc <- openxlsx2_type(data)
  }

  hconvert_date1904 <- grepl('date1904="1"|date1904="true"',
                             stri_join(unlist(wb$workbook), collapse = ""),
                             ignore.case = TRUE)

  # TODO need to tell excel that we have a date, apply some kind of numFmt
  data <- convertToExcelDate(df = data, date1904 = hconvert_date1904)

  if (inherits(data, "data.frame") || inherits(data, "matrix")) {
    is_data_frame <- TRUE

    if (inherits(data, "data.table")) data <- as.data.frame(data)

    sel <- !dc %in% c(4, 5, 10)
    data[sel] <- lapply(data[sel], as.character)

    # add rownames
    if (rowNames) {
      data <- cbind("_rowNames_" = rownames(data), data)
      dc <- c(c("_rowNames_" = openxlsx2_celltype[["character"]]), dc)
    }

    # add colnames
    if (colNames)
      data <- rbind(colnames(data), data)
  }


  sheetno <- wb_validate_sheet(wb, sheet)
  # message("sheet no: ", sheetno)

  # create a data frame
  if (!is_data_frame)
    data <- as.data.frame(t(data))

  data_nrow <- NROW(data)
  data_ncol <- NCOL(data)

  endRow <- (startRow - 1) + data_nrow
  endCol <- (startCol - 1) + data_ncol

  dims <- paste0(
    int2col(startCol), startRow,
    ":",
    int2col(endCol), endRow
  )

  # TODO writing defined name should handle global and local: localSheetId
  # this requires access to wb$workbook.
  # TODO The check for existing names is in write_data()
  # TODO use wb$add_named_region()
  if (!is.null(name)) {

    ## named region
    ex_names <- regmatches(wb$workbook$definedNames, regexpr('(?<=name=")[^"]+', wb$workbook$definedNames, perl = TRUE))
    ex_names <- replaceXMLEntities(ex_names)

    if (name %in% ex_names) {
      stop(sprintf("Named region with name '%s' already exists!", name))
    } else if (grepl("^[A-Z]{1,3}[0-9]+$", name)) {
      stop("name cannot look like a cell reference.")
    }

    sheet_name <- wb$sheet_names[[sheetno]]
    if (grepl(" ", sheet_name)) sheet_name <- shQuote(sheet_name, "sh")

    sheet_dim <- paste0(sheet_name, "!", dims)

    def_name <- xml_node_create("definedName",
                                xml_children = sheet_dim,
                                xml_attributes = c(name = name))

    wb$workbook$definedNames <- c(wb$workbook$definedNames, def_name)
  }

  # from here on only wb$worksheets is required

  # rtyp character vector per row
  # list(c("A1, ..., "k1"), ...,  c("An", ..., "kn"))
  rtyp <- dims_to_dataframe(dims, fill = TRUE)

  rows_attr <- vector("list", data_nrow)

  # create <rows ...>
  want_rows <- startRow:endRow
  rows_attr <- empty_row_attr(n = length(want_rows))
  rows_attr$r <- rownames(rtyp)

  # original cc data frame
  cc <- empty_sheet_data_cc(n = nrow(data) * ncol(data))


  sel <- which(dc == openxlsx2_celltype[["logical"]])
  for (i in sel) {
    if (colNames) {
      data[-1, i] <- as.integer(as.logical(data[-1, i]))
    } else {
      data[, i] <- as.integer(as.logical(data[, i]))
    }
  }

  sel <- which(dc == openxlsx2_celltype[["character"]]) # character
  if (length(sel)) {
    data[sel][is.na(data[sel])] <- "_openxlsx_NA"
  }

  wide_to_long(
    data,
    dc,
    cc,
    ColNames = colNames,
    start_col = startCol,
    start_row = startRow,
    ref = dims
  )

  # if any v is missing, set typ to 'e'. v is only filled for non character
  # values, but contains a string. To avoid issues, set it to the missing
  # value expression

  ## replace NA, NaN, and Inf
  is_na <- which(cc$is == "<is><t>_openxlsx_NA</t></is>" | cc$v == "NA")
  if (length(is_na)) {
    if (missing(na.strings)) {
      cc[is_na, "v"]   <- "#N/A"
      cc[is_na, "c_t"] <- "e"
      cc[is_na, "is"]  <- ""
    } else {
      cc[is_na, "v"]  <- ""
      if (is.null(na.strings)) {
        # do nothing
      } else {
        cc[is_na, "c_t"] <- "inlineStr"
        cc[is_na, "is"] <- txt_to_is(as.character(na.strings),
                                     no_escapes = TRUE, raw = TRUE)
      }
    }
  }

  is_nan <- which(cc$v == "NaN")
  if (length(is_nan)) {
    cc[is_nan, "v"]   <- "#VALUE!"
    cc[is_nan, "c_t"] <- "e"
  }

  is_inf <- which(cc$v == "-Inf" | cc$v == "Inf")
  if (length(is_inf)) {
    cc[is_inf, "v"]   <- "#NUM!"
    cc[is_inf, "c_t"] <- "e"
  }

  # if rownames = TRUE and data_table = FALSE, remove "_rownames_"
  if (!data_table && rowNames && colNames) {
    cc <- cc[cc$r != rtyp[1, 1], ]
  }

  if (is.null(wb$worksheets[[sheetno]]$sheet_data$cc)) {

    wb$worksheets[[sheetno]]$dimension <- paste0("<dimension ref=\"", dims, "\"/>")

    wb$worksheets[[sheetno]]$sheet_data$row_attr <- rows_attr

    wb$worksheets[[sheetno]]$sheet_data$cc <- cc

  } else {
    # update cell(s)
    # message("update_cell()")

    wb <- update_cell(
      x = cc,
      wb =  wb,
      sheet = sheetno,
      cell = dims,
      colNames = colNames,
      removeCellStyle = removeCellStyle,
      na.strings = na.strings
    )
  }

  ### Begin styles

  if (applyCellStyle) {

    ## create a cell style format for specific types at the end of the existing
    # styles. gets the reference an passes it on.
    get_data_class_dims <- function(data_class) {
      sel <- dc == openxlsx2_celltype[[data_class]]
      sel_cols <- names(rtyp[sel == TRUE])
      sel_rows <- rownames(rtyp)

      # ignore first row if colNames
      if (colNames) sel_rows <- sel_rows[-1]

      paste(
        unname(
          unlist(
            rtyp[rownames(rtyp) %in% sel_rows, sel_cols, drop = FALSE]
          )
        ),
        collapse = ";"
      )
    }

    # if hyperlinks are found, Excel sets something like the following font
    # blue with underline
    if (any(dc == openxlsx2_celltype[["hyperlink"]])) {

      dim_sel <- get_data_class_dims("hyperlink")
      # message("hyperlink: ", dim_sel)

      wb$add_font(
          sheet = sheetno,
          dim = dim_sel,
          color = wb_colour(hex = "FF0000FF"),
          name = wb_get_base_font(wb)$name$val,
          u = "single"
      )
    }

    # options("openxlsx2.numFmt" = NULL)
    if (any(dc == openxlsx2_celltype[["numeric"]])) { # numeric or integer
      if (!is.null(unlist(options("openxlsx2.numFmt")))) {

        numfmt_numeric <- unlist(options("openxlsx2.numFmt"))

        dim_sel <- get_data_class_dims("numeric")
        # message("numeric: ", dim_sel)

        wb$add_numfmt(
          sheet = sheetno,
          dim = dim_sel,
          numfmt = numfmt_numeric
        )
      }
    }
    if (any(dc == openxlsx2_celltype[["short_date"]])) { # Date
      if (is.null(unlist(options("openxlsx2.dateFormat")))) {
        numfmt_dt <- 14
      } else {
        numfmt_dt <- unlist(options("openxlsx2.dateFormat"))
      }

      dim_sel <- get_data_class_dims("short_date")
      # message("short_date: ", dim_sel)

      wb$add_numfmt(
        sheet = sheetno,
        dim = dim_sel,
        numfmt = numfmt_dt
      )
    }
    if (any(dc == openxlsx2_celltype[["long_date"]])) {
      if (is.null(unlist(options("openxlsx2.datetimeFormat")))) {
        numfmt_posix <- 22
      } else {
        numfmt_posix <- unlist(options("openxlsx2.datetimeFormat"))
      }

      dim_sel <- get_data_class_dims("long_date")
      # message("long_date: ", dim_sel)

      wb$add_numfmt(
        sheet = sheetno,
        dim = dim_sel,
        numfmt = numfmt_posix
      )
    }
    if (any(dc == openxlsx2_celltype[["accounting"]])) { # accounting
      if (is.null(unlist(options("openxlsx2.accountingFormat")))) {
        numfmt_accounting <- 4
      } else {
        numfmt_accounting <- unlist(options("openxlsx2.accountingFormat"))
      }

      dim_sel <- get_data_class_dims("accounting")
      # message("accounting: ", dim_sel)

      wb$add_numfmt(
        dim = dim_sel,
        numfmt = numfmt_accounting
      )
    }
    if (any(dc == openxlsx2_celltype[["percentage"]])) { # percentage
      if (is.null(unlist(options("openxlsx2.percentageFormat")))) {
        numfmt_percentage <- 10
      } else {
        numfmt_percentage <- unlist(options("openxlsx2.percentageFormat"))
      }

      dim_sel <- get_data_class_dims("percentage")
      # message("percentage: ", dim_sel)

      wb$add_numfmt(
        sheet = sheetno,
        dim = dim_sel,
        numfmt = numfmt_percentage
      )
    }
    if (any(dc == openxlsx2_celltype[["scientific"]])) {
      if (is.null(unlist(options("openxlsx2.scientificFormat")))) {
        numfmt_scientific <- 48
      } else {
        numfmt_scientific <- unlist(options("openxlsx2.scientificFormat"))
      }

      dim_sel <- get_data_class_dims("scientific")
      # message("scientific: ", dim_sel)

      wb$add_numfmt(
        sheet = sheetno,
        dim = dim_sel,
        numfmt = numfmt_scientific
      )
    }
    if (any(dc == openxlsx2_celltype[["comma"]])) {
      if (is.null(unlist(options("openxlsx2.comma")))) {
        numfmt_comma <- 3
      } else {
        numfmt_comma <- unlist(options("openxlsx2.commaFormat"))
      }

      dim_sel <- get_data_class_dims("comma")
      # message("comma: ", dim_sel)

      wb$add_numfmt(
        sheet = sheetno,
        dim = dim_sel,
        numfmt = numfmt_comma
      )
    }
  }
  ### End styles

  return(wb)
}


#' internal driver function to write_data and write_data_table
#' @name write_datatable
#' @title Write to a worksheet as an Excel table
#' @description Write to a worksheet and format as an Excel table
#' @param wb A Workbook object containing a worksheet.
#' @param sheet The worksheet to write to. Can be the worksheet index or name.
#' @param x A data frame.
#' @param startCol A vector specifying the starting column to write df
#' @param startRow A vector specifying the starting row to write df
#' @param dims Spreadsheet dimensions that will determine startCol and startRow: "A1", "A1:B2", "A:B"
#' @param array A bool if the function written is of type array
#' @param xy An alternative to specifying startCol and startRow individually.
#' A vector of the form c(startCol, startRow)
#' @param colNames If `TRUE`, column names of x are written.
#' @param rowNames If `TRUE`, row names of x are written.
#' @param tableStyle Any excel table style name or "none" (see "formatting" vignette).
#' @param tableName name of table in workbook. The table name must be unique.
#' @param withFilter If `TRUE`, columns with have filters in the first row.
#' @param sep Only applies to list columns. The separator used to collapse list columns to a character vector e.g. sapply(x$list_column, paste, collapse = sep).
#' @param firstColumn logical. If TRUE, the first column is bold
#' @param lastColumn logical. If TRUE, the last column is bold
#' @param bandedRows logical. If TRUE, rows are colour banded
#' @param bandedCols logical. If TRUE, the columns are colour banded
#' @param bandedCols logical. If TRUE, a data table is created
#' @param name If not NULL, a named region is defined.
#' @param applyCellStyle apply styles when writing on the sheet
#' @param removeCellStyle if writing into existing cells, should the cell style be removed?
#' @param na.strings optional na.strings argument. if missing #N/A is used. If NULL no cell value is written, if character or numeric this is written (even if NA is part of numeric data)
#' @noRd
write_data_table <- function(
    wb,
    sheet,
    x,
    startCol = 1,
    startRow = 1,
    dims,
    array = FALSE,
    xy = NULL,
    colNames = TRUE,
    rowNames = FALSE,
    tableStyle = "TableStyleLight9",
    tableName = NULL,
    withFilter = TRUE,
    sep = ", ",
    firstColumn = FALSE,
    lastColumn = FALSE,
    bandedRows = TRUE,
    bandedCols = FALSE,
    name = NULL,
    applyCellStyle = TRUE,
    removeCellStyle = FALSE,
    data_table = FALSE,
    na.strings
) {

  op <- openxlsx2_options()
  on.exit(options(op), add = TRUE)

  ## Input validating
  assert_workbook(wb)
  if (missing(x)) stop("`x` is missing")
  assert_class(colNames, "logical")
  assert_class(rowNames, "logical")
  assert_class(withFilter, "logical")
  if (data_table) assert_class(x, "data.frame")
  assert_class(firstColumn, "logical")
  assert_class(lastColumn, "logical")
  assert_class(bandedRows, "logical")
  assert_class(bandedCols, "logical")

  if (!is.null(dims)) {
    dims <- dims_to_rowcol(dims, as_integer = TRUE)
    startCol <- min(dims[[1]])
    startRow <- min(dims[[2]])
  }

  if (missing(na.strings)) na.strings <- substitute()

  ## common part ---------------------------------------------------------------
  if ((!is.character(sep)) || (length(sep) != 1))
    stop("sep must be a character vector of length 1")

  # TODO clean up when moved into wbWorkbook
  sheet <- wb$.__enclos_env__$private$get_sheet_index(sheet)
  # sheet <- wb$validate_sheet(sheet)

  if (wb$isChartSheet[[sheet]]) stop("Cannot write to chart sheet.")

  ## All input conversions/validations
  if (!is.null(xy)) {
    warning("xy is deprecated. Please use dims instead.")
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

  ## special case - vector of hyperlinks
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
      if (!any(grepl("=([\\s]*?)HYPERLINK\\(", x[is_hyperlink], perl = TRUE))) {
        x[is_hyperlink] <- create_hyperlink(text = x[is_hyperlink])
      }
      class(x[is_hyperlink]) <- c("character", "hyperlink")
    } else {
      # check should be in create_hyperlink and that apply should not be required either
      if (!any(grepl("=([\\s]*?)HYPERLINK\\(", x[is_hyperlink], perl = TRUE))) {
        x[is_hyperlink] <- apply(
          x[is_hyperlink], 1,
          FUN = function(str) create_hyperlink(text = str)
        )
      }
      class(x[, is_hyperlink]) <- c("character", "hyperlink")
    }
  }

  ### Create data frame --------------------------------------------------------

  ## special case - formula
  if (inherits(x, "formula")) {
    x <- data.frame("X" = x, stringsAsFactors = FALSE)
    class(x[[1]]) <- if (array) "array_formula" else "formula"
    colNames <- FALSE
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
  }

  nCol <- ncol(x)
  nRow <- nrow(x)

  ### Beg: Only in data --------------------------------------------------------
  if (!data_table) {

    ## write autoFilter, can only have a single filter per worksheet
    if (withFilter) {
      coords <- data.frame("x" = c(startRow, startRow + nRow + colNames - 1L), "y" = c(startCol, startCol + nCol - 1L))
      ref <- stri_join(get_cell_refs(coords), collapse = ":")

      wb$worksheets[[sheet]]$autoFilter <- sprintf('<autoFilter ref="%s"/>', ref)

      l <- int2col(unlist(coords[, 2]))
      dfn <- sprintf("'%s'!%s", wb$get_sheet_names()[sheet], stri_join("$", l, "$", coords[, 1], collapse = ":"))

      dn <- sprintf('<definedName name="_xlnm._FilterDatabase" localSheetId="%s" hidden="1">%s</definedName>', sheet - 1L, dfn)

      if (length(wb$workbook$definedNames) > 0) {
        ind <- grepl('name="_xlnm._FilterDatabase"', wb$workbook$definedNames)
        if (length(ind) > 0) {
          wb$workbook$definedNames[ind] <- dn
        }
      } else {
        wb$workbook$definedNames <- dn
      }
    }
  }
  ### End: Only in data --------------------------------------------------------

  if (data_table) {
    overwrite_nrows <- 1L
    check_tab_head_only <- FALSE
    error_msg <- "Cannot overwrite existing table with another table"
  } else {
    overwrite_nrows <- colNames
    check_tab_head_only <- TRUE
    error_msg <- "Cannot overwrite table headers. Avoid writing over the header row or see wb_get_tables() & wb_remove_tabless() to remove the table object."
  }

  ## Check not overwriting existing table headers
  wb_check_overwrite_tables(
    wb = wb,
    sheet = sheet,
    new_rows = c(startRow, startRow + nRow - 1L + overwrite_nrows),
    new_cols = c(startCol, startCol + nCol - 1L),
    check_table_header_only = check_tab_head_only,
    error_msg = error_msg
  )

  ## actual driver, the rest should not create data used for writing
  wb <- write_data2(
    wb =  wb,
    sheet = sheet,
    data = x,
    name = name,
    colNames = colNames,
    rowNames = rowNames,
    startRow = startRow,
    startCol = startCol,
    applyCellStyle = applyCellStyle,
    removeCellStyle = removeCellStyle,
    na.strings = na.strings,
    data_table = data_table
  )

  ### Beg: Only in datatable ---------------------------------------------------

  # if rowNames is set, write_data2 has added a rowNames column to the sheet.
  # This has to be handled in colnames and in ref.
  if (data_table) {

    ## replace invalid XML characters
    col_names <- replace_legal_chars(colnames(x))
    if (rowNames) col_names <- c("_rowNames_", col_names)

    ## Table name validation
    if (is.null(tableName)) {
      tableName <- paste0("Table", as.character(NROW(wb$tables) + 1L))
    } else {
      tableName <- wb_validate_table_name(wb, tableName)
    }

    ## If 0 rows append a blank row
    validNames <- c("none", paste0("TableStyleLight", seq_len(21)), paste0("TableStyleMedium", seq_len(28)), paste0("TableStyleDark", seq_len(11)))
    if (!tolower(tableStyle) %in% tolower(validNames)) {
      stop("Invalid table style.")
    } else {
      tableStyle <- grep(paste0("^", tableStyle, "$"), validNames, ignore.case = TRUE, value = TRUE)
    }

    tableStyle <- tableStyle[!is.na(tableStyle)]
    if (length(tableStyle) == 0) {
      stop("Unknown table style.")
    }

    ## If zero rows, append an empty row (prevent XML from corrupting)
    if (nrow(x) == 0) {
      x <- rbind(as.data.frame(x), matrix("", nrow = 1, ncol = nCol, dimnames = list(character(), colnames(x))))
      names(x) <- colNames
    }

    ref1 <- paste0(int2col(startCol), startRow)
    ref2 <- paste0(int2col(startCol + nCol - !rowNames), startRow + nRow)
    ref <- paste(ref1, ref2, sep = ":")

    ## create table.xml and assign an id to worksheet tables
    wb$buildTable(
      sheet = sheet,
      colNames = col_names,
      ref = ref,
      showColNames = colNames,
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

  ### End: Only in datatable ---------------------------------------------------

  return(wb)
}


#' @name write_data
#' @title Write an object to a worksheet
#' @description Write an object to worksheet with optional styling.
#' @param wb A Workbook object containing a worksheet.
#' @param sheet The worksheet to write to. Can be the worksheet index or name.
#' @param x Object to be written. For classes supported look at the examples.
#' @param startCol A vector specifying the starting column to write to.
#' @param startRow A vector specifying the starting row to write to.
#' @param dims Spreadsheet dimensions that will determine startCol and startRow: "A1", "A1:B2", "A:B"
#' @param array A bool if the function written is of type array
#' @param xy An alternative to specifying `startCol` and
#' `startRow` individually.  A vector of the form
#' `c(startCol, startRow)`.
#' @param colNames If `TRUE`, column names of x are written.
#' @param rowNames If `TRUE`, data.frame row names of x are written.
#' @param withFilter If `TRUE`, add filters to the column name row. NOTE can only have one filter per worksheet.
#' @param sep Only applies to list columns. The separator used to collapse list columns to a character vector e.g. sapply(x$list_column, paste, collapse = sep).
#' @param name If not NULL, a named region is defined.
#' @param applyCellStyle apply styles when writing on the sheet
#' @param removeCellStyle if writing into existing cells, should the cell style be removed?
#' @param na.strings optional na.strings argument. if missing #N/A is used. If NULL no cell value is written, if character or numeric this is written (even if NA is part of numeric data)
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
#' #####################################################################################
#' ## Create Workbook object and add worksheets
#' wb <- wb_workbook()
#'
#' ## Add worksheets
#' wb$add_worksheet("Cars")
#' wb$add_worksheet("Formula")
#'
#' x <- mtcars[1:6, ]
#' wb$add_data("Cars", x, startCol = 2, startRow = 3, rowNames = TRUE)
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
write_data <- function(
    wb,
    sheet,
    x,
    startCol = 1,
    startRow = 1,
    dims = rowcol_to_dims(startRow, startCol),
    array = FALSE,
    xy = NULL,
    colNames = TRUE,
    rowNames = FALSE,
    withFilter = FALSE,
    sep = ", ",
    name = NULL,
    applyCellStyle = TRUE,
    removeCellStyle = FALSE,
    na.strings
) {

  if (missing(na.strings)) na.strings <- substitute()

  write_data_table(
    wb = wb,
    sheet = sheet,
    x = x,
    startCol = startCol,
    startRow = startRow,
    dims = dims,
    array = array,
    xy = xy,
    colNames = colNames,
    rowNames = rowNames,
    tableStyle = NULL,
    tableName = NULL,
    withFilter = withFilter,
    sep = sep,
    firstColumn = FALSE,
    lastColumn = FALSE,
    bandedRows = FALSE,
    bandedCols = FALSE,
    name = name,
    applyCellStyle = applyCellStyle,
    removeCellStyle = removeCellStyle,
    data_table = FALSE,
    na.strings = na.strings
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
#' @param dims Spreadsheet dimensions that will determine startCol and startRow: "A1", "A1:B2", "A:B"
#' @param array A bool if the function written is of type array
#' @param xy An alternative to specifying `startCol` and
#' `startRow` individually.  A vector of the form
#' `c(startCol, startRow)`.
#' @param applyCellStyle apply styles when writing on the sheet
#' @param removeCellStyle if writing into existing cells, should the cell style be removed?
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
#' ## 4. - Writing internal hyperlinks
#'
#' wb <- wb_workbook()
#' wb$add_worksheet("Sheet1")
#' wb$add_worksheet("Sheet2")
#' write_formula(wb, "Sheet1", x = '=HYPERLINK("#Sheet2!B3", "Text to Display - Link to Sheet2")')
#'
#' ## 5. - Writing array formulas
#'
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
write_formula <- function(
  wb,
  sheet,
  x,
  startCol = 1,
  startRow = 1,
  dims = rowcol_to_dims(startRow, startCol),
  array = FALSE,
  xy = NULL,
  applyCellStyle = TRUE,
  removeCellStyle = FALSE
) {
  assert_class(x, "character")
  # remove xml encoding and reapply it afterwards. until v0.3 encoding was not enforced
  x <- replaceXMLEntities(x)
  x <- vapply(x, function(val) xml_value(xml_node_create("fml", val, escapes = TRUE), "fml"), NA_character_)
  dfx <- data.frame("X" = x, stringsAsFactors = FALSE)
  class(dfx$X) <- c("character", if (array) "array_formula" else "formula")

  if (any(grepl("=([\\s]*?)HYPERLINK\\(", x, perl = TRUE))) {
    class(dfx$X) <- c("character", "formula", "hyperlink")
  }

  write_data(
    wb = wb,
    sheet = sheet,
    x = dfx,
    startCol = startCol,
    startRow = startRow,
    dims = dims,
    array = array,
    xy = xy,
    colNames = FALSE,
    rowNames = FALSE,
    applyCellStyle = applyCellStyle,
    removeCellStyle = removeCellStyle
  )
}


#' @name write_datatable
#' @title Write to a worksheet as an Excel table
#' @description Write to a worksheet and format as an Excel table
#' @param wb A Workbook object containing a worksheet.
#' @param sheet The worksheet to write to. Can be the worksheet index or name.
#' @param x A data frame.
#' @param startCol A vector specifying the starting column to write df
#' @param startRow A vector specifying the starting row to write df
#' @param dims Spreadsheet dimensions that will determine startCol and startRow: "A1", "A1:B2", "A:B"
#' @param xy An alternative to specifying startCol and startRow individually.
#' A vector of the form c(startCol, startRow)
#' @param colNames If `TRUE`, column names of x are written.
#' @param rowNames If `TRUE`, row names of x are written.
#' @param tableStyle Any excel table style name or "none" (see "formatting" vignette).
#' @param tableName name of table in workbook. The table name must be unique.
#' @param withFilter If `TRUE`, columns with have filters in the first row.
#' @param sep Only applies to list columns. The separator used to collapse list columns to a character vector e.g. sapply(x$list_column, paste, collapse = sep).
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
#' @param applyCellStyle apply styles when writing on the sheet
#' @param removeCellStyle if writing into existing cells, should the cell style be removed?
#' @param na.strings optional na.strings argument. if missing #N/A is used. If NULL no cell value is written, if character or numeric this is written (even if NA is part of numeric data)
#' @details columns of x with class Date/POSIXt, currency, accounting,
#' hyperlink, percentage are automatically styled as dates, currency, accounting,
#' hyperlinks, percentages respectively.
#' The string `"_openxlsx_NA"` is reserved for `openxlsx2`. If the data frame
#' contains this string, the output will be broken.
#' @seealso [wb_add_worksheet()]
#' @seealso [write_data()]
#' @seealso [wb_remove_tables()]
#' @seealso [wb_get_tables()]
#' @export
#' @examples
#' ## see package vignettes for further examples.
#'
#' #####################################################################################
#' ## Create Workbook object and add worksheets
#' wb <- wb_workbook()
#' wb$add_worksheet("S1")
#' wb$add_worksheet("S2")
#' wb$add_worksheet("S3")
#'
#' #####################################################################################
#' ## -- write data.frame as an Excel table with column filters
#' ## -- default table style is "TableStyleMedium2"
#'
#' wb$add_data_table("S1", x = iris)
#'
#' wb$add_data_table("S2",
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
#' wb$add_data_table("S3", x = df, startRow = 4, rowNames = TRUE, tableStyle = "TableStyleMedium9")
#'
#' #####################################################################################
#' ## Additional Header Styling and remove column filters
#'
#' write_datatable(wb,
#'   sheet = 1,
#'   x = iris,
#'   startCol = 7,
#'   withFilter = FALSE,
#'   firstColumn = TRUE,
#'   lastColumn	= TRUE,
#'   bandedRows = TRUE,
#'   bandedCols = TRUE
#' )
#'
#' #####################################################################################
#' ## Pre-defined table styles gallery
#'
#' wb <- wb_workbook(paste0("tableStylesGallery.xlsx"))
#' wb$add_worksheet("Style Samples")
#' for (i in 1:21) {
#'   style <- paste0("TableStyleLight", i)
#'   write_datatable(wb,
#'     x = data.frame(style), sheet = 1,
#'     tableStyle = style, startRow = 1, startCol = i * 3 - 2
#'   )
#' }
#'
#' for (i in 1:28) {
#'   style <- paste0("TableStyleMedium", i)
#'   write_datatable(wb,
#'     x = data.frame(style), sheet = 1,
#'     tableStyle = style, startRow = 4, startCol = i * 3 - 2
#'   )
#' }
#'
#' for (i in 1:11) {
#'   style <- paste0("TableStyleDark", i)
#'   write_datatable(wb,
#'     x = data.frame(style), sheet = 1,
#'     tableStyle = style, startRow = 7, startCol = i * 3 - 2
#'   )
#' }
write_datatable <- function(
    wb,
    sheet,
    x,
    startCol = 1,
    startRow = 1,
    dims = rowcol_to_dims(startRow, startCol),
    xy = NULL,
    colNames = TRUE,
    rowNames = FALSE,
    tableStyle = "TableStyleLight9",
    tableName = NULL,
    withFilter = TRUE,
    sep = ", ",
    firstColumn = FALSE,
    lastColumn = FALSE,
    bandedRows = TRUE,
    bandedCols = FALSE,
    applyCellStyle = TRUE,
    removeCellStyle = FALSE,
    na.strings
) {

  if (missing(na.strings)) na.strings <- substitute()

  write_data_table(
    wb = wb,
    sheet = sheet,
    x = x,
    startCol = startCol,
    startRow = startRow,
    dims = dims,
    array = FALSE,
    xy = xy,
    colNames = colNames,
    rowNames = rowNames,
    tableStyle = tableStyle,
    tableName = tableName,
    withFilter = withFilter,
    sep = sep,
    firstColumn = firstColumn,
    lastColumn = lastColumn,
    bandedRows = bandedRows,
    bandedCols = bandedCols,
    name = NULL,
    data_table = TRUE,
    applyCellStyle = applyCellStyle,
    removeCellStyle = removeCellStyle,
    na.strings = na.strings
  )
}
