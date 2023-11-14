#' function to add missing cells to cc and rows
#'
#' Create a cell in the workbook
#'
#' @param wb the workbook update
#' @param sheet_id the sheet to update
#' @param x the newly filled cc frame
#' @param rows the rows needed
#' @param cells_needed the cells needed
#' @param colNames has colNames (only in update_cell)
#' @param removeCellStyle remove the cell style (only in update_cell)
#' @param na.strings Value used for replacing `NA` values from `x`. Default
#'   `na_strings()` uses the special `#N/A` value within the workbook.#' @keywords internal
#' @noRd
inner_update <- function(
    wb,
    sheet_id,
    x,
    rows,
    cells_needed,
    colNames = FALSE,
    removeCellStyle = FALSE,
    na.strings = na_strings()
) {

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

  if (is_na_strings(na.strings)) {
    na.strings <- NULL
  }

  if (removeCellStyle) {
    cell_style <- "c_s"
  } else {
    cell_style <- NULL
  }

  replacement <- c("r", cell_style, "c_t", "c_cm", "c_ph", "c_vm", "v",
                   "f", "f_t", "f_ref", "f_ca", "f_si", "is", "typ")

  sel <- match(x$r, cc$r)
  cc[sel, replacement] <- x[replacement]

  # avoid missings in cc
  if (any(is.na(cc)))
    cc[is.na(cc)] <- ""

  # push everything back to workbook
  wb$worksheets[[sheet_id]]$sheet_data$cc  <- cc

  wb
}

#' Initialize data cell(s)
#'
#' Create a cell in the workbook
#'
#' @param wb the workbook you want to update
#' @param sheet the sheet you want to update
#' @param new_cells the cell you want to update in Excel connotation e.g. "A1"
#'
#' @keywords internal
#' @noRd
initialize_cell <- function(wb, sheet, new_cells) {

  sheet_id <- wb$validate_sheet(sheet)

  # create artificial cc for the missing cells
  x <- empty_sheet_data_cc(n = length(new_cells))
  x$r     <- new_cells
  x$row_r <- gsub("[[:upper:]]", "", new_cells)
  x$c_r   <- gsub("[[:digit:]]", "", new_cells)

  rows <- x$row_r
  cells_needed <- new_cells

  inner_update(wb, sheet_id, x, rows, cells_needed)
}

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

  if (missing(na.strings))
    na.strings <- substitute()

  sheet_id <- wb$validate_sheet(sheet)

  dims <- dims_to_dataframe(cell, fill = TRUE)
  rows <- rownames(dims)

  cells_needed <- unname(unlist(dims))

  inner_update(wb, sheet_id, x, rows, cells_needed, colNames, removeCellStyle, na.strings)
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
#' @param na.strings Value used for replacing `NA` values from `x`. Default
#'   `na_strings()` uses the special `#N/A` value within the workbook.
#' @param data_table logical. if `TRUE` and `rowNames = TRUE`, do not write the cell containing  `"_rowNames_"`
#' @param inline_strings write characters as inline strings
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
#' @noRd
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
    na.strings = na_strings(),
    data_table = FALSE,
    inline_strings = TRUE
) {

  is_data_frame <- FALSE
  #### prepare the correct data formats for openxml
  dc <- openxlsx2_type(data)

  # convert factor to character
  is_factor <- dc == openxlsx2_celltype[["factor"]]

  if (any(is_factor)) {
    fcts <- names(dc[is_factor])
    data[fcts] <- lapply(data[fcts], to_string)
  }

  # remove xml encoding and reapply it afterwards. until v0.3 encoding was not enforced.
  # until 1.1 formula encoding was applied in write_formula() and missed formulas written
  # as data frames with class formula
  is_fml <- dc %in% c(
    openxlsx2_celltype[["formula"]], openxlsx2_celltype[["array_formula"]],
    openxlsx2_celltype[["cm_formula"]], openxlsx2_celltype[["hyperlink"]]
  )

  if (any(is_fml)) {
    fmls <- names(dc[is_fml])
    data[fmls] <- lapply(
      data[fmls],
      function(val) {
        val <- replaceXMLEntities(val)
        vapply(val, function(x) xml_value(xml_node_create("fml", x, escapes = TRUE), "fml"), "")
      }
    )
  }

  hconvert_date1904 <- grepl('date1904="1"|date1904="true"',
                             stri_join(unlist(wb$workbook), collapse = ""),
                             ignore.case = TRUE)

  # TODO need to tell excel that we have a date, apply some kind of numFmt
  data <- convert_to_excel_date(df = data, date1904 = hconvert_date1904)

  # backward compatible
  if (!inherits(data, "data.frame") || inherits(data, "matrix")) {
    data <- as.data.frame(data)
    colNames <- FALSE
  }

  if (inherits(data, "data.frame") || inherits(data, "matrix")) {
    is_data_frame <- TRUE

    if (is.data.frame(data)) data <- as.data.frame(data)

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

  ### ref for formulas from write_formula caring ref attr
  if (!is.null(attr(data, "f_ref"))) {
    ref <- attr(data, "f_ref")
  } else {
    ref <- dims
  }

  if (!is.null(attr(data, "c_cm"))) {
    warning("modifications with cm formulas are experimental. use at own risk")
    c_cm <- attr(data, "c_cm")
  } else {
    c_cm <- ""
  }

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

    sheet_name <- wb$get_sheet_names(escape = TRUE)[[sheetno]]
    if (grepl("[^A-Za-z0-9]", sheet_name)) sheet_name <- shQuote(sheet_name, "sh")

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

  sel <- which(dc == openxlsx2_celltype[["character"]] | dc == openxlsx2_celltype[["factor"]]) # character
  if (length(sel)) {
    data[sel][is.na(data[sel])] <- "_openxlsx_NA"

    if (getOption("openxlsx2.force_utf8_encoding", default = FALSE)) {
      from_enc <- getOption("openxlsx2.native_encoding")
      data[sel] <- lapply(data[sel], stringi::stri_encode, from = from_enc, to = "UTF-8")
    }
  }

  string_nums <- getOption("openxlsx2.string_nums", default = 0)

  na_missing <- FALSE
  na_null    <- FALSE

  if (is_na_strings(na.strings)) {
    na.strings <- ""
    na_missing <- TRUE
  } else if (is.null(na.strings)) {
    na.strings <- ""
    na_null    <- TRUE
  }

  wide_to_long(
    data,
    dc,
    cc,
    ColNames       = colNames,
    start_col      = startCol,
    start_row      = startRow,
    ref            = ref,
    string_nums    = string_nums,
    na_null        = na_null,
    na_missing     = na_missing,
    na_strings     = na.strings,
    inline_strings = inline_strings,
    c_cm           = c_cm
  )

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

      dataframe_to_dims(rtyp[rownames(rtyp) %in% sel_rows, sel_cols, drop = FALSE])
    }

    # if hyperlinks are found, Excel sets something like the following font
    # blue with underline
    if (any(dc == openxlsx2_celltype[["hyperlink"]])) {

      dim_sel <- get_data_class_dims("hyperlink")
      # message("hyperlink: ", dim_sel)

      wb$add_font(
        sheet = sheetno,
        dim = dim_sel,
        color = wb_color(hex = "FF0000FF"),
        name = wb_get_base_font(wb)$name$val,
        u = "single"
      )
    }

    if (any(dc == openxlsx2_celltype[["character"]])) {
      if (any(sel <- cc$typ == openxlsx2_celltype[["string_nums"]])) {

        # # we cannot select every cell like this, because it is terribly slow.
        # dim_sel <- paste0(cc$r[sel], collapse = ";")
        dim_sel <- get_data_class_dims("character")
        # message("character: ", dim_sel)

        wb$add_cell_style(
          sheet = sheetno,
          dim = dim_sel,
          applyNumberFormat = "1",
          quotePrefix = "1",
          numFmtId = "49"
        )
      }
    }

    # options("openxlsx2.numFmt" = NULL)
    if (any(dc == openxlsx2_celltype[["numeric"]])) { # numeric or integer
      if (!is.null(getOption("openxlsx2.numFmt"))) {

        numfmt_numeric <- getOption("openxlsx2.numFmt")

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
      numfmt_dt <- getOption("openxlsx2.dateFormat", 14)

      dim_sel <- get_data_class_dims("short_date")
      # message("short_date: ", dim_sel)

      wb$add_numfmt(
        sheet = sheetno,
        dim = dim_sel,
        numfmt = numfmt_dt
      )
    }
    if (any(dc == openxlsx2_celltype[["long_date"]])) {
      numfmt_posix <- getOption("openxlsx2.datetimeFormat", default = 22)

      dim_sel <- get_data_class_dims("long_date")
      # message("long_date: ", dim_sel)

      wb$add_numfmt(
        sheet = sheetno,
        dim = dim_sel,
        numfmt = numfmt_posix
      )
    }
    if (any(dc == openxlsx2_celltype[["hms_time"]])) {
      numfmt_hms <- getOption("openxlsx2.hmsFormat", default = 21)

      dim_sel <- get_data_class_dims("hms_time")
      # message("hms: ", dim_sel)

      wb$add_numfmt(
        sheet = sheetno,
        dim = dim_sel,
        numfmt = numfmt_hms
      )
    }
    if (any(dc == openxlsx2_celltype[["accounting"]])) { # accounting
      numfmt_accounting <- getOption("openxlsx2.accountingFormat", default = 4)

      dim_sel <- get_data_class_dims("accounting")
      # message("accounting: ", dim_sel)

      wb$add_numfmt(
        dim = dim_sel,
        numfmt = numfmt_accounting
      )
    }
    if (any(dc == openxlsx2_celltype[["percentage"]])) { # percentage
      numfmt_percentage <- getOption("openxlsx2.percentageFormat", default = 10)

      dim_sel <- get_data_class_dims("percentage")
      # message("percentage: ", dim_sel)

      wb$add_numfmt(
        sheet = sheetno,
        dim = dim_sel,
        numfmt = numfmt_percentage
      )
    }
    if (any(dc == openxlsx2_celltype[["scientific"]])) {
      numfmt_scientific <- getOption("openxlsx2.scientificFormat", default = 48)

      dim_sel <- get_data_class_dims("scientific")
      # message("scientific: ", dim_sel)

      wb$add_numfmt(
        sheet = sheetno,
        dim = dim_sel,
        numfmt = numfmt_scientific
      )
    }
    if (any(dc == openxlsx2_celltype[["comma"]])) {
      numfmt_comma <- getOption("openxlsx2.commaFormat", default = 3)

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

  # update shared strings if we use shared strings
  if (!inline_strings) {

    cc <- wb$worksheets[[sheetno]]$sheet_data$cc

    sel <- grepl("<si>", cc$v)
    cc_sst <- stringi::stri_unique(cc[sel, "v"])

    wb$sharedStrings <- stringi::stri_unique(c(wb$sharedStrings, cc_sst))

    sel <- grepl("<si>", cc$v)
    cc$v[sel] <- as.character(match(cc$v[sel], wb$sharedStrings) - 1L)

    text        <- si_to_txt(wb$sharedStrings)
    uniqueCount <- length(wb$sharedStrings)

    attr(wb$sharedStrings, "uniqueCount") <- uniqueCount
    attr(wb$sharedStrings, "text")        <- text

    wb$worksheets[[sheetno]]$sheet_data$cc <- cc

    if (!any(grepl("sharedStrings", wb$workbook.xml.rels))) {
      wb$append(
        "workbook.xml.rels",
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>"
      )
    }
  }

  ### Update calcChain
  if (length(wb$calcChain)) {

    # if we overwrite a formula cell in the calculation chain, we have to update it
    # At the moment we simply remove it from the calculation chain, in the future
    # we might want to keep it if we write a formula.
    xml <- wb$calcChain
    calcChainR <- rbindlist(xml_attr(xml, "calcChain", "c"))
    # according to the documentation there can be cases, without the sheetno reference
    sel <- calcChainR$r %in% rtyp & calcChainR$i == sheetno
    rmCalcChain <- as.integer(rownames(calcChainR[sel, , drop = FALSE]))
    if (length(rmCalcChain)) {
      xml <- xml_rm_child(xml, xml_child = "c", which = rmCalcChain)
      # xml can not be empty, otherwise excel will complain. If xml is empty, remove all
      # calcChain references from the workbook
      if (length(xml_node_name(xml, "calcChain")) == 0) {
        wb$Content_Types <- wb$Content_Types[-grep("/xl/calcChain", wb$Content_Types)]
        wb$workbook.xml.rels <- wb$workbook.xml.rels[-grep("calcChain.xml", wb$workbook.xml.rels)]
        wb$worksheets[[sheetno]]$sheetCalcPr <- character()
        xml <- character()
      }
      wb$calcChain <- xml
    }

  }
  ### End update calcChain

  return(wb)
}

# `write_data_table()` ---------------------------------------------------------
# `write_data_table()` an internal driver function to `write_data` and `write_data_table` ----
#' Write to a worksheet as an Excel table
#'
#' Write to a worksheet and format as an Excel table
#'
#' @param wb A Workbook object containing a worksheet.
#' @param sheet The worksheet to write to. Can be the worksheet index or name.
#' @param x A data frame.
#' @param startCol A vector specifying the starting column to write df
#' @param startRow A vector specifying the starting row to write df
#' @param dims Spreadsheet dimensions that will determine startCol and startRow: "A1", "A1:B2", "A:B"
#' @param array A bool if the function written is of type array
#' @param colNames If `TRUE`, column names of x are written.
#' @param rowNames If `TRUE`, row names of x are written.
#' @param tableStyle Any excel table style name or "none" (see "formatting" vignette).
#' @param tableName name of table in workbook. The table name must be unique.
#' @param withFilter If `TRUE`, columns with have filters in the first row.
#' @param sep Only applies to list columns. The separator used to collapse list columns to a character vector e.g. sapply(x$list_column, paste, collapse = sep).
#' @param firstColumn logical. If TRUE, the first column is bold
#' @param lastColumn logical. If TRUE, the last column is bold
#' @param bandedRows logical. If TRUE, rows are color banded
#' @param bandedCols logical. If TRUE, the columns are color banded
#' @param bandedCols logical. If TRUE, a data table is created
#' @param name If not NULL, a named region is defined.
#' @param applyCellStyle apply styles when writing on the sheet
#' @param removeCellStyle if writing into existing cells, should the cell style be removed?
#' @param na.strings Value used for replacing `NA` values from `x`. Default
#'   `na_strings()` uses the special `#N/A` value within the workbook.
#' @param inline_strings optional write strings as inline strings
#' @noRd
#' @keywords internal
write_data_table <- function(
    wb,
    sheet,
    x,
    startCol        = 1,
    startRow        = 1,
    dims,
    array           = FALSE,
    colNames        = TRUE,
    rowNames        = FALSE,
    tableStyle      = "TableStyleLight9",
    tableName       = NULL,
    withFilter      = TRUE,
    sep             = ", ",
    firstColumn     = FALSE,
    lastColumn      = FALSE,
    bandedRows      = TRUE,
    bandedCols      = FALSE,
    name            = NULL,
    applyCellStyle  = TRUE,
    removeCellStyle = FALSE,
    data_table      = FALSE,
    na.strings      = na_strings(),
    inline_strings  = TRUE
) {

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

  # force with globalenv() options
  x <- force(x)

  op <- default_save_opt()
  on.exit(options(op), add = TRUE)

  if (!is.null(dims)) {
    dims <- dims_to_rowcol(dims, as_integer = TRUE)
    startCol <- min(dims[[1]])
    startRow <- min(dims[[2]])
  }

  # avoid stoi error with NULL
  if (is.null(x)) {
    return(wb)
  }

  if (data_table && nrow(x) < 1) {
    warning("Found data table with zero rows, adding one.",
            " Modify na with na.strings")
    x[1, ] <- NA
  }

  ## common part ---------------------------------------------------------------
  if ((!is.character(sep)) || (length(sep) != 1))
    stop("sep must be a character vector of length 1")

  # TODO clean up when moved into wbWorkbook
  sheet <- wb$.__enclos_env__$private$get_sheet_index(sheet)
  # sheet <- wb$validate_sheet(sheet)

  if (wb$is_chartsheet[[sheet]]) stop("Cannot write to chart sheet.")

  ## convert startRow and startCol
  if (!is.numeric(startCol)) {
    startCol <- col2int(startCol)
  }
  startRow <- as.integer(startRow)

  ## special case - vector of hyperlinks
  is_hyperlink <- FALSE
  if (applyCellStyle) {
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
  }

  ### Create data frame --------------------------------------------------------

  ## special case - formula
  # only for data frame case where a data frame is passed down containing formulas
  if (inherits(x, "formula")) {
    x <- data.frame("X" = x, stringsAsFactors = FALSE)
    class(x[[1]]) <- if (array) "array_formula" else "formula"
    colNames <- FALSE
  }

  if (is.vector(x) || is.factor(x) || inherits(x, "Date") || inherits(x, "POSIXt")) {
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
      dfn <- sprintf("'%s'!%s", wb$get_sheet_names(escape = TRUE)[sheet], stri_join("$", l, "$", coords[, 1], collapse = ":"))

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
    data_table = data_table,
    inline_strings = inline_strings
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
    cstm_tableStyles <- wb$styles_mgr$tableStyle$name
    validNames <- c("none", paste0("TableStyleLight", seq_len(21)), paste0("TableStyleMedium", seq_len(28)), paste0("TableStyleDark", seq_len(11)), cstm_tableStyles)
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

# `write_data()` ---------------------------------------------------------------
#' Write an object to a worksheet
#'
#' Use [wb_add_data()] or [write_xlsx()] in new code.
#'
#' @inheritParams wb_add_data
#' @return invisible(0)
#' @export
#' @keywords internal
write_data <- function(
    wb,
    sheet,
    x,
    dims              = wb_dims(start_row, start_col),
    start_col         = 1,
    start_row         = 1,
    array             = FALSE,
    col_names         = TRUE,
    row_names         = FALSE,
    with_filter       = FALSE,
    sep               = ", ",
    name              = NULL,
    apply_cell_style  = TRUE,
    remove_cell_style = FALSE,
    na.strings        = na_strings(),
    inline_strings    = TRUE,
    ...
) {

  standardize_case_names(...)

  write_data_table(
    wb              = wb,
    sheet           = sheet,
    x               = x,
    dims            = dims,
    startCol        = start_col,
    startRow        = start_row,
    array           = array,
    colNames        = col_names,
    rowNames        = row_names,
    tableStyle      = NULL,
    tableName       = NULL,
    withFilter      = with_filter,
    sep             = sep,
    firstColumn     = FALSE,
    lastColumn      = FALSE,
    bandedRows      = FALSE,
    bandedCols      = FALSE,
    name            = name,
    applyCellStyle  = apply_cell_style,
    removeCellStyle = remove_cell_style,
    data_table      = FALSE,
    na.strings      = na.strings,
    inline_strings  = inline_strings
  )
}

# write_formula() -------------------------------------------
#' Write a character vector as an Excel Formula
#'
#' Write a a character vector containing Excel formula to a worksheet.
#' Use [wb_add_formula()] or `add_formula()` in new code. This function is meant
#' for internal use.
#'
#' @inheritParams wb_add_formula
#' @export
#' @keywords internal
write_formula <- function(
    wb,
    sheet,
    x,
    dims              = wb_dims(start_row, start_col),
    start_col         = 1,
    start_row         = 1,
    array             = FALSE,
    cm                = FALSE,
    apply_cell_style  = TRUE,
    remove_cell_style = FALSE,
    ...
) {

  standardize_case_names(...)

  assert_class(x, "character")
  dfx <- data.frame("X" = x, stringsAsFactors = FALSE)

  formula <- "formula"
  if (array) formula <- "array_formula"
  if (cm)    {
    # need to set cell metadata in wb$metadata
    if (is.null(wb$metadata)) {

      wb$append("Content_Types", "<Override PartName=\"/xl/metadata.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml\"/>")

      wb$metadata <- # danger danger no clue what this means!
        xml_node_create(
          "metadata",
          xml_attributes = c(
            xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "xmlns:xda" = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"
          ),
          xml_children = read_xml(
            "<metadataTypes count=\"1\">
            <metadataType name=\"XLDAPR\" minSupportedVersion=\"120000\" copy=\"1\" pasteAll=\"1\" pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" clearComments=\"1\" assign=\"1\" coerce=\"1\" cellMeta=\"1\"/>
            </metadataTypes>
            <futureMetadata name=\"XLDAPR\" count=\"1\">
            <bk>
            <extLst>
            <ext uri=\"{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}\">
            <xda:dynamicArrayProperties fDynamic=\"1\" fCollapsed=\"0\"/>
            </ext>
            </extLst>
            </bk>
            </futureMetadata>,
            <cellMetadata/>",
            pointer = FALSE
          )
        )
    }

    ## TODO Not sure if there are more cases
    # add new cell metadata record
    cM <- xml_node(wb$metadata, "metadata", "cellMetadata")
    cM <- xml_add_child(cM, xml_child = "<bk><rc t=\"1\" v=\"0\"/></bk>")

    # we need to update count
    cnt <- as_xml_attr(length(xml_node(cM, "cellMetadata", "bk")))
    cM <- xml_attr_mod(cM, xml_attributes = c(count = cnt))

    # remove current cellMetadata update new
    wb$metadata <- xml_rm_child(wb$metadata, "cellMetadata")
    wb$metadata <- xml_add_child(wb$metadata, cM)

    attr(dfx, "c_cm") <- cnt
    formula <- "cm_formula"
  }

  class(dfx$X) <- c("character", formula)
  if (!is.null(dims)) {
    if (array || cm) {
      attr(dfx, "f_ref") <- dims
    }
  }

  if (any(grepl("=([\\s]*?)HYPERLINK\\(", x, perl = TRUE))) {
    class(dfx$X) <- c("character", "formula", "hyperlink")
  }

  write_data(
    wb                = wb,
    sheet             = sheet,
    x                 = dfx,
    start_col         = start_col,
    start_row         = start_row,
    dims              = dims,
    array             = array,
    col_names         = FALSE,
    row_names         = FALSE,
    apply_cell_style  = apply_cell_style,
    remove_cell_style = remove_cell_style
  )

}

# `write_datatable()` ----------------------
#' Write to a worksheet as an Excel table
#'
#' Write to a worksheet and format as an Excel table. Use [wb_add_data_table()] in new code.
#' This function may not be exported
#' @inheritParams wb_add_data_table
#' @inherit wb_add_data_table details
#' @export
#' @keywords internal
write_datatable <- function(
    wb,
    sheet,
    x,
    dims              = wb_dims(start_row, start_col),
    start_col         = 1,
    start_row         = 1,
    col_names         = TRUE,
    row_names         = FALSE,
    table_style       = "TableStyleLight9",
    table_name        = NULL,
    with_filter       = TRUE,
    sep               = ", ",
    first_column      = FALSE,
    last_column       = FALSE,
    banded_rows       = TRUE,
    banded_cols       = FALSE,
    apply_cell_style  = TRUE,
    remove_cell_style = FALSE,
    na.strings        = na_strings(),
    inline_strings    = TRUE,
    ...
) {

  standardize_case_names(...)

  write_data_table(
    wb              = wb,
    sheet           = sheet,
    x               = x,
    startCol        = start_col,
    startRow        = start_row,
    dims            = dims,
    array           = FALSE,
    colNames        = col_names,
    rowNames        = row_names,
    tableStyle      = table_style,
    tableName       = table_name,
    withFilter      = with_filter,
    sep             = sep,
    firstColumn     = first_column,
    lastColumn      = last_column,
    bandedRows      = banded_rows,
    bandedCols      = banded_cols,
    name            = NULL,
    data_table      = TRUE,
    applyCellStyle  = apply_cell_style,
    removeCellStyle = remove_cell_style,
    na.strings      = na.strings,
    inline_strings  = inline_strings
  )
}
