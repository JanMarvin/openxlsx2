# `wb_to_df()` ----------------------------------------
#' Create a data frame from a Workbook
#'
#' Simple function to create a `data.frame` from a sheet in workbook. Simple as
#' in it was simply written down. `read_xlsx()` and `wb_read()` are just
#' internal wrappers of `wb_to_df()` intended for people coming from other
#' packages.
#'
#' @details
#' The returned data frame will have named rows matching the rows of the
#' worksheet. With `col_names = FALSE` the returned data frame will have
#' column names matching the columns of the worksheet. Otherwise the first
#' row is selected as column name.
#'
#' Depending if the R package `hms` is loaded, `wb_to_df()` returns
#' `hms` variables or string variables in the `hh:mm:ss` format.
#'
#' The `types` argument must be a named numeric.
#' * 0: character
#' * 1: numeric
#' * 2: date
#' * 3: posixt (datetime)
#' * 4: logical
#'
#' `wb_to_df()` will not pick up formulas added to a workbook object
#' via [wb_add_formula()]. This is because only the formula is written and left
#' to be evaluated when the file is opened in a spreadsheet software.
#' Opening, saving and closing the file in a spreadsheet software will resolve
#' this.
#'
#' @seealso [wb_get_named_regions()]
#'
#' @param file An xlsx file, [wbWorkbook] object or URL to xlsx file.
#' @param sheet Either sheet name or index. When missing the first sheet in the workbook is selected.
#' @param col_names If `TRUE`, the first row of data will be used as column names.
#' @param row_names If `TRUE`, the first col of data will be used as row names.
#' @param dims Character string of type "A1:B2" as optional dimensions to be imported.
#' @param detect_dates If `TRUE`, attempt to recognize dates and perform conversion.
#' @param show_formula If `TRUE`, the underlying Excel formulas are shown.
#' @param convert If `TRUE`, a conversion to dates and numerics is attempted.
#' @param skip_empty_cols If `TRUE`, empty columns are skipped.
#' @param skip_empty_rows If `TRUE`, empty rows are skipped.
#' @param skip_hidden_cols If `TRUE`, hidden columns are skipped.
#' @param skip_hidden_rows If `TRUE`, hidden rows are skipped.
#' @param start_row first row to begin looking for data.
#' @param start_col first column to begin looking for data.
#' @param rows A numeric vector specifying which rows in the xlsx file to read.
#'   If `NULL`, all rows are read.
#' @param cols A numeric vector specifying which columns in the xlsx file to read.
#'   If `NULL`, all columns are read.
#' @param named_region Character string with a `named_region` (defined name or table).
#'   If no sheet is selected, the first appearance will be selected. See [wb_get_named_regions()]
#' @param types A named numeric indicating, the type of the data.
#'   Names must match the returned data. See **Details** for more.
#' @param na.strings A character vector of strings which are to be interpreted as `NA`.
#'   Blank cells will be returned as `NA`.
#' @param na.numbers A numeric vector of digits which are to be interpreted as `NA`.
#'   Blank cells will be returned as `NA`.
#' @param fill_merged_cells If `TRUE`, the value in a merged cell is given to all cells within the merge.
#' @param keep_attributes If `TRUE` additional attributes are returned.
#'   (These are used internally to define a cell type.)
#' @param ... additional arguments
#'
#' @examples
#' ###########################################################################
#' # numerics, dates, missings, bool and string
#' example_file <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
#' wb1 <- wb_load(example_file)
#'
#' # import workbook
#' wb_to_df(wb1)
#'
#' # do not convert first row to column names
#' wb_to_df(wb1, col_names = FALSE)
#'
#' # do not try to identify dates in the data
#' wb_to_df(wb1, detect_dates = FALSE)
#'
#' # return the underlying Excel formula instead of their values
#' wb_to_df(wb1, show_formula = TRUE)
#'
#' # read dimension without colNames
#' wb_to_df(wb1, dims = "A2:C5", col_names = FALSE)
#'
#' # read selected cols
#' wb_to_df(wb1, cols = c("A:B", "G"))
#'
#' # read selected rows
#' wb_to_df(wb1, rows = c(2, 4, 6))
#'
#' # convert characters to numerics and date (logical too?)
#' wb_to_df(wb1, convert = FALSE)
#'
#' # erase empty rows from dataset
#' wb_to_df(wb1, skip_empty_rows = TRUE)
#'
#' # erase empty columns from dataset
#' wb_to_df(wb1, skip_empty_cols = TRUE)
#'
#' # convert first row to rownames
#' wb_to_df(wb1, sheet = 2, dims = "C6:G9", row_names = TRUE)
#'
#' # define type of the data.frame
#' wb_to_df(wb1, cols = c(2, 5), types = c("Var1" = 0, "Var3" = 1))
#'
#' # start in row 5
#' wb_to_df(wb1, start_row = 5, col_names = FALSE)
#'
#' # na string
#' wb_to_df(wb1, na.strings = "a")
#'
#' ###########################################################################
#' # Named regions
#' file_named_region <- system.file("extdata", "namedRegions3.xlsx", package = "openxlsx2")
#' wb2 <- wb_load(file_named_region)
#'
#' # read dataset with named_region (returns global first)
#' wb_to_df(wb2, named_region = "MyRange", col_names = FALSE)
#'
#' # read named_region from sheet
#' wb_to_df(wb2, named_region = "MyRange", sheet = 4, col_names = FALSE)
#'
#' # read_xlsx() and wb_read()
#' example_file <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
#' read_xlsx(file = example_file)
#' df1 <- wb_read(file = example_file, sheet = 1)
#' df2 <- wb_read(file = example_file, sheet = 1, rows = c(1, 3, 5), cols = 1:3)
#' @export
wb_to_df <- function(
    file,
    sheet,
    start_row         = 1,
    start_col         = NULL,
    row_names         = FALSE,
    col_names         = TRUE,
    skip_empty_rows   = FALSE,
    skip_empty_cols   = FALSE,
    skip_hidden_rows  = FALSE,
    skip_hidden_cols  = FALSE,
    rows              = NULL,
    cols              = NULL,
    detect_dates      = TRUE,
    na.strings        = "#N/A",
    na.numbers        = NA,
    fill_merged_cells = FALSE,
    dims,
    show_formula      = FALSE,
    convert           = TRUE,
    types,
    named_region,
    keep_attributes   = FALSE,
    ...
) {


  xlsx_file <- list(...)$xlsx_file
  standardize_case_names(...)

  if (!is.null(xlsx_file)) {
    .Deprecated(old = "xlsx_file", new = "file", package = "openxlsx2")
    file <- xlsx_file %||% file
  }

  if (!is.null(cols)) cols <- col2int(cols)

  if (inherits(file, "wbWorkbook")) {
    wb <- file
  } else {
    # passes missing further on
    if (missing(sheet))
      sheet <- substitute()

    # possible false positive on current lintr runs
    wb <- wb_load(file, sheet = sheet, data_only = TRUE) # nolint
  }

  if (!missing(named_region)) {

    nr <- wb_get_named_regions(wb, tables = TRUE)

    if ((named_region %in% nr$name) && missing(sheet)) {
      sel   <- nr[nr$name == named_region, ][1, ]
      sheet <- sel$sheet
      dims  <- sel$coords
    } else if (named_region %in% nr$name) {
      sel <- nr[nr$name == named_region & nr$sheet == wb_validate_sheet(wb, sheet), ]
      if (NROW(sel) == 0) {
        stop("no such named_region on selected sheet")
      }
      dims <- sel$coords
    } else {
      stop("no such named_region")
    }
  }

  if (missing(sheet)) {
    # TODO default sheet as 1
    sheet <- 1
  }

  if (is.character(sheet)) {
    sheet <- wb_validate_sheet(wb, sheet)
  }

  # the sheet has no data
  if (is.null(wb$worksheets[[sheet]]$sheet_data$cc) ||
      nrow(wb$worksheets[[sheet]]$sheet_data$cc) == 0) {
    # TODO do we need more checks or do we need to initialize a new cc object?
    message("sheet found, but contains no data")
    return(NULL)
  }

  # # Should be available, but is optional according to openxml-2.8.1. Still some
  # # third party applications are known to require it. Maybe make using
  # # dimensions an optional parameter?
  # if (missing(dims))
  #   dims <- getXML1attr_one(wb$worksheets[[sheet]]$dimension,
  #                           "dimension",
  #                           "ref")

  # If no dims are requested via named_region, simply construct them from min
  # and max columns and row found on worksheet
  # TODO it would be useful to have both named_region and dims?
  if (missing(named_region) && missing(dims)) {

    sd <- wb$worksheets[[sheet]]$sheet_data$cc[c("row_r", "c_r")]
    sd$row <- as.integer(sd$row_r)
    sd$col <- col2int(sd$c_r)

    dims <- paste0(int2col(min(sd$col)), min(sd$row), ":",
                   int2col(max(sd$col)), max(sd$row))

  }

  row_attr  <- wb$worksheets[[sheet]]$sheet_data$row_attr
  cc  <- wb$worksheets[[sheet]]$sheet_data$cc
  sst <- attr(wb$sharedStrings, "text")

  rnams <- row_attr$r

  numfmt_date <- numfmt_is_date(wb$styles_mgr$styles$numFmts)
  xlsx_date_style <- style_is_date(wb$styles_mgr$styles$cellXfs, numfmt_date)

  # exclude if year, month or day are suspected
  numfmt_hms <- numfmt_is_hms(wb$styles_mgr$styles$numFmts)
  xlsx_hms_style <- style_is_hms(wb$styles_mgr$styles$cellXfs, numfmt_hms)

  numfmt_posix <- numfmt_is_posix(wb$styles_mgr$styles$numFmts)
  xlsx_posix_style <- style_is_posix(wb$styles_mgr$styles$cellXfs, numfmt_posix)

  # create temporary data frame. hard copy required
  z  <- dims_to_dataframe(dims)
  tt <- copy(z)

  keep_cols <- colnames(z)
  keep_rows <- rownames(z)

  maxRow <- max(as.numeric(keep_rows))
  maxCol <- max(col2int(keep_cols))

  if (start_row > 1) {
    keep_rows <- as.character(seq(start_row, maxRow))
    if (start_row <= maxRow) {
      z  <- z[rownames(z) %in% keep_rows, , drop = FALSE]
      tt <- tt[rownames(tt) %in% keep_rows, , drop = FALSE]
    } else {
      keep_rows <- as.character(start_row)
      z  <- z[keep_rows, , drop = FALSE]
      tt <- tt[keep_rows, , drop = FALSE]

      rownames(z)  <- as.integer(keep_rows)
      rownames(tt) <- as.integer(keep_rows)
    }
  }

  if (!is.null(rows)) {
    keep_rows <- as.character(rows)

    if (all(keep_rows %in% rownames(z))) {
      z  <- z[rownames(z) %in% keep_rows, , drop = FALSE]
      tt <- tt[rownames(tt) %in% keep_rows, , drop = FALSE]
    } else {
      z  <- z[keep_rows, , drop = FALSE]
      tt <- tt[keep_rows, , drop = FALSE]

      rownames(z)  <- as.integer(keep_rows)
      rownames(tt) <- as.integer(keep_rows)
    }
  }

  if (!is.null(start_col)) {
    keep_cols <- int2col(seq(col2int(start_col), maxCol))

    if (!all(keep_cols %in% colnames(z))) {
      keep_col <- keep_cols[!keep_cols %in% colnames(z)]

      z[keep_col]  <- NA_character_
      tt[keep_col] <- NA_character_

      z  <- z[keep_cols]
      tt <- tt[keep_cols]
    }

    z  <- z[, match(keep_cols, colnames(z)), drop = FALSE]
    tt <- tt[, match(keep_cols, colnames(tt)), drop = FALSE]
  }

  if (!is.null(cols)) {
    keep_cols <- int2col(cols)

    if (!all(keep_cols %in% colnames(z))) {
      keep_col <- keep_cols[!keep_cols %in% colnames(z)]

      z[keep_col] <- NA_character_
      tt[keep_col] <- NA_character_
    }

    z  <- z[, match(keep_cols, colnames(z)), drop = FALSE]
    tt <- tt[, match(keep_cols, colnames(tt)), drop = FALSE]
  }

  keep_rows <- keep_rows[keep_rows %in% rnams]

  # reduce data to selected cases only
  if (length(keep_rows) && length(keep_cols))
    cc <- cc[cc$row_r %in% keep_rows & cc$c_r %in% keep_cols, ]

  cc$val <- NA_character_
  cc$typ <- NA_character_

  cc_tab <- unique(cc$c_t)

  # bool
  if (any(cc_tab == "b")) {
    sel <- cc$c_t %in% "b"
    cc$val[sel] <- as.logical(as.numeric(cc$v[sel]))
    cc$typ[sel] <- "b"
  }
  # text in v
  if (any(cc_tab %in% c("str", "e"))) {
    sel <- cc$c_t %in% c("str", "e")
    cc$val[sel] <- replaceXMLEntities(cc$v[sel])
    cc$typ[sel] <- "s"
  }
  # text in t
  if (any(cc_tab %in% c("inlineStr"))) {
    sel <- cc$c_t %in% c("inlineStr")
    cc$val[sel] <- is_to_txt(cc$is[sel])
    cc$typ[sel] <- "s"
  }
  # test is sst
  if (any(cc_tab %in% c("s"))) {
    sel <- cc$c_t %in% c("s")
    cc$val[sel] <- sst[as.numeric(cc$v[sel]) + 1]
    cc$typ[sel] <- "s"
  }

  has_na_string <- FALSE
  # convert missings
  if (!all(is.na(na.strings))) {
    sel <- cc$val %in% na.strings
    if (any(sel)) {
      cc$val[sel] <- NA_character_
      cc$typ[sel] <- "na_string"
      has_na_string <- TRUE
    }
  }

  has_na_number <- FALSE
  # convert missings.
  # at this stage we only have characters.
  na.numbers <- as.character(na.numbers)
  if (!all(is.na(na.numbers))) {
    sel <- cc$v %in% na.numbers
    if (any(sel)) {
      cc$val[sel] <- NA_character_
      cc$typ[sel] <- "na_number"
      has_na_number <- TRUE
    }
  }

  origin <- get_date_origin(wb)

  # dates
  if (!is.null(cc$c_s)) {

    # if a cell is t="s" the content is a sst and not da date
    if (detect_dates && missing(types)) {
      cc$is_string <- FALSE
      if (!is.null(cc$c_t))
        cc$is_string <- cc$c_t %in% c("s", "str", "b", "inlineStr")

      if (any(sel <- cc$c_s %in% xlsx_date_style)) {
        sel <- sel & !cc$is_string & cc$v != ""
        cc$val[sel] <- suppressWarnings(as.character(convert_date(cc$v[sel], origin = origin)))
        cc$typ[sel]  <- "d"
      }

      if (any(sel <- cc$c_s %in% xlsx_hms_style)) {
        sel <- sel & !cc$is_string & cc$v != ""
        if (isNamespaceLoaded("hms")) {
          # if hms is loaded, we have to avoid applying convert_hms() twice
          cc$val[sel] <- cc$v[sel]
        } else {
          cc$val[sel] <- suppressWarnings(as.character(convert_hms(cc$v[sel])))
        }
        cc$typ[sel]  <- "h"
      }

      if (any(sel <- cc$c_s %in% xlsx_posix_style)) {
        sel <- sel & !cc$is_string & cc$v != ""
        cc$val[sel] <- suppressWarnings(as.character(convert_datetime(cc$v[sel], origin = origin)))
        cc$typ[sel]  <- "p"
      }
    }
  }

  # remaining values are numeric?
  if (any(sel <- is.na(cc$typ))) {
    cc$val[sel] <- cc$v[sel]
    cc$typ[sel] <- "n"
  }

  if (show_formula) {
    sel <- cc$f != ""
    cc$val[sel] <- replaceXMLEntities(cc$f[sel])
    cc$typ[sel] <- "s"
  }

  # convert "na_string" to missing
  if (has_na_string) cc$typ[cc$typ == "na_string"] <- NA
  if (has_na_number) cc$typ[cc$typ == "na_number"] <- NA

  # prepare to create output object z
  zz <- cc[c("val", "typ")]
  zz$cols <- NA_integer_
  zz$rows <- NA_integer_
  # we need to create the correct col and row position as integer starting at 0. Because we allow
  # to select specific rows and columns, we must make sure that our zz cols and rows matches the
  # z data frame.
  zz$cols <- match(cc$c_r, colnames(z)) - 1L
  zz$rows <- match(cc$row_r, rownames(z)) - 1L

  zz <- zz[order(zz[, "cols"], zz[, "rows"]), ]
  if (any(zz$val == "", na.rm = TRUE)) zz <- zz[zz$val != "", ]
  long_to_wide(z, tt, zz)

  # backward compatible option. get the mergedCells dimension and fill it with
  # the value of the first cell in the range. do the same for tt.
  if (fill_merged_cells) {
    mc <- wb$worksheets[[sheet]]$mergeCells
    if (length(mc)) {

      mc <- unlist(xml_attr(mc, "mergeCell"))

      for (i in seq_along(mc)) {
        filler <- stringi::stri_split_fixed(mc[i], pattern = ":")[[1]][1]


        dms <- dims_to_dataframe(mc[i])

        if (any(row_sel <- rownames(z) %in% rownames(dms)) &&
            any(col_sel <- colnames(z) %in% colnames(dms))) {

          # TODO there probably is a better way in not reducing cc above, so
          # that we do not have to go through large xlsx files multiple times
          z_fill <- wb_to_df(
            file            = wb,
            sheet           = sheet,
            dims            = filler,
            na.strings      = na.strings,
            convert         = FALSE,
            col_names       = FALSE,
            detect_dates    = detect_dates,
            show_formula    = show_formula,
            keep_attributes = TRUE
          )

          tt_fill <- attr(z_fill, "tt")

          z[row_sel,  col_sel] <- z_fill
          tt[row_sel, col_sel] <- tt_fill
        }
      }
    }

  }

  # the following two skip hidden columns and row and need a valid keep_rows and
  # keep_cols length.
  if (skip_hidden_rows) {
    sel <- row_attr$hidden == "1" | row_attr$hidden == "true"
    if (any(sel)) {
      hide   <- !keep_rows %in% row_attr$r[sel]

      z  <- z[hide, , drop = FALSE]
      tt <- tt[hide, , drop = FALSE]
    }
  }

  if (skip_hidden_cols) {
    col_attr <- wb$worksheets[[sheet]]$unfold_cols()
    sel <- col_attr$hidden == "1" | col_attr$hidden == "true"
    if (any(sel)) {
      hide     <- col2int(keep_cols) %in% as.integer(col_attr$min[sel])

      z[hide]  <- NULL
      tt[hide] <- NULL
    }
  }

  # is.na needs convert
  if (skip_empty_rows) {
    empty <- vapply(seq_len(nrow(z)), function(x) all(is.na(z[x, ])), NA)

    z  <- z[!empty, , drop = FALSE]
    tt <- tt[!empty, , drop = FALSE]
  }

  if (skip_empty_cols) {

    empty <- vapply(z, function(x) all(is.na(x)), NA)

    if (any(empty)) {
      sel <- which(empty)
      z[sel]  <- NULL
      tt[sel] <- NULL
    }

  }

  # prepare colnames object
  xlsx_cols_names <- colnames(z)
  names(xlsx_cols_names) <- xlsx_cols_names

  # if colNames, then change tt too
  if (col_names) {
    # select first row as colnames, but do not yet assign. it might contain
    # missing values and if assigned, convert below might break with unambiguous
    # names.
    nams <- names(xlsx_cols_names)
    xlsx_cols_names  <- z[1, ]
    names(xlsx_cols_names) <- nams

    z  <- z[-1, , drop = FALSE]
    tt <- tt[-1, , drop = FALSE]
  }

  if (row_names) {
    rownames(z)  <- z[, 1]
    rownames(tt) <- z[, 1]
    xlsx_cols_names <- xlsx_cols_names[-1]

    z  <- z[, -1, drop = FALSE]
    tt <- tt[, -1, drop = FALSE]
  }

  # # faster guess_col_type alternative? to avoid tt
  # types <- ftable(cc$row_r ~ cc$c_r ~ cc$typ)

  date_conv     <- NULL
  datetime_conv <- NULL
  hms_conv      <- convert_hms

  if (missing(types)) {
    types <- guess_col_type(tt)
    date_conv     <- as.Date
    datetime_conv <- as.POSIXct
  } else {
    # assign types the correct column name "A", "B" etc.
    names(types) <- names(xlsx_cols_names[names(types) %in% xlsx_cols_names])
    date_conv     <- convert_date
    datetime_conv <- convert_datetime
  }

  # could make it optional or explicit
  if (convert) {
    sel <- !is.na(names(types))

    if (any(sel)) {
      nums <- names(which(types[sel] == 1))
      dtes <- names(which(types[sel] == 2))
      poxs <- names(which(types[sel] == 3))
      logs <- names(which(types[sel] == 4))
      difs <- names(which(types[sel] == 5))
      # convert "#NUM!" to "NaN" -- then converts to NaN
      # maybe consider this an option to instead return NA?
      if (length(nums)) z[nums] <- lapply(z[nums], function(i) as.numeric(replace(i, i == "#NUM!", "NaN")))
      if (length(dtes)) z[dtes] <- lapply(z[dtes], date_conv, origin = origin)
      if (length(poxs)) z[poxs] <- lapply(z[poxs], datetime_conv, origin = origin)
      if (length(logs)) z[logs] <- lapply(z[logs], as.logical)
      if (isNamespaceLoaded("hms")) z[difs] <- lapply(z[difs], hms_conv)
    } else {
      warning("could not convert. All missing in row used for variable names")
    }
  }

  if (col_names) {
    names(z)  <- xlsx_cols_names
    names(tt) <- xlsx_cols_names
  }

  if (keep_attributes) {
    attr(z, "tt") <- tt
    attr(z, "types") <- types
    # attr(z, "sd") <- sd
    if (!missing(named_region)) attr(z, "dn") <- nr
  }
  z
}

# `read_xlsx()` -----------------------------------------------------------------
# Ignored by roxygen2 when combining documentation
# #' Read from an Excel file or Workbook object
#' @rdname wb_to_df
#' @export
read_xlsx <- function(
  file,
  sheet,
  start_row         = 1,
  start_col         = NULL,
  row_names         = FALSE,
  col_names         = TRUE,
  skip_empty_rows   = FALSE,
  skip_empty_cols   = FALSE,
  rows              = NULL,
  cols              = NULL,
  detect_dates      = TRUE,
  named_region,
  na.strings        = "#N/A",
  na.numbers        = NA,
  fill_merged_cells = FALSE,
  ...
) {

  # keep sheet missing // read_xlsx is the function to replace.
  # dont mess with wb_to_df
  if (missing(file))
    file <- substitute()

  if (missing(sheet))
    sheet <- substitute()

  wb_to_df(
    file              = file,
    sheet             = sheet,
    start_row         = start_row,
    start_col         = start_col,
    row_names         = row_names,
    col_names         = col_names,
    skip_empty_rows   = skip_empty_rows,
    skip_empty_cols   = skip_empty_cols,
    rows              = rows,
    cols              = cols,
    detect_dates      = detect_dates,
    named_region      = named_region,
    na.strings        = na.strings,
    na.numbers        = na.numbers,
    fill_merged_cells = fill_merged_cells,
    ...
  )
}

# `wb_read()` ------------------------------------------------------------------
#' @rdname wb_to_df
#' @export
wb_read <- function(
  file,
  sheet           = 1,
  start_row       = 1,
  start_col       = NULL,
  row_names       = FALSE,
  col_names       = TRUE,
  skip_empty_rows = FALSE,
  skip_empty_cols = FALSE,
  rows            = NULL,
  cols            = NULL,
  detect_dates    = TRUE,
  named_region,
  na.strings      = "NA",
  na.numbers      = NA,
  ...
) {

  # keep sheet missing // read_xlsx is the function to replace.
  # dont mess with wb_to_df
  if (missing(file))
    file <- substitute()

  if (missing(sheet))
    sheet <- substitute()

  wb_to_df(
    file            = file,
    sheet           = sheet,
    start_row       = start_row,
    start_col       = start_col,
    row_names       = row_names,
    col_names       = col_names,
    skip_empty_rows = skip_empty_rows,
    skip_empty_cols = skip_empty_cols,
    rows            = rows,
    cols            = cols,
    detect_dates    = detect_dates,
    named_region    = named_region,
    na.strings      = na.strings,
    na.numbers      = na.numbers,
    ...
  )

}

#' Add the `wb_data` attribute to a data frame in a worksheet
#'
#' provide wb_data object as mschart input
#'
#' @param wb a workbook
#' @param sheet a sheet in the workbook either name or index
#' @param dims the dimensions
#' @param ... additional arguments for `wb_to_df()`. Be aware that not every
#' argument is valid.
#' @returns A data frame of class `wb_data`.
#' @seealso [wb_to_df()] [wb_add_mschart()], [wb_add_pivot_table()]
#' @examples
#'  wb <- wb_workbook() %>%
#'    wb_add_worksheet() %>%
#'    wb_add_data(x = mtcars, dims = "B2")
#'
#'  wb_data(wb, 1, dims = "B2:E6")
#' @export
wb_data <- function(wb, sheet = current_sheet(), dims, ...) {
  assert_workbook(wb)
  sheetno <- wb_validate_sheet(wb, sheet)
  sheetname <- wb$get_sheet_names(escape = TRUE)[[sheetno]]

  if (missing(dims)) {
    dims <- unname(unlist(xml_attr(wb$worksheets[[sheetno]]$dimension, "dimension")))
  }

  z <- wb_to_df(wb, sheet, dims = dims, ...)
  attr(z, "dims")  <- dims_to_dataframe(dims, fill = TRUE)
  attr(z, "sheet") <- sheetname

  class(z) <- c("wb_data", "data.frame")
  z
}

#' Extract or Replace Parts of an `wb_data` Object
#' @method [ wb_data
#' @param x x
#' @param i i
#' @param j j
#' @param drop drop
#' @rdname wb_data
#' @export
"[.wb_data" <- function(x, i, j, drop = ifelse((missing(j) && length(i) > 1) || (!missing(i) && length(j) > 1), FALSE, TRUE)) {

  sheet <- attr(x, "sheet")
  dd    <- attr(x, "dims")

  class(x) <- "data.frame"

  has_colnames <- as.integer(nrow(dd) - nrow(x))

  if (missing(j) && is.character(i)) {
    j <- match(i, colnames(x))
    i <- seq_len(nrow(x))
  }

  if (missing(i)) {
    i <- seq_len(nrow(x))
  }

  if (missing(j)) {
    j <- seq_along(x)
  }

  x <- x[i, j, drop]

  if (inherits(x, "data.frame")) {

    # we have the colnames in the first row
    if (all(i < 0)) {
      sel <- seq_len(nrow(dd))
      i <- sel[!sel %in% (abs(i) + has_colnames)]
    } else {
      i <- c(1, i + has_colnames)
    }

    dd <- dd[i, j, drop]
    attr(x, "dims")  <- dd
    attr(x, "sheet") <- sheet

    class(x) <- c("wb_data", "data.frame")
  }

  x
}
