
#' create dataframe from dimensions
#' @param dims Character vector of expected dimension.
#' @param fill If TRUE, fills the dataframe with variables
#' @examples {
#'   dims_to_dataframe("A1:B2")
#' }
#' @export
dims_to_dataframe <- function(dims, fill = FALSE) {

  if (grepl(";", dims)) {
    dims <- unlist(strsplit(dims, ";"))
  }

  rows_out <- NULL
  cols_out <- NULL
  for (dim in dims) {

    if (!grepl(":", dim)) {
      dim <- paste0(dim, ":", dim)
    }

    if (identical(dim, "Inf:-Inf")) {
      # This should probably be fixed elsewhere?
      stop("dims are inf:-inf")
    } else {
      dimensions <- strsplit(dim, ":")[[1]]

      rows <- as.numeric(gsub("[[:upper:]]", "", dimensions))
      if (all(is.na(rows))) rows <- c(1, 1048576)
      rows <- seq.int(rows[1], rows[2])

      rows_out <- unique(c(rows_out, rows))

      # TODO seq.wb_columns?  make a wb_cols vector?
      cols <- gsub("[[:digit:]]", "", dimensions)
      cols <- int2col(seq.int(col2int(cols[1]), col2int(cols[2])))

      cols_out <- unique(c(cols_out, cols))
    }
  }

  # create data frame from rows/
  dims_to_df(
    rows = rows_out,
    cols = cols_out,
    fill = fill
  )
}

#' create dimensions from dataframe
#' @param df dataframe with spreadsheet columns and rows
#' @examples {
#'   df <- dims_to_dataframe("A1:D5;F1:F6;D8", fill = TRUE)
#'   dataframe_to_dims(df)
#' }
#' @export
dataframe_to_dims <- function(df) {

  # get continuous sequences of columns and rows in df
  v <- as.integer(rownames(df))
  rows <- split(v, cumsum(diff(c(-Inf, v)) != 1))

  v <- col2int(colnames(df))
  cols <- split(colnames(df), cumsum(diff(c(-Inf, v)) != 1))

  # combine columns and rows to construct dims
  out <- NULL
  for (col in seq_along(cols)) {
    for (row in seq_along(rows)) {
      tmp <- paste0(
        cols[[col]][[1]], rows[[row]][[1]],
        ":",
        rev(cols[[col]])[[1]],  rev(rows[[row]])[[1]]
      )
      out <- c(out, tmp)
    }
  }

  paste0(out, collapse = ";")
}

#' function to estimate the column type.
#' 0 = character, 1 = numeric, 2 = date.
#' @param tt dataframe produced by wb_to_df()
#' @export
guess_col_type <- function(tt) {

  # all columns are character
  types <- vector("numeric", NCOL(tt))
  names(types) <- names(tt)

  # but some values are numeric
  col_num <- vapply(tt, function(x) all(x == "n", na.rm = TRUE), NA)
  types[names(col_num[col_num])] <- 1

  # or even date
  col_dte <- vapply(tt[!col_num], function(x) all(x == "d", na.rm = TRUE), NA)
  types[names(col_dte[col_dte])] <- 2

  # or even posix
  col_dte <- vapply(tt[!col_num], function(x) all(x == "p", na.rm = TRUE), NA)
  types[names(col_dte[col_dte])] <- 3

  # there are bools as well
  col_log <- vapply(tt[!col_num], function(x) any(x == "b", na.rm = TRUE), NA)
  types[names(col_log[col_log])] <- 4

  types
}

#' check if numFmt is date. internal function
#' @param numFmt numFmt xml nodes
numfmt_is_date <- function(numFmt) {

  # if numFmt is character(0)
  if (length(numFmt) == 0) return(z <- NULL)

  numFmt_df <- read_numfmt(read_xml(numFmt))
  # we have to drop any square bracket part
  numFmt_df$fC <- gsub("\\[[^\\]]*]", "", numFmt_df$formatCode, perl = TRUE)
  num_fmts <- c(
    "#", as.character(0:9)
  )
  num_or_fmt <- paste0(num_fmts, collapse = "|")
  maybe_num <- grepl(pattern = num_or_fmt, x = numFmt_df$fC)

  date_fmts <- c(
    "yy", "yyyy",
    "m", "mm", "mmm", "mmmm", "mmmmm",
    "d", "dd", "ddd", "dddd"
  )
  date_or_fmt <- paste0(date_fmts, collapse = "|")
  maybe_dates <- grepl(pattern = date_or_fmt, x = numFmt_df$fC)

  z <- numFmt_df$numFmtId[maybe_dates & !maybe_num]
  if (length(z) == 0) z <- NULL
  z
}

#' check if numFmt is posix. internal function
#' @param numFmt numFmt xml nodes
numfmt_is_posix <- function(numFmt) {

  # if numFmt is character(0)
  if (length(numFmt) == 0) return(z <- NULL)

  numFmt_df <- read_numfmt(read_xml(numFmt))
  # we have to drop any square bracket part
  numFmt_df$fC <- gsub("\\[[^\\]]*]", "", numFmt_df$formatCode, perl = TRUE)
  num_fmts <- c(
    "#", as.character(0:9)
  )
  num_or_fmt <- paste0(num_fmts, collapse = "|")
  maybe_num <- grepl(pattern = num_or_fmt, x = numFmt_df$fC)

  posix_fmts <- c(
    # "yy", "yyyy",
    # "m", "mm", "mmm", "mmmm", "mmmmm",
    # "d", "dd", "ddd", "dddd",
    "h", "hh", ":m", ":mm", ":s", ":ss",
    "AM", "PM", "A", "P"
  )
  posix_or_fmt <- paste0(posix_fmts, collapse = "|")
  maybe_posix <- grepl(pattern = posix_or_fmt, x = numFmt_df$fC)

  z <- numFmt_df$numFmtId[maybe_posix & !maybe_num]
  if (length(z) == 0) z <- NULL
  z
}

#' check if style is date. internal function
#'
#' @param cellXfs cellXfs xml nodes
#' @param numfmt_date custom numFmtId dates
style_is_date <- function(cellXfs, numfmt_date) {

  # numfmt_date: some basic date formats and custom formats
  date_numfmts <- as.character(14:17)
  numfmt_date <- c(numfmt_date, date_numfmts)

  cellXfs_df <- read_xf(read_xml(cellXfs))
  z <- rownames(cellXfs_df[cellXfs_df$numFmtId %in% numfmt_date, ])
  if (length(z) == 0) z <- NA
  z
}

#' check if style is posix. internal function
#'
#' @param cellXfs cellXfs xml nodes
#' @param numfmt_date custom numFmtId dates
style_is_posix <- function(cellXfs, numfmt_date) {

  # numfmt_date: some basic date formats and custom formats
  date_numfmts <- as.character(18:22)
  numfmt_date <- c(numfmt_date, date_numfmts)

  cellXfs_df <- read_xf(read_xml(cellXfs))
  z <- rownames(cellXfs_df[cellXfs_df$numFmtId %in% numfmt_date, ])
  if (length(z) == 0) z <- NA
  z
}

#' Create Dataframe from Workbook
#'
#' Simple function to create a dataframe from a workbook. Simple as in simply
#' written down and not optimized etc. The goal was to have something working.
#'
#' @param xlsxFile An xlsx file, Workbook object or URL to xlsx file.
#' @param sheet Either sheet name or index. When missing the first sheet in the workbook is selected.
#' @param colNames If TRUE, the first row of data will be used as column names.
#' @param rowNames If TRUE, the first col of data will be used as row names.
#' @param dims Character string of type "A1:B2" as optional dimensions to be imported.
#' @param detectDates If TRUE, attempt to recognize dates and perform conversion.
#' @param showFormula If TRUE, the underlying Excel formulas are shown.
#' @param convert If TRUE, a conversion to dates and numerics is attempted.
#' @param skipEmptyCols If TRUE, empty columns are skipped.
#' @param skipEmptyRows If TRUE, empty rows are skipped.
#' @param startRow first row to begin looking for data.
#' @param startCol first column to begin looking for data.
#' @param rows A numeric vector specifying which rows in the Excel file to read. If NULL, all rows are read.
#' @param cols A numeric vector specifying which columns in the Excel file to read. If NULL, all columns are read.
#' @param named_region Character string with a named_region (defined name or table). If no sheet is selected, the first appearance will be selected.
#' @param types A named numeric indicating, the type of the data. 0: character, 1: numeric, 2: date, 3: posixt, 4:logical. Names must match the returned data
#' @param na.strings A character vector of strings which are to be interpreted as NA. Blank cells will be returned as NA.
#' @param na.numbers A numeric vector of digits which are to be interpreted as NA. Blank cells will be returned as NA.
#' @param fillMergedCells If TRUE, the value in a merged cell is given to all cells within the merge.
#' @examples
#'
#'   ###########################################################################
#'   # numerics, dates, missings, bool and string
#'   xlsxFile <- system.file("extdata", "readTest.xlsx", package = "openxlsx2")
#'   wb1 <- wb_load(xlsxFile)
#'
#'   # import workbook
#'   wb_to_df(wb1)
#'
#'   # do not convert first row to colNames
#'   wb_to_df(wb1, colNames = FALSE)
#'
#'   # do not try to identify dates in the data
#'   wb_to_df(wb1, detectDates = FALSE)
#'
#'   # return the underlying Excel formula instead of their values
#'   wb_to_df(wb1, showFormula = TRUE)
#'
#'   # read dimension withot colNames
#'   wb_to_df(wb1, dims = "A2:C5", colNames = FALSE)
#'
#'   # read selected cols
#'   wb_to_df(wb1, cols = c(1:2, 7))
#'
#'   # read selected rows
#'   wb_to_df(wb1, rows = c(1, 4, 6))
#'
#'   # convert characters to numerics and date (logical too?)
#'   wb_to_df(wb1, convert = FALSE)
#'
#'   # erase empty Rows from dataset
#'   wb_to_df(wb1, sheet = 3, skipEmptyRows = TRUE)
#'
#'   # erase rmpty Cols from dataset
#'   wb_to_df(wb1, skipEmptyCols = TRUE)
#'
#'   # convert first row to rownames
#'   wb_to_df(wb1, sheet = 3, dims = "C6:G9", rowNames = TRUE)
#'
#'   # define type of the data.frame
#'   wb_to_df(wb1, cols = c(1, 4), types = c("Var1" = 0, "Var3" = 1))
#'
#'   # start in row 5
#'   wb_to_df(wb1, startRow = 5, colNames = FALSE)
#'
#'   # na string
#'   wb_to_df(wb1, na.strings = "")
#'
#'   # read_xlsx(wb1)
#'
#'   ###########################################################################
#'   # inlinestr
#'   xlsxFile <- system.file("extdata", "inline_str.xlsx", package = "openxlsx2")
#'   wb2 <- wb_load(xlsxFile)
#'
#'   # read dataset with inlinestr
#'   wb_to_df(wb2)
#'   # read_xlsx(wb2)
#'
#'   ###########################################################################
#'   # named_region // namedRegion
#'   xlsxFile <- system.file("extdata", "namedRegions3.xlsx", package = "openxlsx2")
#'   wb3 <- wb_load(xlsxFile)
#'
#'   # read dataset with named_region (returns global first)
#'   wb_to_df(wb3, named_region = "MyRange", colNames = FALSE)
#'
#'   # read named_region from sheet
#'   wb_to_df(wb3, named_region = "MyRange", sheet = 4, colNames = FALSE)
#'
#' @export
wb_to_df <- function(
    xlsxFile,
    sheet,
    startRow        = 1,
    startCol        = NULL,
    rowNames        = FALSE,
    colNames        = TRUE,
    skipEmptyRows   = FALSE,
    skipEmptyCols   = FALSE,
    rows            = NULL,
    cols            = NULL,
    detectDates     = TRUE,
    na.strings      = "#N/A",
    na.numbers      = NA,
    fillMergedCells = FALSE,
    dims,
    showFormula     = FALSE,
    convert         = TRUE,
    types,
    named_region
) {

  # .mc <- match.call() # not (yet) used?

  if (!is.null(cols)) cols <- col2int(cols)

  if (inherits(xlsxFile, "wbWorkbook")) {
    wb <- xlsxFile
  } else {
    # passes missing further on
    if (missing(sheet))
      sheet <- substitute()

    # possible false positive on current lintr runs
    wb <- wb_load(xlsxFile, sheet = sheet, data_only = TRUE) # nolint
  }

  if (!missing(named_region)) {

    nr <- wb_get_named_regions(wb)

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

  numfmt_posix <- numfmt_is_posix(wb$styles_mgr$styles$numFmts)
  xlsx_posix_style <- style_is_posix(wb$styles_mgr$styles$cellXfs, numfmt_posix)

  # create temporary data frame. hard copy required
  z  <- dims_to_dataframe(dims)
  tt <- copy(z)

  keep_cols <- colnames(z)
  keep_rows <- rownames(z)

  maxRow <- max(as.numeric(keep_rows))
  maxCol <- max(col2int(keep_cols))

  if (startRow > 1) {
    keep_rows <- as.character(seq(startRow, maxRow))
    if (startRow <= maxRow) {
      z  <- z[rownames(z) %in% keep_rows, , drop = FALSE]
      tt <- tt[rownames(tt) %in% keep_rows, , drop = FALSE]
    } else {
      keep_rows <- as.character(startRow)
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

  if (!is.null(startCol)) {
    keep_cols <- int2col(seq(col2int(startCol), maxCol))

    if (!all(keep_cols %in% colnames(z))) {
      keep_col <- keep_cols[!keep_cols %in% colnames(z)]

      z[keep_col]  <- NA_character_
      tt[keep_col] <- NA_character_

      # return expected order of columns
      z  <- z[keep_cols]
      tt <- tt[keep_cols]
    }


    z  <- z[, colnames(z) %in% keep_cols, drop = FALSE]
    tt <- tt[, colnames(tt) %in% keep_cols, drop = FALSE]
  }

  if (!is.null(cols)) {
    keep_cols <- int2col(cols)

    if (!all(keep_cols %in% colnames(z))) {
      keep_col <- keep_cols[!keep_cols %in% colnames(z)]

      z[keep_col] <- NA_character_
      tt[keep_col] <- NA_character_
    }

    z  <- z[, colnames(z) %in% keep_cols, drop = FALSE]
    tt <- tt[, colnames(tt) %in% keep_cols, drop = FALSE]
  }

  keep_rows <- keep_rows[keep_rows %in% rnams]

  # reduce data to selected cases only
  if (!is.null(cols) && !is.null(rows) && !missing(dims))
    cc <- cc[cc$row_r %in% keep_rows & cc$c_r %in% keep_cols, ]

  cc$val <- NA_character_
  cc$typ <- NA_character_

  cc_tab <- unique(cc$c_t)

  # bool
  if (any(cc_tab == c("b"))) {
    sel <- cc$c_t %in% c("b")
    cc$val[sel] <- as.logical(as.numeric(cc$v[sel]))
    cc$typ[sel] <- "b"
  }
  # text in v
  if (any(cc_tab %in% c("str", "e"))) {
    sel <- cc$c_t %in% c("str", "e")
    cc$val[sel] <- replaceXMLEntities(cc$v[sel])
    cc$typ[sel] <- "s"
  }
  if (showFormula) {
    sel <- cc$f != ""
    cc$val[sel] <- replaceXMLEntities(cc$f[sel])
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

  # dates
  if (!is.null(cc$c_s)) {
    # if a cell is t="s" the content is a sst and not da date
    if (detectDates && missing(types)) {
      cc$is_string <- FALSE
      if (!is.null(cc$c_t))
        cc$is_string <- cc$c_t %in% c("s", "str", "b", "inlineStr")

      if (any(sel <- cc$c_s %in% xlsx_date_style)) {
        sel <- sel & !cc$is_string & cc$v != ""
        cc$val[sel] <- suppressWarnings(as.character(convert_date(cc$v[sel])))
        cc$typ[sel]  <- "d"
      }

      if (any(sel <- cc$c_s %in% xlsx_posix_style)) {
        sel <- sel & !cc$is_string & cc$v != ""
        cc$val[sel] <- suppressWarnings(as.character(convert_datetime(cc$v[sel])))
        cc$typ[sel]  <- "p"
      }
    }
  }

  # remaining values are numeric?
  if (any(sel <- is.na(cc$typ))) {
    cc$val[sel] <- cc$v[sel]
    cc$typ[sel] <- "n"
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
  if (fillMergedCells) {
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
            dims = filler,
            xlsxFile = xlsxFile,
            sheet = sheet,
            na.strings = na.strings,
            convert = FALSE,
            colNames = FALSE,
            detectDates = detectDates,
            showFormula = showFormula
          )

          tt_fill <- attr(z_fill, "tt")

          z[row_sel,  col_sel] <- z_fill
          tt[row_sel, col_sel] <- tt_fill
        }
      }
    }

  }

  # is.na needs convert
  if (skipEmptyRows) {
    empty <- vapply(seq_len(nrow(z)), function(x) all(is.na(z[x, ])), NA)

    z  <- z[!empty, , drop = FALSE]
    tt <- tt[!empty, , drop = FALSE]
  }

  if (skipEmptyCols) {

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
  if (colNames) {
    # select first row as colnames, but do not yet assing. it might contain
    # missing values and if assigned, convert below might break with unambiguous
    # names.
    nams <- names(xlsx_cols_names)
    xlsx_cols_names  <- z[1, ]
    names(xlsx_cols_names) <- nams

    z  <- z[-1, , drop = FALSE]
    tt <- tt[-1, , drop = FALSE]
  }

  if (rowNames) {
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
      # convert "#NUM!" to "NaN" -- then converts to NaN
      # maybe consider this an option to instead return NA?
      if (length(nums)) z[nums] <- lapply(z[nums], function(i) as.numeric(replace(i, i == "#NUM!", "NaN")))
      if (length(dtes)) z[dtes] <- lapply(z[dtes], date_conv)
      if (length(poxs)) z[poxs] <- lapply(z[poxs], datetime_conv)
      if (length(logs)) z[logs] <- lapply(z[logs], as.logical)
    } else {
      warning("could not convert. All missing in row used for variable names")
    }
  }

  if (colNames) {
    names(z)  <- xlsx_cols_names
    names(tt) <- xlsx_cols_names
  }

  attr(z, "tt") <- tt
  attr(z, "types") <- types
  # attr(z, "sd") <- sd
  if (!missing(named_region)) attr(z, "dn") <- nr
  z
}


#' @rdname cleanup
#' @param wb workbook
#' @param sheet sheet to clean
#' @param cols numeric column vector
#' @param rows numeric row vector
#' @export
delete_data <- function(wb, sheet, cols, rows) {

  sheet_id <- wb_validate_sheet(wb, sheet)

  cc <- wb$worksheets[[sheet_id]]$sheet_data$cc

  if (is.numeric(cols)) {
    sel <- cc$row_r %in% as.character(rows) & cc$c_r %in% int2col(cols)
  } else {
    sel <- cc$row_r %in% as.character(rows) & cc$c_r %in% cols
  }

  # clean selected entries of cc
  clean <- names(cc)[!names(cc) %in% c("r", "row_r", "c_r")]
  cc[sel, clean] <- ""

  wb$worksheets[[sheet_id]]$sheet_data$cc <- cc

}


#' Get a worksheet from a `wbWorkbook` object
#'
#' @param wb a [wbWorkbook] object
#' @param sheet A sheet name or index
#' @returns A `wbWorksheet` object
#' @export
wb_get_worksheet <- function(wb, sheet) {
  assert_workbook(wb)
  wb$get_worksheet(sheet)
}

#' @rdname wb_get_worksheet
#' @export
wb_ws <- wb_get_worksheet

#' get and set table of sheets and their state as selected and active
#' @description Multiple sheets can be selected, but only a single one can be
#' active (visible). The visible sheet, must not necessarily be a selected
#' sheet.
#' @param wb a workbook
#' @returns a data frame with tabSelected and names
#' @export
#' @examples
#'   wb <- wb_load(file = system.file("extdata", "loadExample.xlsx", package = "openxlsx2"))
#'   # testing is the selected sheet
#'   wb_get_selected(wb)
#'   # change the selected sheet to IrisSample
#'   wb <- wb_set_selected(wb, "IrisSample")
#'   # get the active sheet
#'   wb_get_active_sheet(wb)
#'   # change the selected sheet to IrisSample
#'   wb <- wb_set_active_sheet(wb, sheet = "IrisSample")
#' @name select_active_sheet
wb_get_active_sheet <- function(wb) {
  assert_workbook(wb)
  at <- rbindlist(xml_attr(wb$workbook$bookViews, "bookViews", "workbookView"))["activeTab"]
  # return c index as R index
  as.numeric(at) + 1
}

#' @rdname select_active_sheet
#' @param sheet a sheet name of the workbook
#' @export
wb_set_active_sheet <- function(wb, sheet) {
  # active tab requires a c index
  assert_workbook(wb)
  sheet <- wb_validate_sheet(wb, sheet)
  wb$clone()$set_bookview(activeTab = sheet - 1L)
}

#' @name select_active_sheet
#' @export
wb_get_selected <- function(wb) {

  assert_workbook(wb)

  len <- length(wb$sheet_names)
  sv <- vector("list", length = len)

  for (i in seq_len(len)) {
    sv[[i]] <- xml_node(wb$worksheets[[i]]$sheetViews, "sheetViews", "sheetView")
  }

  # print(sv)
  z <- rbindlist(xml_attr(sv, "sheetView"))
  z$names <- wb$get_sheet_names()

  z
}

#' @name select_active_sheet
#' @export
wb_set_selected <- function(wb, sheet) {

  sheet <- wb_validate_sheet(wb, sheet)

  for (i in seq_along(wb$sheet_names)) {
    xml_attr <- ifelse(i == sheet, TRUE, FALSE)
    wb$worksheets[[i]]$set_sheetview(tabSelected = xml_attr)
  }

  wb
}

#' Add mschart object to an existing workbook
#' @param wb a workbook
#' @param sheet the sheet on which the graph will appear
#' @param dims the dimensions where the sheet will appear
#' @param graph mschart object
#' @examples
#' if (requireNamespace("mschart")) {
#' require(mschart)
#'
#' ## Add mschart to worksheet (adds data and chart)
#' scatter <- ms_scatterchart(data = iris, x = "Sepal.Length", y = "Sepal.Width", group = "Species")
#' scatter <- chart_settings(scatter, scatterstyle = "marker")
#'
#' wb <- wb_workbook() %>%
#'  wb_add_worksheet() %>%
#'  wb_add_mschart(dims = "F4:L20", graph = scatter)
#'
#' ## Add mschart to worksheet and use available data
#' wb <- wb_workbook() %>%
#'   wb_add_worksheet() %>%
#'   wb_add_data(x = mtcars, dims = "B2")
#'
#' # create wb_data object
#' dat <- wb_data(wb, 1, dims = "B2:E6")
#'
#' # call ms_scatterplot
#' data_plot <- ms_scatterchart(
#'   data = dat,
#'   x = "mpg",
#'   y = c("disp", "hp"),
#'   labels = c("disp", "hp")
#' )
#'
#' # add the scatterplot to the data
#' wb <- wb %>%
#'   wb_add_mschart(dims = "F4:L20", graph = data_plot)
#' }
#' @seealso [wb_data()]
#' @export
wb_add_mschart <- function(
    wb,
    sheet = current_sheet(),
    dims = NULL,
    graph
) {
  assert_workbook(wb)
  wb$clone()$add_mschart(sheet = sheet, dims = dims, graph = graph)
}

#' provide wb_data object as mschart input
#' @param wb a workbook
#' @param sheet a sheet in the workbook either name or index
#' @param dims the dimensions
#' @param ... additional arguments for wb_to_df. Be aware that not every
#' argument is valid.
#' @seealso [wb_to_df()] [wb_add_mschart()]
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
  sheetname <- wb$get_sheet_names()[[sheetno]]

  if (missing(dims)) {
    dims <- unname(unlist(xml_attr(wb$worksheets[[sheetno]]$dimension, "dimension")))
  }

  z <- wb_to_df(wb, sheet, dims = dims, ...)
  attr(z, "dims")  <- dims_to_dataframe(dims, fill = TRUE)
  attr(z, "sheet") <- sheetname

  class(z) <- c("data.frame", "wb_data")
  z
}
