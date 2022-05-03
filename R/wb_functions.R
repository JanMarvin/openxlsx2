
#' create dataframe from dimensions
#' @param dims Character vector of expected dimension.
#' @param fill If TRUE, fills the dataframe with variables
#' @examples {
#'   dims_to_dataframe("A1:B2")
#' }
#' @export
dims_to_dataframe <- function(dims, fill = FALSE) {

  if (!grepl(":", dims)) {
    dims <- paste0(dims, ":", dims)
  }

  if (identical(dims, "Inf:-Inf")) {
    # This should probably be fixed elsewhere?
    stop("dims are inf:-inf")
  } else {
    dimensions <- strsplit(dims, ":")[[1]]

    rows <- as.numeric(gsub("[[:upper:]]","", dimensions))
    rows <- seq.int(rows[1], rows[2])

    # TODO seq.wb_columns?  make a wb_cols vector?
    cols <- gsub("[[:digit:]]","", dimensions)
    cols <- int2col(seq.int(col2int(cols[1]), col2int(cols[2])))
  }

  # create data frame from rows/
  dims_to_df(
    rows = rows,
    cols = cols,
    fill = fill
  )
}

# # similar to all, simply check if most of the values match the condition
# # in guess_col_type not all bools may be "b" some are "s" (missings)
# most <- function(x) {
#   as.logical(names(sort(table(x), decreasing = TRUE)[1]))
# }

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
  if (length(numFmt) ==0) return(z <- NULL)

  numFmt_df <- read_numfmt(read_xml(numFmt))
  num_fmts <- c(
    "#", as.character(0:9)
  )
  num_or_fmt <- paste0(num_fmts, collapse = "|")
  maybe_num <- grepl(pattern = num_or_fmt, x = numFmt_df$formatCode)

  date_fmts <- c(
    "yy", "yyyy",
    "m", "mm", "mmm", "mmmm", "mmmmm",
    "d", "dd", "ddd", "dddd"
  )
  date_or_fmt <- paste0(date_fmts, collapse = "|")
  maybe_dates <- grepl(pattern = date_or_fmt, x = numFmt_df$formatCode)

  z <- numFmt_df$numFmtId[maybe_dates & !maybe_num]
  if (length(z)==0) z <- NULL
  z
}

#' check if numFmt is posix. internal function
#' @param numFmt numFmt xml nodes
numfmt_is_posix <- function(numFmt) {

  # if numFmt is character(0)
  if (length(numFmt) ==0) return(z <- NULL)

  numFmt_df <- read_numfmt(read_xml(numFmt))
  num_fmts <- c(
    "#", as.character(0:9)
  )
  num_or_fmt <- paste0(num_fmts, collapse = "|")
  maybe_num <- grepl(pattern = num_or_fmt, x = numFmt_df$formatCode)

  posix_fmts <- c(
    "yy", "yyyy",
    "m", "mm", "mmm", "mmmm", "mmmmm",
    "d", "dd", "ddd", "dddd",
    "h", "hh", "m", "mm", "s", "ss",
    "AM", "PM", "A", "P"
  )
  posix_or_fmt <- paste0(posix_fmts, collapse = "|")
  maybe_posix <- grepl(pattern = posix_or_fmt, x = numFmt_df$formatCode)

  z <- numFmt_df$numFmtId[maybe_posix & !maybe_num]
  if (length(z)==0) z <- NULL
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
  z <- rownames(cellXfs_df[cellXfs_df$numFmtId %in% numfmt_date,])
  if (length(z)==0) z <- NA
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
  z <- rownames(cellXfs_df[cellXfs_df$numFmtId %in% numfmt_date,])
  if (length(z)==0) z <- NA
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
#' @param rows A numeric vector specifying which rows in the Excel file to read. If NULL, all rows are read.
#' @param cols A numeric vector specifying which columns in the Excel file to read. If NULL, all columns are read.
#' @param definedName Character string with a definedName. If no sheet is selected, the first appearance will be selected.
#' @param types A named numeric indicating, the type of the data. 0: character, 1: numeric, 2: date. Names must match the created
#' @param na.strings A character vector of strings which are to be interpreted as NA. Blank cells will be returned as NA.
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
#'   # definedName // namedRegion
#'   xlsxFile <- system.file("extdata", "namedRegions3.xlsx", package = "openxlsx2")
#'   wb3 <- wb_load(xlsxFile)
#'
#'   # read dataset with definedName (returns global first)
#'   wb_to_df(wb3, definedName = "MyRange", colNames = FALSE)
#'
#'   # read definedName from sheet
#'   wb_to_df(wb3, definedName = "MyRange", sheet = 4, colNames = FALSE)
#'
#' @export
wb_to_df <- function(
  xlsxFile,
  sheet,
  startRow        = 1,
  colNames        = TRUE,
  rowNames        = FALSE,
  detectDates     = TRUE,
  skipEmptyCols   = FALSE,
  skipEmptyRows   = FALSE,
  rows            = NULL,
  cols            = NULL,
  na.strings      = "#N/A",
  fillMergedCells = FALSE,
  dims,
  showFormula     = FALSE,
  convert         = TRUE,
  types,
  definedName
) {

  # .mc <- match.call() # not (yet) used?

  if (is.character(xlsxFile)) {
    # TODO this should instead check for the Workbook class?  Maybe also check
    # if the file exists?

    # passes missing further on
    if (missing(sheet))
      sheet <- substitute()

    # possible false positive on current lintr runs
    wb <- wb_load(xlsxFile, sheet = sheet) # nolint
  } else {
    wb <- xlsxFile
  }


  if (!missing(definedName)) {

    nr <- get_named_regions(wb)

    if ((definedName %in% nr$name) && missing(sheet)) {
      sel   <- nr[nr$name == definedName, ][1,]
      sheet <- sel$sheet
      dims  <- sel$coords
    } else if (definedName %in% nr$name) {
      sel <- nr[nr$name == definedName & nr$sheet == wb_validate_sheet(wb, sheet), ]
      if (NROW(sel) == 0) {
        stop("no such definedName on selected sheet")
      }
      dims <- sel$coords
    } else {
      stop("no such definedName")
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
  if (is.null(wb$worksheets[[sheet]]$sheet_data$cc)) {
    # TODO do we need more checks or do we need to initialize a new cc object?
    # TODO would this also apply of nrow(cc) == 0?
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

  # If no dims are requested via definedName, simply construct them from min
  # and max columns and row found on worksheet
  # in theory it could be useful to have both definedName and dims?
  if (missing(definedName) && missing(dims)) {

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

  # tt <- data.frame(matrix(0, nrow = 4, ncol = ncol(z)))
  # names(tt) <- names(z)
  # rownames(tt) <- c("b", "s", "d", "n")

  keep_cols <- colnames(z)
  keep_rows <- rownames(z)

  if (startRow > 1) {
    keep_rows <- as.character(seq(startRow, max(as.numeric(keep_rows))))

    z  <- z[rownames(z) %in% keep_rows,]
    tt <- tt[rownames(tt) %in% keep_rows,]
  }

  if (!is.null(rows)) {
    keep_rows <- as.character(rows)

    z  <- z[rownames(z) %in% keep_rows,]
    tt <- tt[rownames(tt) %in% keep_rows,]
  }

  if (!is.null(cols)) {
    keep_cols <- int2col(cols)

    z  <- z[keep_cols]
    tt <- tt[keep_cols]
  }

  keep_rows <- keep_rows[keep_rows %in% rnams]

  # reduce data to selected cases only
  cc <- cc[cc$row_r %in% keep_rows & cc$c_r %in% keep_cols, ]

  # if (!nrow(cc)) browser()

  cc$val <- NA
  cc$typ <- NA

  # bool
  sel <- cc$c_t %in% c("b")
  cc$val[sel] <- as.logical(as.numeric(cc$v[sel]))
  cc$typ[sel] <- "b"
  # text in v
  sel <- cc$c_t %in% c("str", "e")
  cc$val[sel] <- cc$v[sel]
  cc$typ[sel] <- "s"
  if (showFormula) {
    sel <- cc$c_t %in% c("e")
    cc$val[sel] <- cc$f[sel]
    cc$typ[sel] <- "s"
  }
  # text in t
  sel <- cc$c_t %in% c("inlineStr")
  cc$val[sel] <- is_to_txt(cc$is[sel])
  cc$typ[sel] <- "s"
  # test is sst
  sel <- cc$c_t %in% c("s")
  cc$val[sel] <- sst[as.numeric(cc$v[sel])+1]
  cc$typ[sel] <- "s"
  # convert missings
  if (!is.na(na.strings) || !missing(na.strings)) {
    sel <- cc$val %in% na.strings
    cc$val[sel] <- NA_character_
    cc$typ[sel] <- "na_string"
  }

  # dates
  if (!is.null(cc$c_s)) {
    # if a cell is t="s" the content is a sst and not da date
    cc$is_string <- FALSE
    if (!is.null(cc$c_t))
      cc$is_string <- cc$c_t %in% c("s", "str", "b", "inlineStr")

    if (detectDates) {
      sel <- (cc$c_s %in% xlsx_date_style) & !cc$is_string & cc$v != "_openxlsx_NA_"
      cc$val[sel] <- suppressWarnings(as.character(convert_date(cc$v[sel])))
      cc$typ[sel]  <- "d"

      sel <- (cc$c_s %in% xlsx_posix_style) & !cc$is_string & cc$v != "_openxlsx_NA_"
      cc$val[sel] <- suppressWarnings(as.character(convert_datetime(cc$v[sel])))
      cc$typ[sel]  <- "p"
    }
  }

  # remaining values are numeric?
  sel <- is.na(cc$typ)
  cc$val[sel] <- cc$v[sel]
  cc$typ[sel] <- "n"

  # convert "na_string" to missing
  cc$typ[cc$typ == "na_string"] <- NA

  # prepare to create output object z
  zz <- cc[c("val", "typ")]
  # we need to create the correct col and row position as integer starting at 0.
  zz$cols <- match(cc$c_r, colnames(z)) - 1
  zz$rows <- match(cc$row_r, rownames(z)) - 1

  zz <- zz[order(zz[, "cols"], zz[,"rows"]),]
  zz <- zz[zz$val != "_openxlsx_NA_",]
  long_to_wide(z, tt, zz)

  # prepare colnames object
  xlsx_cols_names <- colnames(z)
  names(xlsx_cols_names) <- xlsx_cols_names

  # backward compatible option. get the mergedCells dimension and fill it with
  # the value of the first cell in the range. do the same for tt.
  if (fillMergedCells) {
    mc <- wb$worksheets[[sheet]]$mergeCells
    if (length(mc)) {

      mc <- unlist(xml_attr(mc, "mergeCell"))

      for (i in seq_along(mc)) {
        dms <- dims_to_dataframe(mc[i])

        z[rownames(z) %in% rownames(dms),
          colnames(z) %in% colnames(dms)] <- z[rownames(z) %in% rownames(dms[1, 1, drop = FALSE]),
                                               colnames(z) %in% colnames(dms[1, 1, drop = FALSE])]
        tt[rownames(tt) %in% rownames(dms),
           colnames(tt) %in% colnames(dms)] <- tt[rownames(tt) %in% rownames(dms[1, 1, drop = FALSE]),
                                                  colnames(tt) %in% colnames(dms[1, 1, drop = FALSE])]
      }

    }

  }

  # if colNames, then change tt too
  if (colNames) {
    # select first row as colnames, but do not yet assing. it might contain
    # missing values and if assigned, convert below might break with unambiguous
    # names.
    nams <- names(xlsx_cols_names)
    xlsx_cols_names  <- z[1,]
    names(xlsx_cols_names) <- nams

    z  <- z[-1, , drop = FALSE]
    tt <- tt[-1, , drop = FALSE]
  }

  if (rowNames) {
    rownames(z)  <- z[,1]
    rownames(tt) <- z[,1]
    xlsx_cols_names <- xlsx_cols_names[-1]

    z  <- z[ ,-1]
    tt <- tt[ , -1]
  }

  # # faster guess_col_type alternative? to avoid tt
  # types <- ftable(cc$row_r ~ cc$c_r ~ cc$typ)

  if (missing(types)) {
    types <- guess_col_type(tt)
  } else {
    # assign types the correct column name "A", "B" etc.
    names(types) <- names(xlsx_cols_names[names(types) %in% xlsx_cols_names])
  }

  # could make it optional or explicit
  if (convert) {
    sel <- !is.na(names(types))

    if (any(sel)) {
      nums <- names( which(types[sel] == 1) )
      dtes <- names( which(types[sel] == 2) )
      poxs <- names( which(types[sel] == 3) )
      logs <- names( which(types[sel] == 4) )
      # convert "#NUM!" to "NaN" -- then converts to NaN
      # maybe consider this an option to instead return NA?
      z[nums] <- lapply(z[nums], function(i) as.numeric(replace(i, i == "#NUM!", "NaN")))
      z[dtes] <- lapply(z[dtes], as.Date)
      z[poxs] <- lapply(z[poxs], as.POSIXct)
      z[logs] <- lapply(z[logs], as.logical)
    } else {
      warning("could not convert. All missing in row used for variable names")
    }
  }

  if (colNames) {
    names(z) <- xlsx_cols_names
    names(tt) <- xlsx_cols_names
  }

  # is.na needs convert
  if (skipEmptyRows) {
    empty <- apply(z, 1, function(x) all(is.na(x)), simplify = TRUE)

    z  <- z[!empty, , drop = FALSE]
    tt <- tt[!empty, , drop = FALSE]
  }

  if (skipEmptyCols) {

    empty <- vapply(z, function(x) all(is.na(x)), NA)

    if (any(empty)) {
      sel <- which(names(empty) %in% names(empty[empty == TRUE]))
      z[sel]  <- NULL
      tt[sel] <- NULL
    }

  }

  attr(z, "tt") <- tt
  attr(z, "types") <- types
  # attr(z, "sd") <- sd
  if (!missing(definedName)) attr(z, "dn") <- nr
  z
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


#' @rdname cleanup
#' @param wb workbook
#' @param sheet sheet to clean
#' @param cols numeric column vector
#' @param rows numeric row vector
#' @param gridExpand does nothing
#' @export
delete_data <- function(wb, sheet, cols, rows, gridExpand) {

  sheet_id <- wb_validate_sheet(wb, sheet)

  cc <- wb$worksheets[[sheet_id]]$sheet_data$cc

  if (is.numeric(cols)) {
    sel <- cc$row_r %in% as.character(rows) & cc$c_r %in% int2col(cols)
  } else {
    sel <- cc$row_r %in% as.character(rows) & cc$c_r %in% cols
  }

  cc[sel, ] <- "_openxlsx_NA_"

  wb$worksheets[[sheet_id]]$sheet_data$cc <- cc

}

#' clean sheet (remove all values)
#'
#' @param wb workbook
#' @param sheet sheet to clean
#' @param numbers remove all numbers
#' @param characters remove all characters
#' @param styles remove all styles
#' @param merged_cells remove all merged_cells
#' @name cleanup
#' @export
cleanSheet <- function(wb, sheet, numbers = TRUE, characters = TRUE, styles = TRUE, merged_cells = TRUE) {

  sheet_id <- wb_validate_sheet(wb, sheet)

  cc <- wb$worksheets[[sheet_id]]$sheet_data$cc

  if (numbers)
    cc[cc$c_t %in% c("n", "_openxlsx_NA_"), # imported values might be _NA_
       c("c_t", "v", "f", "f_t", "f_ref", "f_ca", "f_si", "is")] <- "_openxlsx_NA_"

  if (characters)
    cc[cc$c_t %in% c("inlineStr", "s"),
       c("v", "f", "f_t", "f_ref", "f_ca", "f_si", "is")] <- ""

  if (styles)
    cc[c("c_s")] <- "_openxlsx_NA_"

  wb$worksheets[[sheet_id]]$sheet_data$cc <- cc

  if (merged_cells)
    wb$worksheets[[sheet_id]]$mergeCells <- character(0)

}



#' little worksheet helper
#' @param wb a workbook
#' @param sheet a worksheet either id or character
#' @export
wb_ws <- function(wb, sheet) {
  wb$ws(sheet)
}


#' little worksheet opener
#' @param wb a workbook
#' @export
wb_open <- function(wb) {
  tmp <- temp_xlsx()
  wb_save(wb, tmp)
  xl_open(tmp)
}
