
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
    cols <- NA
    rows <- NA
  } else {
    dimensions <- strsplit(dims, ":")[[1]]

    rows <- as.numeric(gsub("[[:upper:]]","", dimensions))
    rows <- seq.int(rows[1], rows[2])

    # TODO seq.wb_columns?  make a wb_cols vector?
    cols <- gsub("[[:digit:]]","", dimensions)
    cols <- int2col(seq.int(col2int(cols[1]), col2int(cols[2])))
  }

  if (fill) {
    data <- expand.grid(cols, rows, stringsAsFactors = FALSE)
    data <- paste0(data[[1L]], data[[2L]])
  } else {
    data <- NA_character_
  }

  # matrix as.data.frame
  as.data.frame(matrix(
    data     = data,
    nrow     = max(length(rows), 1L),
    ncol     = max(length(cols), 1L),
    dimnames = list(rows, cols),
    byrow    = TRUE
  ))
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
  col_num <- sapply(tt, function(x) all(x[is.na(x) == FALSE] == "n"))
  types[names(col_num[col_num == TRUE])] <- 1

  # or even date
  col_dte <- sapply(tt[!col_num], function(x) all(x[is.na(x) == FALSE] == "d"))
  types[names(col_dte[col_dte == TRUE])] <- 2

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
    "d", "dd", "ddd", "dddd",
    "h", "hh", "m", "mm", "s", "ss",
    "AM", "PM", "A", "P"
  )
  date_or_fmt <- paste0(date_fmts, collapse = "|")
  maybe_dates <- grepl(pattern = date_or_fmt, x = numFmt_df$formatCode)

  z <- numFmt_df$numFmtId[maybe_dates & !maybe_num]
  if (length(z)==0) z <- NULL
  z
}

#' check if style is date. internal function
#'
#' @param cellXfs cellXfs xml nodes
#' @param numfmt_date custom numFmtId dates
style_is_date <- function(cellXfs, numfmt_date) {

  # numfmt_date: some basic date formats and custom formats
  date_numfmts <- as.character(14:22)
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
#' @param dims Character string of type "A1:B2" as optional dimentions to be imported.
#' @param detectDates If TRUE, attempt to recognise dates and perform conversion.
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
#' @examples
#'
#'   ###########################################################################
#'   # numerics, dates, missings, bool and string
#'   xlsxFile <- system.file("extdata", "readTest.xlsx", package = "openxlsx2")
#'   wb1 <- loadWorkbook(xlsxFile)
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
#'   # read.xlsx(wb1)
#'
#'   ###########################################################################
#'   # inlinestr
#'   xlsxFile <- system.file("extdata", "inline_str.xlsx", package = "openxlsx2")
#'   wb2 <- loadWorkbook(xlsxFile)
#'
#'   # read dataset with inlinestr
#'   wb_to_df(wb2)
#'   # read.xlsx(wb2)
#'
#'   ###########################################################################
#'   # definedName // namedRegion
#'   xlsxFile <- system.file("extdata", "namedRegions3.xlsx", package = "openxlsx2")
#'   wb3 <- loadWorkbook(xlsxFile)
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
  startRow      = 1,
  colNames      = TRUE,
  rowNames      = FALSE,
  detectDates   = TRUE,
  skipEmptyCols = FALSE,
  skipEmptyRows = FALSE,
  rows          = NULL,
  cols          = NULL,
  na.strings    = "#N/A",
  dims,
  showFormula   = FALSE,
  convert       = TRUE,
  types,
  definedName
) {

  .mc <- match.call()

  if (is.character(xlsxFile)) {
    # TODO this should instead check for the Workbook class?  Maybe also check
    # if the file exists?

    # if using it this way, it might be benefitial to only load the sheet we
    # want to read instead of every sheet of the entire xlsx file WHEN we do
    # not even see it
    wb <- loadWorkbook(xlsxFile)
  } else {
    wb <- xlsxFile
  }


  if (!missing(definedName)) {

    dn <- wb$workbook$definedNames
    wo <- xml_value(dn, "definedName")
    # remove dollar sign: $A$1:$B$2
    wo <- gsub("\\$", "", wo)
    wo <- unlist(sapply(wo, strsplit, "!"))
    # remove ' from 'Sheet 1'
    wo <- gsub("'", "", wo)
    wo <- unlist(sapply(wo, strsplit, "!"))

    nr <- as.data.frame(
      matrix(wo,
        ncol = 2,
        byrow = TRUE,
        dimnames = list(seq_len(length(dn)),
          c("sheet", "dims") ))
    )
    dn_attr <- rbindlist(xml_attr(dn, "definedName"))

    nr$name <- dn_attr$name
    if (!is.null(dn_attr$localSheetId)) {
      nr$local <- ifelse(dn_attr$localSheetId == "", 0, 1)
    } else {
      nr$local <- 0
    }
    nr$sheet <- which(wb$sheet_names %in% nr$sheet)

    nr <- nr[order(nr$local, nr$name, nr$sheet),]

    if (definedName %in% nr$name & missing(sheet)) {
      sel   <- nr[nr$name == definedName, ][1,]
      sheet <- sel$sheet
      dims  <- sel$dims
    } else if (definedName %in% nr$name) {
      sel <- nr[nr$name == definedName & nr$sheet == sheet, ]
      if (NROW(sel) == 0) {
        stop("no such definedName on selected sheet")
      }
      dims <- sel$dims
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
  if (missing(definedName) & missing(dims)) {

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

  numfmt_date <- numfmt_is_date(wb$styles$numFmts)
  xlsx_date_style <- style_is_date(wb$styles$cellXfs, numfmt_date)

  # create temporary data frame. hard copy required
  z  <- dims_to_dataframe(dims)
  tt <- dims_to_dataframe(dims)


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
  if (!is.na(na.strings) | !missing(na.strings)) {
    sel <- cc$val %in% na.strings
    cc$val[sel] <- NA
    cc$typ[sel] <- NA
  }
  # dates
  if (!is.null(cc$c_s)) {
    # if a cell is t="s" the content is a sst and not da date
    cc$is_string <- FALSE
    if (!is.null(cc$c_t))
      cc$is_string <- cc$c_t %in% c("s", "str", "b", "inlineStr")

    if (detectDates) {
      sel <- (cc$c_s %in% xlsx_date_style) & !cc$is_string & cc$v != "_openxlsx_NA_"
      cc$val[sel] <- as.character(convertToDate(cc$v[sel]))
      cc$typ[sel]  <- "d"
    }
  }
  # remaining values are numeric?
  sel <- is.na(cc$typ)
  cc$val[sel] <- cc$v[sel]
  cc$typ[sel] <- "n"

  # prepare to create output object z
  zz <- cc[c("val", "typ")]
  # we need to create the correct col and row position as integer starting at 0.
  zz$cols <- match(cc$c_r, colnames(z)) - 1
  zz$rows <- match(cc$row_r, rownames(z)) - 1

  zz <- zz[with(zz, ordered(order(cols, rows))),]
  zz <- zz[zz$val != "_openxlsx_NA_",]
  long_to_wide(z, tt, zz)

  # prepare colnames object
  xlsx_cols_names <- colnames(z)
  names(xlsx_cols_names) <- xlsx_cols_names

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
      # convert "#NUM!" to "NaN" -- then converts to NaN
      # maybe consider this an option to instead return NA?
      z[nums] <- lapply(z[nums], function(i) as.numeric(replace(i, i == "#NUM!", "NaN")))
      z[dtes] <- lapply(z[dtes], as.Date)
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

    z  <- z[!empty, ]
    tt <- tt[!empty,]
  }

  if (skipEmptyCols) {

    empty <- sapply(z, function(x) all(is.na(x)))

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


#' Replace data cell(s)
#'
#' Minimal invasive update of cell(s) inside of imported workbooks.
#'
#' @param x value you want to insert
#' @param wb the workbook you want to update
#' @param sheet the sheet you want to update
#' @param cell the cell you want to update in Excel conotation e.g. "A1"
#' @param data_class optional data class object
#' @param colNames if TRUE colNames are passed down
#' @param removeCellStyle keep the cell style?
#'
#' @examples
#'    xlsxFile <- system.file("extdata", "update_test.xlsx", package = "openxlsx2")
#'    wb <- loadWorkbook(xlsxFile)
#'
#'    # update Cells D4:D6 with 1:3
#'    wb <- update_cell(x = c(1:3), wb = wb, sheet = "Sheet1", cell = "D4:D6")
#'
#'    # update Cells B3:D3 (names())
#'    wb <- update_cell(x = c("x", "y", "z"), wb = wb, sheet = "Sheet1", cell = "B3:D3")
#'
#'    # update D4 again (single value this time)
#'    wb <- update_cell(x = 7, wb = wb, sheet = "Sheet1", cell = "D4")
#'
#'    # add new column on the left of the existing workbook
#'    wb <- update_cell(x = 7, wb = wb, sheet = "Sheet1", cell = "A4")
#'
#'    # add new row on the end of the existing workbook
#'    wb <- update_cell(x = 7, wb = wb, sheet = "Sheet1", cell = "A9")
#'    wb_to_df(wb)
#'
#' @export
update_cell <- function(x, wb, sheet, cell, data_class,
                        colNames = FALSE, removeCellStyle = FALSE) {

  dimensions <- unlist(strsplit(cell, ":"))
  rows <- gsub("[[:upper:]]","", dimensions)
  cols <- gsub("[[:digit:]]","", dimensions)

  if (length(dimensions) == 2) {
    # cols
    cols <- col2int(cols)
    cols <- seq(cols[1], cols[2])
    cols <- int2col(cols)

    rows <- as.character(seq(rows[1], rows[2]))
  }


  if(is.character(sheet)) {
    sheet_id <- which(sheet == wb$sheet_names)
  } else {
    sheet_id <- sheet
  }

  if (missing(data_class))
    data_class <- sapply(x, class)

  # if(identical(sheet_id, integer()))
  #   stop("sheet not in workbook")

  # 1) pull sheet to modify from workbook; 2) modify it; 3) push it back
  cc  <- wb$worksheets[[sheet_id]]$sheet_data$cc
  row_attr <- wb$worksheets[[sheet_id]]$sheet_data$row_attr

  # workbooks contain only entries for values currently present.
  # if A1 is filled, B1 is not filled and C1 is filled the sheet will only
  # contain fields A1 and C1.
  cc$row <- paste0(cc$c_r, cc$row_r)
  cells_in_wb <- cc$rw
  rows_in_wb <- row_attr$r


  # check if there are rows not available
  if (!all(rows %in% rows_in_wb)) {
    # message("row(s) not in workbook")

    # add row to name vector, extend the entire thing
    total_rows <- as.character(sort(unique(as.numeric(c(rows, rows_in_wb)))))

    # new row_attr
    row_attr_new <- empty_row_attr(n = length(total_rows))
    row_attr_new$r <- total_rows

    row_attr_new <- merge(row_attr_new[c("r")], row_attr, all.x = TRUE)
    row_attr_new[is.na(row_attr_new)] <- ""

    wb$worksheets[[sheet_id]]$sheet_data$row_attr <- row_attr_new
    # provide output
    rows_in_wb <- total_rows

  }

  if (!any(cols %in% cells_in_wb)) {
    # all rows are availabe in the dataframe
    for (row in rows){

      # collect all wanted cols and order for excel
      total_cols <- unique(c(cc$c_r[cc$row_r == row], cols))
      total_cols <- int2col(sort(col2int(total_cols)))

      # create candidate
      cc_row_new <- data.frame(matrix("_openxlsx_NA_", nrow = length(total_cols), ncol = 2))
      names(cc_row_new) <- names(cc)[1:2]
      cc_row_new$row_r <- row
      cc_row_new$c_r <- total_cols

      # extract row (easier or maybe only way to change order?)
      cc_row <- cc[cc$row_r == row, ]
      # remove row from cc
      if (nrow(cc_row)>0) cc <- cc[-which(rownames(cc) %in% rownames(cc_row)),]
      # new row
      cc_row <- merge(x = cc_row_new, y = cc_row, all.x = TRUE)

      # assign to cc
      cc <- rbind(cc, cc_row)

    }
  }

  # update dimensions
  cc$row <- paste0(cc$c_r, cc$row_r)
  cells_in_wb <- cc$rw

  all_rows <- unique(cc$row_r)
  all_cols <- unique(cc$c_r)

  min_cell <- trimws(paste0(min(all_cols), min(all_rows)))
  max_cell <- trimws(paste0(max(all_cols), max(all_rows)))

  # i know, i know, i'm lazy
  wb$worksheets[[sheet_id]]$dimension <- paste0("<dimension ref=\"", min_cell, ":", max_cell, "\"/>")

  # if (any(rows %in% rows_in_wb) )
  # message("found cell(s) to update")

  if (all(rows %in% rows_in_wb)) {
    # message("cell(s) to update already in workbook. updating ...")

    i <- 0; n <- 0
    for (row in rows) {

      n <- n+1
      m <- 0

      for (col in cols) {
        i <- i+1
        m <- m+1

        # check if is data frame or matrix
        value <- ifelse(is.null(dim(x)), x[i], x[n, m])

        sel <- cc$row_r == row & cc$c_r == col
        c_s <- NULL
        if (removeCellStyle) c_s <- "c_s"

        cc[sel, c(c_s, "c_t", "v", "f", "f_t", "f_ref", "f_si", "is")] <- "_openxlsx_NA_"


        # for now convert all R-characters to inlineStr (e.g. names() of a dataframe)
        if (data_class[m] %in% c("character", "factor") | (colNames == TRUE & n == 1)) {
          cc[sel, "c_t"] <- "inlineStr"
          cc[sel, "is"]   <- paste0("<is><t>", as.character(value), "</t></is>")
        } else {
          cc[sel, "v"]   <- as.character(value)
        }

      }
    }

  }

  # order cc
  cc$ordered_rows <- col2int(cc$c_r)
  cc$ordered_cols <- as.numeric(cc$row_r)

  cc <- cc[order(cc$ordered_rows, cc$ordered_cols),]
  cc$ordered_cols <- NULL
  cc$ordered_rows <- NULL

  # push everything back to workbook
  wb$worksheets[[sheet_id]]$sheet_data$cc  <- cc

  wb
}


numfmt_class <- function(data) {
  dc <- as.data.frame(Map(class, data))

  # check all columns for the required types
  if (nrow(dc) >= 1) {

    # returns logical vector of length ncol(data)
    is_class <- function(dc, cl) apply(dc, 2, function(x)(any(x == cl)))

    # check the class. all have the basic R classes, some have additional
    # openxml builtin formats
    is_fact <- is_class(dc, "factor")
    is_date <- is_class(dc, "Date")
    is_posi <- is_class(dc, "POSIXct")
    is_logi <- is_class(dc, "logical")
    is_char <- is_class(dc, "character")
    is_inte <- is_class(dc, "integer")
    is_numb <- is_class(dc, "numeric")
    is_hype <- is_class(dc, "hyperlink")
    is_curr <- is_class(dc, "currency")
    is_acco <- is_class(dc, "accounting")
    is_perc <- is_class(dc, "percentage")
    is_scie <- is_class(dc, "scientific")
    is_comm <- is_class(dc, "comma")
    # formulas get no special output class here. They are characters, but are
    # assigned to <f ...>
    is_form <- is_class(dc, "formula")

    # prepare output
    out_class <- dc[1, , drop = FALSE]
    out_class[1,] <- "character"

    # assign the final output
    out_class[is_fact] <- "factor"
    out_class[is_date] <- "date"
    out_class[is_posi] <- "posix"
    out_class[is_logi] <- "logical"
    out_class[is_char] <- "character" # superfluous
    out_class[is_inte] <- "integer"
    out_class[is_numb] <- "numeric"
    out_class[is_hype] <- "hyperlink"
    out_class[is_curr] <- "currency"
    out_class[is_acco] <- "accounting"
    out_class[is_perc] <- "percentage"
    out_class[is_scie] <- "scientific"
    out_class[is_comm] <- "comma"
    out_class[is_form] <- "formula"
  }

  out_class
}


short_dtecell <- function(x,y, short_dtefmt){
  c(v   = as.character(x),
    typ = "n",
    r   = y,
    c_s = as.character(short_dtefmt),
    c_t = "_openxlsx_NA_",
    is  =  "_openxlsx_NA_",
    f   =  "_openxlsx_NA_")
}

long_dtecell <- function(x,y, long_dtefmt){
  c(v   = as.character(x),
    typ = "n",
    r   = y,
    c_s = as.character(long_dtefmt),
    c_t = "_openxlsx_NA_",
    is  =  "_openxlsx_NA_",
    f   =  "_openxlsx_NA_")
}

numcell <- function(x,y, numfmt = "_openxlsx_NA_"){
  c(v   = as.character(x),
    typ = "n",
    r   = y,
    c_s = as.character(numfmt),
    c_t = "_openxlsx_NA_",
    is  =  "_openxlsx_NA_",
    f   =  "_openxlsx_NA_")
}

boolcell <- function(x,y){
  c(v   = as.character(as.integer(as.logical(x))),
    typ = "n",
    r   = y,
    c_s = "_openxlsx_NA_",
    c_t = "b",
    is  =  "_openxlsx_NA_",
    f   =  "_openxlsx_NA_")
}

chrcell <- function(x,y){
  c(v   = "_openxlsx_NA_",
    typ = "c",
    r   = y,
    c_s = "_openxlsx_NA_",
    c_t = "inlineStr",
    is  = paste0("<is><t>", as.character(x), "</t></is>"),
    f   =  "_openxlsx_NA_")
}

fmlcell <- function(x,y){
  c(v   = "_openxlsx_NA_",
    typ = "c",
    r   = y,
    c_s = "_openxlsx_NA_",
    c_t = "str",
    is  =  "_openxlsx_NA_",
    f   = as.character(x))
}

cell <- function(x, y, data_class, special_fmts) {
  z <- NULL
  if (data_class %in% c("date"))
    z <- short_dtecell(x,y, special_fmts$short_date)
  if (data_class %in% c("posix"))
    z <- long_dtecell(x,y, special_fmts$long_date)
  if (data_class %in% c("logical"))
    z <- boolcell(x,y)
  if (data_class %in% c("numeric", "integer"))
    z <- numcell(x,y)
  if (data_class %in% c("character", "factor", "hyperlink", "currency"))
    z <- chrcell(x,y)
  if (data_class %in% c("formula"))
    z <- fmlcell(x,y)
  if (data_class %in% c("accounting"))
    z <- numcell(x,y, special_fmts$accounting)
  if (data_class %in% c("percentage"))
    z <- numcell(x,y, special_fmts$percentage)
  if (data_class %in% c("scientific"))
    z <- numcell(x,y, special_fmts$scientific)
  if (data_class %in% c("comma"))
    z <- numcell(x,y, special_fmts$comma)

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


#' dummy function to write data
#' @param wb workbook
#' @param sheet sheet
#' @param data data to export
#' @param colNames include colnames?
#' @param rowNames include rownames?
#' @param startRow row to place it
#' @param startCol col to place it
#' @param removeCellStyle keep the cell style?
#'
#' @examples
#' # create a workbook and add some sheets
#' wb <- wb_workbook()
#'
#' addWorksheet(wb, "sheet1")
#' writeData2(wb, "sheet1", mtcars, colNames = TRUE, rowNames = TRUE)
#'
#' addWorksheet(wb, "sheet2")
#' writeData2(wb, "sheet2", cars, colNames = FALSE)
#'
#' addWorksheet(wb, "sheet3")
#' writeData2(wb, "sheet3", letters)
#'
#' addWorksheet(wb, "sheet4")
#' writeData2(wb, "sheet4", as.data.frame(Titanic), startRow = 2, startCol = 2)
#'
#' \dontrun{
#' file <- tempfile(fileext = ".xlsx")
#' wb_save(wb, path = file, overwrite = TRUE)
#' file.remove(file)
#' }
#'
#' @export
writeData2 <-function(wb, sheet, data,
                      colNames = TRUE, rowNames = FALSE,
                      startRow = 1, startCol = 1,
                      removeCellStyle = FALSE) {


  is_data_frame <- FALSE
  #### prepare the correct data formats for openxml
  data_class <- as.data.frame(Map(class, data))
  dc <- numfmt_class(data)

  # convert factor to character
  if (any(data_class == "factor")) {
    is_factor <- apply(data_class, 2, function(x)(any(x == "factor")))
    fcts <- names(data_class[is_factor])
    data[fcts] <- lapply(data[fcts], as.character)
  }

  has_date1904 <- grepl('date1904="1"|date1904="true"',
                        stri_join(unlist(wb$workbook), collapse = ""),
                        ignore.case = TRUE)

  # TODO need to tell excel that we have a date, apply some kind of numFmt
  data <- convertToExcelDate(df = data, date1904 = has_date1904)

  if (class(data) == "data.frame" | class(data) == "matrix") {
    is_data_frame <- TRUE

    # add colnames
    if (colNames == TRUE)
      data <- rbind(colnames(data), data)

    # add rownames
    if (rowNames == TRUE) {
      nam <- names(data)
      data <- cbind(rownames(data), data)
      names(data) <- c("", nam)
      data_class <- cbind(c("_rowNames_" = "character"), data_class)
      names(data_class) <- names(data)
      dc <- cbind(c("_rowNames_" = "character"), dc)
      names(dc) <- names(data)

    }
  }


  sheetno <- wb_validate_sheet(wb, sheet)
  # message("sheet no: ", sheetno)

  # create a data frame
  if (!is_data_frame)
    data <- as.data.frame(t(data))

  data_nrow <- NROW(data)
  data_ncol <- NCOL(data)

  endRow <- (startRow -1) + data_nrow
  endCol <- (startCol -1) + data_ncol


  dims <- paste0(int2col(startCol), startRow,
    ":",
    int2col(endCol), endRow)

  if (is.null(wb$worksheets[[sheetno]]$sheet_data$cc)) {


    sheet_data <- list()
    wb$worksheets[[sheetno]]$dimension <- paste0("<dimension ref=\"", dims, "\"/>")

    # rtyp character vector per row
    # list(c("A1, ..., "k1"), ...,  c("An", ..., "kn"))
    rtyp <- dims_to_dataframe(dims, fill = TRUE)

    rows_attr <- cols_attr <- cc_tmp <- vector("list", data_nrow)

    # # create <cols ...>
    # cols_attr <- empty_cols_attr(n = data_nrow)
    # cols_attr$collapsed <- ""
    # cols_attr$customWidth <- ""
    # cols_attr$hidden <- ""
    # cols_attr$outlineLevel <- ""
    # cols_attr$max <- ""
    # cols_attr$min <- ""
    # cols_attr$style <- ""
    # cols_attr$width <- ""
    #
    # wb$worksheets[[sheetno]]$cols_attr <- df_to_xml("col", cols_attr)

    # create <rows ...>
    want_rows <- startRow:endRow
    rows_attr <- empty_row_attr(n = length(want_rows))
    rows_attr$r <- rownames(rtyp)

    wb$worksheets[[sheetno]]$sheet_data$row_attr <- rows_attr

    # original cc dataframe
    # TODO should be empty_sheet_data(n = nrow(data) * ncol(data))
    nams <- c("row_r", "c_r", "c_s", "c_t", "v", "f", "f_t", "f_ref", "f_si", "is")
    cc <- as.data.frame(
      matrix(data = "_openxlsx_NA_",
             nrow = nrow(data) * ncol(data),
             ncol = length(nams))
    )
    names(cc) <- nams

    # update a few styles informations
    wb$styles$numFmts <- character()

    ## create a cell style format for specific types at the end of the existing
    # styles. gets the reference an passes it on.
    short_date_fmt <- long_date_fmt <- accounting_fmt <- percentage_fmt <-
      comma_fmt <- scientific_fmt <- NULL
    if (any(dc == "date")) short_date_fmt <- write_xf(nmfmt_df(14))
    if (any(dc == "posix")) long_date_fmt  <- write_xf(nmfmt_df(22))
    if (any(dc == "accounting")) accounting_fmt <- write_xf(nmfmt_df(4))
    if (any(dc == "percentage")) percentage_fmt <- write_xf(nmfmt_df(10))
    if (any(dc == "scientific")) scientific_fmt <- write_xf(nmfmt_df(48))
    if (any(dc == "comma")) comma_fmt      <- write_xf(nmfmt_df(3))

    wb$styles$cellXfs <- c(wb$styles$cellXfs,
                           short_date_fmt, long_date_fmt,
                           accounting_fmt,
                           percentage_fmt,
                           comma_fmt,
                           scientific_fmt)

    # get the last style id
    style_id <- function(cellfmt) {
      style_id <- which(wb$styles$cellXfs == cellfmt)
      # get the last id, the one we have just written. return them as
      # 0-index and minimum value of 0
      ifelse(length(style_id), max(max(style_id)-1,0) , 0)
    }

    special_fmts <- data.frame(
      short_date = style_id(short_date_fmt),
      long_date = style_id(long_date_fmt),
      accounting = style_id(accounting_fmt),
      percentage = style_id(percentage_fmt),
      scientific = style_id(scientific_fmt),
      comma = style_id(comma_fmt)
    )

    for (i in seq_len(NROW(data))) {

      col_nams <- c("v", "typ", "r", "c_s", "c_t", "is", "f")
      col <- data.frame(matrix(data = "_openxlsx_NA_", nrow = ncol(data), ncol = length(col_nams)))
      names(col) <- col_nams
      for (j in seq_along(data)) {
        val <- data[i, j]
        # if first row and colnames write character else whatever typ is found
        dc_j <- ifelse((colNames & (i == 1)), "character", dc[1, j])
        col[j,] <- cell(val, rtyp[i, j], dc_j, special_fmts)
      }
      col$row_r <- rownames(rtyp)[i]
      col$c_r   <- colnames(rtyp)

      cc_tmp[[i]] <- col
    }

    cc_tmp <- do.call("rbind", cc_tmp)

    cc <- merge(cc[!names(cc) %in% names(cc_tmp)],
      cc_tmp,
      by = "row.names")

    cc <- cc[order(as.numeric(cc$Row.names)),-1]
    cc <- cc[c(nams, "typ", "r")]

    wb$worksheets[[sheetno]]$sheet_data$cc <- cc


    wb$styles$dxfs <- character()

    # now the cc dataframe is similar to an imported cc dataframe. to save the
    # file, we now need to split it up and store it in a list of lists.

  } else {
    # update cell(s)
    wb <- update_cell(x = data, wb, sheetno, dims, data_class, colNames, removeCellStyle)
  }


  wb
}


#' clean sheet (remove all values)
#'
#' @param wb workbook
#' @param sheet sheet to clean
#' @param numbers remove all numbers
#' @param characters remove all characters
#' @param styles remove all styles
#' @param merged_cells remove all merged_cells
#' @export
cleanSheet <- function(wb, sheet, numbers = TRUE, characters = TRUE, styles = TRUE, merged_cells = TRUE) {

  sheet_id <- wb_validate_sheet(wb, sheet)

  cc <- wb$worksheets[[sheet_id]]$sheet_data$cc

  if (numbers)
    cc[cc$c_t %in% c("n", "_openxlsx_NA_"), # imported values might be _NA_
       c("c_t", "v", "f", "f_t", "f_ref", "f_si", "is")] <- "_openxlsx_NA_"

  if (characters)
    cc[cc$c_t %in% c("inlineStr", "s"),
       c("v", "f", "f_t", "f_ref", "f_si", "is")] <- ""

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
  openXL(tmp)
}
