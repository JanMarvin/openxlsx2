
#' create dataframe from dimensions
#' @param dims Character vector of expected dimension.
#' @param fill If TRUE, fills the dataframe with variables
#' @examples {
#'   dims_to_dataframe("A1:B2")
#' }
#' @export
dims_to_dataframe <- function(dims, fill = FALSE) {

  if(!grepl(":", dims))
    dims <- paste0(dims, ":", dims)

  dimensions <- strsplit(dims, ":")
  rows <- as.numeric(gsub("[[:upper:]]","", dimensions[[1]]))
  cols <- gsub("[[:digit:]]","", dimensions[[1]])

  # cols
  col_min <- col2int(cols[1])
  col_max <- col2int(cols[2])

  cols <- seq(col_min, col_max)
  cols <- int2col(cols)

  # rows
  rows_min <- rows[1]
  rows_max <- rows[2]

  rows <- seq(rows_min, rows_max)

  data <- as.character(NA)
  if (fill) {
    data <- expand.grid(cols, rows, stringsAsFactors = FALSE)
    data <- paste0(data$Var1, data$Var2)
  }

  # matrix as.data.frame
  mm <- matrix(data = data,
               nrow = length(rows),
               ncol = length(cols),
               dimnames = list(rows, cols), byrow = TRUE)

  z <- as.data.frame(mm)
  z
}

#' @param rtyp rtyp object
#' @export
get_row_names <- function(rtyp) {
  sapply(rtyp, function(x) gsub("[[:upper:]]","", x[[1]]))
}

#' function to estimate the column type.
#' 0 = character, 1 = numeric, 2 = date.
#' @param tt dataframe produced by wb_to_df()
#' @export
guess_col_type <- function(tt) {

  # everythings character
  types <- vector("numeric", NCOL(tt))
  names(types) <- names(tt)

  # but some values are numerics
  col_num <- sapply(tt, function(x) all(x[is.na(x) == FALSE] == "n"))
  types[names(col_num[col_num == TRUE])] <- 1

  # or even dates
  col_dte <- sapply(tt[!col_num], function(x) all(x[is.na(x) == FALSE] == "d"))
  types[names(col_dte[col_dte == TRUE])] <- 2

  types
}


#' Create Dataframe from Workbook
#'
#' Simple function to create a dataframe from a workbook. Simple as in simply
#' written down and not optimized etc. The goal was to have something working.
#'
#' @param xlsxFile An xlsx file, Workbook object or URL to xlsx file.
#' @param sheet Either sheet name or index. When missing the first sheet in the workbook is selected.
#' @param colnames If TRUE, the first row of data will be used as column names.
#' @param dims Character string of type "A1:B2" as optional dimentions to be imported.
#' @param detectDates If TRUE, attempt to recognise dates and perform conversion.
#' @param showFormula If TRUE, the underlying Excel formulas are shown.
#' @param convert If TRUE, a conversion to dates and numerics is attempted.
#' @param skipEmptyCols If TRUE, empty columns are skipped.
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
#'   xlsxFile <- system.file("extdata", "inlinestr.xlsx", package = "openxlsx2")
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
#'   wb_to_df(wb3, definedName = "MyRange", colNames = F)
#'
#'   # read definedName from sheet
#'   wb_to_df(wb3, definedName = "MyRange", sheet = 4, colNames = F)
#'
#' @export
wb_to_df <- function(xlsxFile,
                     sheet,
                     startRow = 1,
                     colNames = TRUE,
                     rowNames = FALSE,
                     detectDates = TRUE,
                     skipEmptyCols = FALSE,
                     skipEmptyRows = FALSE,
                     rows = NULL,
                     cols = NULL,
                     na.strings = "#N/A",
                     dims,
                     showFormula = FALSE,
                     convert = TRUE,
                     types,
                     definedName) {

  if (is.character(xlsxFile)){
    # if using it this way, it might be benefitial to only load the sheet we
    # want to read instead of every sheet of the entire xlsx file WHEN we do
    # not even see it
    wb <- loadWorkbook(xlsxFile)
  } else {
    wb <- xlsxFile
  }


  if (!missing(definedName)) {

    dn <- wb$workbook$definedNames
    wo <- unlist(lapply(dn, function(x) getXML1val(x, "definedName")))
    wo <- gsub("\\$", "", wo)
    wo <- unlist(sapply(wo, strsplit, "!"))

    nr <- as.data.frame(
      matrix(wo,
             ncol = 2,
             byrow = TRUE,
             dimnames = list(seq_len(length(dn)),
                             c("sheet", "dims") ))
    )
    nr$name <- sapply(dn, function(x) getXML1attr_one(x, "definedName", "name"))
    nr$local <- sapply(dn, function(x) ifelse(
      getXML1attr_one(x,"definedName", "localSheetId") == "", 0, 1)
    )
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
      } else {
        dims <- sel$dims
      }
    } else {
      stop("no such definedName")
    }
  }

  if (missing(sheet)) sheet <- 1

  if (is.character(sheet))
    sheet <- which(wb$sheet_names %in% sheet)

  # must be available
  if (missing(dims))
    dims <- openxlsx2:::getXML1attr_one(wb$worksheets[[sheet]]$dimension,
                            "dimension",
                            "ref")

  row_attr  <- wb$worksheets[[sheet]]$sheet_data$row_attr
  cc  <- wb$worksheets[[sheet]]$sheet_data$cc
  sst <- attr(wb$sharedStrings, "text")

  rnams <- names(row_attr)


  # internet says: numFmtId > 0 and applyNumberFormat == 1
  # https://stackoverflow.com/a/5251032/12340029
  # standard date 14 - 22 || formatted date 164 - 180 & applyNumberFormat
  sd <- as.data.frame(
    do.call(
      "rbind",
      lapply(
        wb$styles$cellXfs,
        FUN= function(x)
          c(
            as.numeric(openxlsx2:::getXML1attr_one(x, "xf", "numFmtId")),
            as.numeric(openxlsx2:::getXML1attr_one(x, "xf", "applyNumberFormat"))
          )
      )
    )
  )
  names(sd) <- c("numFmtId", "applyNumberFormat")

  sd$id <- seq_len(nrow(sd))-1
  sd$isdate <- 0
  sd$isdate[(sd$numFmtId >= 14 & sd$numFmtId <= 22)] <- 1
  sd$isdate[(sd$numFmtId >= 164 & sd$numFmtId <= 180) &
              sd$applyNumberFormat == 1] <- 1

  xlsx_date_style <- sd$id[sd$isdate == 1]

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

  keep_row <- keep_rows[keep_rows %in% rnams]

  # reduce data to selected cases only
  cc <- cc[cc$row_r %in% keep_row & cc$c_r %in% keep_cols, ]

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
  cc$val[sel] <- cc$t[sel]
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
      sel <- (cc$c_s %in% xlsx_date_style) & !cc$is_string
      cc$val[sel] <- as.character(convertToDate(cc$v[sel]))
      cc$typ[sel]  <- "d"
    }
  }
  # remaining values are numeric?
  sel <- is.na(cc$typ)
  cc$val[sel] <- cc$v[sel]
  cc$typ[sel] <- "n"

  openxlsx2:::long_to_wide(z, tt, cc, dimnames(z))

  # if colNames, then change tt too
  if (colNames) {
    colnames(z)  <- z[1,]
    colnames(tt) <- z[1,]

    z  <- z[-1, , drop = FALSE]
    tt <- tt[-1, , drop = FALSE]
  }

  if (rowNames) {
    rownames(z)  <- z[,1]
    rownames(tt) <- z[,1]

    z  <- z[ ,-1]
    tt <- tt[ , -1]
  }

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

  # # faster guess_col_type alternative? to avoid tt
  # types <- ftable(cc$row_r ~ cc$c_r ~ cc$typ)

  if (missing(types))
    types <- guess_col_type(tt)

  # could make it optional or explicit
  if (convert) {
    sel <- !is.na(names(types))

    if (any(sel)) {
      nums <- names( which(types[sel] == 1) )
      dtes <- names( which(types[sel] == 2) )

      z[nums] <- lapply(z[nums], as.numeric)
      z[dtes] <- lapply(z[dtes], as.Date)
    } else {
      warning("could not convert. All missing in row used for variable names")
    }
  }

  attr(z, "tt") <- tt
  attr(z, "types") <- types
  # attr(z, "sd") <- sd
  if (!missing(definedName)) attr(z, "dn") <- nr
  z
}


#' dummy function to write data
#' @param wb workbook
#' @param sheet sheet
#' @param data data to export
#' @param colNames include colnames?
#' @param rowNames include rownames?
#' @param startRow row to place it
#' @param startCol col to place it
#'
#' @examples
#' # create a workbook and add some sheets
#' wb <- createWorkbook()
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
#' saveWorkbook(wb, file = "/tmp/test.xlsx", overwrite = TRUE)
#'
#' @export
writeData2 <-function(wb, sheet, data,
                      colNames = TRUE, rowNames = FALSE,
                      startRow = 1, startCol = 1) {


  is_data_frame <- FALSE
  data_class <- sapply(data, class)

  # convert factor to character
  if (any(data_class == "factor")) {
    fcts <- which(data_class == "factor")
    data[fcts] <- lapply(data[fcts],as.character)
  }

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
      data_class <- c("character", data_class)
    }
  }


  sheetno <- wb$validateSheet(sheet)
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

  if (class(wb$worksheets[[sheetno]]$sheet_data$cc) == "uninitializedField") {


    sheet_data <- list()
    wb$worksheets[[sheetno]]$dimension <- paste0("<dimension ref=\"", dims, "\"/>")

    # rtyp character vector per row
    # list(c("A1, ..., "k1"), ...,  c("An", ..., "kn"))
    rtyp <- dims_to_dataframe(dims, fill = TRUE)

    rows_attr <- cols_attr <- cc_tmp <- vector("list", data_nrow)

    cols_attr <- lapply(seq_len(data_nrow),
                        function(x) list(collapsed="false",
                                         customWidth="true",
                                         hidden="false",
                                         outlineLevel="0",
                                         max="121",
                                         min="1",
                                         style="0",
                                         width="9.14"))

    wb$worksheets[[sheetno]]$cols_attr <- openxlsx2:::list_to_attr(cols_attr, "col")



    rows_attr <- lapply(startRow:endRow,
                        function(x) list("r" = as.character(x),
                                         "spans" = paste0("1:", data_ncol),
                                         "x14ac:dyDescent"="0.25"))
    names(rows_attr) <- rownames(rtyp)

    wb$worksheets[[sheetno]]$sheet_data$row_attr <- rows_attr

    # original cc dataframe
    nams <- c("row_r", "c_r", "c_s", "c_t", "v", "f", "f_t", "t")
    cc <- as.data.frame(
      matrix(data = "NA",
             nrow = nrow(data) * ncol(data),
             ncol = length(nams))
    )
    names(cc) <- nams


    numcell <- function(x,y){
      c(val = as.character(x),
        typ = "n",
        r = y,
        c_t = "v")
    }

    chrcell <- function(x,y){
      c(val = x,
        typ = "c",
        r = y,
        c_t = "str")
    }

    cell <- function(x, y, data_class) {
      z <- NULL
      if (data_class == "numeric")
        z <- numcell(x,y)
      if (data_class %in% c("character", "factor"))
        z <- chrcell(x,y)

      z
    }


    for (i in seq_len(nrow(data))) {

      col <- data.frame(matrix(data = "NA", nrow = ncol(data), ncol = 4))
      names(col) <- c("val", "typ", "r", "c_t")
      for (j in seq_along(data)) {
        dc <- ifelse(colNames && i == 1, "character", data_class[j])
        col[j,] <- cell(data[i, j], rtyp[i, j], dc)
      }
      col$row_r <- colnames(rtyp)
      col$c_r   <- rownames(rtyp)[i]

      cc_tmp[[i]] <- col
    }

    cc_tmp <- do.call("rbind", cc_tmp)

    cc <- merge(cc[!names(cc) %in% names(cc_tmp)],
                cc_tmp,
                by = "row.names")

    cc <- cc[order(as.numeric(cc$Row.names)),-1]
    cc <- cc[c(nams, "val", "typ", "r")]

    wb$worksheets[[sheetno]]$sheet_data$cc <- cc

    # update a few styles informations
    wb$styles$numFmts <- character(0)
    wb$styles$cellXfs <- "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
    wb$styles$dxfs <- character(0)

    # now the cc dataframe is similar to an imported cc dataframe. to save the
    # file, we now need to split it up and store it in a list of lists.

  } else {
    # update cell(s)
    wb <- update_cell(x = data, wb, sheetno, dims, data_class, colNames)
  }

  cc_rows <- unique(cc$c_r)
  cc_out <- vector("list", length = length(cc_rows))
  names(cc_out) <- cc_rows

  for (cc_r in cc_rows) {
    tmp <- cc[cc$c_r == cc_r, c("r", "val", "c_t")]
    nams <- cc[cc$c_r == cc_r, c("row_r")]
    ltmp <- vector("list", nrow(tmp))
    names(ltmp) <- nams

    for (i in seq_len(nrow(tmp))) {
      ltmp[[i]] <- as.list(tmp[i,])
    }

    cc_out[[cc_r]] <- ltmp
  }

  wb$worksheets[[sheetno]]$sheet_data$cc_out <- cc_out


  wb
}

