#' Create dataframe from dimensions
#'
#' @param dims Character vector of expected dimension.
#' @param fill If `TRUE`, fills the dataframe with variables
#' @param empty_rm Logical if empty columns and rows should be included
#' @examples
#' dims_to_dataframe("A1:B2")
#' @keywords internal
#' @export
dims_to_dataframe <- function(dims, fill = FALSE, empty_rm = FALSE) {

  # in R 4.4.0 grepl(",", data.frame(x = paste0("K",))) == TRUE
  if (inherits(dims, "data.frame"))
    dims <- unlist(dims)

  has_dim_sep <- FALSE
  if (any(grepl(";", dims))) {
    dims <- unlist(strsplit(dims, ";"))
    has_dim_sep <- TRUE
  }
  if (any(grepl(",", dims))) {
    dims <- unlist(strsplit(dims, ","))
    has_dim_sep <- TRUE
  }

  # this is only required, if dims is not equal sized
  rows_out  <- NULL
  cols_out  <- NULL
  filled    <- NULL
  full_cols <- NULL

  # condition 1) contains dims separator, but all dims are of
  # equal size: "A1:A5,B1:B5"
  # condition 2) either "A1:B5" or separator, but unequal size or "A1:A2,A4:A6,B1:B5"
  if (has_dim_sep && get_dims(dims, check = TRUE)) {

    full_rows <- get_dims(dims, rows = TRUE)
    full_cols <- sort(get_dims(dims, cols = TRUE))

    rows_out  <- unlist(full_rows)
    rows_out  <- seq.int(rows_out[1], rows_out[2])
    cols_out  <- int2col(full_cols)
    full_cols <- full_cols - min(full_cols) # is always a zero offset

  } else {

    for (dim in dims) {

      if (!grepl(":", dim)) {
        dim <- paste0(dim, ":", dim)
      }

      if (length(dims) > 1)
        filled <- c(filled, needed_cells(dim))

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
  }

  if (has_dim_sep) {
    if (empty_rm) {
      cols_out  <- int2col(sort(col2int(cols_out)))
      rows_out  <- sort(rows_out)
      # with empty_rm the dataframe will contain only needed columns
      if (!is.null(full_cols)) full_cols <- seq_along(cols_out) - 1L
    } else {
      # somehow we have to make sure that all columns are covered
      col_ints <- col2int(cols_out)
      cols_out <- int2col(seq.int(from = min(col_ints), to = max(col_ints)))

      row_ints <- rows_out
      rows_out <- seq.int(from = min(row_ints), to = max(row_ints))
    }
  }

  dims_to_df(
    rows   = rows_out,
    cols   = cols_out,
    filled = filled,
    fill   = fill,
    fcols  = full_cols
  )
}

#' Create dimensions from dataframe
#'
#' Use [wb_dims()]
#' @param df dataframe with spreadsheet columns and rows
#' @param dim_break split the dims?
#' @examples
#'  df <- dims_to_dataframe("A1:D5;F1:F6;D8", fill = TRUE)
#'  dataframe_to_dims(df)
#' @keywords internal
#' @export
dataframe_to_dims <- function(df, dim_break = FALSE) {

  if (dim_break) {

    dims <- dims_to_dataframe(dataframe_to_dims(df, dim_break = FALSE), fill = TRUE)

    mm <- as.matrix(df)
    mm[mm != "" | is.na(mm)] <- 1
    mm[mm == ""] <- 0

    matrix <- matrix(as.numeric(mm), nrow(mm), ncol(mm))
    dimnames(matrix) <- list(rownames(mm), colnames(mm))

    # remove columns and rows not in df
    dims <- dims[, colnames(dims) %in% colnames(matrix)]
    dims <- dims[rownames(dims) %in% rownames(matrix), ]

    out <- dims[matrix == 1]

    return(paste0(out, collapse = ","))

  } else {

    rows <- as.integer(rownames(df))
    cols <- colnames(df)

    if (all(diff(col2int(cols)) == 1L) && all(diff(rows) == 1L)) {
      tmp <- paste0(
        cols[[1]][[1]], rows[[1]][[1]],
        ":",
        rev(cols)[[1]][[1]], rev(rows)[[1]][[1]]
      )
    } else {
      tmp <- con_dims(col2int(cols), rows)
    }

    return(tmp)

  }
}

#' function to estimate the column type.
#' 0 = character, 1 = numeric, 2 = date.
#' @param tt dataframe produced by wb_to_df()
#'
#' @noRd
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
  col_log <- vapply(tt[!col_num], function(x) all(x == "b", na.rm = TRUE), NA)
  types[names(col_log[col_log])] <- 4

  # or even hms
  col_dte <- vapply(tt[!col_num], function(x) all(x == "h", na.rm = TRUE), NA)
  types[names(col_dte[col_dte])] <- 5

  # or formula
  col_fml <- vapply(tt[!col_num], function(x) all(x == "f", na.rm = TRUE), NA)
  types[names(col_fml[col_fml])] <- 6

  types
}

#' check if numFmt is date. internal function
#' @param numFmt numFmt xml nodes
#' @noRd
numfmt_is_date <- function(numFmt) {

  # if numFmt is character(0)
  if (length(numFmt) == 0) return(NULL)

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
#' @noRd
numfmt_is_posix <- function(numFmt) {

  # if numFmt is character(0)
  if (length(numFmt) == 0) return(NULL)

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

#' check if numFmt is posix. internal function
#' @param numFmt numFmt xml nodes
#' @noRd
numfmt_is_hms <- function(numFmt) {

  # if numFmt is character(0)
  if (length(numFmt) == 0) return(NULL)

  numFmt_df <- read_numfmt(read_xml(numFmt))
  # we have to drop any square bracket part
  numFmt_df$fC <- gsub("\\[[^\\]]*]", "", numFmt_df$formatCode, perl = TRUE)
  num_fmts <- c(
    "#", as.character(0:9)
  )
  num_or_fmt <- paste0(num_fmts, collapse = "|")
  maybe_num <- grepl(pattern = num_or_fmt, x = numFmt_df$fC)

  hms_fmts <- c(
    "?!^yy$", "?!^yyyy$",
    "?!^mmm$", "?!^mmmm$", "?!^mmmmm$",
    "?!^d$", "?!^dd$", "?!^ddd$", "?!^dddd$",
    "h", "hh", ":m", ":mm", ":s", ":ss",
    "AM", "PM", "A", "P"
  )
  hms_or_fmt <- paste0(hms_fmts, collapse = "|")
  maybe_hms <- grepl(pattern = hms_or_fmt, x = numFmt_df$fC)

  z <- numFmt_df$numFmtId[maybe_hms & !maybe_num]
  if (length(z) == 0) z <- NULL
  z
}

#' check if style is date. internal function
#'
#' @param cellXfs cellXfs xml nodes
#' @param numfmt_date custom numFmtId dates
#' @noRd
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
#' @noRd
style_is_posix <- function(cellXfs, numfmt_date) {

  # numfmt_date: some basic date formats and custom formats
  date_numfmts <- as.character(22)
  numfmt_date <- c(numfmt_date, date_numfmts)

  cellXfs_df <- read_xf(read_xml(cellXfs))
  z <- rownames(cellXfs_df[cellXfs_df$numFmtId %in% numfmt_date, ])
  if (length(z) == 0) z <- NA
  z
}

#' check if style is hms. internal function
#'
#' @param cellXfs cellXfs xml nodes
#' @param numfmt_date custom numFmtId dates
#' @noRd
style_is_hms <- function(cellXfs, numfmt_date) {

  # numfmt_date: some basic date formats and custom formats
  date_numfmts <- as.character(18:21)
  numfmt_date <- c(numfmt_date, date_numfmts)

  cellXfs_df <- read_xf(read_xml(cellXfs))
  z <- rownames(cellXfs_df[cellXfs_df$numFmtId %in% numfmt_date, ])
  if (length(z) == 0) z <- NA
  z
}


#' Delete data
#'
#' This function is deprecated. Use [wb_clean_sheet()].
#' @param wb workbook
#' @param sheet sheet to clean
#' @param cols numeric column vector
#' @param rows numeric row vector
#' @export
#' @keywords internal
delete_data <- function(wb, sheet, cols, rows) {

  .Deprecated(old = "delete_data", new = "wb_clean_sheet", package = "openxlsx2")

  sheet_id <- wb_validate_sheet(wb, sheet)

  cc <- wb$worksheets[[sheet_id]]$sheet_data$cc

  if (is.numeric(cols)) {
    sel <- cc$row_r %in% as.character(as.integer(rows)) & cc$c_r %in% int2col(cols)
  } else {
    sel <- cc$row_r %in% as.character(as.integer(rows)) & cc$c_r %in% cols
  }

  # clean selected entries of cc
  clean <- names(cc)[!names(cc) %in% c("r", "row_r", "c_r")]
  cc[sel, clean] <- ""

  wb$worksheets[[sheet_id]]$sheet_data$cc <- cc

}
