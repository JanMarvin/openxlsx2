#' Create dataframe from dimensions
#'
#' Non consecutive decreasing dims will return an increasing data frame.
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

  if (any(grepl("$", dims))) {
    dims <- gsub("\\$", "", dims)
  }

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

    ldims     <- get_dims(dims)
    full_rows <- ldims$rows[[1]]
    full_cols <- ldims$cols

    rows_out  <- seq.int(full_rows[1], full_rows[2])
    cols_out  <- full_cols
    full_cols <- full_cols - min(full_cols) # is always a zero offset

  } else {
    ll <- dims_to_row_col_fill(dims, fill)

    filled   <- ll$fill
    cols_out <- ll$cols
    rows_out <- ll$rows
  }

  if (has_dim_sep) {
    if (empty_rm) {
      # with empty_rm the dataframe will contain only needed columns
      if (!is.null(full_cols)) full_cols <- seq_along(cols_out) - 1L
    } else {
      # somehow we have to make sure that all columns are covered
      col_ints <- range(cols_out)
      cols_out <- seq.int(from = col_ints[1], to = col_ints[2])

      row_ints <- range(rows_out)
      rows_out <- seq.int(from = row_ints[1], to = row_ints[2])
    }
  }

  dims_to_df(
    rows   = rows_out,
    cols   = int2col(cols_out),
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
  # Initialize types vector with numeric type (default to 0 for character)
  types <- vector("numeric", NCOL(tt))
  names(types) <- names(tt)

  # Identify the unique types present in the data frame
  uu <- lapply(tt, unique)
  unique_types <- unique(unlist(uu))
  unique_types[is.na(unique_types)] <- "n"

  # Function to check column type
  check_col_type <- function(x, type_char) {
    all(unique(x) == type_char, na.rm = TRUE)
  }

  check_type <- function(x) vapply(uu, check_col_type, NA, type_char = x)

  # Check for each type and update types vector accordingly
  if ("n" %in% unique_types) {
    col_num <- check_type("n")
    types[col_num] <- 1
  }

  if ("d" %in% unique_types) {
    col_dte <- check_type("d")
    types[col_dte & types == 0] <- 2
  }

  if ("p" %in% unique_types) {
    col_posix <- check_type("p")
    types[col_posix & types == 0] <- 3
  }

  if ("b" %in% unique_types) {
    col_log <- check_type("b")
    types[col_log & types == 0] <- 4
  }

  if ("h" %in% unique_types) {
    col_hms <- check_type("h")
    types[col_hms & types == 0] <- 5
  }

  if ("f" %in% unique_types) {
    col_fml <- check_type("f")
    types[col_fml & types == 0] <- 6
  }

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
  numFmt_df$fC[tolower(numFmt_df$fC) == "general"] <- ""
  num_fmts <- c(
    "#", as.character(0:9)
  )
  num_or_fmt <- paste0(num_fmts, collapse = "|")
  maybe_num <- grepl(pattern = num_or_fmt, x = numFmt_df$fC)

  date_fmts <- c(
    "yy", "yyyy",
    "YY", "YYYY",
    "m", "mm", "mmm", "mmmm", "mmmmm",
    "M", "MM", "MMM", "MMMM", "MMMMM",
    "d", "dd", "ddd", "dddd",
    "D", "DD", "DDD", "DDDD"
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
  numFmt_df$fC[tolower(numFmt_df$fC) == "general"] <- ""
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
    "H", "HH", ":M", ":MM", ":S", ":SS",
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
  numFmt_df$fC[tolower(numFmt_df$fC) == "general"] <- ""
  num_fmts <- c(
    "#", as.character(0:9)
  )
  num_or_fmt <- paste0(num_fmts, collapse = "|")
  maybe_num <- grepl(pattern = num_or_fmt, x = numFmt_df$fC)

  hms_fmts <- c(
    "?!^yy$", "?!^yyyy$",
    "?!^YY$", "?!^YYYY$",
    "?!^mmm$", "?!^mmmm$", "?!^mmmmm$",
    "?!^MMM$", "?!^MMMM$", "?!^MMMMM$",
    "?!^d$", "?!^dd$", "?!^ddd$", "?!^dddd$",
    "?!^D$", "?!^DD$", "?!^DDD$", "?!^DDDD$",
    "h", "hh", ":m", ":mm", ":s", ":ss",
    "H", "HH", ":M", ":MM", ":S", ":SS",
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
  # # TODO this might not work as expected in xlsb
  # # if the number format is not applied, ignore it
  # cellXfs_df <- cellXfs_df[cellXfs_df$applyNumberFormat %in% c("1", "true"), ]
  z <- rownames(cellXfs_df[cellXfs_df$numFmtId %in% numfmt_date, ])
  if (length(z) == 0) z <- NA
  z
}
