is_integer_ish <- function(x) {
  if (is.integer(x)) {
    return(TRUE)
  }

  if (!is.numeric(x)) {
    return(FALSE)
  }

  all(x[!is.na(x)] %% 1 == 0)
}

naToNULLList <- function(x) {
  lapply(x, function(i) if (is.na(i)) NULL else i)
}

# useful for replacing multiple x <- paste(x, new) where the length is checked
# multiple times.  This combines all elements in ... and removes anything that
# is zero length.  Much faster than multiple if/else (re)assignments
paste_c <- function(..., sep = "", collapse = " ", unlist = FALSE) {
  x <- c(...)
  if (unlist) x <- unlist(x, use.names = FALSE)
  stri_join(x[nzchar(x)], sep = sep, collapse = collapse)
}

`%||%` <- function(x, y) if (is.null(x)) y else x
`%|||%` <- function(x, y) if (length(x)) x else y

na_to_null <- function(x) {
  if (is.null(x)) {
    return(NULL)
  }

  if (isTRUE(is.na(x))) {
    return(NULL)
  }

  x
}

# opposite of %in%
`%out%` <- function(x, table) match(x, table, nomatch = 0L) == 0L



#' helper function to create temporary directory for testing purpose
#' @param name for the temp file
#' @param macros logical if the file extension is xlsm or xlsx
#' @export
temp_xlsx <- function(name = "temp_xlsx", macros = FALSE) {
  fileext <- ifelse(macros, ".xlsm", ".xlsx")
  tempfile(pattern = paste0(name, "_"), fileext = fileext)
}

openxlsx2_options <- function() {
  options(
    # increase scipen to avoid writing in scientific
    scipen = 200,
    OutDec = ".",
    digits = 22
  )
}

unapply <- function(x, FUN, ..., .recurse = TRUE, .names = FALSE) {
  FUN <- match.fun(FUN)
  unlist(lapply(X = x, FUN = FUN, ...), recursive = .recurse, use.names = .names)
}

reg_match0 <- function(x, pat) regmatches(x, gregexpr(pat, x, perl = TRUE))
reg_match  <- function(x, pat) regmatches(x, gregexpr(pat, x, perl = TRUE))[[1]]

apply_reg_match  <- function(x, pat) unapply(x, reg_match,  pat = pat)
apply_reg_match0 <- function(x, pat) unapply(x, reg_match0, pat = pat)

wapply <- function(x, FUN, ...) {
  FUN <- match.fun(FUN)
  which(vapply(x, FUN, FUN.VALUE = NA, ...))
}

has_chr <- function(x, na = FALSE) {
  # na controls if NA is returned as TRUE or FALSE
  vapply(nzchar(x, keepNA = !na), isTRUE, NA)
}

dir_create <- function(..., warn = TRUE, recurse = TRUE) {
  # create path and directory -- returns path
  path <- file.path(...)
  dir.create(path, showWarnings = warn, recursive = recurse)
  path
}

random_string <- function(size = NULL) {
  # creates a random string using tempfile() which does better to not affect the
  # random seed
  # https://github.com/ycphs/openxlsx/issues/183
  # https://github.com/ycphs/openxlsx/pull/224
  res <- basename(tempfile(""))

  if (is.null(size)) {
    return(res)
  }

  size <- as.integer(size)
  stopifnot(size >= 0L)
  while (nchar(res) < size) {
    res <- paste0(res, random_string())
  }

  substr(res, 1L, size)
}

dims_to_rowcol <- function(x, as_integer = FALSE) {
  dimensions <- unlist(strsplit(x, ":"))
  cols <- gsub("[[:digit:]]","", dimensions)
  rows <- gsub("[[:upper:]]","", dimensions)

  # if "A:B"
  has_row <- TRUE
  if (any(rows == "")) {
    has_row <- FALSE
    rows[rows == ""] <- "1"
  }

  # convert cols to integer
  cols_int <- col2int(cols)
  rows_int <- as.integer(rows)

  if (length(dimensions) == 2) {
    # needs integer to create sequence
    cols <- int2col(seq(cols_int[1], cols_int[2]))
    rows_int <- seq(rows_int[1], rows_int[2])
  }

  if (as_integer) {
    cols <- cols_int
    rows <- rows_int
  } else {
    rows <- as.character(rows)
  }

  list(cols, rows)
}
