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
  paste(x[nzchar(x)], sep = "", collapse = " ")
}

`%||%` <- function(x, y) if (is.null(x)) y else x

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



#' helper function to create tempory directory for testing purpose
#' @param name for the temp file
#' @export
temp_xlsx <- function(name = "temp_xlsx") {
  tempfile(pattern = paste0(name, "_"), fileext = ".xlsx")
}


openxlsx_options <- function() {
  # op <- openxlsx_options()
  # on.exit(options(op), add = TRUE)
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
