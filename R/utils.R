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
`%out%` <- function(x, table) {
  match(x, table, nomatch = 0L) == 0L
}

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

as_binary <- function(x) {
  # To be used within a function
  if (any(x %out% list(0, 1, FALSE, TRUE))) {
    stop(deparse(x), " must be 0, 1, FALSE, or TRUE", call. = FALSE)
  }

  as.integer(x)
}

#' random string function that does not alter the seed.
#' 
#' simple wrapper around `stringi::stri_rand_strings()``
#' 
#' @param n n
#' @param length length
#' @param pattern pattern
#' @keywords internal
#' @noRd
random_string <- function(n = 1, length, pattern = "[A-Za-z0-9]") {
  # https://github.com/ycphs/openxlsx/issues/183
  # https://github.com/ycphs/openxlsx/pull/224
  
  if (missing(length)) length <- substitute()

  rng <- get(".Random.seed", .GlobalEnv)

  # try to get a previous openxlsx2 seed and use this as random seed
  try(openxlsx2_seed <- get(".openxlsx2.seed", .GlobalEnv), TRUE)

  if (exists("openxlsx2_seed")) {
    # found one, change the global seed for stri_rand_strings
    assign(".Random.seed", openxlsx2_seed, .GlobalEnv)
  } else { 
    # non found, change the seed to the current
    openxlsx2_seed <- rng
  }

  # create random string, this alters the seed
  res <- stringi::stri_rand_strings(n = 1, length = length, pattern = pattern)

  # store the altered seed and reset the
  assign(".openxlsx2.seed", .Random.seed, .GlobalEnv)
  assign(".Random.seed", rng, .GlobalEnv)

  return (res)
}
