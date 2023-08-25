# Use arg_nm to override the default name of the argument in case of an error message.
assert_class <- function(x, class, or_null = FALSE, all = FALSE, package = NULL, envir = parent.frame(), arg_nm = NULL) {

  sx <- as.character(substitute(x, envir))
  if (length(sx) == 0 || !is.null(arg_nm)) {
    sx <- arg_nm %||% "argument"
  }

  if (missing(x)) {
    stop("input ", sx, " is missing", call. = FALSE)
  }

  ok <- if (all) {
    all(vapply(class, function(i) inherits(x, i), NA))
  } else {
    inherits(x, class)
  }

  if (!is.null(package)) {
    ok <- ok & isTRUE(attr(class(x), "package") == package)
  }

  if (or_null) {
    ok <- ok | is.null(x)
    class <- c(class, "null")
  }

  if (!ok) {
    msg <- sprintf("%s must be class %s", sx, paste(class, collapse = " or "))
    stop(simpleError(msg))
  }

  invisible(NULL)
}

assert_chart_sheet <- function(x) assert_class(x, c("wbChartSheet", "R6"), all = TRUE)
assert_comment     <- function(x) assert_class(x, c("wbComment",    "R6"), all = TRUE)
assert_color       <- function(x) assert_class(x, c("wbColour"),           all = TRUE)
assert_hyperlink   <- function(x) assert_class(x, c("wbHyperlink",  "R6"), all = TRUE)
assert_sheet_data  <- function(x) assert_class(x, c("wbSheetData",  "R6"), all = TRUE)
assert_workbook    <- function(x) assert_class(x, c("wbWorkbook",   "R6"), all = TRUE)
assert_worksheet   <- function(x) assert_class(x, c("wbWorksheet",  "R6"), all = TRUE)

match_oneof <- function(x, y, or_null = FALSE, several = FALSE, envir = parent.frame()) {
  sx <- as.character(substitute(x, envir))

  if (or_null && is.null(x)) return(NULL)

  m <- match(x, y, nomatch = NA_integer_)
  m <- m[!is.na(m)]
  if (!several) m <- m[1]

  if (anyNA(m) || !length(m)) {
    msg <- sprintf("%s must be one of: '%s'", sx, paste(y, collapse = "', '"))
    stop(simpleError(msg))
  }

  y[m]
}
