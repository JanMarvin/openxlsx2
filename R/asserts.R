
assert_class <- function(x, class, or_null = FALSE, all = FALSE, package = NULL, envir = parent.frame()) {
  sx <- as.character(substitute(x, envir))

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

validate_color <- function(color = NULL, or_null = FALSE, envir = parent.frame()) {
  sx <- as.character(substitute(color, envir))

  if (identical(color, "none") && or_null) {
    return(NULL)
  }

  # returns black
  if (is.null(color)) {
    if (or_null) return(NULL)
    return("FF000000")
  }

  if (ind <- any(color %in% grDevices::colors())) {
    color[ind] <- col2hex(color[ind])
  }

  if (any(!grepl("^#[A-Fa-f0-9]{6}$", color))) {
    msg <- sprintf("`%s` ['%s'] is not a valid color", sx, color)
    stop(simpleError(msg))
  }

  gsub("^#", "FF", toupper(color))
}
