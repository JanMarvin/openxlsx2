
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


# With R6 use assert_class(x, c("class", "R6"), all = TRUE)
assert_chart_sheet <- function(x) assert_class(x, "ChartSheet", package = "openxlsx2")
assert_comment     <- function(x) assert_class(x, "Comment",    package = "openxlsx2")
assert_hyperlink   <- function(x) assert_class(x, "Hyperlink",  package = "openxlsx2")
assert_sheet_data  <- function(x) assert_class(x, "SheetData",  package = "openxlsx2")
assert_style       <- function(x) assert_class(x, "Style",      package = "openxlsx2")
assert_workbook    <- function(x) assert_class(x, "Workbook",   package = "openxlsx2")
assert_worksheet   <- function(x) assert_class(x, "Worksheet",  package = "openxlsx2")


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

match_allof <- function(x, y, or_null = FALSE, envir = parent.frame()) {
  sx <- as.character(substitute(x, envir))

  if (or_null && is.null(x)) return(NULL)

  m <- match(x, y, nomatch = NA_integer_)

  if (anyNA(m)) {
    msg <- sprintf("%s must be: '%s'", sx, paste(y, collapse = "', '"))
    stop(simpleError(msg))
  }

  invisible(x)
}

validate_colour <- function(colour = NULL, or_null = FALSE, envir = parent.frame()) {
  sx <- as.character(substitute(colour, envir))

  if (identical(colour, "none") && or_null) {
    return(NULL)
  }

  # returns black
  if (is.null(colour)) {
    if (or_null) return(NULL)
    return("FF000000")
  }

  if (ind <- any(colour %in% grDevices::colours())) {
    colour[ind] <- col2hex(colour[ind])
  }

  if (any(!grepl("^#[A-Fa-f0-9]{6}$", colour))) {
    msg <- sprintf("`%s` ['%s'] is not a valid colour", sx, colour)
    stop(simpleError(msg))
  }

  gsub("^#", "FF", toupper(colour))
}
