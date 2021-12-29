
assert_class <- function(x, class, or_null = FALSE, all = TRUE, package = NULL, envir = parent.frame()) {
  sx <- as.character(substitute(x, envir))

  if (is.null(package)) {
    ok <- if (all) {
      all(vapply(class, function(i) inherits(x, i), NA))
    } else {
      inherits(x, class)
    }
  } else {
    cl <- class(x)

    ok <- if (all) {
      all(class %in% cl)
    } else {
      any(cl == class)
    }

    ok <- ok && attr(cl, "package") == package
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
