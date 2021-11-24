
assert_class <- function(x, class, or_null = FALSE, envir = parent.frame()) {
  sx <- as.character(substitute(x, envir))
  ok <- inherits(x, class)

  if (or_null) {
    ok <- ok | is.null(x)
    class <- c(class, "null")
  }

  if (!ok) {
    msg <- sprintf("%s must be class %s", sx, paste(class, collapse = " or "))
    stop(simpleError(msg))
  }
}

assert_chart_sheet <- function(x) assert_class(x, "ChartSheet")
assert_comment     <- function(x) assert_class(x, "Comment")
assert_hyperlink   <- function(x) assert_class(x, "Hyperlink")
assert_sheet_data  <- function(x) assert_class(x, "SheetData")
assert_sytle       <- function(x) assert_class(x, "Style")
assert_workbook    <- function(x) assert_class(x, "Workbook")
assert_worksheet   <- function(x) assert_class(x, "Worksheet")
