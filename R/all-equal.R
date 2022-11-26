
expect_equal_workbooks <- function(object, expected, ..., ignore_fields = NULL) {
  # Quick internal expectation function.  Easier to ignore fields to be set

  requireNamespace("testthat")
  requireNamespace("waldo")

  # object <- as.list(object)
  # expected <- as.list(expected)

  assert_workbook(object)
  assert_workbook(expected)

  fields <- c(names(wbWorkbook$public_fields), names(wbWorkbook$private_fields))
  bad <- setdiff(ignore_fields, fields)

  if (length(bad)) {
    stop("Invalid fields: ", toString(bad))
  }

  for (i in ignore_fields) {
    object[[i]] <- NULL
    expected[[i]] <- NULL
  }

  bad <- waldo::compare(
    x     = object,
    y     = expected,
    x_arg = "object",
    y_arg = "expected",
    ...
  )

  if (length(bad)) {
    testthat::fail(bad)
    return(invisible())
  }

  testthat::succeed()
  return(invisible())
}
