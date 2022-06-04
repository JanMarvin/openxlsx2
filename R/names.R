
#' get or set worksheet names
#' get or set worksheet names
#'
#' @param x A `wbWorkbook` object
#' @export
names.wbWorkbook <- function(x) {
  .Deprecated("$get_sheet_names()")
  assert_workbook(x)
  x$get_sheet_names()
}

#' @rdname names.wbWorkbook
#' @param value a character vector the same length as `x`
#' @export
`names<-.wbWorkbook` <- function(x, value) {
  .Deprecated("$set_sheet_names()")
  assert_workbook(x)
  x$clone()$set_sheet_names(new = value)
}
