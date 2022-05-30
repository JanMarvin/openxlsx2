
#' get or set worksheet names
#' get or set worksheet names
#'
#' @param x A `wbWorkbook` object
#' @export
#' @examples
#'
#' wb <- wb_workbook()
#' wb$add_worksheet("S1")
#' wb$add_worksheet("S2")
#' wb$add_worksheet("S3")
#'
#' names(wb)
#' names(wb)[[2]] <- "S2a"
#' names(wb)
#' names(wb) <- paste("Sheet", 1:3)
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
