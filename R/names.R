
#' get or set worksheet names
#' get or set worksheet names
#'
#' @param x A `wbWorkbook` object
#' @export
#' @examples
#'
#' wb <- createWorkbook()
#' addWorksheet(wb, "S1")
#' addWorksheet(wb, "S2")
#' addWorksheet(wb, "S3")
#'
#' names(wb)
#' names(wb)[[2]] <- "S2a"
#' names(wb)
#' names(wb) <- paste("Sheet", 1:3)
names.wbWorkbook <- function(x) {
  replaceXMLEntities(x$sheet_names)
}

#' @rdname names.wbWorkbook
#' @param value a character vector the same length as `x`
#' @export
`names<-.wbWorkbook` <- function(x, value) {
  op <- openxlsx_options()
  on.exit(options(op), add = TRUE)

  if (any(duplicated(tolower(value)))) {
    stop("Worksheet names must be unique.")
  }

  existing_sheets <- x$sheet_names
  inds <- which(value != existing_sheets)

  if (length(inds) == 0) {
    return(invisible(x))
  }

  if (length(value) != length(x$worksheets)) {
    stop(sprintf("names vector must have length equal to number of worksheets in Workbook [%s]", length(existing_sheets)))
  }

  if (any(nchar(value) > 31)) {
    warning("Worksheet names must less than 32 characters. Truncating names...")
    value[nchar(value) > 31] <- vapply(value[nchar(value) > 31], substr, NA_character_, start = 1, stop = 31)
  }

  for (i in inds) {
    invisible(x$setSheetName(i, value[[i]]))
  }

  invisible(x)
}
