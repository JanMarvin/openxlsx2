
#' @name names
#' @title get or set worksheet names
#' @description get or set worksheet names
#' @aliases names.Workbook
#' @export
#' @method names Workbook
#' @param x A \code{Workbook} object
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
names.Workbook <- function(x) {
  nms <- x$sheet_names
  nms <- replaceXMLEntities(nms)
}

#' @rdname names
#' @param value a character vector the same length as wb
#' @export
`names<-.Workbook` <- function(x, value) {
  od <- getOption("OutDec")
  options("OutDec" = ".")
  on.exit(expr = options("OutDec" = od), add = TRUE)

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
    value[nchar(value) > 31] <- sapply(value[nchar(value) > 31], substr, start = 1, stop = 31)
  }

  for (i in inds) {
    invisible(x$setSheetName(i, value[[i]]))
  }

  invisible(x)
}
