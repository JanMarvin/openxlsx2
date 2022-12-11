workbook_validate_sheet <- function(self, private, sheet) {
  # workbook has no sheets
  if (!length(self$sheet_names)) {
    return(NA_integer_)
  }

  # write_comment uses wb_validate and bails otherwise
  if (inherits(sheet, "openxlsx2_waiver")) {
    sheet <- private$get_sheet_index(sheet)
  }

  # input is number
  if (is.numeric(sheet)) {
    badsheet <- !sheet %in% seq_along(self$sheet_names)
    if (any(badsheet)) sheet[badsheet] <- NA_integer_
    return(sheet)
  }

  if (!sheet %in% replaceXMLEntities(self$sheet_names)) {
    return(NA_integer_)
  }

  which(replaceXMLEntities(self$sheet_names) == sheet)
}
