#' Validate sheet
#'
#' @param wb A workbook
#' @param sheetName The sheet name to validate
#' @return The sheet name -- or the position?  This should be consistent
#' @noRd
wb_validate_sheet <- function(wb, sheetName) {
  assert_workbook(wb)

  if (!is.numeric(sheetName)) {
    if (is.null(wb$sheet_names)) {
      stop("wb does not contain any worksheets.", call. = FALSE)
    }
  }

  if (is.numeric(sheetName)) {
    if (sheetName > length(wb$sheet_names)) {
      msg <- sprintf("wb only contains %i sheets.", length(wb$sheet_names))
      stop(msg, call. = FALSE)
    }

    return(sheetName)
  }

  if (!sheetName %in% replaceXMLEntities(wb$sheet_names)) {
    msg <- sprintf("Sheet '%s' does not exist.", replaceXMLEntities(sheetName))
    stop(msg, call. = FALSE)
  }


  which(replaceXMLEntities(wb$sheet_names) == sheetName)
}
