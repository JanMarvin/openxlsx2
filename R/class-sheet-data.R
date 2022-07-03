
#' R6 class for a Workbook Hyperlink
#'
#' A hyperlink
#'
#' @export
wbSheetData <- R6::R6Class(
  "wbSheetData",
  public = list(
    # TODO which fields should be moved to private?

    #' @field row_attr row_attr
    row_attr = NULL,

    #' @field cc cc
    cc = NULL,

    #' @field cc_out cc_out
    cc_out = NULL,

    #' @description
    #' Creates a new `wbSheetData` object
    #' @return a `wbSheetData` object
    initialize = function() {
      invisible(self)
    }
  )
)
    
#' @rdname wbSheetData
#' @export
wb_sheet_data <- function() {
  wbSheetData$new()
}


# helpers -----------------------------------------------------------------

# Consider making some helpers for the cc stuff.

empty_sheet_data_cc <- function(n = 0) {
  value <- rep.int(NA_character_, n)
  data.frame(
    row_r = value,
    c_r   = value,
    c_s   = value,
    c_t   = value,
    v     = value,
    f     = value,
    f_t   = value,
    f_ref = value,
    f_si  = value,
    is    = value,
    typ   = value,
    r     = value,
    check.names = FALSE,
    stringsAsFactors = FALSE
  )
}


empty_row_attr <- function(n = 0) {
  value <- character(n)
  data.frame(
    collapsed         = value,
    customFormat      = value,
    customHeight      = value,
    `x14ac:dyDescent` = value,
    ht                = value,
    hidden            = value,
    outlineLevel      = value,
    r                 = value,
    ph                = value,
    spans             = value,
    s                 = value,
    thickBot          = value,
    thickTop          = value,
    check.names = FALSE,
    stringsAsFactors = FALSE
  )
}
