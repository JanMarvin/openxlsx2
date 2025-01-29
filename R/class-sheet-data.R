
#' R6 class for a Workbook Hyperlink
#'
#' A hyperlink
#'
#' @noRd
wbSheetData <- R6::R6Class(
  "wbSheetData",
  public = list(
    # TODO which fields should be moved to private?

    #' @field row_attr row_attr
    row_attr = NULL,

    #' @field cc cc
    cc = NULL,

    #' @description
    #' Creates a new `wbSheetData` object
    #' @return a `wbSheetData` object
    initialize = function() {
      invisible(self)
    }
  )
)

## TODO is this even used?
#' @rdname wbSheetData
#' @noRd
wb_sheet_data <- function() {
  wbSheetData$new()
}


# helpers -----------------------------------------------------------------

empty_row_attr <- function(n) {
  create_char_dataframe(
    colnames = c("collapsed", "customFormat", "customHeight", "x14ac:dyDescent",
                 "ht", "hidden", "outlineLevel", "r", "ph", "spans", "s",
                 "thickBot", "thickTop"),
    n = n
  )
}
