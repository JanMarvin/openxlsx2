
#' R6 class for a Workbook Chart Sheet
#'
#' A chart sheet
#'
#' @export
wbChartSheet <- R6::R6Class(
  "wbChartSheet",

  # TODO add field definitions
  public = list(
    #' @field sheetPr Sheet something?
    sheetPr = character(),

    #' @field sheetViews Something
    sheetViews = character(),

    #' @field pageMargins page margins
    pageMargins = character(),

    #' @field drawing drawing
    drawing = character(),


    #' @field hyperlinks hyperlinks
    hyperlinks = NULL,

    #' @description
    #' Create a new workbook chart sheet object
    #' @param tabSelected `logical`, if `TRUE` ...
    #' @param tabColour `character` a tab colour to set
    #' @param zoom The zoom level as a single integer
    #' @return The `wbChartSheet` object
    initialize = function(
      tabSelected = FALSE,
      tabColour = character(),
      zoom = 100
    ) {
      if (length(tabColour)) {
        tabColour <- sprintf("<sheetPr>%s</sheetPr>", tabColour)
      } else {
        tabColour <- character()
      }

      # nocov start
      if (zoom < 10) {
        zoom <- 10
      } else if (zoom > 400) {
        zoom <- 400
      }
      #nocov end

      self$sheetPr     <- tabColour
      self$sheetViews  <- sprintf('<sheetViews><sheetView workbookViewId="0" zoomScale="%s" tabSelected="%s"/></sheetViews>', as.integer(zoom), as.integer(tabSelected))
      self$pageMargins <- '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
      self$drawing     <- '<drawing r:id=\"rId1\"/>'
      self$hyperlinks  <- NULL

      invisible(self)
    },

    # TODO should this be `get_sheet_data()`?  or `to_xml()`?
    #' @description
    #' get (prior) sheet data
    #' @returns A character vector of xml
    get_prior_sheet_data = function() {
      paste_c(
        '<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">>',
        self$sheetPr,
        self$sheetViews,
        self$pageMargins,
        self$drawing,
        "</chartsheet>",
        sep = " "
      )
    }
  )
)

#' New chart sheet
#'
#' Create a new chart sheet
#'
#' @returns a `wbChartSheet` object
#' @export
wb_chart_sheet <- function() {
  wbChartSheet$new()
}
