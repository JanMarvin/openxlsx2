

ChartSheet <- setRefClass(
  "ChartSheet",
  fields = c(
    "sheetPr" = "character",
    "sheetViews" = "character",
    "pageMargins" = "character",
    "drawing" = "character",
    "hyperlinks" = "ANY"
  ),

  methods = list(
    initialize = function(
      tabSelected = FALSE,
      tabColour = character(0),
      zoom = 100
    ) {
      if (length(tabColour) > 0) {
        tabColour <- sprintf("<sheetPr>%s</sheetPr>", tabColour)
      } else {
        tabColour <- character(0)
      }
      if (zoom < 10) {
        zoom <- 10
      } else if (zoom > 400) {
        zoom <- 400
      }

      .self$sheetPr <- tabColour
      .self$sheetViews <- sprintf('<sheetViews><sheetView workbookViewId="0" zoomScale="%s" tabSelected="%s"/></sheetViews>', as.integer(zoom), as.integer(tabSelected))
      .self$pageMargins <- '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
      .self$drawing <- '<drawing r:id=\"rId1\"/>'
      .self$hyperlinks <- character(0)

      invisible(.self)
    },

    get_prior_sheet_data = function() {
      paste_c(
        '<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">>',
        .self$sheetPr,
        .self$sheetViews,
        .self$pageMargins,
        .self$drawing,
        "</chartsheet>",
        sep = " "
      )
    }
  )
)

new_chart_sheet <- function() {
  ChartSheet$new()
}
