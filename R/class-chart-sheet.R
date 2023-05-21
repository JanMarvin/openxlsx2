
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

    #' @field sheetProtection sheetProtection
    sheetProtection = character(),

    #' @field customSheetViews customSheetViews
    customSheetViews = character(),

    #' @field pageMargins page margins
    pageMargins = character(),

    #' @field pageSetup pageSetup
    pageSetup = character(),

    #' @field headerFooter headerFooter
    headerFooter = character(),

    #' @field drawing drawing
    drawing = character(),

    #' @field drawingHF drawingHF
    drawingHF = character(),

    #' @field picture picture
    picture = character(),

    #' @field webPublishItems webPublishItems
    webPublishItems = character(),

    #' #' @field hyperlinks hyperlinks
    #' hyperlinks = NULL,

    #' @field relships relships
    relships = NULL,

    #' @description
    #' Create a new workbook chart sheet object
    #' @param tabColor `character` a tab color to set
    #' @return The `wbChartSheet` object
    initialize = function(tabColor = tabColor) {
      if (length(tabColor)) {
        tabColor <- sprintf('<sheetPr><tabColor rgb="%s"/></sheetPr>', tabColor)
      } else {
        tabColor <- character()
      }

      self$sheetPr     <- tabColor
      self$sheetViews  <- character()
      self$pageMargins <- '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
      self$drawing     <- '<drawing r:id=\"rId1\"/>'
      self$relships              <- list(
        comments         = integer(),
        drawing          = integer(),
        pivotTable       = integer(),
        slicer           = integer(),
        table            = integer(),
        threadedComments = integer(),
        vmlDrawing       = integer()
      )

      invisible(self)
    },

    # TODO should this be `get_sheet_data()`?  or `to_xml()`?
    #' @description
    #' get (prior) sheet data
    #' @returns A character vector of xml
    get_prior_sheet_data = function() {
      paste_c(
        '<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" mc:Ignorable="xr xr3">',
        self$sheetPr,
        self$sheetViews,
        self$customSheetViews,
        # self$hyperlinks,
        self$pageMargins,
        self$pageSetup,
        self$headerFooter,
        self$drawing,
        self$drawingHF,
        self$picture,
        self$webPublishItems,
        "</chartsheet>",
        sep = " "
      )
    },


    #' @description add sheetview
    #' @param colorId colorId
    #' @param defaultGridColor defaultGridColor
    #' @param rightToLeft rightToLeft
    #' @param showFormulas showFormulas
    #' @param showGridLines showGridLines
    #' @param showOutlineSymbols showOutlineSymbols
    #' @param showRowColHeaders showRowColHeaders
    #' @param showRuler showRuler
    #' @param showWhiteSpace showWhiteSpace
    #' @param showZeros showZeros
    #' @param tabSelected tabSelected
    #' @param topLeftCell topLeftCell
    #' @param view view
    #' @param windowProtection windowProtection
    #' @param workbookViewId workbookViewId
    #' @param zoomScale zoomScale
    #' @param zoomScaleNormal zoomScaleNormal
    #' @param zoomScalePageLayoutView zoomScalePageLayoutView
    #' @param zoomScaleSheetLayoutView zoomScaleSheetLayoutView
    #' @return The `wbWorksheetObject`, invisibly
    set_sheetview = function(
      colorId                  = NULL,
      defaultGridColor         = NULL,
      rightToLeft              = NULL,
      showFormulas             = NULL,
      showGridLines            = NULL,
      showOutlineSymbols       = NULL,
      showRowColHeaders        = NULL,
      showRuler                = NULL,
      showWhiteSpace           = NULL,
      showZeros                = NULL,
      tabSelected              = NULL,
      topLeftCell              = NULL,
      view                     = NULL,
      windowProtection         = NULL,
      workbookViewId           = NULL,
      zoomScale                = NULL,
      zoomScaleNormal          = NULL,
      zoomScalePageLayoutView  = NULL,
      zoomScaleSheetLayoutView = NULL
    ) {

      # all zoom scales must be in the range of 10 - 400

      # get existing sheetView
      sheetView <- xml_node(self$sheetViews, "sheetViews", "sheetView")

      if (length(sheetView) == 0)
        sheetView <- xml_node_create("sheetView")

      sheetView <- xml_attr_mod(
        sheetView,
        xml_attributes = c(
          # order according to Ecma Office Open XML Part 1. p3929
          windowProtection         = as_xml_attr(windowProtection),
          showFormulas             = as_xml_attr(showFormulas),
          showGridLines            = as_xml_attr(showGridLines),
          showZeros                = as_xml_attr(showZeros),
          rightToLeft              = as_xml_attr(rightToLeft),
          tabSelected              = as_xml_attr(tabSelected),
          showRuler                = as_xml_attr(showRuler),
          showOutlineSymbols       = as_xml_attr(showOutlineSymbols),
          defaultGridColor         = as_xml_attr(defaultGridColor),
          showWhiteSpace           = as_xml_attr(showWhiteSpace),
          view                     = as_xml_attr(view),
          topLeftCell              = as_xml_attr(topLeftCell),
          colorId                  = as_xml_attr(colorId),
          zoomScale                = as_xml_attr(zoomScale),
          showRowColHeaders        = as_xml_attr(showRowColHeaders),
          zoomScaleNormal          = as_xml_attr(zoomScaleNormal),
          zoomScalePageLayoutView  = as_xml_attr(zoomScalePageLayoutView),
          zoomScaleSheetLayoutView = as_xml_attr(zoomScaleSheetLayoutView),
          workbookViewId           = as_xml_attr(workbookViewId)
        )
      )

      self$sheetViews <- xml_node_create(
        "sheetViews",
        xml_children = sheetView
      )

      invisible(self)
    }
  )
)
