
#' R6 class for a Workbook Chart Sheet
#'
#' @description
#' A chart sheet
#' @noRd
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
    #' @param tab_color tabColor
    #' @return The `wbChartSheet` object
    initialize = function(tab_color = NULL) {

      if (!is.null(tab_color)) {
        tab_color <- xml_node_create("tabColor", xml_attributes = tab_color)
        tabColor <- sprintf('<sheetPr>%s</sheetPr>', tab_color)
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
    #' @param sheet sheet
    #' @param color_id,default_grid_color Integer: A color, default is 64
    #' @param right_to_left Logical: if TRUE column ordering is right  to left
    #' @param show_formulas Logical: if TRUE cell formulas are shown
    #' @param show_grid_lines Logical: if TRUE the worksheet grid is shown
    #' @param show_outline_symbols Logical: if TRUE outline symbols are shown
    #' @param show_row_col_headers Logical: if TRUE row and column headers are shown
    #' @param show_ruler Logical: if TRUE a ruler is shown in page layout view
    #' @param show_white_space Logical: if TRUE margins are shown in page layout view
    #' @param show_zeros Logical: if FALSE cells containing zero are shown blank if !showFormulas
    #' @param tab_selected Integer: zero vector indicating the selected tab
    #' @param top_left_cell Cell: the cell shown in the top left corner / or top right with rightToLeft
    #' @param view View: "normal", "pageBreakPreview" or "pageLayout"
    #' @param window_protection Logical: if TRUE the panes are protected
    #' @param workbook_view_id integer: Pointing to some other view inside the workbook
    #' @param zoom_scale,zoom_scale_normal,zoom_scale_page_layout_view,zoom_scale_sheet_layout_view Integer: the zoom scale should be between 10 and 400. These are values for current, normal etc.
    #' @param ... additional arguments
    #' @return The `wbWorksheetObject`, invisibly
    set_sheetview = function(
      color_id                     = NULL,
      default_grid_color           = NULL,
      right_to_left                = NULL,
      show_formulas                = NULL,
      show_grid_lines              = NULL,
      show_outline_symbols         = NULL,
      show_row_col_headers         = NULL,
      show_ruler                   = NULL,
      show_white_space             = NULL,
      show_zeros                   = NULL,
      tab_selected                 = NULL,
      top_left_cell                = NULL,
      view                         = NULL,
      window_protection            = NULL,
      workbook_view_id             = NULL,
      zoom_scale                   = NULL,
      zoom_scale_normal            = NULL,
      zoom_scale_page_layout_view  = NULL,
      zoom_scale_sheet_layout_view = NULL,
      ...
    ) {

      standardize(...)

      # all zoom scales must be in the range of 10 - 400

      # get existing sheetView
      sheetView <- xml_node(self$sheetViews, "sheetViews", "sheetView")

      if (length(sheetView) == 0)
        sheetView <- xml_node_create("sheetView")

      sheetView <- xml_attr_mod(
        sheetView,
        xml_attributes = c(
          colorId                  = as_xml_attr(color_id),
          defaultGridColor         = as_xml_attr(default_grid_color),
          rightToLeft              = as_xml_attr(right_to_left),
          showFormulas             = as_xml_attr(show_formulas),
          showGridLines            = as_xml_attr(show_grid_lines),
          showOutlineSymbols       = as_xml_attr(show_outline_symbols),
          showRowColHeaders        = as_xml_attr(show_row_col_headers),
          showRuler                = as_xml_attr(show_ruler),
          showWhiteSpace           = as_xml_attr(show_white_space),
          showZeros                = as_xml_attr(show_zeros),
          tabSelected              = as_xml_attr(tab_selected),
          topLeftCell              = as_xml_attr(top_left_cell),
          view                     = as_xml_attr(view),
          windowProtection         = as_xml_attr(window_protection),
          workbookViewId           = as_xml_attr(workbook_view_id),
          zoomScale                = as_xml_attr(zoom_scale),
          zoomScaleNormal          = as_xml_attr(zoom_scale_normal),
          zoomScalePageLayoutView  = as_xml_attr(zoom_scale_page_layout_view),
          zoomScaleSheetLayoutView = as_xml_attr(zoom_scale_sheet_layout_view)
        ),
        remove_empty_attr = FALSE
      )

      self$sheetViews <- xml_node_create(
        "sheetViews",
        xml_children = sheetView
      )

      invisible(self)
    }
  )
)
