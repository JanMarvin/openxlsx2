
# class -------------------------------------------------------------------

#' R6 class for a Workbook Worksheet
#'
#' A Worksheet
#'
#' @noRd
wbWorksheet <- R6::R6Class(
  "wbWorksheet",

  ## public ----
  public = list(

    # TODO can any of these be private?

    #' @field sheetPr sheetPr
    sheetPr = character(),

    #' @field dimension dimension
    dimension = character(),

    #' @field sheetViews sheetViews
    sheetViews = character(),

    #' @field sheetFormatPr sheetFormatPr
    sheetFormatPr = character(),

    #' @field sheet_data sheet_data
    sheet_data = NULL,

    #' @field cols_attr cols_attr
    cols_attr  = NULL,

    #' @field autoFilter autoFilter
    autoFilter = character(),

    #' @field mergeCells mergeCells
    mergeCells = NULL,

    #' @field conditionalFormatting conditionalFormatting
    conditionalFormatting = character(),

    #' @field dataValidations dataValidations
    dataValidations = NULL,

    #' @field freezePane freezePane
    freezePane = character(),

    #' @field hyperlinks hyperlinks
    hyperlinks = NULL,

    #' @field sheetProtection sheetProtection
    sheetProtection = character(),

    #' @field pageMargins pageMargins
    pageMargins = character(),

    #' @field pageSetup pageSetup
    pageSetup = character(),

    #' @field headerFooter headerFooter
    headerFooter = NULL,

    #' @field rowBreaks rowBreaks
    rowBreaks = character(),

    #' @field colBreaks colBreaks
    colBreaks = character(),

    #' @field drawing drawing
    drawing = character(),

    #' @field legacyDrawing legacyDrawing
    legacyDrawing = character(),

    #' @field legacyDrawingHF legacyDrawingHF
    legacyDrawingHF = character(),

    #' @field oleObjects oleObjects
    oleObjects = character(),

    #' @field tableParts tableParts
    tableParts = character(),

    #' @field extLst extLst
    extLst = character(),

    ### list with imported openxml-2.8.1 nodes
    #' @field cellWatches cellWatches
    cellWatches = character(),

    #' @field controls controls
    controls = character(),

    #' @field customProperties customProperties
    customProperties = character(),

    #' @field customSheetViews customSheetViews
    customSheetViews = character(),

    #' @field dataConsolidate dataConsolidate
    dataConsolidate = character(),

    #' @field drawingHF drawingHF
    drawingHF = character(),

    #' @field relships relships
    relships = NULL,

    #' @field ignoredErrors ignoredErrors
    ignoredErrors = character(),

    #' @field phoneticPr phoneticPr
    phoneticPr = character(),

    #' @field picture picture
    picture = character(),

    #' @field printOptions printOptions
    printOptions = character(),

    #' @field protectedRanges protectedRanges
    protectedRanges = character(),

    #' @field scenarios scenarios
    scenarios = character(),

    #' @field sheetCalcPr sheetCalcPr
    sheetCalcPr = character(),

    #' @field smartTags smartTags
    smartTags = character(),

    #' @field sortState sortState
    sortState = character(),

    #' @field webPublishItems webPublishItems
    webPublishItems = character(),

    #' @description
    #' Creates a new `wbWorksheet` object
    #' @param tab_color tabColor
    #' @param odd_header oddHeader
    #' @param odd_footer oddFooter
    #' @param even_header evenHeader
    #' @param even_footer evenFooter
    #' @param first_header firstHeader
    #' @param first_footer firstFooter
    #' @param paper_size paperSize
    #' @param orientation orientation
    #' @param hdpi hdpi
    #' @param vdpi vdpi
    #' @param grid_lines printGridLines
    #' @param ... additional arguments
    #' @return a `wbWorksheet` object
    initialize = function(
      tab_color    = NULL,
      odd_header   = NULL,
      odd_footer   = NULL,
      even_header  = NULL,
      even_footer  = NULL,
      first_header = NULL,
      first_footer = NULL,
      paper_size   = 9,
      orientation  = "portrait",
      hdpi         = 300,
      vdpi         = 300,
      grid_lines   = FALSE,
      ...
    ) {

      standardize_case_names(...)

      if (!is.null(tab_color)) {
        tabColor <- sprintf('<sheetPr><tabColor rgb="%s"/></sheetPr>', tab_color)
      } else {
        tabColor <- character()
      }

      hf <- list(
        oddHeader   = na_to_null(odd_header),
        oddFooter   = na_to_null(odd_footer),
        evenHeader  = na_to_null(even_header),
        evenFooter  = na_to_null(even_footer),
        firstHeader = na_to_null(first_header),
        firstFooter = na_to_null(first_footer)
      )

      if (all(lengths(hf) == 0)) {
        hf <- list()
      }

      # only add if printGridLines not TRUE. The openxml default is TRUE
      if (grid_lines) {
       self$set_print_options(gridLines = grid_lines, gridLinesSet = grid_lines)
      }

      ## list of all possible children
      self$sheetPr               <- tabColor
      self$dimension             <- '<dimension ref="A1"/>'
      self$sheetViews            <- character()
      self$sheetFormatPr         <- '<sheetFormatPr baseColWidth="8.43" defaultRowHeight="16" x14ac:dyDescent="0.2"/>'
      self$cols_attr             <- character()
      self$autoFilter            <- character()
      self$mergeCells            <- character()
      self$conditionalFormatting <- character()
      self$dataValidations       <- NULL
      self$hyperlinks            <- list()
      self$pageMargins           <- '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
      self$pageSetup             <- sprintf('<pageSetup paperSize="%s" orientation="%s" horizontalDpi="%s" verticalDpi="%s"/>', paper_size, orientation, hdpi, vdpi)
      self$headerFooter          <- hf
      self$rowBreaks             <- character()
      self$colBreaks             <- character()
      self$drawing               <- character()
      self$legacyDrawing         <- character()
      self$legacyDrawingHF       <- character()
      self$oleObjects            <- character()
      self$tableParts            <- character()
      self$extLst                <- character()
      self$freezePane            <- character()
      self$sheet_data            <- wbSheetData$new()
      self$relships              <- list(
        comments         = integer(),
        drawing          = integer(),
        pivotTable       = integer(),
        slicer           = integer(),
        table            = integer(),
        threadedComment  = integer(),
        vmlDrawing       = integer()
      )

      invisible(self)
    },

    #' @description
    #' Get prior sheet data
    #' @return A character vector of xml
    get_prior_sheet_data = function() {

      # apparently every sheet needs to have a sheetView
      sheetViews <- self$sheetViews

      if (length(self$freezePane)) {
        if (length(xml_node(sheetViews, "sheetViews", "sheetView")) == 1) {
          # get sheetView node and append freezePane
          # TODO Can we unfreeze a pane? It should be possible to simply null freezePane
          sheetViews <- xml_add_child(sheetViews, xml_child = self$freezePane, level = "sheetView")
        } else {
          message("Sheet contains multiple sheetViews. Could not freeze pane") #nocov
        }
      }

      paste_c(
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3">',

        # sheetPr
        if (length(self$sheetPr) && !any(xml_node_name(self$sheetPr) == "sheetPr")) {
          xml_node_create("sheetPr", xml_children = self$sheetPr)
        } else {
          self$sheetPr
        },

        self$dimension,

        sheetViews,

        self$sheetFormatPr,
        # cols_attr
        # is this fine if it's just <cols></cols>?
        if (length(self$cols_attr)) {
          paste(c("<cols>", self$cols_attr, "</cols>"), collapse = "")
        },
        '</worksheet>',
        sep = ""
      )
    },

    #' @description
    #' Get post sheet data
    #' @return A character vector of xml
    get_post_sheet_data = function() {
      paste_c(
        # self$sheetCalcPr -- we do not provide calcPr
        self$sheetProtection,
        self$protectedRanges,
        self$scenarios,
        self$autoFilter,
        self$sortState,
        self$dataConsolidate,
        self$customSheetViews,

        # mergeCells
        if (length(self$mergeCells)) {
          paste0(
            sprintf('<mergeCells count="%i">', length(self$mergeCells)),
            pxml(self$mergeCells),
            "</mergeCells>"
          )
        },

        self$phoneticPr,

        # conditionalFormatting
        if (length(self$conditionalFormatting)) {
          nms <- names(self$conditionalFormatting)
          paste(
            vapply(
              unique(nms),
              function(i) {
                paste0(
                  sprintf('<conditionalFormatting sqref="%s">', i),
                  pxml(self$conditionalFormatting[nms == i]),
                  "</conditionalFormatting>"
                )
              },
              NA_character_
            ),
            collapse = ""
          )
        },

        # dataValidations
        if (length(self$dataValidations)) {
          paste0(
            sprintf('<dataValidations count="%i">', length(self$dataValidations)),
            pxml(self$dataValidations),
            "</dataValidations>"
          )
        },

        # hyperlinks
        if (n <- length(self$hyperlinks)) {
          h_inds <- paste0(seq_len(n), "h")
          paste(
            "<hyperlinks>",
            paste(
              vapply(
                seq_along(h_inds),
                function(i)  {
                  self$hyperlinks[[i]]$to_xml(h_inds[i])
                },
                NA_character_
              ),
              collapse = ""
            ),
            "</hyperlinks>"
          )
        },

        self$printOptions,
        self$pageMargins,
        self$pageSetup,

        # headerFooter
        # should return NULL when !length(self$headerFooter)
        genHeaderFooterNode(self$headerFooter),

        # rowBreaks
        if (n <- length(self$rowBreaks)) {
          paste0(
            sprintf('<rowBreaks count="%i" manualBreakCount="%i">', n, n),
            paste(self$rowBreaks, collapse = ""),
            "</rowBreaks>"
          )
        },

        # colBreaks
        if (n <- length(self$colBreaks)) {
          paste0(
            sprintf('<colBreaks count="%i" manualBreakCount="%i">', n, n),
            paste(self$colBreaks, collapse = ""),
            "</colBreaks>"
          )
        },

        self$customProperties,
        self$cellWatches,
        self$ignoredErrors,
        self$smartTags,
        self$drawing,
        self$drawingHF,
        self$legacyDrawing,   # these appear to be removed in 2.8.1
        self$legacyDrawingHF, # these appear to be removed in 2.8.1
        self$picture,
        self$oleObjects,
        self$controls,
        self$webPublishItems,

        # tableParts
        if (n <- length(self$tableParts)) {
          paste0(sprintf('<tableParts count="%i">', n), pxml(self$tableParts), "</tableParts>")
        },

        # extLst
        if (length(self$extLst)) {
          sprintf(
            "<extLst>%s</extLst>",
            paste0(
              pxml(self$extLst)
            )
          )
        },

        # end
        sep = ""
      )

    },

    #' @description
    #' unfold `<cols ..>` node to dataframe. `<cols><col ..>` are compressed.
    #' Only columns with attributes are written to the file. This function
    #' unfolds them so that each cell beginning with the "A" to the last one
    #' found in cc gets a value.
    #' TODO might extend this to match either largest cc or largest col. Could
    #' be that "Z" is formatted, but the last value is written to "Y".
    #' TODO might replace the xml nodes with the data frame?
    #' @return The column data frame
    unfold_cols = function() {

      # avoid error and return empty data frame
      if (length(self$cols_attr) == 0)
        return(empty_cols_attr())

      col_df <- col_to_df(read_xml(self$cols_attr))
      col_df$min <- as.numeric(col_df$min)
      col_df$max <- as.numeric(col_df$max)

      max_col <- max(col_df$max)

      # always begin at 1, even if 1 is not in the dataset. fold_cols requires this
      key <- seq(1, max_col)

      # merge against this data frame
      tmp_col_df <- data.frame(
        key = key,
        stringsAsFactors = FALSE
      )

      out <- NULL
      for (i in seq_len(nrow(col_df))) {
        z <- col_df[i, ]
        for (j in seq(z$min, z$max)) {
          z$key <- j
          out <- rbind(out, z)
        }
      }

      # merge and convert to character, remove key
      col_df <- merge(x = tmp_col_df, y = out, by = "key", all.x = TRUE)
      col_df$min <- as.character(col_df$key)
      col_df$max <- as.character(col_df$key)
      col_df[is.na(col_df)] <- ""
      col_df$key <- NULL

      col_df
    },

    #' @description
    #' fold the column dataframe back into a node.
    #' @param col_df the column data frame
    #' @return The `wbWorksheetObject`, invisibly
    fold_cols = function(col_df) {

      # remove min and max columns and create merge identifier: string
      col_df <- col_df[-which(names(col_df) %in% c("min", "max"))]
      col_df$string <- apply(col_df, 1, paste, collapse = "")

      # run length
      out <- with(
        rle(col_df$string),
        data.frame(
          string = values,
          min = cumsum(lengths) - lengths + 1,
          max = cumsum(lengths))
      )

      # remove duplicates pre merge
      col_df <- unique(col_df)

      # merge with string variable, drop empty string and clean up
      col_df <- merge(out, col_df, by = "string", all.x = TRUE)
      col_df <- col_df[col_df$string != "", ]
      col_df$string <- NULL

      # order and return
      col_df <- col_df[order(col_df$min), ]
      col_df$min <- as.character(col_df$min)
      col_df$max <- as.character(col_df$max)

      # assign as xml-nodes
      self$cols_attr <- df_to_xml("col", col_df)

      invisible(self)
    },

    #' @description
    #' Set cell merging for a sheet
    #' @param rows,cols Row and column specifications.
    #' @return The `wbWorkbook` object, invisibly
    merge_cells = function(rows = NULL, cols = NULL) {

      rows <- range(as.integer(rows))
      cols <- range(col2int(cols))

      sqref <- paste0(int2col(cols), rows)
      sqref <- stri_join(sqref, collapse = ":", sep = " ")

      current <- rbindlist(xml_attr(xml = self$mergeCells, "mergeCell"))$ref

      # regmatch0 will return character(0) when x is NULL
      if (length(current)) {

        new_merge     <- unname(unlist(dims_to_dataframe(sqref, fill = TRUE)))
        current_cells <- lapply(current, function(x) unname(unlist(dims_to_dataframe(x, fill = TRUE))))
        intersects    <- vapply(current_cells, function(x) any(x %in% new_merge), NA)

        # Error if merge intersects
        if (any(intersects)) {
          msg <- sprintf(
            "Merge intersects with existing merged cells: \n\t\t%s.\nRemove existing merge first.",
            stri_join(current[intersects], collapse = "\n\t\t")
          )
          stop(msg, call. = FALSE)
        }
      }

      # TODO does this have to be xml?  Can we just save the data.frame or
      # matrix and then check that?  This would also simplify removing the
      # merge specifications
      self$append("mergeCells", sprintf('<mergeCell ref="%s"/>', sqref))

      invisible(self)

    },

    #' @description
    #' Removes cell merging for a sheet
    #' @param rows,cols Row and column specifications.
    #' @return The `wbWorkbook` object, invisibly
    unmerge_cells = function(rows = NULL, cols = NULL) {

      rows <- range(as.integer(rows))
      cols <- range(col2int(cols))

      sqref <- paste0(int2col(cols), rows)
      sqref <- stri_join(sqref, collapse = ":", sep = " ")

      current <- rbindlist(xml_attr(xml = self$mergeCells, "mergeCell"))$ref

      if (!is.null(current)) {
        new_merge     <- unname(unlist(dims_to_dataframe(sqref, fill = TRUE)))
        current_cells <- lapply(current, function(x) unname(unlist(dims_to_dataframe(x, fill = TRUE))))
        intersects    <- vapply(current_cells, function(x) any(x %in% new_merge), NA)

        # Remove intersection
        self$mergeCells <- self$mergeCells[!intersects]
      }

      invisible(self)

    },

    #' @description clean sheet (remove all values)
    #' @param dims dimensions
    #' @param numbers remove all numbers
    #' @param characters remove all characters
    #' @param styles remove all styles
    #' @param merged_cells remove all merged_cells
    #' @return The `wbWorksheetObject`, invisibly
    clean_sheet = function(dims = NULL, numbers = TRUE, characters = TRUE, styles = TRUE, merged_cells = TRUE) {

      cc <- self$sheet_data$cc

      if (NROW(cc) == 0) return(invisible(self))
      sel <- rep(TRUE, nrow(cc))

      if (!is.null(dims)) {
        ddims <- dims_to_dataframe(dims, fill = TRUE)
        rows <- rownames(ddims)
        cols <- colnames(ddims)

        dims <- unname(unlist(ddims))
        sel <- cc$r %in% dims
      }

      if (numbers)
        cc[sel & cc$c_t %in% c("n", ""),
          c("c_t", "v", "f", "f_t", "f_ref", "f_ca", "f_si", "is")] <- ""

      if (characters)
        cc[sel & cc$c_t %in% c("inlineStr", "s"),
          c("v", "f", "f_t", "f_ref", "f_ca", "f_si", "is")] <- ""

      if (styles)
        cc[sel, c("c_s")] <- ""

      self$sheet_data$cc <- cc

      if (merged_cells) {
        if (is.null(dims)) {
          self$mergeCells <- character(0)
        } else {
          self$unmerge_cells(cols = cols, rows = rows)
        }
      }

      invisible(self)

    },

    #' @description add page break
    #' @param row row
    #' @param col col
    #' @returns The `wbWorksheet` object
    add_page_break = function(row = NULL, col = NULL) {
      if (!xor(is.null(row), is.null(col))) {
        stop("either `row` or `col` must be NULL but not both")
      }

      if (!is.null(row)) {
        if (!is.numeric(row)) stop("`row` must be numeric")
        self$append("rowBreaks", sprintf('<brk id="%i" max="16383" man="1"/>', round(row)))
      } else if (!is.null(col)) {
        if (!is.numeric(col)) stop("`col` must be numeric")
        self$append("colBreaks", sprintf('<brk id="%i" max="1048575" man="1"/>', round(col)))
      }

      invisible(self)
    },

    #' @description add print options
    #' @param gridLines gridLines
    #' @param gridLinesSet gridLinesSet
    #' @param headings If TRUE prints row and column headings
    #' @param horizontalCentered If TRUE the page is horizontally centered
    #' @param verticalCentered If TRUE the page is vertically centered
    #' @returns The `wbWorksheet` object
    set_print_options = function(
        gridLines          = NULL,
        gridLinesSet       = NULL,
        headings           = NULL,
        horizontalCentered = NULL,
        verticalCentered   = NULL
    ) {
      self$printOptions <- xml_node_create(
        xml_name = "printOptions",
        xml_attributes = c(
          gridLines          = as_xml_attr(gridLines),
          gridLinesSet       = as_xml_attr(gridLinesSet),
          headings           = as_xml_attr(headings),
          horizontalCentered = as_xml_attr(horizontalCentered),
          verticalCentered   = as_xml_attr(verticalCentered)
        )
      )
    },

    #' @description append a field.  Intended for internal use only.  Not
    #'   guaranteed to remain a public method.
    #' @param field a field name
    #' @param value a new value
    #' @return The `wbWorksheetObject`, invisibly
    append = function(field, value = NULL) {
      self[[field]] <- c(self[[field]], value)
      invisible(self)
    },

    #' @description add sparkline
    #' @param sparklines sparkline created by `create_sparkline()`
    #' @return The `wbWorksheetObject`, invisibly
    add_sparklines = function(
      sparklines
    ) {

      private$do_append_x14(sparklines, "x14:sparklineGroup", "x14:sparklineGroups")

      invisible(self)
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
    },

    #' @description Ignore error on worksheet
    #' @param dims dims
    #' @param calculatedColumn calculatedColumn
    #' @param emptyCellReference emptyCellReference
    #' @param evalError evalError
    #' @param formula formula
    #' @param formulaRange formulaRange
    #' @param listDataValidation listDataValidation
    #' @param numberStoredAsText numberStoredAsText
    #' @param twoDigitTextYear twoDigitTextYear
    #' @param unlockedFormula unlockedFormula
    #' @return The `wbWorksheetObject`, invisibly
    ignore_error = function(
      dims               = "A1",
      calculatedColumn   = FALSE,
      emptyCellReference = FALSE,
      evalError          = FALSE,
      formula            = FALSE,
      formulaRange       = FALSE,
      listDataValidation = FALSE,
      numberStoredAsText = FALSE,
      twoDigitTextYear   = FALSE,
      unlockedFormula    = FALSE
    ) {

      dims <- unname(unlist(dims_to_dataframe(dims, fill = TRUE)))

      iEs <- self$ignoredErrors
      if (xml_node_name(iEs) == "ignoredErrors") {
        iE <- xml_node(iEs, "ignoredErrors", "ignoredError")
        iE_df <- rbindlist(xml_attr(iE, "ignoredError"))

        need <- dims[!dims %in% iE_df$sref]
        need_df <- as.data.frame(
          matrix("", ncol = ncol(iE_df), nrow = length(need)),
          stringsAsFactors = FALSE
        )
        names(need_df) <- names(iE_df)
        need_df$sqref <- need

        iE_df <- rbind(iE_df, need_df)

      } else {
        iE_df <- data.frame(sqref = dims, stringsAsFactors = FALSE)
      }

      sel <- match(dims, iE_df$sqref)

      if (calculatedColumn) {
        if (is.null(iE_df[["calculatedColumn"]])) iE_df$calculatedColumn <- ""
        iE_df[sel, "calculatedColumn"]   <- as_xml_attr(calculatedColumn)
      }

      if (emptyCellReference) {
        if (is.null(iE_df[["emptyCellReference"]])) iE_df$emptyCellReference <- ""
        iE_df[sel, "emptyCellReference"] <- as_xml_attr(emptyCellReference)
      }

      if (evalError) {
        if (is.null(iE_df[["evalError"]])) iE_df$evalError <- ""
        iE_df[sel, "evalError"]          <- as_xml_attr(evalError)
      }

      if (formula) {
        if (is.null(iE_df[["formula"]])) iE_df$formula <- ""
        iE_df[sel, "formula"]            <- as_xml_attr(formula)
      }

      if (formulaRange) {
        if (is.null(iE_df[["formulaRange"]])) iE_df$formulaRange <- ""
        iE_df[sel, "formulaRange"]       <- as_xml_attr(formulaRange)
      }

      if (listDataValidation) {
        if (is.null(iE_df[["listDataValidation"]])) iE_df$listDataValidation <- ""
        iE_df[sel, "listDataValidation"] <- as_xml_attr(listDataValidation)
      }

      if (numberStoredAsText) {
        if (is.null(iE_df[["numberStoredAsText"]])) iE_df$numberStoredAsText <- ""
        iE_df[sel, "numberStoredAsText"] <- as_xml_attr(numberStoredAsText)
      }

      if (twoDigitTextYear) {
        if (is.null(iE_df[["twoDigitTextYear"]])) iE_df$twoDigitTextYear <- ""
        iE_df[sel, "twoDigitTextYear"]   <- as_xml_attr(twoDigitTextYear)
      }

      if (unlockedFormula) {
        if (is.null(iE_df[["unlockedFormula"]])) iE_df$unlockedFormula <- ""
        iE_df[sel, "unlockedFormula"]    <- as_xml_attr(unlockedFormula)
      }

      iE <- df_to_xml("ignoredError", iE_df)

      self$ignoredErrors <- xml_node_create("ignoredErrors", iE)

      invisible(self)
    }
  ),

  ## private ----
  private = list(
    # These were commented out in the RC object -- not sure if they're needed
    cols                  = NULL,
    sheetData             = NULL,

    # @description add data_validation_lst
    # @param datavalidation datavalidation
    do_append_x14 = function(
      x,
      s_name,
      l_name
    ) {

      if (!all(xml_node_name(x) == s_name))
        stop(sprintf("all nodes must match %s. Got %s", s_name, xml_node_name(x)))

      # can have length > 1 for multiple xmlns attributes. we take this extLst,
      # inspect it, update if needed and return it
      extLst <- xml_node(self$extLst, "ext")
      is_xmlns_x14 <- grepl(pattern = "xmlns:x14", extLst)

      # different ext types have different uri ids. We support dataValidations
      # and sparklineGroups.
      uri <- ""
      # if (l_name == "x14:dataValidations") uri <- "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"
      if (l_name == "x14:sparklineGroups") uri <- "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"

      is_needed_uri <- grepl(pattern = uri, extLst, fixed = TRUE)

      # check if any <ext xmlns:x14 ...> node exists, else add it
      if (length(extLst) == 0 || !any(is_xmlns_x14) || !any(is_needed_uri)) {
        ext <- xml_node_create(
          "ext",
          xml_attributes = c("xmlns:x14" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
                             uri = uri)
        )

        # update extLst
        extLst <- c(extLst, ext)
        is_needed_uri <- c(is_needed_uri, TRUE)
      } else {
        ext <- extLst[is_needed_uri]
      }

      # check again and should be exactly one ext node
      is_xmlns_x14 <- grepl(pattern = "xmlns:x14", extLst)

      # check for l_name and add one if none is found
      if (length(xml_node(ext, "ext", l_name)) == 0) {
        ext <- xml_add_child(
          ext,
          xml_node_create(
            l_name,
            xml_attributes = c("xmlns:xm" = "http://schemas.microsoft.com/office/excel/2006/main"))
        )
      }

      # add new x to exisisting l_name
      ext <- xml_add_child(
        ext,
        level = c(l_name),
        x
      )

      # update counts for dataValidations
      # count is all matching nodes. not sure if required
      if (l_name == "x14:dataValidations") {

        outer <- xml_attr(ext, "ext")
        inner <- getXMLPtr1con(read_xml(ext))

        xdv <- grepl(l_name, inner)
        inner <- xml_attr_mod(
          inner[xdv],
          xml_attributes = c(count = as.character(length(xml_node_name(inner[xdv], l_name))))
        )

        ext <- xml_node_create("ext", xml_children = inner, xml_attributes = unlist(outer))
      }

      # update extLst and add it back to worksheet
      extLst[is_needed_uri] <- ext
      self$extLst <- extLst

      invisible(self)
    },

    data_validation = function(
      type,
      operator,
      value,
      allowBlank,
      showInputMsg,
      showErrorMsg,
      errorStyle,
      errorTitle,
      error,
      promptTitle,
      prompt,
      origin,
      sqref
    ) {

      header <- xml_node_create(
        "dataValidation",
        xml_attributes = c(
          type = type,
          operator = operator,
          allowBlank = allowBlank,
          showInputMessage = showInputMsg,
          showErrorMessage = showErrorMsg,
          sqref = sqref,
          errorStyle = errorStyle,
          errorTitle = errorTitle,
          error = error,
          promptTitle = promptTitle,
          prompt = prompt
        )
      )

      if (type == "date") {
        value <- as.integer(value) + origin
      }

      if (type == "time") {
        t <- format(value[1], "%z")
        offSet <-
          suppressWarnings(
            ifelse(substr(t, 1, 1) == "+", 1L, -1L) * (
              as.integer(substr(t, 2, 3)) + as.integer(substr(t, 4, 5)) / 60
            ) / 24
          )
        if (is.na(offSet)) {
          offSet[i] <- 0
        }

        value <- as.numeric(as.POSIXct(value)) / 86400 + origin + offSet
      }

      form <- sapply(
        seq_along(value),
        function(i) {
          sprintf("<formula%s>%s</formula%s>", i, value[i], i)
        }
      )

      self$append("dataValidations", xml_add_child(header, form))
      invisible(self)
    }
  )
)



wb_worksheet <- function() {
  wbWorksheet$new()
}

empty_cols_attr <- function(n = 0, beg, end) {
  # make make this a specific class/object?

  if (!missing(beg) && !missing(end)) {
    n_seq <- seq.int(beg, end, by = 1)
    n <- length(n_seq)
  } else {
     n_seq <- seq_len(n)
  }

  cols_attr_nams <- c("bestFit", "collapsed", "customWidth", "hidden", "max",
                      "min", "outlineLevel", "phonetic", "style", "width")

  z <- data.frame(
    matrix("", nrow = n, ncol = length(cols_attr_nams)),
    stringsAsFactors = FALSE
  )
  names(z) <- cols_attr_nams

  if (n > 0) {
    z$min <- n_seq
    z$max <- n_seq
    z$width <- "8.43"
  }

  z
}
