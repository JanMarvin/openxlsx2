
Worksheet <- setRefClass(
  "Worksheet",
  fields = c(
    "sheetPr" = "character",
    "dimension" = "character",
    "sheetViews" = "character",
    "sheetFormatPr" = "character",

    "sheet_data" = "SheetData",
    "rows_attr" = "ANY",
    "cols_attr" = "ANY",
    "cols" = "ANY",

    "autoFilter" = "character",
    "mergeCells" = "ANY",
    "conditionalFormatting" = "character",
    "dataValidations" = "ANY",
    "dataValidationsLst" = "character",

    "freezePane" = "character",
    "hyperlinks" = "ANY",

    "sheetProtection" = "character",
    "pageMargins" = "character",
    "pageSetup" = "character",
    "headerFooter" = "ANY",
    "rowBreaks" = "character",
    "colBreaks" = "character",
    "drawing" = "character",
    "legacyDrawing" = "character",
    "legacyDrawingHF" = "character",
    "oleObjects" = "character",
    "tableParts" = "character",
    "extLst" = "character"
  ),

  methods = list(
    initialize = function(
      showGridLines = TRUE,
      tabSelected = FALSE,
      tabColour = NULL,
      zoom = 100,

      oddHeader = NULL,
      oddFooter = NULL,
      evenHeader = NULL,
      evenFooter = NULL,
      firstHeader = NULL,
      firstFooter = NULL,

      paperSize = 9,
      orientation = "portrait",
      hdpi = 300,
      vdpi = 300
    ) {
      if (!is.null(tabColour)) {
        tabColour <- sprintf('<sheetPr><tabColor rgb="%s"/></sheetPr>', tabColour)
      } else {
        tabColour <- character(0)
      }

      if (zoom < 10) {
        zoom <- 10
      } else if (zoom > 400) {
        zoom <- 400
      }

      hf <- list(
        oddHeader = naToNULLList(oddHeader),
        oddFooter = naToNULLList(oddFooter),
        evenHeader = naToNULLList(evenHeader),
        evenFooter = naToNULLList(evenFooter),
        firstHeader = naToNULLList(firstHeader),
        firstFooter = naToNULLList(firstFooter)
      )

      if (all(sapply(hf, length) == 0)) {
        hf <- list()
      }

      ## list of all possible children
      .self$sheetPr <- tabColour
      .self$dimension <- '<dimension ref="A1"/>'
      .self$sheetViews <- sprintf('<sheetViews><sheetView workbookViewId="0" zoomScale="%s" showGridLines="%s" tabSelected="%s"/></sheetViews>', as.integer(zoom), as.integer(showGridLines), as.integer(tabSelected))
      .self$sheetFormatPr <- '<sheetFormatPr defaultRowHeight="15.0" baseColWidth="10"/>'
      .self$cols_attr <- character(0)

      .self$autoFilter <- character(0)
      .self$mergeCells <- character(0)
      .self$conditionalFormatting <- character(0)
      .self$dataValidations <- NULL
      .self$dataValidationsLst <- character(0)
      .self$hyperlinks <- list()
      .self$pageMargins <- '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
      .self$pageSetup <- sprintf('<pageSetup paperSize="%s" orientation="%s" horizontalDpi="%s" verticalDpi="%s" r:id="rId2"/>', paperSize, orientation, hdpi, vdpi) ## will always be 2
      .self$headerFooter <- hf
      .self$rowBreaks <- character(0)
      .self$colBreaks <- character(0)
      .self$drawing <- '<drawing r:id=\"rId1\"/>' ## will always be 1
      .self$legacyDrawing <- character(0)
      .self$legacyDrawingHF <- character(0)
      .self$oleObjects <- character(0)
      .self$tableParts <- character(0)
      .self$extLst <- character(0)

      .self$freezePane <- character(0)

      .self$sheet_data <- SheetData$new()
    },

    get_prior_sheet_data = function() {
      xml <- '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3">'

      if (length(.self$sheetPr) > 0) {
        tmp <- .self$sheetPr
        if (!any(grepl("<sheetPr>", tmp, fixed = TRUE))) {
          tmp <- paste0("<sheetPr>", paste(tmp, collapse = ""), "</sheetPr>")
        }

        xml <- paste(xml, tmp, collapse = "")
      }

      if (length(.self$dimension) > 0) {
        xml <- paste(xml, .self$dimension, collapse = "")
      }

      ## sheetViews handled here
      if (length(.self$freezePane) > 0) {
        xml <- paste(xml, gsub("/></sheetViews>", paste0(">", .self$freezePane, "</sheetView></sheetViews>"), .self$sheetViews, fixed = TRUE), collapse = "")
      } else if (length(.self$sheetViews) > 0) {
        xml <- paste(xml, .self$sheetViews, collapse = "")
      }

      if (length(.self$sheetFormatPr) > 0) {
        xml <- paste(xml, .self$sheetFormatPr, collapse = "")
      }

      # TODO write function for new cols_attr
      if (length(.self$cols_attr) > 0) {
        # xml <- paste0(xml, paste(c("<cols>", setXMLcols(cols_attr), "</cols>"), collapse = ""))
        xml <- paste0(xml, paste(c("<cols>", .self$cols_attr, "</cols>"), collapse = ""))
      }


      return(xml)
    },

    get_post_sheet_data = function() {
      # TODO use paste_c()
      xml <- ""

      if (length(.self$sheetProtection) > 0) {
        xml <- paste0(xml, .self$sheetProtection, collapse = "")
      }

      if (length(.self$autoFilter) > 0) {
        xml <- paste0(xml, .self$autoFilter, collapse = "")
      }

      if (length(.self$mergeCells) > 0) {
        xml <- paste0(xml, paste0(sprintf('<mergeCells count="%s">', length(.self$mergeCells)), pxml(.self$mergeCells), "</mergeCells>"), collapse = "")
      }



      if (length(.self$conditionalFormatting) > 0) {
        nms <- names(.self$conditionalFormatting)
        xml <- paste0(xml,
          paste(
            sapply(unique(nms), function(x) {
              paste0(
                sprintf('<conditionalFormatting sqref="%s">', x),
                pxml(.self$conditionalFormatting[nms == x]),
                "</conditionalFormatting>"
              )
            }),
            collapse = ""
          ),
          collapse = ""
        )
      }


      if (length(.self$dataValidations) > 0) {
        xml <- paste0(xml, paste0(sprintf('<dataValidations count="%s">', length(.self$dataValidations)), pxml(.self$dataValidations), "</dataValidations>"))
      }



      if (length(.self$hyperlinks) > 0) {
        # TODO use seq_along?  seq_len?
        h_inds <- paste0(1:length(.self$hyperlinks), "h")
        xml <- paste(xml, paste("<hyperlinks>", paste(sapply(1:length(h_inds), function(i) .self$hyperlinks[[i]]$to_xml(h_inds[i])), collapse = ""), "</hyperlinks>"), collapse = "")
      }


      if (length(.self$pageMargins) > 0) {
        xml <- paste0(xml, .self$pageMargins, collapse = "")
      }

      if (length(.self$pageSetup) > 0) {
        xml <- paste0(xml, .self$pageSetup, collapse = "")
      }

      if (length(.self$headerFooter) > 0) {
        xml <- paste0(xml, genHeaderFooterNode(.self$headerFooter), collapse = "")
      }


      ## rowBreaks and colBreaks
      if (length(.self$rowBreaks) > 0) {
        # TODO just save length as n object
        xml <- paste0(xml,
          paste0(sprintf('<rowBreaks count="%s" manualBreakCount="%s">', length(.self$rowBreaks), length(.self$rowBreaks)), paste(.self$rowBreaks, collapse = ""), "</rowBreaks>"),
          collapse = ""
        )
      }


      if (length(.self$colBreaks) > 0) {
        xml <- paste0(xml,
          paste0(sprintf('<colBreaks count="%s" manualBreakCount="%s">', length(.self$colBreaks), length(.self$colBreaks)), paste(.self$colBreaks, collapse = ""), "</colBreaks>"),
          collapse = ""
        )
      }

      if (length(.self$drawing) > 0) {
        xml <- paste0(xml, .self$drawing, collapse = "")
      }

      if (length(.self$legacyDrawing) > 0) {
        xml <- paste0(xml, .self$legacyDrawing, collapse = "")
      }

      if (length(.self$legacyDrawingHF) > 0) {
        xml <- paste0(xml, .self$legacyDrawingHF, collapse = "")
      }

      if (length(.self$oleObjects) > 0) {
        xml <- paste0(xml, .self$oleObjects, collapse = "")
      }

      if (length(.self$tableParts) > 0) {
        xml <- paste0(xml,
          paste0(sprintf('<tableParts count="%s">', length(.self$tableParts)), pxml(.self$tableParts), "</tableParts>"),
          collapse = ""
        )
      }


      if (length(.self$dataValidationsLst) > 0) {
        dataValidationsLst_xml <- paste0(sprintf('<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataValidations count="%s" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">', length(dataValidationsLst)),
          paste0(pxml(.self$dataValidationsLst), "</x14:dataValidations></ext>"),
          collapse = ""
        )
      } else {
        dataValidationsLst_xml <- character(0)
      }


      if (length(.self$extLst) > 0 || length(.self$dataValidationsLst) > 0) {
        xml <- paste0(xml, sprintf("<extLst>%s</extLst>", paste0(pxml(.self$extLst), dataValidationsLst_xml)))
      }

      xml <- paste0(xml, "</worksheet>")

      return(xml)
    },

    order_sheetdata = function() {
      if (.self$sheet_data$n_elements == 0) {
        return(invisible(.self))
      }

      if (.self$sheet_data$data_count > 1) {
        ord <- order(.self$sheet_data$rows, .self$sheet_data$cols, method = "radix", na.last = TRUE)
        .self$sheet_data$rows <- .self$sheet_data$rows[ord]
        .self$sheet_data$cols <- .self$sheet_data$cols[ord]
        .self$sheet_data$t <- .self$sheet_data$t[ord]
        .self$sheet_data$v <- .self$sheet_data$v[ord]
        .self$sheet_data$f <- .self$sheet_data$f[ord]

        .self$sheet_data$style_id <- .self$sheet_data$style_id[ord]

        .self$sheet_data$data_count <- 1L

        dm1 <- paste0(int_2_cell_ref(cols = .self$sheet_data$cols[1]), .self$sheet_data$rows[1])
        dm2 <- paste0(int_2_cell_ref(cols = .self$sheet_data$cols[.self$sheet_data$n_elements]), .self$sheet_data$rows[sheet_data$n_elements])

        if (length(dm1) == 1 & length(dm2) != 1) {
          if (!is.na(dm1) & !is.na(dm2) & dm1 != "NA" & dm2 != "NA") {
            .self$dimension <- sprintf("<dimension ref=\"%s:%s\"/>", dm1, dm2)
          }
        }
      }


      invisible(.self)
    }
  )
)

new_worksheet <- function() {
  Worksheet$new()
}

