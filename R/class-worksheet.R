
Worksheet <- setRefClass(
  "Worksheet",
  fields = c(
    "sheetPr"       = "character",
    "dimension"     = "character",
    "sheetViews"    = "character",
    "sheetFormatPr" = "character",

    "sheet_data" = "SheetData",
    "rows_attr"  = "ANY",
    "cols_attr"  = "ANY",
    "cols"       = "ANY",

    "autoFilter"            = "character",
    "mergeCells"            = "ANY",
    "conditionalFormatting" = "character",
    "dataValidations"       = "ANY",
    "dataValidationsLst"    = "character",

    "freezePane" = "character",
    "hyperlinks" = "ANY",

    "sheetProtection" = "character",
    "pageMargins"     = "character",
    "pageSetup"       = "character",
    "headerFooter"    = "ANY",
    "rowBreaks"       = "character",
    "colBreaks"       = "character",
    "drawing"         = "character",
    "legacyDrawing"   = "character",
    "legacyDrawingHF" = "character",
    "oleObjects"      = "character",
    "tableParts"      = "character",
    "extLst"          = "character"
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
        tabColour <- character()
      }

      if (zoom < 10) {
        zoom <- 10
      } else if (zoom > 400) {
        zoom <- 400
      }

      hf <- list(
        oddHeader   = na_to_null(oddHeader),
        oddFooter   = na_to_null(oddFooter),
        evenHeader  = na_to_null(evenHeader),
        evenFooter  = na_to_null(evenFooter),
        firstHeader = na_to_null(firstHeader),
        firstFooter = na_to_null(firstFooter)
      )

      if (all(sapply(hf, length) == 0)) {
        hf <- list()
      }

      ## list of all possible children
      .self$sheetPr <- tabColour
      .self$dimension <- '<dimension ref="A1"/>'
      .self$sheetViews <- sprintf('<sheetViews><sheetView workbookViewId="0" zoomScale="%s" showGridLines="%s" tabSelected="%s"/></sheetViews>', as.integer(zoom), as.integer(showGridLines), as.integer(tabSelected))
      .self$sheetFormatPr <- '<sheetFormatPr defaultRowHeight="15.0" baseColWidth="10"/>'
      .self$cols_attr <- character()

      .self$autoFilter <- character()
      .self$mergeCells <- character()
      .self$conditionalFormatting <- character()
      .self$dataValidations <- NULL
      .self$dataValidationsLst <- character()
      .self$hyperlinks <- list()
      .self$pageMargins <- '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
      .self$pageSetup <- sprintf('<pageSetup paperSize="%s" orientation="%s" horizontalDpi="%s" verticalDpi="%s" r:id="rId2"/>', paperSize, orientation, hdpi, vdpi) ## will always be 2
      .self$headerFooter <- hf
      .self$rowBreaks <- character()
      .self$colBreaks <- character()
      .self$drawing <- '<drawing r:id=\"rId1\"/>' ## will always be 1
      .self$legacyDrawing <- character()
      .self$legacyDrawingHF <- character()
      .self$oleObjects <- character()
      .self$tableParts <- character()
      .self$extLst <- character()

      .self$freezePane <- character()

      .self$sheet_data <- SheetData$new()
    },

    get_prior_sheet_data = function() {
      paste_c(
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3">',

        # sheetPr
        if (length(.self$sheetPr) && !any(grepl("<sheetPr>", .self$sheetPr, fixed = TRUE))) {
          paste0("<sheetPr>", paste(.self$sheetPr, collapse = ""), "</sheetPr>")
        } else {
          .self$sheetPr
        },

        .self$dimension,

        if (length(.self$freezePane)) {
          gsub("/></sheetViews>", paste0(">", .self$freezePane, "</sheetView></sheetViews>"), .self$sheetViews, fixed = TRUE)
        } else if (length(.self$sheetViews)) {
          .self$sheetViews
        },
        .self$sheetFormatPr,
        # cols_attr
        # is this fine if it's just <cols></cols>?
        if (length(.self$cols_attr)) {
          paste(c("<cols>", .self$cols_attr, "</cols>"), collapse = "")
        },
        sep = " "
      )
    },

    get_post_sheet_data = function() {
      paste_c(
        .self$sheetProtection,
        .self$autoFilter,

        # mergeCells
        if (length(.self$mergeCells)) {
          paste0(
            sprintf('<mergeCells count="%i">', length(.self$mergeCells)),
            pxml(.self$mergeCells),
            "</mergeCells>"
          )
        },

        # conditionalFormatting
        if (length(.self$conditionalFormatting)) {
          nms <- names(.self$conditionalFormatting)
          paste(
            vapply(
              unique(nms),
              function(i) {
                paste0(
                  sprintf('<conditionalFormatting sqref="%s">', i),
                  pxml(.self$conditionalFormatting[nms == i]),
                  "</conditionalFormatting>"
                )
              },
              NA_character_
            ),
            collapse = ""
          )
        },

        # dataValidations
        if (length(.self$dataValidations)) {
          paste0(
            sprintf('<dataValidations count="%i">', length(.self$dataValidations)),
            pxml(.self$dataValidations),
            "</dataValidations>"
          )
        },

        # hyperlinks
        if (n <- length(.self$hyperlinks)) {
          h_inds <- paste0(seq_len(n), "h")
          paste(
            "<hyperlinks>",
            paste(
              vapply(
                seq_along(h_inds),
                function(i)  {
                  .self$hyperlinks[[i]]$to_xml(h_inds[i])
                },
                NA_character_
              ),
              collapse = ""
            ),
            "</hyperlinks>"
          )
        },

        .self$pageMargins,
        .self$pageSetup,

        # headerFooter
        # should return NULL when !length(.self$headerFooter)
        genHeaderFooterNode(.self$headerFooter),

        # rowBreaks
        if (n <- length(.self$rowBreaks)) {
          paste0(
            sprintf('<rowBreaks count="%i" manualBreakCount="%i">', n, n),
            paste(.self$rowBreaks, collapse = ""),
            "</rowBreaks>"
          )
        },

        # colBreaks
        if (n <- length(.self$colBreaks)) {
          paste0(
            sprintf('<colBreaks count="%i" manualBreakCount="%i">', n, n),
            paste(.self$colBreaks, collapse = ""),
            "</colBreaks>"
          )
        },

        .self$drawing,
        .self$legacyDrawing,
        .self$legacyDrawingHF,
        .self$oleObjects,

        # tableParts
        if (n <- length(.self$tableParts)) {
          paste0(sprintf('<tableParts count="%i">', n), pxml(.self$tableParts), "</tableParts>")
        },

        # extLst, dataValidationsLst
        # parenthese or R gets confused with the ||
        if ((length(.self$extLst)) || (n <- length(.self$dataValidationsLst))) {
          sprintf(
            "<extLst>%s</extLst>",
            paste0(
              pxml(.self$extLst),
              # dataValidationsLst_xml
              if (n) {
                paste0(
                  sprintf('<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataValidations count="%i" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">', n),
                  paste0(pxml(.self$dataValidationsLst), "</x14:dataValidations></ext>"),
                  collapse = ""
                )
              }
            )
          )
        },

        "</worksheet>",

        # end
        sep = " "
      )

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
