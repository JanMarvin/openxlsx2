
# class -------------------------------------------------------------------

#' R6 class for a Workbook Worksheet
#'
#' A Worksheet
#'
#' @export
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

    #' @field dataValidationsLst dataValidationsLst
    dataValidationsLst = character(),

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
    #' @param gridLines gridLines
    #' @param tabSelected tabSelected
    #' @param tabColour tabColour
    #' @param zoom zoom
    #' @param oddHeader oddHeader
    #' @param oddFooter oddFooter
    #' @param evenHeader evenHeader
    #' @param evenFooter evenFooter
    #' @param firstHeader firstHeader
    #' @param firstFooter firstFooter
    #' @param paperSize paperSize
    #' @param orientation orientation
    #' @param hdpi hdpi
    #' @param vdpi vdpi
    #' @return a `wbWorksheet` object
    initialize = function(
      gridLines   = TRUE,
      tabSelected = FALSE,
      tabColour   = NULL,
      zoom        = 100,
      oddHeader   = NULL,
      oddFooter   = NULL,
      evenHeader  = NULL,
      evenFooter  = NULL,
      firstHeader = NULL,
      firstFooter = NULL,
      paperSize   = 9,
      orientation = "portrait",
      hdpi        = 300,
      vdpi        = 300
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

      if (all(lengths(hf) == 0)) {
        hf <- list()
      }

      ## list of all possible children
      self$sheetPr               <- tabColour
      self$dimension             <- '<dimension ref="A1"/>'
      self$sheetViews            <- sprintf('<sheetViews><sheetView workbookViewId="0" zoomScale="%s" gridLines="%s" tabSelected="%s"/></sheetViews>', as.integer(zoom), as.integer(gridLines), as.integer(tabSelected))
      self$sheetFormatPr         <- '<sheetFormatPr defaultRowHeight="15.0"/>'
      self$cols_attr             <- character()
      self$autoFilter            <- character()
      self$mergeCells            <- character()
      self$conditionalFormatting <- character()
      self$dataValidations       <- NULL
      self$dataValidationsLst    <- character()
      self$hyperlinks            <- list()
      self$pageMargins           <- '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
      self$pageSetup             <- sprintf('<ws_page_setup paperSize="%s" orientation="%s" horizontalDpi="%s" verticalDpi="%s" r:id="rId2"/>', paperSize, orientation, hdpi, vdpi) ## will always be 2
      self$headerFooter          <- hf
      self$rowBreaks             <- character()
      self$colBreaks             <- character()
      self$drawing               <- '<drawing r:id=\"rId1\"/>' ## will always be 1
      self$legacyDrawing         <- character()
      self$legacyDrawingHF       <- character()
      self$oleObjects            <- character()
      self$tableParts            <- character()
      self$extLst                <- character()
      self$freezePane            <- character()
      self$sheet_data            <- wbSheetData$new()

      invisible(self)
    },

    #' @description
    #' Get prior sheet data
    #' @return A character vector of xml
    get_prior_sheet_data = function() {
      paste_c(
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3">',

        # sheetPr
        if (length(self$sheetPr) && !any(grepl("<sheetPr>", self$sheetPr, fixed = TRUE))) {
          paste0("<sheetPr>", paste(self$sheetPr, collapse = ""), "</sheetPr>")
        } else {
          self$sheetPr
        },

        self$dimension,

        if (length(self$freezePane)) {
          gsub("/></sheetViews>", paste0(">", self$freezePane, "</sheetView></sheetViews>"), self$sheetViews, fixed = TRUE)
        } else if (length(self$sheetViews)) {
          self$sheetViews
        },
        self$sheetFormatPr,
        # cols_attr
        # is this fine if it's just <cols></cols>?
        if (length(self$cols_attr)) {
          paste(c("<cols>", self$cols_attr, "</cols>"), collapse = "")
        },
        sep = ""
      )
    },

    #' @description
    #' Get post sheet data
    #' @return A character vector of xml
    get_post_sheet_data = function() {
      paste_c(
        self$sheetProtection,
        self$autoFilter,

        # mergeCells
        if (length(self$mergeCells)) {
          paste0(
            sprintf('<mergeCells count="%i">', length(self$mergeCells)),
            pxml(self$mergeCells),
            "</mergeCells>"
          )
        },

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

        self$drawing,
        self$legacyDrawing,
        self$legacyDrawingHF,
        self$oleObjects,

        # tableParts
        if (n <- length(self$tableParts)) {
          paste0(sprintf('<tableParts count="%i">', n), pxml(self$tableParts), "</tableParts>")
        },

        # extLst, dataValidationsLst
        # parenthese or R gets confused with the ||
        if ((length(self$extLst)) || (length(self$dataValidationsLst))) {
          sprintf(
            "<extLst>%s</extLst>",
            paste0(
              pxml(self$extLst),
              # dataValidationsLst_xml
              if (length(self$dataValidationsLst)) {
                paste0(
                  sprintf('<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataValidations count="%i" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">', n),
                  paste0(pxml(self$dataValidationsLst), "</x14:dataValidations></ext>"),
                  collapse = ""
                )
              }
            )
          )
        },

        "</worksheet>",

        # end
        sep = ""
      )

    },

    #' @description
    #' order sheet data
    #' @return The `wbWorksheetObject`, invisibly
    order_sheetdata = function() {
      if (self$sheet_data$n_elements == 0) {
        return(invisible(self))
      }

      if (self$sheet_data$data_count > 1) {
        ord <- order(self$sheet_data$rows, self$sheet_data$cols, method = "radix", na.last = TRUE)
        self$sheet_data$rows <- self$sheet_data$rows[ord]
        self$sheet_data$cols <- self$sheet_data$cols[ord]
        self$sheet_data$t <- self$sheet_data$t[ord]
        self$sheet_data$v <- self$sheet_data$v[ord]
        self$sheet_data$f <- self$sheet_data$f[ord]

        self$sheet_data$style_id <- self$sheet_data$style_id[ord]

        self$sheet_data$data_count <- 1L

        dm1 <- paste0(int_to_col(self$sheet_data$cols[1]), self$sheet_data$rows[1])
        dm2 <- paste0(int_to_col(self$sheet_data$cols[self$sheet_data$n_elements]), self$sheet_data$rows[sheet_data$n_elements])

        if (length(dm1) == 1 & length(dm2) != 1) {
          if (!is.na(dm1) & !is.na(dm2) & dm1 != "NA" & dm2 != "NA") {
            self$dimension <- sprintf("<dimension ref=\"%s:%s\"/>", dm1, dm2)
          }
        }
      }

      invisible(self)
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
        return (empty_cols_attr())

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
        z <- col_df[i,]
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
      col_df <- col_df[col_df$string != "",]
      col_df$string <- NULL

      # order and return
      col_df <- col_df[order(col_df$min),]
      col_df$min <- as.character(col_df$min)
      col_df$max <- as.character(col_df$max)

      # assign as xml-nodes
      self$cols_attr <- df_to_xml("col", col_df)

      invisible(self)
    },


    #' @description clean sheet (remove all values)
    #' @param numbers remove all numbers
    #' @param characters remove all characters
    #' @param styles remove all styles
    #' @param merged_cells remove all merged_cells
    #' @return The `wbWorksheetObject`, invisibly
    clean_sheet = function(numbers = FALSE, characters = FALSE, styles = FALSE, merged_cells = FALSE) {

      cc <- self$sheet_data$cc

      if (numbers)
        cc[cc$c_t %in% c("n", "_openxlsx_NA_"), # imported values might be _NA_
          c("c_t", "v", "f", "f_t", "f_ref", "f_ca", "f_si", "is")] <- "_openxlsx_NA_"

      if (characters)
        cc[cc$c_t %in% c("inlineStr", "s"),
          c("v", "f", "f_t", "f_ref", "f_ca", "f_si", "is")] <- ""

      if (styles)
        cc[c("c_s")] <- "_openxlsx_NA_"

      self$sheet_data$cc <- cc

      if (merged_cells)
        self$mergeCells <- character(0)

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

      self
    },

    #' @description append a field.  Intended for internal use only.  Not
    #'   guaranteed to remain a public method.
    #' @param field a field name
    #' @param value a new value
    append = function(field, value = NULL) {
      self[[field]] <- c(self[[field]], value)
      self
    }
  ),

  ## private ----
  private = list(
    # These were commented out in the RC object -- not sure if they're needed
    cols                  = NULL,
    sheetData             = NULL
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
    z$width <- "8.43" # default width in ms365
  }

  z
}
