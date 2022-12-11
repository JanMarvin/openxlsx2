page_setup = function(
    self,
    private,
    sheet          = current_sheet(),
    orientation    = NULL,
    scale          = 100,
    left           = 0.7,
    right          = 0.7,
    top            = 0.75,
    bottom         = 0.75,
    header         = 0.3,
    footer         = 0.3,
    fitToWidth     = FALSE,
    fitToHeight    = FALSE,
    paperSize      = NULL,
    printTitleRows = NULL,
    printTitleCols = NULL,
    summaryRow     = NULL,
    summaryCol     = NULL
) {

  sheet <- private$get_sheet_index(sheet)
  xml <- self$worksheets[[sheet]]$pageSetup

  if (!is.null(orientation)) {
    orientation <- tolower(orientation)
    if (!orientation %in% c("portrait", "landscape")) stop("Invalid page orientation.")
  } else {
    # if length(xml) == 1 then use if () {} else {}
    orientation <- ifelse(grepl("landscape", xml), "landscape", "portrait") ## get existing
  }

  if ((scale < 10) || (scale > 400)) {
    stop("Scale must be between 10 and 400.")
  }

  if (!is.null(paperSize)) {
    paperSizes <- 1:68
    paperSizes <- paperSizes[!paperSizes %in% 48:49]
    if (!paperSize %in% paperSizes) {
      stop("paperSize must be an integer in range [1, 68]. See ?ws_page_setup details.")
    }
    paperSize <- as.integer(paperSize)
  } else {
    paperSize <- regmatches(xml, regexpr('(?<=paperSize=")[0-9]+', xml, perl = TRUE)) ## get existing
  }

  ## Keep defaults on orientation, hdpi, vdpi, paperSize ----
  hdpi <- regmatches(xml, regexpr('(?<=horizontalDpi=")[0-9]+', xml, perl = TRUE))
  vdpi <- regmatches(xml, regexpr('(?<=verticalDpi=")[0-9]+', xml, perl = TRUE))

  ## Update ----
  self$worksheets[[sheet]]$pageSetup <- sprintf(
    '<pageSetup paperSize="%s" orientation="%s" scale = "%s" fitToWidth="%s" fitToHeight="%s" horizontalDpi="%s" verticalDpi="%s"/>',
    paperSize, orientation, scale, as.integer(fitToWidth), as.integer(fitToHeight), hdpi, vdpi
  )

  if (fitToHeight || fitToWidth) {
    self$worksheets[[sheet]]$sheetPr <- unique(c(self$worksheets[[sheet]]$sheetPr, '<pageSetupPr fitToPage="1"/>'))
  }

  self$worksheets[[sheet]]$pageMargins <-
    sprintf('<pageMargins left="%s" right="%s" top="%s" bottom="%s" header="%s" footer="%s"/>', left, right, top, bottom, header, footer)

  validRow <- function(summaryRow) {
    return(tolower(summaryRow) %in% c("above", "below"))
  }
  validCol <- function(summaryCol) {
    return(tolower(summaryCol) %in% c("left", "right"))
  }

  outlinepr <- ""

  if (!is.null(summaryRow)) {

    if (!validRow(summaryRow)) {
      stop("Invalid \`summaryRow\` option. Must be one of \"Above\" or \"Below\".")
    } else if (tolower(summaryRow) == "above") {
      outlinepr <- ' summaryBelow=\"0\"'
    } else {
      outlinepr <- ' summaryBelow=\"1\"'
    }
  }

  if (!is.null(summaryCol)) {

    if (!validCol(summaryCol)) {
      stop("Invalid \`summaryCol\` option. Must be one of \"Left\" or \"Right\".")
    } else if (tolower(summaryCol) == "left") {
      outlinepr <- paste0(outlinepr, ' summaryRight=\"0\"')
    } else {
      outlinepr <- paste0(outlinepr, ' summaryRight=\"1\"')
    }
  }

  if (!stri_isempty(outlinepr)) {
    self$worksheets[[sheet]]$sheetPr <- unique(c(self$worksheets[[sheet]]$sheetPr, paste0("<outlinePr", outlinepr, "/>")))
  }

  ## print Titles ----
  if (!is.null(printTitleRows) && is.null(printTitleCols)) {
    if (!is.numeric(printTitleRows)) {
      stop("printTitleRows must be numeric.")
    }

    private$create_named_region(
      ref1 = paste0("$", min(printTitleRows)),
      ref2 = paste0("$", max(printTitleRows)),
      name = "_xlnm.Print_Titles",
      sheet = self$get_sheet_names()[[sheet]],
      localSheetId = sheet - 1L
    )
  } else if (!is.null(printTitleCols) && is.null(printTitleRows)) {
    if (!is.numeric(printTitleCols)) {
      stop("printTitleCols must be numeric.")
    }

    cols <- int2col(range(printTitleCols))
    private$create_named_region(
      ref1 = paste0("$", cols[1]),
      ref2 = paste0("$", cols[2]),
      name = "_xlnm.Print_Titles",
      sheet = self$get_sheet_names()[[sheet]],
      localSheetId = sheet - 1L
    )
  } else if (!is.null(printTitleCols) && !is.null(printTitleRows)) {
    if (!is.numeric(printTitleRows)) {
      stop("printTitleRows must be numeric.")
    }

    if (!is.numeric(printTitleCols)) {
      stop("printTitleCols must be numeric.")
    }

    cols <- int2col(range(printTitleCols))
    rows <- range(printTitleRows)

    cols <- paste(paste0("$", cols[1]), paste0("$", cols[2]), sep = ":")
    rows <- paste(paste0("$", rows[1]), paste0("$", rows[2]), sep = ":")
    localSheetId <- sheet - 1L
    sheet <- self$get_sheet_names()[[sheet]]

    self$workbook$definedNames <- c(
      self$workbook$definedNames,
      sprintf('<definedName name="_xlnm.Print_Titles" localSheetId="%s">\'%s\'!%s,\'%s\'!%s</definedName>', localSheetId, sheet, cols, sheet, rows)
    )

  }

  invisible(self)
}
