workbook_add_worksheet <- function(
    self,
    private,
    sheet       = next_sheet(),
    gridLines   = TRUE,
    rowColHeaders = TRUE,
    tabColour   = NULL,
    zoom        = 100,
    header      = NULL,
    footer      = NULL,
    oddHeader   = header,
    oddFooter   = footer,
    evenHeader  = header,
    evenFooter  = footer,
    firstHeader = header,
    firstFooter = footer,
    visible     = c("true", "false", "hidden", "visible", "veryhidden"),
    hasDrawing  = FALSE,
    paperSize   = getOption("openxlsx2.paperSize", default = 9),
    orientation = getOption("openxlsx2.orientation", default = "portrait"),
    hdpi        = getOption("openxlsx2.hdpi", default = getOption("openxlsx2.dpi", default = 300)),
    vdpi        = getOption("openxlsx2.vdpi", default = getOption("openxlsx2.dpi", default = 300))
) {
  visible <- tolower(as.character(visible))
  visible <- match.arg(visible)
  orientation <- match.arg(orientation, c("portrait", "landscape"))

  # set up so that a single error can be reported on fail
  fail <- FALSE
  msg <- NULL

  private$validate_new_sheet(sheet)

  if (is_waiver(sheet)) {
    if (sheet == "current_sheet") {
      stop("cannot add worksheet to current sheet")
    }

    # TODO openxlsx2.sheet.default_name is undocumented. should incorporate
    # a better check for this
    sheet <- paste0(
      getOption("openxlsx2.sheet.default_name", "Sheet "),
      length(self$sheet_names) + 1L
    )
  }

  sheet <- as.character(sheet)
  sheet_name <- replace_legal_chars(sheet)
  private$validate_new_sheet(sheet_name)

  if (!is.logical(gridLines) | length(gridLines) > 1) {
    fail <- TRUE
    msg <- c(msg, "gridLines must be a logical of length 1.")
  }

  if (!is.null(tabColour)) {
    tabColour <- validateColour(tabColour, "Invalid tabColour in add_worksheet.")
  }

  if (!is.numeric(zoom)) {
    fail <- TRUE
    msg <- c(msg, "zoom must be numeric")
  }

  # nocov start
  if (zoom < 10) {
    zoom <- 10
  } else if (zoom > 400) {
    zoom <- 400
  }
  #nocov end

  if (!is.null(oddHeader) & length(oddHeader) != 3) {
    fail <- TRUE
    msg <- c(msg, lcr("header"))
  }

  if (!is.null(oddFooter) & length(oddFooter) != 3) {
    fail <- TRUE
    msg <- c(msg, lcr("footer"))
  }

  if (!is.null(evenHeader) & length(evenHeader) != 3) {
    fail <- TRUE
    msg <- c(msg, lcr("evenHeader"))
  }

  if (!is.null(evenFooter) & length(evenFooter) != 3) {
    fail <- TRUE
    msg <- c(msg, lcr("evenFooter"))
  }

  if (!is.null(firstHeader) & length(firstHeader) != 3) {
    fail <- TRUE
    msg <- c(msg, lcr("firstHeader"))
  }

  if (!is.null(firstFooter) & length(firstFooter) != 3) {
    fail <- TRUE
    msg <- c(msg, lcr("firstFooter"))
  }

  vdpi <- as.integer(vdpi)
  hdpi <- as.integer(hdpi)

  if (is.na(vdpi)) {
    fail <- TRUE
    msg <- c(msg, "vdpi must be numeric")
  }

  if (is.na(hdpi)) {
    fail <- TRUE
    msg <- c(msg, "hdpi must be numeric")
  }

  if (fail) {
    stop(msg, call. = FALSE)
  }

  newSheetIndex <- length(self$worksheets) + 1L
  private$set_current_sheet(newSheetIndex)
  sheetId <- private$get_sheet_id_max() # checks for self$worksheet length

  # check for errors ----

  visible <- switch(
    visible,
    true = "visible",
    false = "hidden",
    veryhidden = "veryHidden",
    visible
  )

  # Order matters: if a sheet is added to a blank workbook, we add a default style. If we already have
  # sheets in the workbook, we do not add a new style. This could confuse Excel which will complain.
  # This fixes output of the example in wb_load.
  if (length(self$sheet_names) == 0) {
    # TODO this should live wherever the other default values for an empty worksheet are initialized
    empty_cellXfs <- data.frame(numFmtId = "0", fontId = "0", fillId = "0", borderId = "0", xfId = "0", stringsAsFactors = FALSE)
    self$styles_mgr$styles$cellXfs <- write_xf(empty_cellXfs)
  }

  self$append_sheets(
    sprintf(
      '<sheet name="%s" sheetId="%s" state="%s" r:id="rId%s"/>',
      sheet_name,
      sheetId,
      visible,
      newSheetIndex
    )
  )

  ## append to worksheets list
  self$append("worksheets",
              wbWorksheet$new(
                tabColour   = tabColour,
                oddHeader   = oddHeader,
                oddFooter   = oddFooter,
                evenHeader  = evenHeader,
                evenFooter  = evenFooter,
                firstHeader = firstHeader,
                firstFooter = firstFooter,
                paperSize   = paperSize,
                orientation = orientation,
                hdpi        = hdpi,
                vdpi        = vdpi,
                printGridLines = gridLines
              )
  )

  # NULL or TRUE/FALSE
  rightToLeft <- getOption("openxlsx2.rightToLeft")

  # set preselected set for sheetview
  self$worksheets[[newSheetIndex]]$set_sheetview(
    workbookViewId    = 0,
    zoomScale         = zoom,
    showGridLines     = gridLines,
    showRowColHeaders = rowColHeaders,
    tabSelected       = newSheetIndex == 1,
    rightToLeft       = rightToLeft
  )


  ## update content_tyes
  ## add a drawing.xml for the worksheet
  if (hasDrawing) {
    self$append("Content_Types", c(
      sprintf('<Override PartName="/xl/worksheets/sheet%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>', newSheetIndex),
      sprintf('<Override PartName="/xl/drawings/drawing%s.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>', newSheetIndex)
    ))
  } else {
    self$append("Content_Types",
                sprintf(
                  '<Override PartName="/xl/worksheets/sheet%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
                  newSheetIndex
                )
    )
  }

  ## Update xl/rels
  self$append("workbook.xml.rels",
              sprintf(
                '<Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%s.xml"/>',
                newSheetIndex
              )
  )

  ## create sheet.rels to simplify id assignment
  new_drawings_idx <- length(self$drawings) + 1
  self$drawings[[new_drawings_idx]]      <- ""
  self$drawings_rels[[new_drawings_idx]] <- ""

  self$worksheets_rels[[newSheetIndex]]  <- genBaseSheetRels(newSheetIndex)
  self$vml_rels[[newSheetIndex]]         <- list()
  self$vml[[newSheetIndex]]              <- list()
  self$isChartSheet[[newSheetIndex]]     <- FALSE
  self$comments[[newSheetIndex]]         <- list()
  self$threadComments[[newSheetIndex]]   <- list()

  self$append("sheetOrder", as.integer(newSheetIndex))
  private$set_single_sheet_name(newSheetIndex, sheet_name, sheet)

  invisible(self)

}
