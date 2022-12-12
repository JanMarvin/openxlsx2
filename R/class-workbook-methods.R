
workbook_append <- function(self, private, field, value) {
  self[[field]] <- c(self[[field]], value)
  invisible(self)
}

workbook_append_sheets <- function(self, private, value) {
  self$workbook$sheets <- c(self$workbook$sheets, value)
  invisible(self)
}

workbook_add_chartsheet <- function(self, private, sheet, tabColour, zoom) {
  # TODO private$new_sheet_index()?
  newSheetIndex <- length(self$worksheets) + 1L
  sheetId <- private$get_sheet_id_max() # checks for length of worksheets

  ##  Add sheet to workbook.xml
  self$append_sheets(
    sprintf(
      '<sheet name="%s" sheetId="%s" r:id="rId%s"/>',
      sheet,
      sheetId,
      newSheetIndex
    )
  )

  ## append to worksheets list
  self$append("worksheets",
              wbChartSheet$new(tabColour = tabColour)
  )


  # nocov start
  if (zoom < 10) {
    zoom <- 10
  } else if (zoom > 400) {
    zoom <- 400
  }
  #nocov end

  self$worksheets[[newSheetIndex]]$set_sheetview(
    workbookViewId = 0,
    zoomScale      = zoom,
    tabSelected    = newSheetIndex == 1
  )

  self$append("sheet_names", sheet)

  ## update content_tyes
  self$append("Content_Types",
              sprintf(
                '<Override PartName="/xl/chartsheets/sheet%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>',
                newSheetIndex
              )
  )

  ## Update xl/rels
  self$append("workbook.xml.rels",
              sprintf(
                '<Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet%s.xml"/>',
                newSheetIndex
              )
  )

  ## add a drawing.xml for the worksheet
  self$append("Content_Types",
              sprintf(
                '<Override PartName="/xl/drawings/drawing%s.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>',
                newSheetIndex
              )
  )

  ## create sheet.rels to simplify id assignment
  new_drawings_idx <- length(self$drawings) + 1
  self$drawings[[new_drawings_idx]]      <- ""
  self$drawings_rels[[new_drawings_idx]] <- ""

  self$worksheets_rels[[newSheetIndex]]  <- genBaseSheetRels(newSheetIndex)
  self$isChartSheet[[newSheetIndex]]     <- TRUE
  self$vml_rels[[newSheetIndex]]         <- list()
  self$vml[[newSheetIndex]]              <- list()
  self$append("sheetOrder", newSheetIndex)

  # invisible(newSheetIndex)
  invisible(self)
}

workbook_add_style <- function(self, private, style = NULL, style_name = NULL) {
  assert_class(style, "character")

  if (is.null(style_name)) {
    style_name <- deparse1(substitute(style, parent.frame()))
  } else {
    assert_class(style_name, "character")
  }

  self$styles_mgr$add(style, style_name)

  invisible(self)
}

workbook_open <- function(self, private, interactive = NA) {
  xl_open(self, interactive = interactive)
  invisible(self)
}

workbook_get_base_font <- function(self, private) {
  baseFont <- self$styles_mgr$styles$fonts[[1]]

  sz     <- unlist(xml_attr(baseFont, "font", "sz"))
  colour <- unlist(xml_attr(baseFont, "font", "color"))
  name   <- unlist(xml_attr(baseFont, "font", "name"))

  if (length(sz[[1]]) == 0) {
    sz <- list("val" = "11")
  } else {
    sz <- as.list(sz)
  }

  if (length(colour[[1]]) == 0) {
    colour <- list("rgb" = "#000000")
  } else {
    colour <- as.list(colour)
  }

  if (length(name[[1]]) == 0) {
    name <- list("val" = "Calibri")
  } else {
    name <- as.list(name)
  }

  list(
    size   = sz,
    colour = colour,
    name   = name
  )
}

workbook_set_base_font <- function(
    self,
    private,
    fontSize = 11,
    fontColour = wb_colour(theme = "1"),
    fontName = "Calibri"
) {
  if (fontSize < 0) {
    stop("Invalid fontSize")
  }

  if (is.character(fontColour) && is.null(names(fontColour))) {
    fontColour <- wb_colour(fontColour)
  }

  self$styles_mgr$styles$fonts[[1]] <- create_font(
    sz = as.character(fontSize),
    color = fontColour,
    name = fontName
  )

  invisible(self)
}

workbook_set_bookview <- function(
    self,
    private,
    activeTab              = NULL,
    autoFilterDateGrouping = NULL,
    firstSheet             = NULL,
    minimized              = NULL,
    showHorizontalScroll   = NULL,
    showSheetTabs          = NULL,
    showVerticalScroll     = NULL,
    tabRatio               = NULL,
    visibility             = NULL,
    windowHeight           = NULL,
    windowWidth            = NULL,
    xWindow                = NULL,
    yWindow                = NULL
) {
  wbv <- self$workbook$bookViews

  if (is.null(wbv)) {
    wbv <- xml_node_create("workbookView")
  } else {
    wbv <- xml_node(wbv, "bookViews", "workbookView")
  }

  wbv <- xml_attr_mod(
    wbv,
    xml_attributes = c(
      activeTab              = as_xml_attr(activeTab),
      autoFilterDateGrouping = as_xml_attr(autoFilterDateGrouping),
      firstSheet             = as_xml_attr(firstSheet),
      minimized              = as_xml_attr(minimized),
      showHorizontalScroll   = as_xml_attr(showHorizontalScroll),
      showSheetTabs          = as_xml_attr(showSheetTabs),
      showVerticalScroll     = as_xml_attr(showVerticalScroll),
      tabRatio               = as_xml_attr(tabRatio),
      visibility             = as_xml_attr(visibility),
      windowHeight           = as_xml_attr(windowHeight),
      windowWidth            = as_xml_attr(windowWidth),
      xWindow                = as_xml_attr(xWindow),
      yWindow                = as_xml_attr(yWindow)
    )
  )

  self$workbook$bookViews <- xml_node_create(
    "bookViews",
    xml_children = wbv
  )

  invisible(self)
}

workbook_print <- function(self, private) {
  exSheets <- self$get_sheet_names()
  nSheets <- length(exSheets)
  nImages <- length(self$media)
  nCharts <- length(self$charts)

  showText <- "A Workbook object.\n"

  ## worksheets
  if (nSheets > 0) {
    showText <- c(showText, "\nWorksheets:\n")
    sheetTxt <- sprintf("Sheets: %s", paste(names(exSheets), collapse = " "))

    showText <- c(showText, sheetTxt, "\n")
  } else {
    showText <-
      c(showText, "\nWorksheets:\n", "No worksheets attached\n")
  }

  if (nSheets > 0) {
    showText <-
      c(showText, sprintf(
        "Write order: %s",
        stri_join(self$sheetOrder, sep = " ", collapse = ", ")
      ))
  }

  cat(unlist(showText))
  invisible(self)
}

workbook_set_last_modified_by <- function(self, private, LastModifiedBy = NULL) {
  # TODO rename to wb_set_last_modified_by() ?
  if (!is.null(LastModifiedBy)) {
    current_LastModifiedBy <-
      stri_match(self$core, regex = "<cp:lastModifiedBy>(.*?)</cp:lastModifiedBy>")[1, 2]
    self$core <-
      stri_replace_all_fixed(
        self$core,
        pattern = current_LastModifiedBy,
        replacement = LastModifiedBy
      )
  }

  invisible(self)
}

workbook_grid_lines <- function(
    self,
    private,
    sheet   = current_sheet(),
    show    = FALSE,
    print   = show
) {
  sheet <- private$get_sheet_index(sheet)

  assert_class(show, "logical")
  assert_class(print, "logical")

  ## show
  sv <- self$worksheets[[sheet]]$sheetViews
  sv <- xml_attr_mod(sv, c(showGridLines = as_xml_attr(show)))
  self$worksheets[[sheet]]$sheetViews <- sv

  ## print
  if (print)
    self$worksheets[[sheet]]$set_print_options(gridLines = print, gridLinesSet = print)

  invisible(self)
}

workbook_set_order <- function(self, private, sheets) {
  sheets <- private$get_sheet_index(sheet = sheets)

  if (anyDuplicated(sheets)) {
    stop("`sheets` cannot have duplicates")
  }

  if (length(sheets) != length(self$worksheets)) {
    stop(sprintf("Worksheet order must be same length as number of worksheets [%s]", length(self$worksheets)))
  }

  if (any(sheets > length(self$worksheets))) {
    stop("Elements of order are greater than the number of worksheets")
  }

  self$sheetOrder <- sheets
  invisible(self)
}

workbook_add_page_breaks <- function(
    self,
    private,
    sheet = current_sheet(),
    row = NULL,
    col = NULL
) {
  sheet <- private$get_sheet_index(sheet)
  self$worksheets[[sheet]]$add_page_break(row = row, col = col)
  invisible(self)
}

workbook_clean_sheet <- function(
    self,
    private,
    sheet        = current_sheet(),
    numbers      = TRUE,
    characters   = TRUE,
    styles       = TRUE,
    merged_cells = TRUE
) {
  sheet <- private$get_sheet_index(sheet)
  self$worksheets[[sheet]]$clean_sheet(
    numbers      = numbers,
    characters   = characters,
    styles       = styles,
    merged_cells = merged_cells
  )

  invisible(self)
}
