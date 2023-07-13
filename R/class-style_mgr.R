#' style manager
#'
#' @examples
#' xlsxFile <- system.file("extdata", "oxlsx2_sheet.xlsx", package = "openxlsx2")
#' wb <- wb_load(xlsxFile)
#'
#' # ## start style mgr
#' # style <- style_mgr$new(wb)
#' # style$initialize(wb)
#'
#' # wb$styles_mgr$get_numfmt() |> print()
#' # wb$styles_mgr$next_numfmt_id() |> print()
#' # wb$styles_mgr$get_numfmt_id("numFmt-166")
#'
#' # create new number format
#' new_numfmt <- create_numfmt(numFmtId = wb$styles_mgr$next_numfmt_id(), formatCode = "#,#")
#'
#' # add it via stylemgr
#' wb$styles_mgr$add(new_numfmt, "test")
#'
#' ## get numfmts (invisible)
#' # z <- wb$styles_mgr$get_numfmt()
#' # z
#' wb$styles_mgr$styles$numFmts
#'
#' ## create and add huge font
#' new_huge_font <- create_font(sz = "20", name = "Arial", b = "1",
#'                              color = wb_color(hex = "FFFFFFFF"))
#' wb$styles_mgr$add(new_huge_font, "arial_huge")
#'
#' ## create another font
#' new_font <- create_font(name = "Arial")
#' wb$styles_mgr$add(new_font, "arial")
#'
#' ## create and add new fill
#' new_fill <- create_fill(patternType = "solid", fgColor = wb_color(hex = "FF00224B"))
#' wb$styles_mgr$add(new_fill, "blue")
#'
#' # create new style with numfmt enabled
#' head_xf <- create_cell_style(
#'   horizontal = "center",
#'   textRotation = "45",
#'   numFmtId = "0",
#'   fontId = wb$styles_mgr$get_font_id("arial_huge"),
#'   fillId = wb$styles_mgr$get_fill_id("blue")
#' )
#'
#' new_xf <- create_cell_style(
#'   numFmtId = wb$styles_mgr$get_numfmt_id("test"),
#'   fontId = wb$styles_mgr$get_font_id("arial")
#' )
#'
#' ## add new styles
#' wb$styles_mgr$add(head_xf, "head_xf")
#' wb$styles_mgr$add(new_xf, "new_xf")
#'
#' ## get cell style ids (invisible)
#' # z <- wb$styles_mgr$get_xf()
#'
#' ## get cell style id
#' # wb$styles_mgr$get_xf_id("new_xf")
#'
#'  ## assign styles to cells
#' wb$set_cell_style("SUM", "B3:I3", wb$styles_mgr$get_xf_id("head_xf"))
#' wb$set_cell_style("SUM", "C7:C16", wb$styles_mgr$get_xf_id("new_xf"))
#' # wb_open(wb)
#'
#' @noRd
style_mgr <- R6::R6Class("wbStylesMgr", {

  public <- list(

    #' @field numfmt numfmt-ids
    numfmt = NULL,

    #' @field font font-ids
    font = NULL,

    #' @field fill fill-ids
    fill = NULL,

    #' @field border border-ids
    border = NULL,

    #' @field xf xf-ids
    xf = NULL,

    #' @field cellStyle cellStyle-ids
    cellStyle = NULL,

    #' @field cellStyleXf cellStyleXf-ids
    cellStyleXf = NULL,

    #' @field dxf dxf-ids
    dxf = NULL,

    #' @field tableStyle tableStyle-ids
    tableStyle = NULL,

    #' @field defaultTableStyle defaultTableStyle
    defaultTableStyle = "TableStyleMedium2",

    #' @field defaultPivotStyle defaultPivotStyle
    defaultPivotStyle = "PivotStyleLight16",

    #' @field styles styles as xml
    styles = NULL,

    #' @description
    #' Creates a new `wbStylesMgr` object
    #' @param numfmt numfmt
    #' @param font font
    #' @param fill fill
    #' @param border border
    #' @param xf xf
    #' @param cellStyle cellStyles
    #' @param cellStyleXf cellStylesXf
    #' @param dxf dxf
    #' @param styles styles
    #' @return a `wbStylesMgr` object
    initialize = function(numfmt = NA, font = NA, fill = NA, border = NA, xf = NA, cellStyle = NA, cellStyleXf = NA, dxf = NA, styles = NA) {

      numfmts <- self$styles$numFmts
      if (length(numfmts)) {
        typ <- xml_node_name(numfmts)
        id  <- rbindlist(xml_attr(numfmts, "numFmt"))[["numFmtId"]]
        name <- paste0(typ, "-", id)

        self$numfmt <- data.frame(
          typ,
          id,
          name
        )
      }

      fonts <- self$styles$fonts
      if (length(fonts)) {
        typ <- xml_node_name(fonts)
        id  <- rownames(read_font(read_xml(fonts)))
        name <- paste0(typ, "-", id)

        self$font <- data.frame(
          typ,
          id,
          name
        )
      }

      fills <- self$styles$fills
      if (length(fills)) {
        typ <- xml_node_name(fills)
        id  <- rownames(read_fill(read_xml(fills)))
        name <- paste0(typ, "-", id)

        self$fill <- data.frame(
          typ,
          id,
          name
        )
      }

      borders <- self$styles$borders
      if (length(borders)) {
        typ <- xml_node_name(borders)
        id  <- rownames(read_border(read_xml(borders)))
        name <- paste0(typ, "-", id)

        self$border <- data.frame(
          typ,
          id,
          name
        )
      }

      xfs <- self$styles$cellXfs
      if (length(xfs)) {
        typ <- xml_node_name(xfs)
        id  <- rownames(read_xf(read_xml(xfs)))
        name <- paste0(typ, "-", id)

        self$xf <- data.frame(
          typ,
          id,
          name
        )
      }

      cellStyles <- self$styles$cellStyles
      if (length(cellStyles)) {

        typ <- xml_node_name(cellStyles)
        id  <- rownames(read_cellStyle(read_xml(cellStyles)))
        name <- rbindlist(xml_attr(cellStyles, "cellStyle"))$name

        self$cellStyle <- data.frame(
          typ,
          id,
          name
        )
      }

      cellStyleXfs <- self$styles$cellStyleXfs
      if (length(cellStyleXfs)) {

        typ <- xml_node_name(cellStyleXfs)
        id  <- rownames(read_xf(read_xml(cellStyleXfs)))
        name <- paste0(typ, "-", id)

        self$cellStyleXf <- data.frame(
          typ,
          id,
          name
        )
      }

      dxfs <- self$styles$dxfs
      if (length(dxfs)) {
        typ <- xml_node_name(dxfs)
        id  <- rownames(read_dxf(read_xml(dxfs)))
        name <- paste0(typ, "-", id)

        self$dxf <- data.frame(
          typ,
          id,
          name
        )
      }

      tableStyles <- self$styles$tableStyles
      if (length(tableStyles)) {

        tab_attrs <- rbindlist(xml_attr(tableStyles, "tableStyles"))

        if (!is.null(tab_attrs$defaultTableStyle))
          self$defaultTableStyle <- tab_attrs$defaultTableStyle

        if (!is.null(tab_attrs$defaultPivotStyle))
          self$defaultPivotStyle <- tab_attrs$defaultPivotStyle

        tableStyles <- self$styles$tableStyles <-
          xml_node(tableStyles, "tableStyles", "tableStyle")

        if (length(tableStyles)) {

          typ <- xml_node_name(tableStyles)
          id  <- rownames(read_tableStyle(read_xml(tableStyles)))
          name <- rbindlist(xml_attr(tableStyles, "tableStyle"))$name

          self$tableStyle <- data.frame(
            typ,
            id,
            name
          )
        }
      }

      invisible(self)
    },

    #' @description get numfmt ids
    get_numfmt = function() {
      invisible(self$numfmt)
    },

    #' @description get font ids
    get_font = function() {
      invisible(self$font)
    },

    #' @description get fill ids
    get_fill = function() {
      invisible(self$fill)
    },

    #' @description get border ids
    get_border = function() {
      invisible(self$border)
    },

    #' @description get xf ids
    get_xf = function() {
      invisible(self$xf)
    },

    #' @description get cellstyle ids
    get_cellStyle = function() {
      invisible(self$cellStyle)
    },

    #' @description get cellstylexf ids
    get_cellStyleXf = function() {
      invisible(self$cellStyleXf)
    },

    #' @description get dxf ids
    get_dxf = function() {
      invisible(self$dxf)
    },

    #' @description get numfmt id by name
    #' @param name name
    get_numfmt_id = function(name) {
      numfmt <- self$numfmt
      id <- numfmt$id[numfmt$name == name]
      if (length(id)) id else NULL
    },

    #' @description get font id by name
    #' @param name name
    get_font_id = function(name) {
      font <- self$font
      id <- font$id[font$name == name]
      if (length(id)) id else NULL
    },

    #' @description get fill id by name
    #' @param name name
    get_fill_id = function(name) {
      fill <- self$fill
      id <- fill$id[fill$name == name]
      if (length(id)) id else NULL
    },

    #' @description get border id by name
    #' @param name name
    get_border_id = function(name) {
      border <- self$border
      id <- border$id[border$name == name]
      if (length(id)) id else NULL
    },

    #' @description get xf id by name
    #' @param name name
    get_xf_id = function(name) {
      xf <- self$xf
      id <- xf$id[match(name, xf$name)]
      if (length(id)) id else NULL
    },

    #' @description get cellstyle id by name
    #' @param name name
    get_cellStyle_id = function(name) {
      cellstyle <- self$cellStyle
      id <- cellstyle$id[match(name, cellstyle$name)]
      if (length(id)) id else NULL
    },

    #' @description get cellstyleXf id by name
    #' @param name name
    get_cellStyleXf_id = function(name) {
      cellstylexf <- self$cellStyleXf
      id <- cellstylexf$id[match(name, cellstylexf$name)]
      if (length(id)) id else NULL
    },

    #' @description get dxf id by name
    #' @param name name
    get_dxf_id = function(name) {
      dxf <- self$dxf
      id <- dxf$id[match(name, dxf$name)]
      if (length(id)) id else NULL
    },

    #' @description get tableStyle id by name
    #' @param name name
    get_tableStyle_id = function(name) {
      tableStyle <- self$tableStyles
      id <- tableStyle$id[match(name, tableStyle$name)]
      if (length(id)) id else NULL
    },

    #' @description get next numfmt id
    next_numfmt_id = function() {
      # TODO check: first free custom format begins at 165?
      invisible(as.character(max(as.numeric(self$numfmt$id), 164) + 1))
    },

    #' @description get next font id
    next_font_id = function() {
      invisible(as.character(max(as.numeric(self$font$id), 0) + 1))
    },

    #' @description get next fill id
    next_fill_id = function() {
      invisible(as.character(max(as.numeric(self$fill$id), 0) + 1))
    },

    #' @description get next border id
    next_border_id = function() {
      invisible(as.character(max(as.numeric(self$border$id), 0) + 1))
    },

    #' @description get next xf id
    next_xf_id = function() {
      invisible(as.character(max(as.numeric(self$xf$id), -1) + 1))
    },

    #' @description get next xf id
    next_cellstyle_id = function() {
      invisible(as.character(max(as.numeric(self$cellStyle$id), -1) + 1))
    },

    #' @description get next xf id
    next_cellstylexf_id = function() {
      invisible(as.character(max(as.numeric(self$cellStyleXf$id), -1) + 1))
    },

    #' @description get next dxf id
    next_dxf_id = function() {
      invisible(as.character(max(as.numeric(self$dxf$id), -1) + 1))
    },

    #' @description get next tableStyle id
    next_tableStyle_id = function() {
      invisible(as.character(max(as.numeric(self$tableStyle$id), -1) + 1))
    },

    #' @description get named style ids
    #' @param name name
    getstyle_ids = function(name) {
      cellstyle_id     <- as.integer(self$get_cellStyle_id(name)) + 1L
      cellstyles_xfid  <- as.integer(rbindlist(xml_attr(self$styles$cellStyles[cellstyle_id], "cellStyle"))[["xfId"]]) + 1L
      cellstylexfs_ids <- rbindlist(xml_attr(self$styles$cellStyleXfs[cellstyles_xfid], "xf"))
      cellstylexfs_ids$titleId   <- cellstyle_id - 1L
      vars <- c("borderId", "fillId", "fontId", "numFmtId", "titleId")
      for (var in vars) {
        if (is.null(cellstylexfs_ids[[var]])) cellstylexfs_ids[var] <- "0"
      }
      cellstylexfs_ids <- sapply(cellstylexfs_ids[vars], as.integer)
      cellstylexfs_ids
    },

    ### adds
    #' @description
    #' add entry
    #' @param style new_style
    #' @param style_name a unique name identifying the style
    #' @param skip_duplicates should duplicates be added?
    add = function(style, style_name, skip_duplicates = TRUE) {

      # make sure that style and style_name length are equal
      if (length(style) != length(style_name))
        stop("style length and name do not match")

      for (sty in seq_along(style)) {

        typ <- NULL
        id  <- NULL

        is_numfmt <- any(ifelse(xml_node_name(style[sty]) == "numFmt", TRUE, FALSE))
        is_font   <- any(ifelse(xml_node_name(style[sty]) == "font", TRUE, FALSE))
        is_fill   <- any(ifelse(xml_node_name(style[sty]) == "fill", TRUE, FALSE))
        is_border <- any(ifelse(xml_node_name(style[sty]) == "border", TRUE, FALSE))
        is_xf     <- any(ifelse(xml_node_name(style[sty]) == "xf", TRUE, FALSE))
        is_celSty <- any(ifelse(xml_node_name(style[sty]) == "cellStyle", TRUE, FALSE))
        is_dxf    <- any(ifelse(xml_node_name(style[sty]) == "dxf", TRUE, FALSE))
        is_tabSty <- any(ifelse(xml_node_name(style[sty]) == "tableStyle", TRUE, FALSE))

        is_xf_fr  <- isTRUE(attr(style, "cellStyleXf"))

        if (skip_duplicates && is_numfmt && style_name[sty] %in% self$numfmt$name) next
        if (skip_duplicates && is_font   && style_name[sty] %in% self$font$name) next
        if (skip_duplicates && is_fill   && style_name[sty] %in% self$fill$name) next
        if (skip_duplicates && is_border && style_name[sty] %in% self$border$name) next
        if (skip_duplicates && is_xf     && style_name[sty] %in% self$xf$name) next
        if (skip_duplicates && is_celSty && style_name[sty] %in% self$cellStyle$name) next
        if (skip_duplicates && is_xf_fr  && style_name[sty] %in% self$cellStyleXf$name) next
        if (skip_duplicates && is_dxf    && style_name[sty] %in% self$dxf$name) next
        if (skip_duplicates && is_tabSty && style_name[sty] %in% self$tableStyle$name) next

        if (is_numfmt) {
          typ <- "numFmt"
          id  <- unname(unlist(xml_attr(style[sty], "numFmt"))["numFmtId"])
          self$styles$numFmts <- c(self$styles$numFmts, style[sty])
        }

        if (is_font) {
          typ <- "font"
          fonts <- c(self$styles$fonts, style[sty])
          id  <- rownames(read_font(read_xml(fonts)))
          self$styles$fonts <- fonts
        }

        if (is_fill) {
          typ <- "fill"
          fills <- c(self$styles$fills, style[sty])
          id  <- rownames(read_fill(read_xml(fills)))
          self$styles$fills <- fills
        }

        if (is_border) {
          typ <- "border"
          borders <- c(self$styles$borders, style[sty])
          id  <- rownames(read_border(read_xml(borders)))
          self$styles$borders <- borders
        }

        if (is_xf) {
          typ <- "xf"
          xfs <- c(self$styles$cellXfs, style[sty])
          id  <- rownames(read_xf(read_xml(xfs)))
          self$styles$cellXfs <- xfs
        }

        if (is_celSty) {
          typ <- "cellStyle"
          cellStyles <- c(self$styles$cellStyles, style[sty])
          id  <- rownames(read_cellStyle(read_xml(cellStyles)))
          self$styles$cellStyles <- cellStyles
        }

        if (is_xf_fr) {
          typ <- "xf"
          xfs <- c(self$styles$cellStyleXfs, style[sty])
          id  <- rownames(read_xf(read_xml(xfs)))
          self$styles$cellStyleXfs <- xfs
        }

        if (is_dxf) {
          typ <- "dxf"
          dxfs <- c(self$styles$dxfs, style[sty])
          id  <- rownames(read_dxf(read_xml(dxfs)))
          self$styles$dxfs <- dxfs
        }

        if (is_tabSty) {
          typ <- "tableStyle"
          tableStyles <- c(self$styles$tableStyles, style[sty])
          id  <- rownames(read_tableStyle(read_xml(tableStyles)))
          self$styles$tableStyles <- tableStyles
        }

        new_entry <- data.frame(
          typ = typ,
          id = id[length(id)],
          name = style_name[sty]
        )

        if (is_numfmt) self$numfmt      <- rbind(self$numfmt, new_entry)
        if (is_font)   self$font        <- rbind(self$font, new_entry)
        if (is_fill)   self$fill        <- rbind(self$fill, new_entry)
        if (is_border) self$border      <- rbind(self$border, new_entry)
        if (is_xf)     self$xf          <- rbind(self$xf, new_entry)
        if (is_celSty) self$cellStyle   <- rbind(self$cellStyle, new_entry)
        if (is_xf_fr)  self$cellStyleXf <- rbind(self$cellStyleXf, new_entry)
        if (is_dxf)    self$dxf         <- rbind(self$dxf, new_entry)
        if (is_tabSty) self$tableStyle  <- rbind(self$tableStyle, new_entry)

      }

      invisible(self)
    },

    #' @param wb wbWorkbook
    #' @param name style name
    #' @param font_name,font_size optional else the default of the theme
    #' @details
    #' possible styles are:
    #' "20% - Accent1"
    #' "20% - Accent2"
    #' "20% - Accent3"
    #' "20% - Accent4"
    #' "20% - Accent5"
    #' "20% - Accent6"
    #' "40% - Accent1"
    #' "40% - Accent2"
    #' "40% - Accent3"
    #' "40% - Accent4"
    #' "40% - Accent5"
    #' "40% - Accent6"
    #' "60% - Accent1"
    #' "60% - Accent2"
    #' "60% - Accent3"
    #' "60% - Accent4"
    #' "60% - Accent5"
    #' "60% - Accent6"
    #' "Accent1"
    #' "Accent2"
    #' "Accent3"
    #' "Accent4"
    #' "Accent5"
    #' "Accent6"
    #' "Bad"
    #' "Calculation"
    #' "Check Cell"
    #' "Comma"
    #' "Comma \[0\]"
    #' "Currency"
    #' "Currency \[0\]"
    #' "Explanatory Text"
    #' "Good"
    #' "Heading 1"
    #' "Heading 2"
    #' "Heading 3"
    #' "Heading 4"
    #' "Input"
    #' "Linked Cell"
    #' ”Neutral"
    #' "Normal"
    #' "Note"
    #' "Output"
    #' "Per cent"
    #' "Title"
    #' "Total"
    #' "Warning Text"
    init_named_style = function(name, font_name = "Arial", font_size = 11) {

      # we probably should only have unique named styles. check if style is found.
      # if yes, abort style initialization.
      got <- self$get_cellStyle_id(name)

      if (!is.null(got) && !is.na(got))
        return(self)

      font_xml <- NULL
      fill_xml <- NULL
      border_xml <- NULL
      cell_style_xml <- NULL

      numFmtId  <- ""
      builtinId <- ""

      if (name == "Bad") {

        font_xml <- create_font(sz = font_size, color = wb_color(hex = "FF9C0006"), name = font_name, family = "2", scheme = "minor")

        fill_xml <- create_fill(patternType = "solid", fgColor = wb_color(hex = "FFFFC7CE"))

        builtinId <- "27"

      }

      if (name == "Good") {

        font_xml <- create_font(sz = font_size, color = wb_color(hex = "FF006100"), name = font_name, family = "2", scheme = "minor")

        fill_xml <- create_fill(patternType = "solid", fgColor = wb_color(hex = "FFC6EFCE"))

        builtinId <- "26"

      }

      if (name == "Neutral") {

        font_xml <- create_font(sz = font_size, color = wb_color(hex = "FF9C5700"), name = font_name, family = "2", scheme = "minor")

        fill_xml <- create_fill(patternType = "solid", fgColor = wb_color(hex = "FFFFEB9C"))

        builtinId <- "28"

      }

      if (name == "Calculation") {

        font_xml <- create_font(b = TRUE, sz = font_size, color = wb_color(hex = "FFFA7D00"), name = font_name, family = "2", scheme = "minor")

        fill_xml <- create_fill(patternType = "solid", fgColor = wb_color(hex = "FFF2F2F2"))

        border_xml <- create_border(
          left = "thin",  left_color = wb_color(hex = "FF7F7F7F"),
          right = "thin",  right_color = wb_color(hex = "FF7F7F7F"),
          top = "thin",  top_color = wb_color(hex = "FF7F7F7F"),
          bottom = "thin",  bottom_color = wb_color(hex = "FF7F7F7F")
        )

        builtinId <- "22"
      }

      if (name == "Check Cell") {

        font_xml <- create_font(sz = font_size, color = wb_color(theme = 0), name = font_name, family = "2", scheme = "minor")

        fill_xml <- create_fill(patternType = "solid", fgColor = wb_color(hex = "FFA5A5A5"))

        border_xml <- create_border(
          left = "double",  left_color = wb_color(hex = "FF3F3F3F"),
          right = "double",  right_color = wb_color(hex = "FF3F3F3F"),
          top = "double",  top_color = wb_color(hex = "FF3F3F3F"),
          bottom = "double",  bottom_color = wb_color(hex = "FF3F3F3F")
        )

        builtinId <- "23"
      }

      if (name == "Explanatory Text") {

        font_xml <- create_font(i = TRUE, sz = font_size, color = wb_color(hex = "FF7F7F7F"), name = font_name, family = "2", scheme = "minor")

        fill_xml <- create_fill(patternType = "none")

        builtinId <- "53"
      }

      if (name == "Input") {

        font_xml <- create_font(sz = font_size, color = wb_color(hex = "FF3F3F76"), name = font_name, family = "2", scheme = "minor")

        fill_xml <- create_fill(patternType = "solid", fgColor = wb_color(hex = "FFFFCC99"))

        border_xml <- create_border(
          left = "thin",  left_color = wb_color(hex = "FF7F7F7F"),
          right = "thin",  right_color = wb_color(hex = "FF7F7F7F"),
          top = "thin",  top_color = wb_color(hex = "FF7F7F7F"),
          bottom = "thin",  bottom_color = wb_color(hex = "FF7F7F7F")
        )

        builtinId <- "20"
      }

      if (name == "Linked Cell") {

        font_xml <- create_font(sz = font_size, color = wb_color(hex = "FFFA7D00"), name = font_name, family = "2", scheme = "minor")

        fill_xml <- create_fill(patternType = "none")

        border_xml <- create_border(
          left = NULL,
          right = NULL,
          top = NULL,
          bottom = "double",  bottom_color = wb_color(hex = "FFFF8001")
        )

        builtinId <- "24"
      }

      if (name == "Note") {

        font_xml <- create_font(sz = font_size, color = wb_color(theme = 1), name = font_name, family = "2", scheme = "minor")

        fill_xml <- create_fill(patternType = "solid", fgColor = wb_color(hex = "FFFFFFCC"))

        border_xml <- create_border(
          left = "thin",  left_color = wb_color(hex = "FFB2B2B2"),
          right = "thin",  right_color = wb_color(hex = "FFB2B2B2"),
          top = "thin",  top_color = wb_color(hex = "FFB2B2B2"),
          bottom = "thin",  bottom_color = wb_color(hex = "FFB2B2B2")
        )

        builtinId <- "10"

      }

      if (name == "Output") {

        font_xml <- create_font(b = TRUE, sz = font_size, color = wb_color(hex = "FF3F3F3F"), name = font_name, family = "2", scheme = "minor")

        fill_xml <- create_fill(patternType = "solid", fgColor = wb_color(hex = "FFF2F2F2"))

        border_xml <- create_border(
          left = "thin",  left_color = wb_color(hex = "FF3F3F3F"),
          right = "thin",  right_color = wb_color(hex = "FF3F3F3F"),
          top = "thin",  top_color = wb_color(hex = "FF3F3F3F"),
          bottom = "thin",  bottom_color = wb_color(hex = "FF3F3F3F")
        )

        builtinId <- "21"
      }

      if (name == "Warning Text") {

        font_xml <- create_font(sz = font_size, color = wb_color(hex = "FFFF0000"), name = font_name, family = "2", scheme = "minor")

        fill_xml <- create_fill(patternType = "none")

        builtinId <- "11"
      }

      if (name == "Heading 1") {

        font_xml <- create_font(b = TRUE, sz = 15, color = wb_color(theme = 3), name = font_name, family = "2", scheme = "minor")

        fill_xml <- create_fill(patternType = "none")

        border_xml <- create_border(
          left = NULL,
          right = NULL,
          top = NULL,
          bottom = "thick",  bottom_color = wb_color(theme = 4)
        )

        builtinId <- "16"
      }

      if (name == "Heading 2") {

        font_xml <- create_font(b = TRUE, sz = 13, color = wb_color(theme = 3), name = font_name, family = "2", scheme = "minor")

        fill_xml <- create_fill(patternType = "none")

        border_xml <- create_border(
          left = NULL,
          right = NULL,
          top = NULL,
          bottom = "thick",  bottom_color = wb_color(theme = 4, tint = "0.499984740745262")
        )

        builtinId <- "17"
      }

      if (name == "Heading 3") {

        font_xml <- create_font(b = TRUE, sz = 11, color = wb_color(theme = 3), name = font_name, family = "2", scheme = "minor")

        fill_xml <- create_fill(patternType = "none")

        border_xml <- create_border(
          left = NULL,
          right = NULL,
          top = NULL,
          bottom = "medium",  bottom_color = wb_color(theme = 4, tint = "0.39997558519241921")
        )

        builtinId <- "18"
      }

      if (name == "Heading 4") {

        font_xml <- create_font(b = TRUE, sz = 11, color = wb_color(theme = 3), name = font_name, family = "2", scheme = "minor")

        fill_xml <- create_fill(patternType = "none")

        builtinId <- "19"
      }

      if (name == "Title") {

        font_xml <- create_font(sz = 18, color = wb_color(theme = 3), name = "Calibri Light", family = "2", scheme = "major")

        builtinId <- "15"

      }

      if (name == "Total") {

        font_xml <- create_font(b = TRUE, sz = font_size, color = wb_color(theme = 1), name = font_name, family = "2", scheme = "minor")

        fill_xml <- create_fill(patternType = "none")

        border_xml <- create_border(
          left = NULL,
          right = NULL,
          top = "thin", top_color = wb_color(theme = 4),
          bottom = "double",  bottom_color = wb_color(theme = 4)
        )

        builtinId <- "25"
      }

      if (name %in% paste0("Accent", 1:6)) {

        accent_id <- gsub("\\D+", "", name)

        font_xml <- create_font(sz = font_size, color = wb_color(theme = 0), name = font_name, family = "2", scheme = "minor")

        theme_id <- as.integer(accent_id) + 3L
        fill_xml <- create_fill(patternType = "solid", fgColor = wb_color(theme = theme_id))

        if (accent_id == "1") builtinId <- "29"
        if (accent_id == "2") builtinId <- "33"
        if (accent_id == "3") builtinId <- "37"
        if (accent_id == "4") builtinId <- "41"
        if (accent_id == "5") builtinId <- "45"
        if (accent_id == "6") builtinId <- "49"
      }

      if (name %in% paste0("20% - Accent", 1:6)) {

        accent_id <- gsub("\\D+", "", strsplit(name, " - ")[[1]][2])

        font_xml <- create_font(sz = font_size, color = wb_color(theme = 1), name = font_name, family = "2", scheme = "minor")

        theme_id <- as.integer(accent_id) + 3L
        fill_xml <- create_fill(patternType = "solid", fgColor = wb_color(theme = theme_id, tint = "0.79998168889431442"), bgColor = wb_color(indexed = 65))

        if (accent_id == "1") builtinId <- "30"
        if (accent_id == "2") builtinId <- "34"
        if (accent_id == "3") builtinId <- "38"
        if (accent_id == "4") builtinId <- "42"
        if (accent_id == "5") builtinId <- "46"
        if (accent_id == "6") builtinId <- "50"
      }

      if (name %in% paste0("40% - Accent", 1:6)) {

        accent_id <- gsub("\\D+", "", strsplit(name, " - ")[[1]][2])

        font_xml <- create_font(sz = font_size, color = wb_color(theme = 1), name = font_name, family = "2", scheme = "minor")

        theme_id <- as.integer(accent_id) + 3L
        fill_xml <- create_fill(patternType = "solid", fgColor = wb_color(theme = theme_id, tint = "0.59999389629810485"), bgColor = wb_color(indexed = 65))

        if (accent_id == "1") builtinId <- "31"
        if (accent_id == "2") builtinId <- "35"
        if (accent_id == "3") builtinId <- "39"
        if (accent_id == "4") builtinId <- "43"
        if (accent_id == "5") builtinId <- "47"
        if (accent_id == "6") builtinId <- "51"
      }

      if (name %in% paste0("60% - Accent", 1:6)) {

        accent_id <- gsub("\\D+", "", strsplit(name, " - ")[[1]][2])

        font_xml <- create_font(sz = font_size, color = wb_color(theme = 1), name = font_name, family = "2", scheme = "minor")

        theme_id <- as.integer(accent_id) + 3L
        fill_xml <- create_fill(patternType = "solid", fgColor = wb_color(theme = theme_id, tint = "0.39997558519241921"), bgColor = wb_color(indexed = 65))

        if (accent_id == "1") builtinId <- "32"
        if (accent_id == "2") builtinId <- "36"
        if (accent_id == "3") builtinId <- "40"
        if (accent_id == "4") builtinId <- "44"
        if (accent_id == "5") builtinId <- "48"
        if (accent_id == "6") builtinId <- "52"
      }

      if (name == "Comma") {

        font_xml <- create_font(sz = font_size, color = wb_color(theme = 1), name = font_name, family = "2", scheme = "minor")

        numFmtId  <- "43"
        builtinId <- "3"
      }

      if (name == "Comma [0]") {

        font_xml <- create_font(sz = font_size, color = wb_color(theme = 1), name = font_name, family = "2", scheme = "minor")

        numFmtId  <- "41"
        builtinId <- "6"
      }

      if (name == "Currency") {

        font_xml <- create_font(sz = font_size, color = wb_color(theme = 1), name = font_name, family = "2", scheme = "minor")

        numFmtId  <- "44"
        builtinId <- "4"
      }

      if (name == "Currency [0]") {

        font_xml <- create_font(sz = font_size, color = wb_color(theme = 1), name = font_name, family = "2", scheme = "minor")

        numFmtId  <- "42"
        builtinId <- "7"
      }

      if (name == "Per cent") {

        font_xml <- create_font(sz = font_size, color = wb_color(theme = 1), name = font_name, family = "2", scheme = "minor")

        numFmtId  <- "9"
        builtinId <- "5"
      }

      font_id <- ""
      if (!is.null(font_xml)) {
        self$add(font_xml, font_xml)
        font_id <- self$get_font_id(font_xml)
      }

      fill_id <- ""
      if (!is.null(fill_xml)) {
        self$add(fill_xml, fill_xml)
        fill_id <- self$get_fill_id(fill_xml)
      }

      border_id <- ""
      if (!is.null(border_xml)) {
        self$add(border_xml, border_xml)
        border_id <- self$get_border_id(border_xml)
      }

      cell_style_xml <- create_cell_style(num_fmt_id = numFmtId, font_id = font_id, fill_id = fill_id, border_id = border_id, is_cell_style_xf = TRUE)
      attr(cell_style_xml, "cellStyleXf") <- TRUE
      self$add(cell_style_xml, name)
      xf_fr_id <- self$get_cellStyleXf_id(name)

      cell_style <- xml_node_create("cellStyle", xml_attributes = c(name = name, xfId = xf_fr_id, builtinId = builtinId))
      self$add(cell_style, name)

      invisible(self)
    }
  )

})
