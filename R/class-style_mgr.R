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
#'                              color = wb_colour(hex = "FFFFFFFF"))
#' wb$styles_mgr$add(new_huge_font, "arial_huge")
#'
#' ## create another font
#' new_font <- create_font(name = "Arial")
#' wb$styles_mgr$add(new_font, "arial")
#'
#' ## create and add new fill
#' new_fill <- create_fill(patternType = "solid", fgColor = wb_colour(hex = "FF00224B"))
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
#' @export
style_mgr <- R6::R6Class("wbStylesMgr", {

  public = list(

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

    #' @field dxf dxf-ids
    dxf = NULL,

    #' @field styles styles as xml
    styles = NULL,

    #' @description
    #' Creates a new `wbStylesMgr` object
    #' @param numfmt numfmt
    #' @param font font
    #' @param fill fill
    #' @param border border
    #' @param xf xf
    #' @param dxf dxf
    #' @param styles styles
    #' @return a `wbStylesMgr` object
    initialize = function(numfmt = NA, font = NA, fill = NA, border = NA, xf = NA, dxf = NA, styles = NA) {

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

      dxfs <- self$styles$dxf
      if (length(dxfs)) {
        typ <- xml_node_name(dxfs)
        id  <- rownames(read_xf(read_xml(dxfs)))
        name <- paste0(typ, "-", id)

        self$dxf <- data.frame(
          typ,
          id,
          name
        )
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

    #' @description get dxf ids
    get_dxf = function() {
      invisible(self$dxf)
    },

    #' @description get numfmt id by name
    #' @param name name
    get_numfmt_id = function(name) {
      numfmt <- self$numfmt
      numfmt$id[numfmt$name == name]
    },

    #' @description get font id by name
    #' @param name name
    get_font_id = function(name) {
      font <- self$font
      font$id[font$name == name]
    },

    #' @description get fill id by name
    #' @param name name
    get_fill_id = function(name) {
      fill <- self$fill
      fill$id[fill$name == name]
    },

    #' @description get border id by name
    #' @param name name
    get_border_id = function(name) {
      border <- self$border
      border$id[border$name == name]
    },

    #' @description get xf id by name
    #' @param name name
    get_xf_id = function(name) {
      xf <- self$xf
      xf$id[match(name, xf$name)]
    },

    #' @description get dxf id by name
    #' @param name name
    get_dxf_id = function(name) {
      dxf <- self$dxf
      dxf$id[match(name, dxf$name)]
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
    next_dxf_id = function() {
      invisible(as.character(max(as.numeric(self$dxf$id), -1) + 1))
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
        is_dxf    <- any(ifelse(xml_node_name(style[sty]) == "dxf", TRUE, FALSE))

        if (skip_duplicates && is_numfmt && style_name[sty] %in% self$numfmt$name) next
        if (skip_duplicates && is_font   && style_name[sty] %in% self$font$name) next
        if (skip_duplicates && is_fill   && style_name[sty] %in% self$fill$name) next
        if (skip_duplicates && is_border && style_name[sty] %in% self$border$name) next
        if (skip_duplicates && is_xf     && style_name[sty] %in% self$xf$name) next
        if (skip_duplicates && is_dxf    && style_name[sty] %in% self$dxf$name) next

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

        if (is_dxf) {
          typ <- "dxf"
          dxfs <- c(self$styles$dxfs, style[sty])
          id  <- rownames(read_dxf(read_xml(dxfs)))
          self$styles$dxfs <- dxfs
        }

        new_entry <- data.frame(
          typ = typ,
          id = id[length(id)],
          name = style_name[sty]
        )

        if (is_numfmt) self$numfmt <- rbind(self$numfmt, new_entry)
        if (is_font)   self$font   <- rbind(self$font, new_entry)
        if (is_fill)   self$fill   <- rbind(self$fill, new_entry)
        if (is_border) self$border <- rbind(self$border, new_entry)
        if (is_xf)     self$xf     <- rbind(self$xf, new_entry)
        if (is_dxf)    self$dxf    <- rbind(self$dxf, new_entry)

      }

      invisible(self)
    }
  )

})
