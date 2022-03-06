#' style manager
#'
#' @examples
#' xlsxFile <- system.file("extdata", "oxlsx2_sheet.xlsx", package = "openxlsx2")
#' wb <- loadWorkbook(xlsxFile)
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
#' new_huge_font <- create_font(sz = "20", name = "Arial", b = "1", color = c(rgb = "FFFFFFFF"))
#' wb$styles_mgr$add(new_huge_font, "arial_huge")
#'
#' ## create another font
#' new_font <- create_font(name = "Arial")
#' wb$styles_mgr$add(new_font, "arial")
#'
#' ## create and add new fill
#' new_fill <- create_fill(patternType = "solid", fgColor = c(rgb = "FF00224B"))
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
#' set_cell_style(wb, "SUM", "B3:I3", wb$styles_mgr$get_xf_id("head_xf"))
#' set_cell_style(wb, "SUM", "C7:C16", wb$styles_mgr$get_xf_id("new_xf"))
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

    #' @field styles styles as xml
    styles = NULL,

    #' @description
    #' Creates a new `wbStylesMgr` object
    #' @param numfmt numfmt
    #' @param font font
    #' @param fill fill
    #' @param border border
    #' @param xf xf
    #' @param styles styles
    #' @return a `wbStylesMgr` object
    initialize = function(numfmt = NA, font = NA, fill = NA, border = NA, xf = NA, styles = NA) {

      numfmts <- self$styles$numFmts
      if (length(numfmts)) {
        typ <- xml_node_name(numfmts)
        id  <- openxlsx2:::rbindlist(xml_attr(numfmts, "numFmt"))[["numFmtId"]]
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
        id  <- rownames(openxlsx2:::read_font(read_xml(fonts)))
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
        id  <- rownames(openxlsx2:::read_fill(read_xml(fills)))
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
        id  <- rownames(openxlsx2:::read_border(read_xml(borders)))
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
        id  <- rownames(openxlsx2:::read_xf(read_xml(xfs)))
        name <- paste0(typ, "-", id)

        self$xf <- data.frame(
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

    #' @description get borer ids
    get_border = function() {
      invisible(self$border)
    },

    #' @description get xf ids
    get_xf = function() {
      invisible(self$xf)
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
      xf$id[xf$name == name]
    },

    #' @description get next numfmt id
    next_numfmt_id = function() {
      invisible(as.character(max(as.numeric(self$numfmt$id)) + 1))
    },

    #' @description get next font id
    next_font_id = function() {
      invisible(as.character(max(as.numeric(self$font$id)) + 1))
    },

    #' @description get next fill id
    next_fill_id = function() {
      invisible(as.character(max(as.numeric(self$fill$id)) + 1))
    },

    #' @description get next border id
    next_border_id = function() {
      invisible(as.character(max(as.numeric(self$border$id)) + 1))
    },

    #' @description get next xf id
    next_xf_id = function() {
      invisible(as.character(max(as.numeric(self$xf$id)) + 1))
    },

    ### adds
    #' @description
    #' add entry
    #' @param style new_style
    #' @param style_name a unique name identifying the style
    add = function(style, style_name) {

      typ <- NULL
      id  <- NULL

      is_numfmt <- ifelse(xml_node_name(style) == "numFmt", TRUE, FALSE)
      is_font   <- ifelse(xml_node_name(style) == "font", TRUE, FALSE)
      is_fill   <- ifelse(xml_node_name(style) == "fill", TRUE, FALSE)
      is_border <- ifelse(xml_node_name(style) == "border", TRUE, FALSE)
      is_xf     <- ifelse(xml_node_name(style) == "xf", TRUE, FALSE)

      if (is_numfmt) {
        typ <- "numFmt"
        id  <- unname(unlist(xml_attr(style, "numFmt"))["numFmtId"])
        self$styles$numFmts <- c(self$styles$numFmts, style)
      }

      if (is_font) {
        typ <- "font"
        fonts <- c(self$styles$fonts, style)
        id  <- rownames(openxlsx2:::read_font(read_xml(fonts)))
        self$styles$fonts <- fonts
      }

      if (is_fill) {
        typ <- "fill"
        fills <- c(self$styles$fills, style)
        id  <- rownames(openxlsx2:::read_fill(read_xml(fills)))
        self$styles$fills <- fills
      }

      if (is_border) {
        typ <- "border"
        borders <- c(self$styles$borders, style)
        id  <- rownames(openxlsx2:::read_fill(read_xml(borders)))
        self$styles$borders <- borders
      }

      if (is_xf) {
        typ <- "xf"
        xfs <- c(self$styles$cellXfs, style)
        id  <- rownames(openxlsx2:::read_xf(read_xml(xfs)))
        self$styles$cellXfs <- xfs
      }

      new_entry <- data.frame(
        typ = typ,
        id = id[length(id)],
        name = style_name
      )

      if (is_numfmt) self$numfmt <- rbind(self$numfmt, new_entry)
      if (is_font)   self$font <- rbind(self$font, new_entry)
      if (is_fill)   self$fill <- rbind(self$fill, new_entry)
      if (is_border) self$borders <- rbind(self$border, new_entry)
      if (is_xf)     self$xf <- rbind(self$xf, new_entry)

      invisible(self)
    }
  )

})
