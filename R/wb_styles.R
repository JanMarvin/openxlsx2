# functions interacting with styles

# internal function used in wb_load
# @param x character string containing styles.xml
import_styles <- function(x) {
  sxml <- openxlsx2::read_xml(x)

  z <- NULL

  # Number Formats numFmtId
  # overrides for
  # + numFmtId
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-2.8.1
  z$numFmts <- xml_node(sxml, "styleSheet", "numFmts", "numFmt")

  # Fonts fontId
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.font?view=openxml-2.8.1
  z$fonts <- xml_node(sxml, "styleSheet", "fonts", "font")

  # Fills fillId
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fill?view=openxml-2.8.1
  z$fills <- xml_node(sxml, "styleSheet", "fills", "fill")

  # Borders borderId
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.border?view=openxml-2.8.1
  z$borders <- xml_node(sxml, "styleSheet", "borders", "border")

  # Formatting Records: Cell Style Formats
  # links
  # + numFmtId
  # + fontId
  # + fillId
  # + borderId
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellstyleformats?view=openxml-2.8.1
  z$cellStyleXfs <- xml_node(sxml, "styleSheet", "cellStyleXfs", "xf")

  # Cell Formats
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformat?view=openxml-2.8.1
  #
  # Position is 0 index s="value" used in worksheet <c ...>
  # links
  # + numFmtId
  # + fontId
  # + fillId
  # + borderId
  # + xfId
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformats?view=openxml-2.8.1
  z$cellXfs <- xml_node(sxml, "styleSheet", "cellXfs", "xf")

  # Cell Styles (no clue)
  # links
  # + xfId
  # + builtinId
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellstyle?view=openxml-2.8.1
  z$cellStyles <- xml_node(sxml, "styleSheet", "cellStyles", "cellStyle")

  # Differential Formats (No clue?)
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.differentialformat?view=openxml-2.8.1
  z$dxfs <- xml_node(sxml, "styleSheet", "dxfs", "dxf")

  # Table Styles Maybe position Id?
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tablestyle?view=openxml-2.8.1
  z$tableStyles <- xml_node(sxml, "styleSheet", "tableStyles")

  # Colors
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.colors?view=openxml-2.8.1
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.indexedcolors?view=openxml-2.8.1
  z$colors <- xml_node(sxml, "styleSheet", "colors")

  # Future extensions No clue, some special styles
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.extensionlist?view=openxml-2.8.1
  z$extLst <- xml_node(sxml, "styleSheet", "extLst")

  z
}

# TODO guessing here
#' Helper to create a border
#' @description
#' Border styles can any of the following: "thin", "thick", "slantDashDot", "none", "mediumDashed", "mediumDashDot", "medium", "hair", "double", "dotted", "dashed", "dashedDotDot", "dashDot"
#' Border colors can be created with [wb_color()]
#' @seealso [wb_add_border()]
#' @param diagonal_down x
#' @param diagonal_up x
#' @param outline x
#' @param bottom X
#' @param bottom_color,diagonal_color,left_color,right_color,top_color a color created with [wb_color()]
#' @param diagonal X
#' @param end x,
#' @param horizontal x
#' @param left x
#' @param right x
#' @param start x
#' @param top x
#' @param vertical x
#' @param ... x
#'
#' @export
create_border <- function(
    diagonal_down  = "",
    diagonal_up    = "",
    outline        = "",
    bottom         = NULL,
    bottom_color   = NULL,
    diagonal       = NULL,
    diagonal_color = NULL,
    end            = "",
    horizontal     = "",
    left           = NULL,
    left_color     = NULL,
    right          = NULL,
    right_color    = NULL,
    start          = "",
    top            = NULL,
    top_color      = NULL,
    vertical       = "",
    ...
) {

  standardize(...)

  if (!is.null(left_color))     left_color     <- xml_node_create("color", xml_attributes = left_color)
  if (!is.null(right_color))    right_color    <- xml_node_create("color", xml_attributes = right_color)
  if (!is.null(top_color))      top_color      <- xml_node_create("color", xml_attributes = top_color)
  if (!is.null(bottom_color))   bottom_color   <- xml_node_create("color", xml_attributes = bottom_color)
  if (!is.null(diagonal_color)) diagonal_color <- xml_node_create("color", xml_attributes = diagonal_color)

  # excel dies on style=\"\"
  if (!is.null(left))     left     <- c(style = left)
  if (!is.null(right))    right    <- c(style = right)
  if (!is.null(top))      top      <- c(style = top)
  if (!is.null(bottom))   bottom   <- c(style = bottom)
  if (!is.null(diagonal)) diagonal <- c(style = diagonal)

  left     <- xml_node_create("left", xml_children = left_color, xml_attributes = left)
  right    <- xml_node_create("right", xml_children = right_color, xml_attributes = right)
  top      <- xml_node_create("top", xml_children = top_color, xml_attributes = top)
  bottom   <- xml_node_create("bottom", xml_children = bottom_color, xml_attributes = bottom)
  diagonal <- xml_node_create("diagonal", xml_children = diagonal_color, xml_attributes = diagonal)

  if (left     == "<left/>")     left     <- ""
  if (right    == "<right/>")    right    <- ""
  if (top      == "<top/>")      top      <- ""
  if (bottom   == "<bottom/>")   bottom   <- ""
  if (diagonal == "<diagonal/>") diagonal <- ""

  df_border <- data.frame(
    start            = start,
    end              = end,
    left             = left,
    right            = right,
    top              = top,
    bottom           = bottom,
    diagonalDown     = diagonal_down,
    diagonalUp       = diagonal_up,
    diagonal         = diagonal,
    vertical         = vertical,
    horizontal       = horizontal,
    outline          = outline,      # unknown position in border
    stringsAsFactors = FALSE
  )
  border <- write_border(df_border)

  return(border)
}

#' create number format
#' @param numFmtId an id, the list can be found in the **Details** of [create_cell_style()]
#' @param formatCode a format code
#' @seealso [wb_add_numfmt()]
#' @export
create_numfmt <- function(numFmtId, formatCode) {

  df_numfmt <- data.frame(
    numFmtId         = as_xml_attr(numFmtId),
    formatCode       = as_xml_attr(formatCode),
    stringsAsFactors = FALSE
  )
  numfmt <- write_numfmt(df_numfmt)

  return(numfmt)
}

#' create font format
#' @param b bold
#' @param charset charset
#' @param color rgb color: default "FF000000"
#' @param condense condense
#' @param extend extend
#' @param family font family: default "2"
#' @param i italic
#' @param name font name: default "Calibri"
#' @param outline outline
#' @param scheme font scheme: default "minor"
#' @param shadow shadow
#' @param strike strike
#' @param sz font size: default "11",
#' @param u underline
#' @param vert_align vertical alignment
#' @param ... ...
#' @seealso [wb_add_font()]
#' @examples
#' font <- create_font()
#' # openxml has the alpha value leading
#' hex8 <- unlist(xml_attr(read_xml(font), "font", "color"))
#' hex8 <- paste0("#", substr(hex8, 3, 8), substr(hex8, 1, 2))
#'
#' # # write test color
#' # col <- crayon::make_style(col2rgb(hex8, alpha = TRUE))
#' # cat(col("Test"))
#' @export
create_font <- function(
    b          = "",
    charset    = "",
    color      = wb_color(hex = "FF000000"),
    condense   = "",
    extend     = "",
    family     = "2",
    i          = "",
    name       = "Calibri",
    outline    = "",
    scheme     = "minor",
    shadow     = "",
    strike     = "",
    sz         = "11",
    u          = "",
    vert_align = "",
    ...
) {

  standardize(...)

  if (b != "") {
    b <- xml_node_create("b", xml_attributes = c("val" = as_xml_attr(b)))
  }

  if (charset != "") {
    charset <- xml_node_create("charset", xml_attributes = c("val" = charset))
  }

  if (!is.null(color) && !all(color == "")) {
    # alt xml_attributes(theme:)
    color <- xml_node_create("color", xml_attributes = color)
  }

  if (condense != "") {
    condense <- xml_node_create("condense", xml_attributes = c("val" = as_xml_attr(condense)))
  }

  if (extend != "") {
    extend <- xml_node_create("extend", xml_attributes = c("val" = as_xml_attr(extend)))
  }

  if (family != "") {
    family <- xml_node_create("family", xml_attributes = c("val" = family))
  }

  if (i != "") {
    i <- xml_node_create("i", xml_attributes = c("val" = as_xml_attr(i)))
  }

  if (is.null(name)) name <- ""
  if (name != "") {
    name <- xml_node_create("name", xml_attributes = c("val" = name))
  }

  if (outline != "") {
    outline <- xml_node_create("outline", xml_attributes = c("val" = as_xml_attr(outline)))
  }

  if (scheme != "") {
    scheme <- xml_node_create("scheme", xml_attributes = c("val" = scheme))
  }

  if (shadow != "") {
    shadow <- xml_node_create("shadow", xml_attributes = c("val" = as_xml_attr(shadow)))
  }

  if (strike != "") {
    strike <- xml_node_create("strike", xml_attributes = c("val" = as_xml_attr(strike)))
  }

  if (sz != "") {
    sz <- xml_node_create("sz", xml_attributes = c("val" = as_xml_attr(sz)))
  }

  if (u != "") {
    u <- xml_node_create("u", xml_attributes = c("val" = as_xml_attr(u)))
  }

  if (vert_align != "") {
    vert_align <- xml_node_create("vertAlign", xml_attributes = c("val" = as_xml_attr(vert_align)))
  }

  df_font <- data.frame(
    b                = b,
    charset          = charset,
    color            = color,
    condense         = condense,
    extend           = extend,
    family           = family,
    i                = i,
    name             = name,
    outline          = outline,
    scheme           = scheme,
    shadow           = shadow,
    strike           = strike,
    sz               = sz,
    u                = u,
    vertAlign        = vert_align,
    stringsAsFactors = FALSE
  )
  font <- write_font(df_font)

  if (font == "<font/>") {
    font <- ""
  }

  return(font)
}

#' create fill
#'
#' @param gradientFill complex fills
#' @param patternType various: default is "none", but also "solid", or a color like "gray125"
#' @param bgColor hex8 color with alpha, red, green, blue only for patternFill
#' @param fgColor hex8 color with alpha, red, green, blue only for patternFill
#' @param ... ...
#' @seealso [wb_add_fill()]
#'
#' @export
create_fill <- function(
    gradientFill = "",
    patternType  = "",
    bgColor      = NULL,
    fgColor      = NULL,
    ...
) {

  standardize_color_names(...)

  if (!is.null(bgColor) && !all(bgColor == "")) {
    bgColor <- xml_node_create("bgColor", xml_attributes = bgColor)
  }

  if (!is.null(fgColor) && !all(fgColor == "")) {
    fgColor <- xml_node_create("fgColor", xml_attributes = fgColor)
  }

  # if gradient fill is specified we can not have patternFill too. otherwise
  # we end up with a solid black fill
  if (gradientFill == "") {
    patternFill <- xml_node_create("patternFill",
      xml_children   = c(fgColor, bgColor),
      xml_attributes = c(patternType = patternType)
    )
  } else {
    patternFill <- ""
  }

  df_fill <- data.frame(
    gradientFill     = gradientFill,
    patternFill      = patternFill,
    stringsAsFactors = FALSE
  )
  fill <- write_fill(df_fill)

  return(fill)
}

# TODO can be further generalized with additional xf attributes and children
#' Helper to create a cell style
#'
#' Create_cell_style with [wb_add_cell_style()]
#' @param border_id dummy
#' @param fill_id dummy
#' @param font_id dummy
#' @param num_fmt_id a numFmt ID for a builtin style
#' @param pivot_button dummy
#' @param quote_prefix dummy
#' @param xf_id dummy
#' @param horizontal dummy
#' @param indent dummy
#' @param justify_last_line dummy
#' @param reading_order dummy
#' @param relative_indent dummy
#' @param shrink_to_fit dummy
#' @param text_rotation dummy
#' @param vertical dummy
#' @param wrap_text dummy
#' @param ext_lst dummy
#' @param hidden dummy
#' @param locked dummy
#' @param horizontal alignment can be "", "center", "right"
#' @param vertical alignment can be "", "center", "right"
#' @param ... reserved for additional arguments
#'
#' @details
#'  | "ID" | "numFmt"                    |
#'  |------|-----------------------------|
#'  | "0"  | "General"                   |
#'  | "1"  | "0"                         |
#'  | "2"  | "0.00"                      |
#'  | "3"  | "#,##0"                     |
#'  | "4"  | "#,##0.00"                  |
#'  | "9"  | "0%"                        |
#'  | "10" | "0.00%"                     |
#'  | "11" | "0.00E+00"                  |
#'  | "12" | "# ?/?"                     |
#'  | "13" | "# ??/??"                   |
#'  | "14" | "mm-dd-yy"                  |
#'  | "15" | "d-mmm-yy"                  |
#'  | "16" | "d-mmm"                     |
#'  | "17" | "mmm-yy"                    |
#'  | "18" | "h:mm AM/PM"                |
#'  | "19" | "h:mm:ss AM/PM"             |
#'  | "20" | "h:mm"                      |
#'  | "21" | "h:mm:ss"                   |
#'  | "22" | "m/d/yy h:mm"               |
#'  | "37" | "#,##0 ;(#,##0)"            |
#'  | "38" | "#,##0 ;\[Red\](#,##0)"     |
#'  | "39" | "#,##0.00;(#,##0.00)"       |
#'  | "40" | "#,##0.00;\[Red\](#,##0.00)"|
#'  | "45" | "mm:ss"                     |
#'  | "46" | "\[h\]:mm:ss"               |
#'  | "47" | "mmss.0"                    |
#'  | "48" | "##0.0E+0"                  |
#'  | "49" | "@"                         |
#' @export
create_cell_style <- function(
    border_id         = "",
    fill_id           = "",
    font_id           = "",
    num_fmt_id        = "",
    pivot_button      = "",
    quote_prefix      = "",
    xf_id             = "",
    horizontal        = "",
    indent            = "",
    justify_last_line = "",
    reading_order     = "",
    relative_indent   = "",
    shrink_to_fit     = "",
    text_rotation     = "",
    vertical          = "",
    wrap_text         = "",
    ext_lst           = "",
    hidden            = "",
    locked            = "",
    ...
) {
  n <- length(num_fmt_id)

  arguments <- c(ls(), "is_cell_style_xf")
  standardize_case_names(..., arguments = arguments)
  args <- list(...)

  is_cell_style_xf <- isTRUE(args$is_cell_style_xf)

  applyAlignment <- ""
  if (any(horizontal != "") || any(text_rotation != "") || any(vertical != "")) applyAlignment <- "1"
  if (is_cell_style_xf) applyAlignment <- "0"

  applyBorder <- ""
  if (any(border_id != "")) applyBorder <- "1"
  if (is_cell_style_xf && isTRUE(border_id == "")) applyBorder <- "0"

  applyFill <- ""
  if (any(fill_id != "")) applyFill <- "1"
  if (is_cell_style_xf && isTRUE(fill_id == "")) applyFill <- "0"

  applyFont <- ""
  if (any(font_id != "")) applyFont <- "1"
  if (is_cell_style_xf && isTRUE(font_id > "1")) applyFont <- "0"

  applyNumberFormat <- ""
  if (any(num_fmt_id != "")) applyNumberFormat <- "1"
  if (is_cell_style_xf) applyNumberFormat <- "0"

  applyProtection <- ""
  if (any(hidden != "") || any(locked != "")) applyProtection <- "1"
  if (is_cell_style_xf) applyProtection <- "0"


  df_cellXfs <- data.frame(
    applyAlignment    = rep(applyAlignment, n),
    applyBorder       = rep(applyBorder, n),
    applyFill         = rep(applyFill, n),
    applyFont         = rep(applyFont, n),
    applyNumberFormat = rep(applyNumberFormat, n),
    applyProtection   = rep(applyProtection, n),
    borderId          = rep(as_xml_attr(border_id), n),
    fillId            = rep(as_xml_attr(fill_id), n),
    fontId            = rep(as_xml_attr(font_id), n),
    numFmtId          = as_xml_attr(num_fmt_id),
    pivotButton       = rep(as_xml_attr(pivot_button), n),
    quotePrefix       = rep(as_xml_attr(quote_prefix), n),
    xfId              = rep(as_xml_attr(xf_id), n),
    horizontal        = rep(as_xml_attr(horizontal), n),
    indent            = rep(as_xml_attr(indent), n),
    justifyLastLine   = rep(as_xml_attr(justify_last_line), n),
    readingOrder      = rep(as_xml_attr(reading_order), n),
    relativeIndent    = rep(as_xml_attr(relative_indent), n),
    shrinkToFit       = rep(as_xml_attr(shrink_to_fit), n),
    textRotation      = rep(as_xml_attr(text_rotation), n),
    vertical          = rep(as_xml_attr(vertical), n),
    wrapText          = rep(as_xml_attr(wrap_text), n),
    extLst            = rep(ext_lst, n),
    hidden            = rep(as_xml_attr(hidden), n),
    locked            = rep(as_xml_attr(locked), n),
    stringsAsFactors  = FALSE
  )

  cellXfs <- write_xf(df_cellXfs)
  cellXfs
}

#' internal function to set border to a style
#' @param xf_node some xf node
#' @param border_id some numeric value as character
#' @noRd
set_border <- function(xf_node, border_id) {
  z <- read_xf(read_xml(xf_node))
  z$applyBorder <- "1"
  z$borderId <- as_xml_attr(border_id)
  write_xf(z)
}

#' internal function to set fill to a style
#' @param xf_node some xf node
#' @param fill_id some numeric value as character
#' @noRd
set_fill <- function(xf_node, fill_id) {
  z <- read_xf(read_xml(xf_node))
  z$applyFill <- "1"
  z$fillId <- as_xml_attr(fill_id)
  write_xf(z)
}

#' internal function to set font to a style
#' @param xf_node some xf node
#' @param font_id some numeric value as character
#' @noRd
set_font <- function(xf_node, font_id) {
  z <- read_xf(read_xml(xf_node))
  z$applyFont <- "1"
  z$fontId <- as_xml_attr(font_id)
  write_xf(z)
}

#' internal function to set numfmt to a style
#' @param xf_node some xf node
#' @param numfmt_id some numeric value as character
#' @noRd
set_numfmt <- function(xf_node, numfmt) {
  z <- read_xf(read_xml(xf_node))
  z$applyNumberFormat <- "1"
  z$numFmtId <- as_xml_attr(numfmt)
  write_xf(z)
}

#' internal function to set cellstyle
#' @noRd
set_cellstyle <- function(
  xf_node,
  applyAlignment,
  applyBorder,
  applyFill,
  applyFont,
  applyNumberFormat,
  applyProtection,
  borderId,
  extLst,
  fillId,
  fontId,
  hidden,
  horizontal,
  indent,
  justifyLastLine,
  locked,
  numFmtId,
  pivotButton,
  quotePrefix,
  readingOrder,
  relativeIndent,
  shrinkToFit,
  textRotation,
  vertical,
  wrapText,
  xfId
) {
  z <- read_xf(read_xml(xf_node))

  if (!is.null(applyAlignment))    z$applyAlignment    <- as_xml_attr(applyAlignment)
  if (!is.null(applyBorder))       z$applyBorder       <- as_xml_attr(applyBorder)
  if (!is.null(applyFill))         z$applyFill         <- as_xml_attr(applyFill)
  if (!is.null(applyFont))         z$applyFont         <- as_xml_attr(applyFont)
  if (!is.null(applyNumberFormat)) z$applyNumberFormat <- as_xml_attr(applyNumberFormat)
  if (!is.null(applyProtection))   z$applyProtection   <- as_xml_attr(applyProtection)
  if (!is.null(borderId))          z$borderId          <- as_xml_attr(borderId)
  if (!is.null(extLst))            z$extLst            <- as_xml_attr(extLst)
  if (!is.null(fillId))            z$fillId            <- as_xml_attr(fillId)
  if (!is.null(fontId))            z$fontId            <- as_xml_attr(fontId)
  if (!is.null(hidden))            z$hidden            <- as_xml_attr(hidden)
  if (!is.null(horizontal))        z$horizontal        <- as_xml_attr(horizontal)
  if (!is.null(indent))            z$indent            <- as_xml_attr(indent)
  if (!is.null(justifyLastLine))   z$justifyLastLine   <- as_xml_attr(justifyLastLine)
  if (!is.null(locked))            z$locked            <- as_xml_attr(locked)
  if (!is.null(numFmtId))          z$numFmtId          <- as_xml_attr(numFmtId)
  if (!is.null(pivotButton))       z$pivotButton       <- as_xml_attr(pivotButton)
  if (!is.null(quotePrefix))       z$quotePrefix       <- as_xml_attr(quotePrefix)
  if (!is.null(readingOrder))      z$readingOrder      <- as_xml_attr(readingOrder)
  if (!is.null(relativeIndent))    z$relativeIndent    <- as_xml_attr(relativeIndent)
  if (!is.null(shrinkToFit))       z$shrinkToFit       <- as_xml_attr(shrinkToFit)
  if (!is.null(textRotation))      z$textRotation      <- as_xml_attr(textRotation)
  if (!is.null(vertical))          z$vertical          <- as_xml_attr(vertical)
  if (!is.null(wrapText))          z$wrapText          <- as_xml_attr(wrapText)
  if (!is.null(xfId))              z$xfId              <- as_xml_attr(xfId)

  write_xf(z)
}

#' get all styles on a sheet
#'
#' @param wb workbook
#' @param sheet worksheet
#'
#' @export
styles_on_sheet <- function(wb, sheet) {
  sheet_id <- wb_validate_sheet(wb, sheet)
  z <- unique(wb$worksheets[[sheet_id]]$sheet_data$cc$c_s)
  as.numeric(z)
}

#' get xml node for a specific style of a cell. function for internal use
#' @param wb workbook
#' @param sheet worksheet
#' @param cell cell
#' @noRd
get_cell_styles <- function(wb, sheet, cell) {
  cellstyles <- wb$get_cell_style(sheet, cell)

  out <- NULL
  for (cellstyle in cellstyles) {
    if (cellstyle == "") {
      tmp <- wb$styles_mgr$styles$cellXfs[1]
    } else {
      tmp <- wb$styles_mgr$styles$cellXfs[as.numeric(cellstyle) + 1]
    }

    out <- c(out, tmp)
  }

  out
}

#' Create a custom formatting style
#'
#' @description
#' Create a new style to apply to worksheet cells. Created styles have to be
#' assigned to a workbook to use them
#'
#' @details
#' It is possible to override border_color and border_style with \{left, right, top, bottom\}_color, \{left, right, top, bottom\}_style.
#'
#' @seealso [wb_add_style()]
# TODO maybe font_name,font_size could be documented together.
#' @param font_name A name of a font. Note the font name is not validated.
#'   If `font_name` is `NULL`, the workbook `base_font` is used. (Defaults to Calibri), see [wb_get_base_font()]
#' @param font_color Color of text in cell.  A valid hex color beginning with "#"
#'   or one of colors(). If `font_color` is NULL, the workbook base font colors is used.
#' (Defaults to black)
#' @param font_size Font size. A numeric greater than 0.
#'   By default, the workbook base font size is used. (Defaults to 11)
#' @param num_fmt Cell formatting. Some custom openxml format
#' @param border `NULL` or `TRUE`
#' @param border_color "black"
#' @param border_style "thin"
#' @param bg_fill Cell background fill color.
#' @param fg_color Cell foreground fill color.
#' @param gradient_fill An xml string beginning with `<gradientFill>` ...
#' @param text_bold bold
#' @param text_strike strikeout
#' @param text_italic italic
#' @param text_underline underline 1, true, single or double
#' @param ... Additional arguments
#' @return A dxfs style node
#' @examples
#' # do not apply anything
#' style1 <- create_dxfs_style()
#'
#' # change font color and background color
#' style2 <- create_dxfs_style(
#'   font_color = wb_color(hex = "FF9C0006"),
#'   bgFill = wb_color(hex = "FFFFC7CE")
#' )
#'
#' # change font (type, size and color) and background
#' # the old default in openxlsx and openxlsx2 <= 0.3
#' style3 <- create_dxfs_style(
#'   font_name = "Calibri",
#'   font_size = 11,
#'   font_color = wb_color(hex = "FF9C0006"),
#'   bgFill = wb_color(hex = "FFFFC7CE")
#' )
#'
#' ## See package vignettes for further examples
#' @export
create_dxfs_style <- function(
    font_name      = NULL,
    font_size      = NULL,
    font_color     = NULL,
    num_fmt        = NULL,
    border         = NULL,
    border_color   = wb_color(getOption("openxlsx2.borderColor", "black")),
    border_style   = getOption("openxlsx2.borderStyle", "thin"),
    bg_fill        = NULL,
    fg_color       = NULL,
    gradient_fill  = NULL,
    text_bold      = NULL,
    text_strike    = NULL,
    text_italic    = NULL,
    text_underline = NULL, # "true" or "double"
    ...
) {

  arguments <- c(..., ls(),
                 "left_color", "left_style", "right_color", "right_style",
                 "top_color", "top_style", "bottom_color", "bottom_style"
  )
  standardize(..., arguments = arguments)

  if (is.null(font_color)) font_color <- ""
  if (is.null(font_size)) font_size <- ""
  if (is.null(text_bold)) text_bold <- ""
  if (is.null(text_strike)) text_strike <- ""
  if (is.null(text_italic)) text_italic <- ""
  if (is.null(text_underline)) text_underline <- ""

  # found numFmtId=3 in MS365 xml not sure if this should be increased
  if (!is.null(num_fmt)) num_fmt <- create_numfmt(3, num_fmt)

  font <- create_font(
    color = font_color, name = font_name,
    sz = as.character(font_size),
    b = text_bold, i = text_italic, strike = text_strike,
    u = text_underline,
    family = "", scheme = ""
  )

  if (exists("pattern_type")) {
    pattern_type <- pattern_type
  } else {
    pattern_type <- "solid"
  }

  if (!is.null(bg_fill) && !all(bg_fill == "") || !is.null(gradient_fill)) {
    if (is.null(gradient_fill)) {
      # gradientFill is an xml string
      gradient_fill <- ""
    } else {
      pattern_type  <- ""
    }
    fill <- create_fill(patternType = pattern_type, bgColor = bg_fill, fgColor = fg_color, gradientFill = gradient_fill)
  } else {
    fill <- NULL
  }

  # untested
  if (!is.null(border)) {
    left_color   <- if (exists("left_color"))   left_color   else border_color
    left_style   <- if (exists("left_style"))   left_style   else border_style
    right_color  <- if (exists("right_color"))  right_color  else border_color
    right_style  <- if (exists("right_style"))  right_style  else border_style
    top_color    <- if (exists("top_color"))    top_color    else border_color
    top_style    <- if (exists("top_style"))    top_style    else border_style
    bottom_color <- if (exists("bottom_color")) bottom_color else border_color
    bottom_style <- if (exists("bottom_style")) bottom_style else border_style

    border <- create_border(
      left         = left_style,
      left_color   = left_color,
      right        = right_style,
      right_color  = right_color,
      top          = top_style,
      top_color    = top_color,
      bottom       = bottom_style,
      bottom_color = bottom_color
    )
  }

  xml_node_create(
    "dxf",
    xml_children = c(
      font,
      num_fmt,
      fill,
      border
    )
  )

}

#' create tableStyle
#'
#' Create a custom table style.
#' This function is for expert use only. Use other styling functions instead.
#'
#' @param name name
#' @param whole_table wholeTable
#' @param header_row headerRow
#' @param total_row totalRow
#' @param first_column firstColumn
#' @param last_column lastColumn
#' @param first_row_stripe firstRowStripe
#' @param second_row_stripe secondRowStripe
#' @param first_column_stripe firstColumnStripe
#' @param second_column_stripe secondColumnStripe
#' @param first_header_cell firstHeaderCell
#' @param last_header_cell lastHeaderCell
#' @param first_total_cell firstTotalCell
#' @param last_total_cell lastTotalCell
#' @param ... additional arguments
#' @export
create_tablestyle <- function(
    name,
    whole_table          = NULL,
    header_row           = NULL,
    total_row            = NULL,
    first_column         = NULL,
    last_column          = NULL,
    first_row_stripe     = NULL,
    second_row_stripe    = NULL,
    first_column_stripe  = NULL,
    second_column_stripe = NULL,
    first_header_cell    = NULL,
    last_header_cell     = NULL,
    first_total_cell     = NULL,
    last_total_cell      = NULL,
    ...
) {

  standardize_case_names(...)

  tab_wholeTable <- NULL
  if (length(whole_table)) {
    tab_wholeTable <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "wholeTable", dxfId = whole_table))
  }
  tab_headerRow <- NULL
  if (length(header_row)) {
    tab_headerRow <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "headerRow", dxfId = header_row))
  }
  tab_totalRow <- NULL
  if (length(total_row)) {
    tab_totalRow <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "totalRow", dxfId = total_row))
  }
  tab_firstColumn <- NULL
  if (length(first_column)) {
    tab_firstColumn <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "firstColumn", dxfId = first_column))
  }
  tab_lastColumn <- NULL
  if (length(last_column)) {
    tab_lastColumn <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "lastColumn", dxfId = last_column))
  }
  tab_firstRowStripe <- NULL
  if (length(first_row_stripe)) {
    tab_firstRowStripe <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "firstRowStripe", dxfId = first_row_stripe))
  }
  tab_secondRowStripe <- NULL
  if (length(second_row_stripe)) {
    tab_secondRowStripe <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "secondRowStripe", dxfId = second_row_stripe))
  }
  tab_firstColumnStripe <- NULL
  if (length(first_column_stripe)) {
    tab_firstColumnStripe <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "firstColumnStripe", dxfId = first_column_stripe))
  }
  tab_secondColumnStripe <- NULL
  if (length(second_column_stripe)) {
    tab_secondColumnStripe <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "secondColumnStripe", dxfId = second_column_stripe))
  }
  tab_firstHeaderCell <- NULL
  if (length(first_header_cell)) {
    tab_firstHeaderCell <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "firstHeaderCell", dxfId = first_header_cell))
  }
  tab_lastHeaderCell <- NULL
  if (length(last_header_cell)) {
    tab_lastHeaderCell <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "lastHeaderCell", dxfId = last_header_cell))
  }
  tab_firstTotalCell <- NULL
  if (length(first_total_cell)) {
    tab_firstTotalCell <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "firstTotalCell", dxfId = first_total_cell))
  }
  tab_lastTotalCell <- NULL
  if (length(last_total_cell)) {
    tab_lastTotalCell <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "lastTotalCell", dxfId = last_total_cell))
  }

  xml_elements <- c(
    tab_wholeTable,
    tab_headerRow,
    tab_totalRow,
    tab_firstColumn,
    tab_lastColumn,
    tab_firstRowStripe,
    tab_secondRowStripe,
    tab_firstColumnStripe,
    tab_secondColumnStripe,
    tab_firstHeaderCell,
    tab_lastHeaderCell,
    tab_firstTotalCell,
    tab_lastTotalCell
  )

  rand_str <- random_string(length = 12, pattern = "[A-Z0-9]")

  xml_node_create(
    "tableStyle",
    xml_attributes = c(
      name      = name,
      pivot     = "0", # pivot uses different styles
      count     = as_xml_attr(length(xml_elements)),
       # possible st_guid()
      `xr9:uid` = sprintf("{CE23B8CA-E823-724F-9713-%s}", rand_str)
    ),
    xml_children = xml_elements
  )
}
