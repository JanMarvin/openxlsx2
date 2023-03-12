
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
#' create border
#' @description
#' Border styles can any of the following: "thin", "thick", "slantDashDot", "none", "mediumDashed", "mediumDashDot", "medium", "hair", "double", "dotted", "dashed", "dashedDotDot", "dashDot"
#' Border colors are of the following type: c(rgb="FF000000")
#' @param diagonalDown x
#' @param diagonalUp x
#' @param outline x
#' @param bottom X
#' @param bottom_color X
#' @param diagonal X
#' @param diagonal_color X,
#' @param end x,
#' @param horizontal x
#' @param left x
#' @param left_color x
#' @param right x
#' @param right_color x
#' @param start x
#' @param top x
#' @param top_color x
#' @param vertical x
#' @param ... x
#'
#' @export
create_border <- function(
    diagonalDown = "",
    diagonalUp = "",
    outline = "",
    bottom = NULL,
    bottom_color = NULL,
    diagonal = NULL,
    diagonal_color = NULL,
    end = "",
    horizontal = "",
    left = NULL,
    left_color = NULL,
    right = NULL,
    right_color = NULL,
    start = "",
    top = NULL,
    top_color = NULL,
    vertical = "",
    ...
) {

  standardize_color_names(...)

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

  df_border <- data.frame(
    start = start,
    end = end,
    left = left,
    right = right,
    top = top,
    bottom = bottom,
    diagonalDown = diagonalDown,
    diagonalUp = diagonalUp,
    diagonal = diagonal,
    vertical = vertical,
    horizontal = horizontal,
    outline = outline, # unknown position in border
    stringsAsFactors = FALSE
  )
  border <- write_border(df_border)

  return(border)
}

#' create number format
#' @param numFmtId an id
#' @param formatCode a format code
#' @export
create_numfmt <- function(numFmtId, formatCode) {

  df_numfmt <- data.frame(
    numFmtId = as_xml_attr(numFmtId),
    formatCode  = as_xml_attr(formatCode),
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
#' @param vertAlign vertical alignment
#' @param ... ...
#' @examples
#' font <- create_font()
#' # openxml has the alpha value leading
#' hex8 <- unlist(xml_attr(read_xml(font), "font", "color"))
#' hex8 <- paste0("#", substr(hex8, 3, 8), substr(hex8, 1,2))
#'
#' # # write test color
#' # col <- crayon::make_style(col2rgb(hex8, alpha = TRUE))
#' # cat(col("Test"))
#' @export
create_font <- function(
    b = "",
    charset = "",
    color = wb_color(hex = "FF000000"),
    condense = "",
    extend = "",
    family = "2",
    i = "",
    name = "Calibri",
    outline = "",
    scheme = "minor",
    shadow = "",
    strike = "",
    sz = "11",
    u = "",
    vertAlign = "",
    ...
) {

  standardize_color_names(...)

  if (b != "") {
    b <- xml_node_create("b", xml_attributes = c("val" = as_xml_attr(b)))
  }

  if (charset != "") {
    charset <- xml_node_create("charset", xml_attributes = c("val" = charset))
  }

  if (!is.null(color) && all(color != "")) {
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

  if (vertAlign != "") {
    vertAlign <- xml_node_create("vertAlign", xml_attributes = c("val" = as_xml_attr(vertAlign)))
  }

  df_font <- data.frame(
    b = b,
    charset = charset,
    color = color,
    condense = condense,
    extend = extend,
    family = family,
    i = i,
    name = name,
    outline = outline,
    scheme = scheme,
    shadow = shadow,
    strike = strike,
    sz = sz,
    u = u,
    vertAlign = vertAlign,
    stringsAsFactors = FALSE
  )
  font <- write_font(df_font)

  return(font)
}

#' create fill
#'
#' @param gradientFill complex fills
#' @param patternType various: default is "none", but also "solid", or a color like "gray125"
#' @param bgColor hex8 color with alpha, red, green, blue only for patternFill
#' @param fgColor hex8 color with alpha, red, green, blue only for patternFill
#' @param ... ...
#'
#' @export
create_fill <- function(
    gradientFill = "",
    patternType = "",
    bgColor = NULL,
    fgColor = NULL,
    ...
) {

  standardize_color_names(...)

  if (!is.null(bgColor) && all(bgColor != "")) {
    bgColor <- xml_node_create("bgColor", xml_attributes = bgColor)
  }

  if (!is.null(fgColor) && all(fgColor != "")) {
    fgColor <- xml_node_create("fgColor", xml_attributes = fgColor)
  }

  # if gradient fill is specified we can not have patternFill too. otherwise
  # we end up with a solid black fill
  if (gradientFill == "") {
    patternFill <- xml_node_create("patternFill",
                                   xml_children = c(bgColor, fgColor),
                                   xml_attributes = c(patternType = patternType))
  } else {
    patternFill <- ""
  }

  df_fill <- data.frame(
    gradientFill = gradientFill,
    patternFill = patternFill,
    stringsAsFactors = FALSE
  )
  fill <- write_fill(df_fill)

  return(fill)
}

# TODO can be further generalized with additional xf attributes and children
#' create_cell_style
#' @param borderId dummy
#' @param fillId dummy
#' @param fontId dummy
#' @param numFmtId a numFmt ID for a builtin style
#' @param pivotButton dummy
#' @param quotePrefix dummy
#' @param xfId dummy
#' @param horizontal dummy
#' @param indent dummy
#' @param justifyLastLine dummy
#' @param readingOrder dummy
#' @param relativeIndent dummy
#' @param shrinkToFit dummy
#' @param textRotation dummy
#' @param vertical dummy
#' @param wrapText dummy
#' @param extLst dummy
#' @param hidden dummy
#' @param locked dummy
#' @param horizontal alignment can be "", "center", "right"
#' @param vertical alignment can be "", "center", "right"
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
    borderId = "",
    fillId = "",
    fontId = "",
    numFmtId = "",
    pivotButton = "",
    quotePrefix = "",
    xfId = "",
    horizontal = "",
    indent = "",
    justifyLastLine = "",
    readingOrder = "",
    relativeIndent = "",
    shrinkToFit = "",
    textRotation = "",
    vertical = "",
    wrapText = "",
    extLst = "",
    hidden = "",
    locked = ""
) {
  n <- length(numFmtId)

  applyAlignment <- ""
  if (any(horizontal != "") || any(textRotation != "") || any(vertical != "")) applyAlignment <- "1"

  applyBorder <- ""
  if (any(borderId != "")) applyBorder <- "1"

  applyFill <- ""
  if (any(fillId != "")) applyFill <- "1"

  applyFont <- ""
  if (any(fontId != "")) applyFont <- "1"

  applyNumberFormat <- ""
  if (any(numFmtId != "")) applyNumberFormat <- "1"

  applyProtection <- ""
  if (any(hidden != "") || any(locked != "")) applyProtection <- "1"


  df_cellXfs <- data.frame(
    applyAlignment = rep(applyAlignment, n),
    applyBorder = rep(applyBorder, n),
    applyFill = rep(applyFill, n),
    applyFont = rep(applyFont, n),
    applyNumberFormat = rep(applyNumberFormat, n),
    applyProtection = rep(applyProtection, n),

    borderId = rep(as_xml_attr(borderId), n),
    fillId = rep(as_xml_attr(fillId), n),
    fontId = rep(as_xml_attr(fontId), n),
    numFmtId = as_xml_attr(numFmtId),
    pivotButton = rep(as_xml_attr(pivotButton), n),
    quotePrefix = rep(as_xml_attr(quotePrefix), n),
    xfId = rep(as_xml_attr(xfId), n),
    horizontal = rep(as_xml_attr(horizontal), n),
    indent = rep(as_xml_attr(indent), n),
    justifyLastLine = rep(as_xml_attr(justifyLastLine), n),
    readingOrder = rep(as_xml_attr(readingOrder), n),
    relativeIndent = rep(as_xml_attr(relativeIndent), n),
    shrinkToFit = rep(as_xml_attr(shrinkToFit), n),
    textRotation = rep(as_xml_attr(textRotation), n),
    vertical = rep(as_xml_attr(vertical), n),
    wrapText = rep(as_xml_attr(wrapText), n),
    extLst = rep(extLst, n),
    hidden = rep(as_xml_attr(hidden), n),
    locked = rep(as_xml_attr(locked), n),
    stringsAsFactors = FALSE
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

  if (!is.null(applyAlignment))    z$applyAlignment <- applyAlignment
  if (!is.null(applyBorder))       z$applyBorder <- applyBorder
  if (!is.null(applyFill))         z$applyFill <- applyFill
  if (!is.null(applyFont))         z$applyFont <- applyFont
  if (!is.null(applyNumberFormat)) z$applyNumberFormat <- applyNumberFormat
  if (!is.null(applyProtection))   z$applyProtection <- applyProtection
  if (!is.null(borderId))          z$borderId <- borderId
  if (!is.null(extLst))            z$extLst <- extLst
  if (!is.null(fillId))            z$fillId <- fillId
  if (!is.null(fontId))            z$fontId <- fontId
  if (!is.null(hidden))            z$hidden <- hidden
  if (!is.null(horizontal))        z$horizontal <- horizontal
  if (!is.null(indent))            z$indent <- indent
  if (!is.null(justifyLastLine))   z$justifyLastLine <- justifyLastLine
  if (!is.null(locked))            z$locked <- locked
  if (!is.null(numFmtId))          z$numFmtId <- numFmtId
  if (!is.null(pivotButton))       z$pivotButton <- pivotButton
  if (!is.null(quotePrefix))       z$quotePrefix <- quotePrefix
  if (!is.null(readingOrder))      z$readingOrder <- readingOrder
  if (!is.null(relativeIndent))    z$relativeIndent <- relativeIndent
  if (!is.null(shrinkToFit))       z$shrinkToFit <- shrinkToFit
  if (!is.null(textRotation))      z$textRotation <- textRotation
  if (!is.null(vertical))          z$vertical <- vertical
  if (!is.null(wrapText))          z$wrapText <- wrapText
  if (!is.null(xfId))              z$xfId <- xfId

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
    if (cellstyle == "")
      tmp <- wb$styles_mgr$styles$cellXfs[1]
    else
      tmp <- wb$styles_mgr$styles$cellXfs[as.numeric(cellstyle) + 1]

    out <- c(out, tmp)
  }

  out
}

#' @name create_dxfs_style
#' @title Create a custom formatting style
#' @description Create a new style to apply to worksheet cells. Created styles have to be
#' assigned to a workbook to use them
#' @param font_name A name of a font. Note the font name is not validated. If fontName is NULL,
#' the workbook base font is used. (Defaults to Calibri)
#' @param font_color Color of text in cell.  A valid hex color beginning with "#"
#' or one of colors(). If fontColor is NULL, the workbook base font colors is used.
#' (Defaults to black)
#' @param font_size Font size. A numeric greater than 0.
#' If fontSize is NULL, the workbook base font size is used. (Defaults to 11)
#' @param numFmt Cell formatting. Some custom openxml format
#' @param border NULL or TRUE
#' @param border_color "black"
#' @param border_style "thin"
#' @param bgFill Cell background fill color.
#' @param text_bold bold
#' @param text_strike strikeout
#' @param text_italic italic
#' @param text_underline underline 1, true, single or double
#' @param ... ...
#' @return A dxfs style node
#' @seealso [wb_add_style()]
#' @examples
#' # do not apply anthing
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
    numFmt         = NULL,
    border         = NULL,
    border_color   = wb_color(getOption("openxlsx2.borderColor", "black")),
    border_style   = getOption("openxlsx2.borderStyle", "thin"),
    bgFill         = NULL,
    text_bold      = NULL,
    text_strike    = NULL,
    text_italic    = NULL,
    text_underline = NULL, # "true" or "double"
    ...
) {

  standardize_color_names(...)

  if (is.null(font_color)) font_color <- ""
  if (is.null(font_size)) font_size <- ""
  if (is.null(text_bold)) text_bold <- ""
  if (is.null(text_strike)) text_strike <- ""
  if (is.null(text_italic)) text_italic <- ""
  if (is.null(text_underline)) text_underline <- ""

  # found numFmtId=3 in MS365 xml not sure if this should be increased
  if (!is.null(numFmt)) numFmt <- create_numfmt(3, numFmt)

  font <- create_font(color = font_color, name = font_name,
                      sz = as.character(font_size),
                      b = text_bold, i = text_italic, strike = text_strike,
                      u = text_underline,
                      family = "", scheme = "")

  if (!is.null(bgFill) && bgFill != "")
    fill <- create_fill(patternType = "solid", bgColor = bgFill)
  else
    fill <- NULL

  # untested
  if (!is.null(border))
    border <- create_border(left = border_style,
                            left_color = border_color,
                            right = border_style,
                            right_color = border_color,
                            top = border_style,
                            top_color = border_color,
                            bottom = border_style,
                            bottom_color = border_color)

  xml_node_create(
    "dxf",
    xml_children = c(
      font,
      numFmt,
      fill,
      border
    )
  )

}
