# functions interacting with styles

# internal function used in wb_load
# @param x character string containing styles.xml
import_styles <- function(x) {
  sxml <- read_xml(x)

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

#' Create border format
#'
#' @description
#' This function creates border styles for a cell in a spreadsheet. Border styles can be any of the following: "none", "thin", "medium", "dashed", "dotted", "thick", "double", "hair", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot". Border colors can be created with [wb_color()].
#'
#' @param bottom,left,right,top,diagonal Character, the style of the border.
#' @param bottom_color,left_color,right_color,top_color,diagonal_color,start_color,end_color,horizontal_color,vertical_color a [wb_color()], the color of the border.
#' @param diagonal_down,diagonal_up Logical, whether the diagonal border goes from the bottom left to the top right, or top left to bottom right.
#' @param outline Logical, whether the border is.
#' @param horizontal,vertical Character, the style of the inner border (only for dxf objects).
#' @param start,end leading and trailing edge of a border.
#' @param ... Additional arguments passed to other methods.
#'
#' @return A formatted border object to be used in a spreadsheet.
#'
#' @seealso [wb_add_border()]
#' @family style creating functions
#'
#' @examples
#' # Create a border with a thick bottom and thin top
#' border <- create_border(
#'   bottom = "thick",
#'   bottom_color = wb_color("FF0000"),
#'   top = "thin",
#'   top_color = wb_color("00FF00")
#' )
#'
#' @export
create_border <- function(
    diagonal_down    = "",
    diagonal_up      = "",
    outline          = "",
    bottom           = NULL,
    bottom_color     = NULL,
    diagonal         = NULL,
    diagonal_color   = NULL,
    end              = "",
    horizontal       = "",
    left             = NULL,
    left_color       = NULL,
    right            = NULL,
    right_color      = NULL,
    start            = "",
    top              = NULL,
    top_color        = NULL,
    vertical         = "",
    start_color      = NULL,
    end_color        = NULL,
    horizontal_color = NULL,
    vertical_color   = NULL,
    ...
) {

  # sml_CT_Border
  standardize(...)

  assert_xml_bool(outline)

  assert_class(left_color,       "wbColour", or_null = TRUE)
  assert_class(right_color,      "wbColour", or_null = TRUE)
  assert_class(top_color,        "wbColour", or_null = TRUE)
  assert_class(bottom_color,     "wbColour", or_null = TRUE)
  assert_class(diagonal_color,   "wbColour", or_null = TRUE)
  assert_class(start_color,      "wbColour", or_null = TRUE)
  assert_class(end_color,        "wbColour", or_null = TRUE)
  assert_class(horizontal_color, "wbColour", or_null = TRUE)
  assert_class(vertical_color,   "wbColour", or_null = TRUE)

  if (!is.null(left_color))       left_color       <- xml_node_create("color", xml_attributes = left_color)
  if (!is.null(right_color))      right_color      <- xml_node_create("color", xml_attributes = right_color)
  if (!is.null(top_color))        top_color        <- xml_node_create("color", xml_attributes = top_color)
  if (!is.null(bottom_color))     bottom_color     <- xml_node_create("color", xml_attributes = bottom_color)
  if (!is.null(diagonal_color))   diagonal_color   <- xml_node_create("color", xml_attributes = diagonal_color)
  if (!is.null(start_color))      start_color      <- xml_node_create("color", xml_attributes = start_color)
  if (!is.null(end_color))        end_color        <- xml_node_create("color", xml_attributes = end_color)
  if (!is.null(horizontal_color)) horizontal_color <- xml_node_create("color", xml_attributes = horizontal_color)
  if (!is.null(vertical_color))   vertical_color   <- xml_node_create("color", xml_attributes = vertical_color)

  valid_borders <- c("none",  "thin",  "medium",  "dashed",  "dotted",  "thick",  "double",  "hair",  "mediumDashed",  "dashDot",  "mediumDashDot",  "dashDotDot",  "mediumDashDotDot", "slantDashDot", "")
  borders       <- c(left, right, top, bottom, diagonal, start, end, horizontal, vertical)
  match.arg_wrapper(borders, valid_borders, fn_name = "create_border")

  # excel dies on style=\"\"
  if (!is.null(left))       left       <- c(style = left)
  if (!is.null(right))      right      <- c(style = right)
  if (!is.null(top))        top        <- c(style = top)
  if (!is.null(bottom))     bottom     <- c(style = bottom)
  if (!is.null(diagonal))   diagonal   <- c(style = diagonal)
  if (!is.null(start))      start      <- c(style = start)
  if (!is.null(end))        end        <- c(style = end)
  if (!is.null(horizontal)) horizontal <- c(style = horizontal)
  if (!is.null(vertical))   vertical   <- c(style = vertical)

  left       <- xml_node_create("left",       xml_children = left_color,       xml_attributes = left)
  right      <- xml_node_create("right",      xml_children = right_color,      xml_attributes = right)
  top        <- xml_node_create("top",        xml_children = top_color,        xml_attributes = top)
  bottom     <- xml_node_create("bottom",     xml_children = bottom_color,     xml_attributes = bottom)
  diagonal   <- xml_node_create("diagonal",   xml_children = diagonal_color,   xml_attributes = diagonal)
  start      <- xml_node_create("start",      xml_children = diagonal_color,   xml_attributes = start)
  end        <- xml_node_create("end",        xml_children = diagonal_color,   xml_attributes = end)
  horizontal <- xml_node_create("horizontal", xml_children = horizontal_color, xml_attributes = horizontal)
  vertical   <- xml_node_create("vertical",   xml_children = vertical_color,   xml_attributes = vertical)

  if (left       == "<left/>")       left       <- ""
  if (right      == "<right/>")      right      <- ""
  if (top        == "<top/>")        top        <- ""
  if (bottom     == "<bottom/>")     bottom     <- ""
  if (diagonal   == "<diagonal/>")   diagonal   <- ""
  if (start      == "<start/>")      start      <- "" # in the ecma spec this is called <begin/>
  if (end        == "<end/>")        end        <- ""
  if (horizontal == "<horizontal/>") horizontal <- ""
  if (vertical   == "<vertical/>")   vertical   <- ""

  df_border <- data.frame(
    start            = as_xml_attr(start),
    end              = as_xml_attr(end),
    left             = left,
    right            = right,
    top              = top,
    bottom           = bottom,
    diagonalDown     = as_xml_attr(diagonal_down),
    diagonalUp       = as_xml_attr(diagonal_up),
    diagonal         = diagonal,
    vertical         = vertical,
    horizontal       = horizontal,
    outline          = as_xml_attr(outline),      # unknown position in border
    stringsAsFactors = FALSE
  )

  write_border(df_border)
}

#' Create number format
#'
#' @description
#' This function creates a number format for a cell in a spreadsheet. Number formats define how numeric values are displayed, including dates, times, currencies, percentages, and more.
#'
#' @param numFmtId An ID representing the number format. The list of valid IDs can be found in the **Details** section of [create_cell_style()].
#' @param formatCode A format code that specifies the display format for numbers. This can include custom formats for dates, times, and other numeric values.
#'
#' @return A formatted number format object to be used in a spreadsheet.
#'
#' @seealso [wb_add_numfmt()]
#' @family style creating functions
#'
#' @examples
#' # Create a number format for currency
#' numfmt <- create_numfmt(
#'   numFmtId = 164,
#'   formatCode = "$#,##0.00"
#' )
#'
#' @export
create_numfmt <- function(numFmtId = 164, formatCode = "#,##0.00") {

  formatCode <- escape_format_code(formatCode)

  df_numfmt <- data.frame(
    numFmtId         = as_xml_attr(numFmtId),
    formatCode       = as_xml_attr(formatCode),
    stringsAsFactors = FALSE
  )
  write_numfmt(df_numfmt)
}


#' Create font format
#'
#' @description
#' This function creates font styles for a cell in a spreadsheet. It allows customization of various font properties including bold, italic, color, size, underline, and more.
#'
#' @param b Logical, whether the font should be bold.
#' @param charset Character, the character set to be used. The list of valid IDs can be found in the **Details** section of [fmt_txt()].
#' @param color A [wb_color()], the color of the font. Default is "FF000000".
#' @param condense Logical, whether the font should be condensed.
#' @param extend Logical, whether the font should be extended.
#' @param family Character, the font family. Default is "2" (modern). "0" (auto), "1" (roman), "2" (swiss), "3" (modern), "4" (script), "5" (decorative). # 6-14 unused
#' @param i Logical, whether the font should be italic.
#' @param name Character, the name of the font. Default is "Aptos Narrow".
#' @param outline Logical, whether the font should have an outline.
#' @param scheme Character, the font scheme. Valid values are "minor", "major", "none". Default is "minor".
#' @param shadow Logical, whether the font should have a shadow.
#' @param strike Logical, whether the font should have a strikethrough.
#' @param sz Character, the size of the font. Default is "11".
#' @param u Character, the underline style. Valid values are "single", "double", "singleAccounting", "doubleAccounting", "none".
#' @param vert_align Character, the vertical alignment of the font. Valid values are "baseline", "superscript", "subscript".
#' @param ... Additional arguments passed to other methods.
#'
#' @return A formatted font object to be used in a spreadsheet.
#'
#' @seealso [wb_add_font()]
#' @family style creating functions
#'
#' @examples
#' # Create a font with bold and italic styles
#' font <- create_font(
#'   b = TRUE,
#'   i = TRUE,
#'   color = wb_color(hex = "FF00FF00"),
#'   name = "Arial",
#'   sz = "12"
#' )
#'
#' # openxml has the alpha value leading
#' hex8 <- unlist(xml_attr(read_xml(font), "font", "color"))
#' hex8 <- paste0("#", substr(hex8, 3, 8), substr(hex8, 1, 2))
#'
#' # # write test color
#' # col <- crayon::make_style(col2rgb(hex8, alpha = TRUE))
#' # cat(col("Test"))
#'
#' @export
create_font <- function(
    b          = "",
    charset    = "",
    color      = wb_color(hex = "FF000000"),
    condense   = "",
    extend     = "",
    family     = "2",
    i          = "",
    name       = "Aptos Narrow",
    outline    = "",
    scheme     = "minor",
    shadow     = "",
    strike     = "",
    sz         = "11",
    u          = "",
    vert_align = "",
    ...
) {

  # sml_CT_RPrElt
  standardize(...)

  if (b != "") {
    assert_xml_bool(b)
    b <- xml_node_create("b", xml_attributes = c("val" = as_xml_attr(b)))
  }

  if (charset != "") {
    charset <- xml_node_create("charset", xml_attributes = c("val" = as_xml_attr(charset)))
  }

  if (!is.null(color) && !all(color == "")) {
    # alt xml_attributes(theme:)
    assert_class(color, "wbColour")
    color <- xml_node_create("color", xml_attributes = color)
  }

  if (condense != "") {
    assert_xml_bool(condense)
    condense <- xml_node_create("condense", xml_attributes = c("val" = as_xml_attr(condense)))
  }

  if (extend != "") {
    assert_xml_bool(extend)
    extend <- xml_node_create("extend", xml_attributes = c("val" = as_xml_attr(extend)))
  }

  if (family != "") {
    if (!family %in% as.character(0:14))
      stop("family needs to be in the range of 0 to 14", call. = FALSE)
    family <- xml_node_create("family", xml_attributes = c("val" = family))
  }

  if (i != "") {
    assert_xml_bool(i)
    i <- xml_node_create("i", xml_attributes = c("val" = as_xml_attr(i)))
  }

  if (is.null(name)) name <- ""
  if (name != "") {
    name <- xml_node_create("name", xml_attributes = c("val" = name))
  }

  if (outline != "") {
    assert_xml_bool(outline)
    outline <- xml_node_create("outline", xml_attributes = c("val" = as_xml_attr(outline)))
  }

  if (scheme != "") {
    valid_scheme <- c("minor", "major", "none")
    match.arg_wrapper(scheme, valid_scheme, fn_name = "create_font")
    scheme <- xml_node_create("scheme", xml_attributes = c("val" = scheme))
  }

  if (shadow != "") {
    assert_xml_bool(shadow)
    shadow <- xml_node_create("shadow", xml_attributes = c("val" = as_xml_attr(shadow)))
  }

  if (strike != "") {
    assert_xml_bool(strike)
    strike <- xml_node_create("strike", xml_attributes = c("val" = as_xml_attr(strike)))
  }

  if (sz != "") {
    sz <- xml_node_create("sz", xml_attributes = c("val" = as_xml_attr(sz)))
  }

  if (u != "") {
    valid_underlines <- c("single", "double", "singleAccounting", "doubleAccounting", "none")
    match.arg_wrapper(u, valid_underlines, fn_name = "create_font")
    u <- xml_node_create("u", xml_attributes = c("val" = as_xml_attr(u)))
  }

  if (vert_align != "") {
    valid_underlines <- c("baseline", "superscript", "subscript")
    match.arg_wrapper(vert_align, valid_underlines, fn_name = "create_font")
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

  font
}

#' Create fill pattern
#'
#' @description
#' This function creates fill patterns for a cell in a spreadsheet. Fill patterns can be simple solid colors or more complex gradient fills. For certain pattern types, two colors are needed.
#'
#' @param gradient_fill Character, specifying complex gradient fills.
#' @param pattern_type Character, specifying the fill pattern type. Valid values are "none" (default), "solid", "mediumGray", "darkGray", "lightGray", "darkHorizontal", "darkVertical", "darkDown", "darkUp", "darkGrid", "darkTrellis", "lightHorizontal", "lightVertical", "lightDown", "lightUp", "lightGrid", "lightTrellis", "gray125", "gray0625".
#' @param bg_color Character, specifying the background color in hex8 format (alpha, red, green, blue) for pattern fills.
#' @param fg_color Character, specifying the foreground color in hex8 format (alpha, red, green, blue) for pattern fills.
#' @param ... Additional arguments passed to other methods.
#'
#' @return A formatted fill pattern object to be used in a spreadsheet.
#'
#' @seealso [wb_add_fill()]
#' @family style creating functions
#'
#' @examples
#' # Create a solid fill pattern with foreground color
#' fill <- create_fill(
#'   pattern_type = "solid",
#'   fg_color = wb_color(hex = "FFFF0000")
#' )
#'
#' @export
create_fill <- function(
    gradient_fill = "",
    pattern_type  = "",
    bg_color      = NULL,
    fg_color      = NULL,
    ...
) {

  standardize(...)
  assert_class(bg_color, "wbColour", or_null = TRUE)
  assert_class(fg_color, "wbColour", or_null = TRUE)

  if (!is.null(bg_color) && !all(bg_color == "")) {
    bg_color <- xml_node_create("bgColor", xml_attributes = bg_color)
  }

  if (!is.null(fg_color) && !all(fg_color == "")) {
    fg_color <- xml_node_create("fgColor", xml_attributes = fg_color)
  }

  # if gradient fill is specified we can not have patternFill too. otherwise
  # we end up with a solid black fill
  if (gradient_fill == "") {
    if (pattern_type == "") pattern_type <- "none"
    valid_pattern <- c("none", "solid", "mediumGray", "darkGray", "lightGray", "darkHorizontal", "darkVertical", "darkDown", "darkUp", "darkGrid", "darkTrellis", "lightHorizontal", "lightVertical", "lightDown", "lightUp", "lightGrid", "lightTrellis", "gray125", "gray0625")
    match.arg_wrapper(pattern_type, valid_pattern, fn_name = "create_fill")

    pattern_fill <- xml_node_create("patternFill",
      xml_children   = c(fg_color, bg_color),
      xml_attributes = c(patternType = pattern_type)
    )
  } else {
    pattern_fill <- ""
  }

  df_fill <- data.frame(
    gradientFill     = gradient_fill,
    patternFill      = pattern_fill,
    stringsAsFactors = FALSE
  )

  write_fill(df_fill)
}

#' Create cell style
#'
#' @description
#' This function creates a cell style for a spreadsheet, including attributes such as borders, fills, fonts, and number formats.
#'
#' @param border_id,fill_id,font_id,num_fmt_id IDs for style elements.
#' @param pivot_button Logical parameter for the pivot button.
#' @param quote_prefix Logical parameter for the quote prefix. (This way a number in a character cell will not cause a warning).
#' @param xf_id Dummy parameter for the xf ID. (Used only with named format styles).
#' @param indent Integer parameter for the indent.
#' @param justify_last_line Logical for justifying the last line.
#' @param reading_order Logical parameter for reading order. 0 (Left to right; default) or 1 (right to left).
#' @param relative_indent Dummy parameter for relative indent.
#' @param shrink_to_fit Logical parameter for shrink to fit.
#' @param text_rotation Integer parameter for text rotation (-180 to 180).
#' @param wrap_text Logical parameter for wrap text. (Required for linebreaks).
#' @param ext_lst Dummy parameter for extension list.
#' @param hidden Logical parameter for hidden.
#' @param locked Logical parameter for locked. (Impacts the cell only).
#' @param horizontal Character, alignment can be '', 'general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'.
#' @param vertical Character, alignment can be '', 'top', 'center', 'bottom', 'justify', 'distributed'.
#' @param ... Reserved for additional arguments.
#'
#' @return A formatted cell style object to be used in a spreadsheet.
#'
#' @seealso [wb_add_cell_style()]
#' @family style creating functions
#'
#' @details
#' A single cell style can make use of various other styles like border, fill, and font. These styles are independent of the cell style and must be registered with the style manager separately.
#' This allows multiple cell styles to share a common font type, for instance. The used style elements are passed to the cell style via their IDs. An example of this can be seen below.
#' The number format can be a custom one created by [create_numfmt()], or a built-in style from the formats table below.
#'
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
#'
#' @examples
#' foo_fill <- create_fill(pattern_type = "lightHorizontal",
#'                         fg_color = wb_color("blue"),
#'                         bg_color = wb_color("orange"))
#' foo_font <- create_font(sz = 36, b = TRUE, color = wb_color("yellow"))
#'
#' wb <- wb_workbook()
#' wb$styles_mgr$add(foo_fill, "foo")
#' wb$styles_mgr$add(foo_font, "foo")
#'
#' foo_style <- create_cell_style(
#'   fill_id = wb$styles_mgr$get_fill_id("foo"),
#'   font_id = wb$styles_mgr$get_font_id("foo")
#' )
#'
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

  # CT_Xf
  n <- length(num_fmt_id)

  arguments <- c(ls(), "is_cell_style_xf")
  standardize_case_names(..., arguments = arguments)
  args <- list(...)

  is_cell_style_xf <- isTRUE(args$is_cell_style_xf)

  applyAlignment <- "" # CT_CellAlignment
  if (any(horizontal != "") || any(text_rotation != "") || any(vertical != "") || any(wrap_text != "") || any(indent != "") || any(relative_indent != "") || any(justify_last_line != "") || any(shrink_to_fit != "") || any(reading_order != "")) applyAlignment <- "1"
  if (is_cell_style_xf) applyAlignment <- "0"

  applyBorder <- ""
  if (any(border_id != "")) applyBorder <- "1"
  if (is_cell_style_xf && isTRUE(border_id == "")) applyBorder <- "0"

  applyFill <- ""
  if (any(fill_id != "")) applyFill <- "1"
  if (is_cell_style_xf && isTRUE(fill_id == "")) applyFill <- "0"

  applyFont <- ""
  if (any(font_id != "")) applyFont <- "1"
  if (is_cell_style_xf && isTRUE(font_id == "")) applyFont <- "0"

  applyNumberFormat <- ""
  if (any(num_fmt_id != "")) applyNumberFormat <- "1"
  if (is_cell_style_xf) applyNumberFormat <- "0"

  applyProtection <- ""
  if (any(hidden != "") || any(locked != "")) applyProtection <- "1"
  if (is_cell_style_xf) applyProtection <- "0"


  if (nchar(horizontal)) {
    valid_horizontal <- c("general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed")
    match.arg_wrapper(horizontal, valid_horizontal, fn_name = "create_cell_style")
  }

  if (nchar(vertical)) {
    valid_vertical <- c("top", "center", "bottom", "justify", "distributed")
    match.arg_wrapper(vertical, valid_vertical, fn_name = "create_cell_style")
  }

  assert_xml_bool(applyAlignment)
  assert_xml_bool(applyBorder)
  assert_xml_bool(applyFill)
  assert_xml_bool(applyFont)
  assert_xml_bool(applyNumberFormat)
  assert_xml_bool(applyProtection)
  assert_xml_bool(pivot_button)
  assert_xml_bool(quote_prefix)
  assert_xml_bool(justify_last_line)
  assert_xml_bool(shrink_to_fit)
  assert_xml_bool(wrap_text)
  assert_xml_bool(hidden)
  assert_xml_bool(locked)

  df_cellXfs <- data.frame(
    applyAlignment    = rep(applyAlignment, n), # bool
    applyBorder       = rep(applyBorder, n), # bool
    applyFill         = rep(applyFill, n), # bool
    applyFont         = rep(applyFont, n), # bool
    applyNumberFormat = rep(applyNumberFormat, n), # bool
    applyProtection   = rep(applyProtection, n), # bool
    borderId          = rep(as_xml_attr(border_id), n), # int
    fillId            = rep(as_xml_attr(fill_id), n), # int
    fontId            = rep(as_xml_attr(font_id), n), # int
    numFmtId          = as_xml_attr(num_fmt_id), # int
    pivotButton       = rep(as_xml_attr(pivot_button), n), # bool
    quotePrefix       = rep(as_xml_attr(quote_prefix), n), # bool
    xfId              = rep(as_xml_attr(xf_id), n), # int
    horizontal        = rep(as_xml_attr(horizontal), n),
    indent            = rep(as_xml_attr(indent), n), # int
    justifyLastLine   = rep(as_xml_attr(justify_last_line), n), # bool
    readingOrder      = rep(as_xml_attr(reading_order), n), # int
    relativeIndent    = rep(as_xml_attr(relative_indent), n), # int
    shrinkToFit       = rep(as_xml_attr(shrink_to_fit), n), # bool
    textRotation      = rep(as_xml_attr(text_rotation), n), # int
    vertical          = rep(as_xml_attr(vertical), n),
    wrapText          = rep(as_xml_attr(wrap_text), n), # bool
    extLst            = rep(ext_lst, n),
    hidden            = rep(as_xml_attr(hidden), n), # bool
    locked            = rep(as_xml_attr(locked), n), # bool
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

#' internal function to remove border from a style
#' @param xf_node some xf node
#' @noRd
remove_border <- function(xf_node) {
  z <- read_xf(read_xml(xf_node))
  z$applyBorder <- ""
  z$borderId <- "0"
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

#' internal function to remove fill from a style
#' @param xf_node some xf node
#' @noRd
remove_fill <- function(xf_node) {
  z <- read_xf(read_xml(xf_node))
  z$applyFill <- ""
  z$fillId <- "0"
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

#' internal function to remove font from a style
#' @param xf_node some xf node
#' @noRd
remove_font <- function(xf_node) {
  z <- read_xf(read_xml(xf_node))
  z$applyFont <- ""
  z$fontId <- "0"
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

#' internal function to remove numfmt from a style
#' @param xf_node some xf node
#' @noRd
remove_numfmt <- function(xf_node) {
  z <- read_xf(read_xml(xf_node))
  z$applyNumberFormat <- ""
  z$numFmtId <- "0"
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

#' Get all styles on a sheet
#'
#' @param wb workbook
#' @param sheet worksheet
#'
#' @export
styles_on_sheet <- function(wb, sheet) {
  sheet_id <- wb$clone()$.__enclos_env__$private$get_sheet_index(sheet)
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
#' Create a new style to apply to worksheet cells. These styles are used in conditional formatting and in (pivot) table styles.
#'
#' @details
#' It is possible to override border_color and border_style with \{left, right, top, bottom\}_color, \{left, right, top, bottom\}_style.
#'
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
#' @seealso [wb_add_style()] [wb_add_dxfs_style()]
#' @family style creating functions
#' @examples
#' # do not apply anything
#' style1 <- create_dxfs_style()
#'
#' # change font color and background color
#' style2 <- create_dxfs_style(
#'   font_color = wb_color(hex = "FF9C0006"),
#'   bg_fill = wb_color(hex = "FFFFC7CE")
#' )
#'
#' # change font (type, size and color) and background
#' # the old default in openxlsx and openxlsx2 <= 0.3
#' style3 <- create_dxfs_style(
#'   font_name = "Aptos Narrow",
#'   font_size = 11,
#'   font_color = wb_color(hex = "FF9C0006"),
#'   bg_fill = wb_color(hex = "FFFFC7CE")
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
    text_underline = NULL,
    ...
) {

  # TODO diagonal requires up or down
  arguments <- c(..., ls(),
                 "left_color", "left_style", "right_color", "right_style",
                 "top_color", "top_style", "bottom_color", "bottom_style",
                 "start_color", "start_style", "end_color", "end_style",
                 "horizontal_color", "horizontal_style", "vertical_color", "vertical_style",
                 "diagonal_color", "diagonal_style"
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
    fill <- create_fill(pattern_type = pattern_type, bg_color = bg_fill, fg_color = fg_color, gradient_fill = gradient_fill)
  } else {
    fill <- NULL
  }

  # untested
  if (!is.null(border)) {
    left_color       <- if (exists("left_color"))       left_color   else border_color
    left_style       <- if (exists("left_style"))       left_style   else border_style
    right_color      <- if (exists("right_color"))      right_color  else border_color
    right_style      <- if (exists("right_style"))      right_style  else border_style
    top_color        <- if (exists("top_color"))        top_color    else border_color
    top_style        <- if (exists("top_style"))        top_style    else border_style
    bottom_color     <- if (exists("bottom_color"))     bottom_color else border_color
    bottom_style     <- if (exists("bottom_style"))     bottom_style else border_style
    start_color      <- if (exists("start_color"))      start_color else border_color
    start_style      <- if (exists("start_style"))      start_style else border_style
    end_color        <- if (exists("end_color"))        end_color else border_color
    end_style        <- if (exists("end_style"))        end_style else border_style
    horizontal_color <- if (exists("horizontal_color")) horizontal_color else border_color
    horizontal_style <- if (exists("horizontal_style")) horizontal_style else border_style
    vertical_color   <- if (exists("vertical_color"))   vertical_color else border_color
    vertical_style   <- if (exists("vertical_style"))   vertical_style else border_style
    diagonal_color   <- if (exists("diagonal_color"))   diagonal_color else border_color
    diagonal_style   <- if (exists("diagonal_style"))   diagonal_style else border_style

    border <- create_border(
      left              = left_style,
      left_color        = left_color,
      right             = right_style,
      right_color       = right_color,
      top               = top_style,
      top_color         = top_color,
      bottom            = bottom_style,
      bottom_color      = bottom_color,
      start             = start_style,
      start_color       = start_color,
      end               = end_style,
      end_color         = end_color,
      horizontal        = horizontal_style,
      horizontal_color  = horizontal_color,
      vertical          = vertical_style,
      vertical_color    = vertical_color
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

#' Helper function used to update borders styles
#' picks a style from a worksheet dimension and adds the required border
#' @noRd
update_border <- function(wb, sheet = current_sheet(), dims = "A1", new_border) {

  XFs <- wb$get_cell_style(sheet = sheet, dims = dims)
  # access the ids in a data frame
  c_s <- read_xf(read_xml(wb$styles_mgr$styles$cellXfs))[XFs, ]

  ## the current border for the cell
  get_border <- c_s$borderId
  borders <- read_border(read_xml(wb$styles_mgr$styles$borders))[get_border, ]

  # While an update was requested, the original cell might have been borderless
  borders[is.na(borders)] <- ""

  new_border <- read_border(read_xml(new_border))

  ## update the existing style
  sel <- which(new_border != "")
  borders[, sel] <- new_border[, sel]

  nms <- c(
    "start", "end", "left", "right", "top", "bottom", "diagonalDown",
    "diagonalUp", "diagonal", "vertical", "horizontal", "outline"
  )

  write_border(borders[nms])
}

## The functions below are only partially covered, but span over 300+ LOC ##

# nocov start

#' Create custom (pivot) table styles
#'
#' Create a custom (pivot) table style.
#' These functions are for expert use only. Use other styling functions instead.
#'
#' @param name name
#' @param whole_table wholeTable
#' @param header_row,total_row ...Row
#' @param first_column,last_column ...Column
#' @param first_row_stripe,second_row_stripe ...RowStripe
#' @param first_column_stripe,second_column_stripe ...ColumnStripe
#' @param first_header_cell,last_header_cell ...HeaderCell
#' @param first_total_cell,last_total_cell ...TotalCell
#' @param grand_total_row totalRow
#' @param grand_total_column lastColumn
#' @param first_subtotal_column,second_subtotal_column,third_subtotal_column ...SubtotalColumn
#' @param first_subtotal_row,second_subtotal_row,third_subtotal_row ...SubtotalRow
#' @param first_column_subheading,second_column_subheading,third_column_subheading ...ColumnSubheading
#' @param first_row_subheading,second_row_subheading,third_row_subheading ...RowSubheading
#' @param blank_row blankRow
#' @param page_field_labels pageFieldLabels
#' @param page_field_values pageFieldValues
#' @param ... additional arguments
#' @name create_tablestyle
#' @family style creating functions
#' @export
# TODO: combine both functions and simply set pivot=0 or table=0
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

#' @rdname create_tablestyle
#' @export
create_pivottablestyle <- function(
    name,
    whole_table              = NULL,
    header_row               = NULL,
    grand_total_row          = NULL,
    first_column             = NULL,
    grand_total_column       = NULL,
    first_row_stripe         = NULL,
    second_row_stripe        = NULL,
    first_column_stripe      = NULL,
    second_column_stripe     = NULL,
    first_header_cell        = NULL,
    first_subtotal_column    = NULL,
    second_subtotal_column   = NULL,
    third_subtotal_column    = NULL,
    first_subtotal_row       = NULL,
    second_subtotal_row      = NULL,
    third_subtotal_row       = NULL,
    blank_row                = NULL,
    first_column_subheading  = NULL,
    second_column_subheading = NULL,
    third_column_subheading  = NULL,
    first_row_subheading     = NULL,
    second_row_subheading    = NULL,
    third_row_subheading     = NULL,
    page_field_labels        = NULL,
    page_field_values        = NULL,
    ...
) {

  standardize_case_names(...)

  tab_wholeTable <- NULL
  if (length(whole_table)) {
    tab_wholeTable <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "wholeTable", dxfId = whole_table))
  }
  tab_pageFieldLabels <- NULL
  if (length(page_field_labels)) {
    tab_pageFieldLabels <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "pageFieldLabels", dxfId = page_field_labels))
  }
  tab_pageFieldValues <- NULL
  if (length(page_field_values)) {
    tab_pageFieldValues <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "pageFieldValues", dxfId = page_field_values))
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
  tab_firstColumn <- NULL
  if (length(first_column)) {
    tab_firstColumn <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "firstColumn", dxfId = first_column))
  }
  tab_headerRow <- NULL
  if (length(header_row)) {
    tab_headerRow <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "headerRow", dxfId = header_row))
  }
  tab_firstHeaderCell <- NULL
  if (length(first_header_cell)) {
    tab_firstHeaderCell <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "firstHeaderCell", dxfId = first_header_cell))
  }
  tab_subtotalColumn1 <- NULL
  if (length(first_subtotal_column)) {
    tab_subtotalColumn1 <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "firstSubtotalColumn", dxfId = first_subtotal_column))
  }
  tab_subtotalColumn2 <- NULL
  if (length(second_subtotal_column)) {
    tab_subtotalColumn2 <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "secondSubtotalColumn", dxfId = second_subtotal_column))
  }
  tab_subtotalColumn3 <- NULL
  if (length(third_subtotal_column)) {
    tab_subtotalColumn3 <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "thirdSubtotalColumn", dxfId = third_subtotal_column))
  }
  tab_blankRow <- NULL
  if (length(blank_row)) {
    tab_blankRow <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "blankRow", dxfId = blank_row))
  }
  tab_subtotalRow1 <- NULL
  if (length(first_subtotal_row)) {
    tab_subtotalRow1 <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "firstSubtotalRow", dxfId = first_subtotal_row))
  }
  tab_subtotalRow2 <- NULL
  if (length(second_subtotal_row)) {
    tab_subtotalRow2 <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "secondSubtotalRow", dxfId = second_subtotal_row))
  }
  tab_subtotalRow3 <- NULL
  if (length(third_subtotal_row)) {
    tab_subtotalRow3 <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "thirdSubtotalRow", dxfId = third_subtotal_row))
  }
  tab_columnSubheading1 <- NULL
  if (length(first_column_subheading)) {
    tab_columnSubheading1 <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "firstColumnSubheading", dxfId = first_column_subheading))
  }
  tab_columnSubheading2 <- NULL
  if (length(second_column_subheading)) {
    tab_columnSubheading2 <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "secondColumnSubheading", dxfId = second_column_subheading))
  }
  tab_columnSubheading3 <- NULL
  if (length(third_column_subheading)) {
    tab_columnSubheading3 <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "thirdColumnSubheading", dxfId = third_column_subheading))
  }
  tab_rowSubheading1 <- NULL
  if (length(first_row_subheading)) {
    tab_rowSubheading1 <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "firstRowSubheading", dxfId = first_row_subheading))
  }
  tab_rowSubheading2 <- NULL
  if (length(second_row_subheading)) {
    tab_rowSubheading2 <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "secondRowSubheading", dxfId = second_row_subheading))
  }
  tab_rowSubheading3 <- NULL
  if (length(third_row_subheading)) {
    tab_rowSubheading3 <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "thirdRowSubheading", dxfId = third_row_subheading))
  }
  tab_grandTotalColumn <- NULL
  if (length(grand_total_column)) {
    tab_grandTotalColumn <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "lastColumn", dxfId = grand_total_column))
  }
  tab_grandTotalRow <- NULL
  if (length(grand_total_row)) {
    tab_grandTotalRow <- xml_node_create(
      "tableStyleElement",
      xml_attributes = c(type = "totalRow", dxfId = grand_total_row))
  }

  xml_elements <- c(
    tab_wholeTable,
    tab_headerRow,
    tab_firstColumn,
    tab_grandTotalColumn,
    tab_firstRowStripe,
    tab_secondRowStripe,
    tab_firstColumnStripe,
    tab_secondColumnStripe,
    tab_firstHeaderCell,
    tab_subtotalColumn1,
    tab_subtotalColumn2,
    tab_subtotalColumn3,
    tab_subtotalRow1,
    tab_subtotalRow2,
    tab_subtotalRow3,
    tab_blankRow,
    tab_columnSubheading1,
    tab_columnSubheading2,
    tab_columnSubheading3,
    tab_rowSubheading1,
    tab_rowSubheading2,
    tab_rowSubheading3,
    tab_grandTotalRow,
    tab_pageFieldLabels,
    tab_pageFieldValues
  )

  rand_str <- random_string(length = 12, pattern = "[A-Z0-9]")

  xml_node_create(
    "tableStyle",
    xml_attributes = c(
      name      = name,
      table     = "0", # table uses different styles
      count     = as_xml_attr(length(xml_elements)),
       # possible st_guid()
      `xr9:uid` = sprintf("{ADA2BA8A-2FEC-2C46-A01D-%s}", rand_str)
    ),
    xml_children = xml_elements
  )
}

#' Create custom color xml schemes
#'
#' Create custom color themes that can be used with [wb_set_base_colors()]. The color input will be checked with [wb_color()], so it must be either a color R from [grDevices::colors()] or a hex value.
#' Default values for the dark argument are: `black`, `white`, `darkblue` and `lightgray`. For the accent argument, the six inner values of [grDevices::palette()]. The link argument uses `blue` and `purple` by default for active and visited links.
#' @name create_colors_xml
#' @param name the color name
#' @param dark four colors: dark, light, brighter dark, darker light
#' @param accent six accent colors
#' @param link two link colors: link and visited link
#' @family style creating functions
#' @examples
#' colors <- create_colors_xml()
#' wb <- wb_workbook()$add_worksheet()$set_base_colors(xml = colors)
#' @export
create_colors_xml <- function(
    name   = "Base R",
    dark   = NULL,
    accent = NULL,
    link   = NULL
  ) {

  if (is.null(dark))
    dark <- c("black", "white", "darkblue", "lightgray")

  if (is.null(accent))
    accent <- grDevices::palette()[2:7]

  if (is.null(link))
    link <- c("blue", "purple")

  if (length(dark) != 4) {
    stop("dark vector must be of length 4")
  }
  if (length(accent) != 6) {
    stop("accent vector must be of length 6")
  }
  if (length(link) != 2) {
    stop("link vector must be of length 2")
  }

  colors <- c(dark, accent, link)
  colors <- as.character(wb_color(colors))
  if (length(colors) != 12) {
    stop("colors vector must be of length 12")
  }

  colors <- vapply(
    colors,
    function(x) substring(x, 3, nchar(x)),
    NA_character_
  )

  read_xml(
    sprintf(
      "<a:clrScheme name=\"%s\">
      <a:dk1><a:sysClr val=\"windowText\" lastClr=\"%s\"/></a:dk1>
      <a:lt1><a:sysClr val=\"window\" lastClr=\"%s\"/></a:lt1>
      <a:dk2><a:srgbClr val=\"%s\"/></a:dk2>
      <a:lt2><a:srgbClr val=\"%s\"/></a:lt2>
      <a:accent1><a:srgbClr val=\"%s\"/></a:accent1>
      <a:accent2><a:srgbClr val=\"%s\"/></a:accent2>
      <a:accent3><a:srgbClr val=\"%s\"/></a:accent3>
      <a:accent4><a:srgbClr val=\"%s\"/></a:accent4>
      <a:accent5><a:srgbClr val=\"%s\"/></a:accent5>
      <a:accent6><a:srgbClr val=\"%s\"/></a:accent6>
      <a:hlink><a:srgbClr val=\"%s\"/></a:hlink>
      <a:folHlink><a:srgbClr val=\"%s\"/></a:folHlink>
      </a:clrScheme>",
      name,
      colors[1], # dk1
      colors[2], # lt1
      colors[3], # dk2
      colors[4], # lt2
      colors[5], # accent1
      colors[6], # accent2
      colors[7], # accent3
      colors[8], # accent4
      colors[9], # accent5
      colors[10], # accent6
      colors[11], # hlink
      colors[12] # folHlink
    ), pointer = FALSE
  )
}

#' @export
#' @rdname create_colors_xml
#' @usage NULL
create_colours_xml <- create_colors_xml

# OOXML builtin formats
## in spreadsheet software the accounting style is shown with the symbol of the
## local currency and its either in front or behind the digits. We cannot do this
## therefore the symbol is dropped.
builtin_fmts_df <- data.frame(
  numFmtId = c(
    # 0,
    1, 2, 3, 4, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22,
    37, 38, 39, 40, 44, 45, 46, 47, 48, 49
  ),
  formatCode = c(
    # "General",                               # 0
    "0",                                     # 1
    "0.00",                                  # 2
    "#,##0",                                 # 3
    "#,##0.00",                              # 4
    "0%",                                    # 9
    "0.00%",                                 # 10
    "0.00E+00",                              # 11
    "# ?/?",                                 # 12
    "# ??/??",                               # 13
    "mm-dd-yy",                              # 14
    "d-mmm-yy",                              # 15
    "d-mmm",                                 # 16
    "mmm-yy",                                # 17
    "h:mm AM/PM",                            # 18
    "h:mm:ss AM/PM",                         # 19
    "h:mm",                                  # 20
    "h:mm:ss",                               # 21
    "m/d/yy h:mm",                           # 22
    "#,##0 ;(#,##0)",                        # 37
    "#,##0 ;[Red](#,##0)",                   # 38
    "#,##0.00;(#,##0.00)",                   # 39
    "#,##0.00;[Red](#,##0.00)",              # 40
    "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)", # 44 (Accounting)
    "mm:ss",                                 # 45
    "[h]:mm:ss",                             # 46
    "mmss.0",                                # 47
    "##0.0E+0",                              # 48
    "@"                                      # 49 (Text)
  ),
  stringsAsFactors = FALSE
)

# nocov end

# style helper used in wb_to_df()
get_numfmt_style <- function(wb, cc) {
  nms_uni_styles <- unique(cc$c_s)
  uni_styles <- as.numeric(nms_uni_styles) + 1L
  names(uni_styles) <- nms_uni_styles
  uni_styles <- uni_styles[!is.na(uni_styles)]

  cc$num_fmt <- rep_len("", nrow(cc))

  if (length(uni_styles) == 0) {
    return(cc)
  }

  styles <- rbindlist(
    xml_attr(
      wb$styles_mgr$styles$cellXfs[uni_styles],
      "xf"
    )
  )
  styles$style_id <- names(uni_styles)

  # in an xlsx file created by LibreOffice there were duplicates in
  # numFmtId for some reason
  sel <- !duplicated(styles[c("numFmtId", "style_id")]) & styles$applyNumberFormat %in% c("1", "true")
  styles <- styles[sel, c("style_id", "numFmtId")]

  numfmts <- rbindlist(
    xml_attr(
      wb$styles_mgr$styles$numFmts,
      "numFmt"
    )
  )

  numfmts <- rbind(numfmts, builtin_fmts_df)

  styles <- merge(
    styles,
    numfmts,
    by = "numFmtId",
    all.x = TRUE,
    all.y = FALSE
  )

  # 1. Calculate the match indices
  idx <- match(cc$c_s, styles$style_id)

  # 2. Only map values where a match was found
  # This prevents errors if idx contains NAs
  matches <- !is.na(idx)
  cc$num_fmt[matches] <- styles$formatCode[idx[matches]]

  cc
}

# --- Helper: Split OOXML Sections Safely ---
split_ooxml_sections <- function(fmt) {
  unlist(strsplit(fmt, ";(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)", perl = TRUE))
}

# --- Helper: Process Literals and Alignment ---
process_literals <- function(fmt, value = NULL) {
  res <- fmt
  res <- gsub('\"(.*?)\"', "\\1", res)
  res <- gsub("\\\\(.)", "\\1", res)
  res <- gsub("\\*.", " ", res)
  res <- gsub("_.", "", res)

  # Alignment placeholders (?) become spaces in non-fraction contexts
  if (!grepl("/", res)) {
    res <- gsub("\\?", " ", res)
  }

  if (!is.null(value) && grepl("@", res)) {
    res <- gsub("@", as.character(value), res)
  }

  res <- gsub("\\s+", " ", res)
  res <- trimws(res)
  Encoding(res) <- "UTF-8"

  enc2native(res)
}

# --- Helper: Date and Time Formatter ---
format_date_time <- function(value, format_code) {

  # 1. Determine if this is a calendar date or a pure duration
  # If it contains brackets or is numeric, it's likely a duration.
  is_duration_fmt <- grepl("\\[", format_code)

  # Coerce dt_val ONLY if we aren't dealing with a pure numeric duration
  # to avoid the format() crash you are seeing.
  dt_val <- NA
  if (!is.numeric(value)) {
    dt_val <- if (inherits(value, "POSIXct") || inherits(value, "Date")) {
      value
    } else if (is.character(value)) {
      # for a hms format we need to add a date
      if (grepl("^\\d{1,2}:\\d{2}", value)) {
        value <- paste("1900-01-01", value)
      }
      # Try to parse, but keep as NA if it's just a time string
      tryCatch(as.POSIXct(value, tz = "UTC"), error = function(e) NA)
    }
  }

  res <- format_code

  # 2. Handle Elapsed Units [h], [m], [s]
  # We do this first and use FIXED matching to avoid regex/bracket confusion
  if (is_duration_fmt) {

    if (is.character(value)) {
      value <- as.POSIXct(value, "UTC")
    }

    if (inherits(dt_val, "POSIXct")) {
      origin <- as.POSIXct("1899-12-31", tz = "UTC")
      value <- as.numeric(difftime(dt_val, origin, units = "days"))
    }

    excel_res <- convert_to_excel_date(list(v = value))
    serial_num <- excel_res[[1]]

    if (is.numeric(serial_num)) {
      # Literal tags for durations. Must be greedy (hh before h).
      tags <- c("[hh]", "[h]", "[HH]", "[H]", "[mm]", "[m]", "[MM]", "[M]", "[ss]", "[s]", "[SS]", "[S]")
      for (tag in tags) {
        if (grepl(tag, res, fixed = TRUE)) {
          val_str <- ""
          unit <- tolower(substr(tag, 2, 2))
          if (unit == "h") {
            val <- floor((serial_num * 24) + 1e-7)
            val_str <- if (nchar(tag) > 3) sprintf("%02d", val) else as.character(val)
          } else if (unit == "m") {
            val <- floor((serial_num * 1440) + 1e-7)
            val_str <- if (nchar(tag) > 3) sprintf("%02d", val) else as.character(val)
          } else if (unit == "s") {
            val <- floor((serial_num * 86400) + 1e-7)
            val_str <- if (nchar(tag) > 3) sprintf("%02d", val) else as.character(val)
          }
          res <- gsub(tag, val_str, res, fixed = TRUE)
        }
      }
    }
  }

  # 3. Calendar & Clock Formatting
  # ONLY run if we have a valid date and the format isn't purely elapsed seconds/minutes
  if (!all(is.na(dt_val))) {

    # Standard Date Parts
    res <- gsub("yyyy", format(dt_val, "%Y"), res, ignore.case = TRUE)
    res <- gsub("yy", format(dt_val, "%y"), res, ignore.case = TRUE)
    res <- gsub("dddd", format(dt_val, "%A"), res, ignore.case = TRUE)
    res <- gsub("ddd", format(dt_val, "%a"), res, ignore.case = TRUE)
    res <- gsub("dd", format(dt_val, "%d"), res, ignore.case = TRUE)
    res <- gsub("\\bd\\b", as.numeric(format(dt_val, "%d")), res, perl = TRUE, ignore.case = TRUE)

    # Minute Protection logic
    res <- gsub("(?<=[Hh\\d]):mm", ":_MINP_", res, perl = TRUE)
    res <- gsub("(?<=[Hh\\d]):m", ":_MIN_", res, perl = TRUE)
    res <- gsub("mm(?=:ss|s)", "_MINP_", res, perl = TRUE)
    res <- gsub("m(?=:ss|s)", "_MIN_", res, perl = TRUE)

    # Contextual minutes for durations
    if (is_duration_fmt) {
      res <- gsub("\\bmm\\b(?!(.*_MIN))", "_MINP_", res, perl = TRUE)
      res <- gsub("\\bm\\b(?!(.*_MIN))", "_MIN_", res, perl = TRUE)
    }

    # Months
    res <- gsub("mmmmm", substr(format(dt_val, "%B"), 1, 1), res, ignore.case = TRUE)
    res <- gsub("mmmm", format(dt_val, "%B"), res, ignore.case = TRUE)
    res <- gsub("mmm", format(dt_val, "%b"), res, ignore.case = TRUE)
    res <- gsub("mm", format(dt_val, "%m"), res, ignore.case = TRUE)
    res <- gsub("\\bm\\b", as.numeric(format(dt_val, "%m")), res, perl = TRUE, ignore.case = TRUE)

    # Restore Minutes
    res <- gsub("_MINP_", format(dt_val, "%M"), res)
    res <- gsub("_MIN_", as.numeric(format(dt_val, "%M")), res)

    # Clock Time
    is_12h <- grepl("AM/PM|A/P", res, ignore.case = TRUE)
    # \\bh\\b boundary is vital to not overwrite the 296 we just put in
    res <- gsub("hh", format(dt_val, if (is_12h) "%I" else "%H"), res, ignore.case = TRUE)
    res <- gsub("\\bh\\b", as.numeric(format(dt_val, if (is_12h) "%I" else "%H")), res, perl = TRUE, ignore.case = TRUE)
    res <- gsub("ss", format(dt_val, "%S"), res, ignore.case = TRUE)
    res <- gsub("AM/PM", format(dt_val, "%p"), res, ignore.case = TRUE)
    res <- gsub("A/P", substr(format(dt_val, "%p"), 1, 1), res, ignore.case = TRUE)

  } else if (is.numeric(value) || is_duration_fmt) {
    # If we have a numeric value but no valid dt_val (like 3600),
    # handle the remaining :mm:ss or m using serial math to avoid the prettyNum crash.
    excel_res <- convert_to_excel_date(list(v = value))
    s <- excel_res[[1]]

    # Simple replacement for clock-parts in durations when no calendar date exists
    mins <- sprintf("%02d", floor((s * 1440) %% 60 + 1e-7))
    secs <- sprintf("%02d", floor((s * 86400) %% 60 + 1e-7))

    res <- gsub(":mm", paste0(":", mins), res, ignore.case = TRUE)
    res <- gsub(":ss", paste0(":", secs), res, ignore.case = TRUE)
    # Handle single m if it's the only thing left
    res <- gsub("\\bm\\b", as.numeric(mins), res, perl = TRUE)
  }

  process_literals(res)
}

# --- Helper: Number Formatter ---
format_number <- function(value, format_code) {
  is_scientific <- grepl("E\\+", format_code, ignore.case = TRUE)

  mask_regex <- "[#0,\\.E\\+]+"
  matches <- gregexpr(mask_regex, format_code)[[1]]
  if (matches[1] == -1) return(process_literals(format_code))

  match_idx <- matches[length(matches)]
  match_len <- attr(matches, "match.length")[length(matches)]
  num_pattern <- substr(format_code, match_idx, match_idx + match_len - 1)

  if (is_scientific) {
    prec_match <- regmatches(num_pattern, regexec("\\.(.*?)E", num_pattern, ignore.case = TRUE))
    precision <- if (length(prec_match[[1]]) > 1) nchar(prec_match[[1]][2]) else 0

    formatted_num <- sprintf(paste0("%.", precision, "E"), value)

    # OOXML usually uses uppercase 'E+'
    formatted_num <- gsub("e", "E", formatted_num)
  } else {
    parts <- strsplit(num_pattern, "\\.")[[1]]
    int_pat <- parts[1]
    frac_pat <- if (length(parts) > 1) parts[2] else ""

    m <- gregexpr(",+$", int_pat, perl = TRUE)
    if (m[[1]][1] != -1) value <- value / (1000 ^ nchar(regmatches(int_pat, m)[[1]]))

    # Calculate mandatory integer digits
    mandatory_int <- nchar(gsub("[^0]", "", gsub(",", "", int_pat)))
    precision <- nchar(gsub("[^0#]", "", frac_pat))

    rounded_val <- abs(round(value, precision))

    # If mandatory_int is 0 and the value is < 1, the integer part should be blank
    int_part_val <- floor(rounded_val)
    if (mandatory_int == 0 && int_part_val == 0) {
      formatted_int <- ""
    } else {
      # Standard padding
      formatted_int <- sprintf(paste0("%0", mandatory_int, "d"), as.integer(int_part_val))
    }

    # Handle the fractional part
    if (precision > 0) {
      # Use a simple %f to get the decimals, then strip the leading "0" or "1" before the dot
      raw_frac <- sprintf(paste0("%.", precision, "f"), rounded_val)
      formatted_frac <- gsub("^.*\\.", ".", raw_frac)

      # Apply your existing # trailing-zero-stripping logic to formatted_frac here...

      formatted_num <- paste0(formatted_int, formatted_frac)
    } else {
      formatted_num <- formatted_int
    }

    if (grepl("#", frac_pat)) {
      # If the entire fractional pattern is #, and there's no fractional value, remove the decimal point too
      if (!grepl("0", frac_pat) && (rounded_val %% 1 == 0)) {
        formatted_num <- gsub("\\.0+$", "", formatted_num)
      } else {
        # Otherwise, only remove the zeros that correspond to the # positions
        # This is a simple way: if frac_pat is '0.0#', we keep one zero but drop the second
        # For now, a common robust approach is trimming trailing zeros if they aren't 'mandatory'
        num_parts <- strsplit(formatted_num, "\\.")[[1]]
        if (length(num_parts) > 1) {
          # Calculate how many zeros are mandatory (the '0's in frac_pat)
          mandatory_frac_len <- nchar(gsub("[^0]", "", frac_pat))
          # Trim trailing zeros but keep up to mandatory_frac_len
          frac_str <- num_parts[2]
          while (nchar(frac_str) > mandatory_frac_len && substr(frac_str, nchar(frac_str), nchar(frac_str)) == "0") {
            frac_str <- substr(frac_str, 1, nchar(frac_str) - 1)
          }
          formatted_num <- if (nchar(frac_str) > 0) paste0(num_parts[1], ".", frac_str) else num_parts[1]
        }
      }
    }

    if (grepl(",", int_pat)) {
      num_parts <- strsplit(formatted_num, "\\.")[[1]]
      num_parts[1] <- gsub("(?<=\\d)(?=(\\d{3})+$)", ",", num_parts[1], perl = TRUE)
      formatted_num <- paste(num_parts, collapse = ".")
    }
  }

  prefix <- substr(format_code, 1, match_idx - 1)
  suffix <- substr(format_code, match_idx + match_len, nchar(format_code))

  process_literals(paste0(prefix, formatted_num, suffix))
}

# --- Helper: Fraction Formatter (Farey Algorithm) ---
format_fraction <- function(value, format_code) {
  clean_fmt <- process_literals(format_code)

  val_abs <- abs(value)
  whole_part <- floor(val_abs + 1e-10) # epsilon for precision
  frac_part <- val_abs - whole_part

  if (frac_part < 1e-10) return(as.character(if (value < 0) -whole_part else whole_part))

  # Determine denominator limit
  denom_match <- regmatches(clean_fmt, regexpr("/([\\?\\d]+)", clean_fmt))
  denom_str <- gsub("/", "", denom_match)

  limit <- if (grepl("^[\\d]+$", denom_str)) as.numeric(denom_str)
  else if (nchar(denom_str) == 3) 999
  else if (nchar(denom_str) == 2) 99
  else 9

  # Find best rational approximation within limit (Farey Sequence)
  low_n <- 0
  low_d <- 1
  high_n <- 1
  high_d <- 0 # Represents 1/0

  while (TRUE) {
    num <- low_n + high_n
    den <- low_d + high_d
    if (den > limit) break

    if (frac_part > (num / den)) {
      low_n <- num
      low_d <- den
    } else if (frac_part < (num / den)) {
      high_n <- num
      high_d <- den
    } else {
      low_n <- num
      low_d <- den
      break
    }
  }

  # Pick the closest of the two boundaries
  if (abs(frac_part - (low_n / low_d)) <= abs(frac_part - (high_n / high_d)) || high_d > limit) {
    n_res <- low_n
    d_res <- low_d
  } else {
    n_res <- high_n
    d_res <- high_d
  }

  sign_str <- if (value < 0) "-" else ""

  if (grepl("#", clean_fmt)) {
    if (whole_part == 0) return(paste0(sign_str, n_res, "/", d_res))
    paste0(sign_str, whole_part, " ", n_res, "/", d_res)
  } else {
    paste0(sign_str, (whole_part * d_res) + n_res, "/", d_res)
  }
}

#' Format Values using OOXML (Spreadsheet) Number Format Codes
#'
#' @description
#' This function emulates a spreadsheet formatting engine. It takes numeric,
#' date, or character values and applies an OOXML format code to produce
#' a formatted string. It supports standard number formatting, date/time,
#' elapsed durations, fractions, and conditional formatting sections.
#'
#' @param value A vector of values to format. Supports `numeric`, `Date`,
#'   `POSIXct`, `character` (ISO dates or times), `hms`, or `difftime`.
#' @param format_code A character vector of format strings (e.g., `"#,##0.00"`,
#'   `"yyyy-mm-dd"`, or `"[h]:mm:ss"`).
#'
#' @details
#' The function splits the `format_code` into up to four sections separated
#' by semicolons (Positive; Negative; Zero; Text).
#'
#' \itemize{
#'   \item \bold{Date and Time:} Supports standard tokens (`yyyy`, `mm`, `dd`,
#'     `hh`, `ss`) and AM/PM toggles.
#'   \item \bold{Durations:} Supports elapsed time tokens in square brackets,
#'     such as `[h]`, `[m]`, and `[s]`, calculating total units.
#'   \item \bold{HMS Handling:} Character strings in "HH:MM:SS" format are
#'     automatically coerced to a date-time object using a base date
#'     to allow clock and duration formatting.
#'   \item \bold{Numbers:} Supports thousands separators (`,`), precision,
#'     percentage conversion, and scientific notation (`E+`).
#'   \item \bold{Fractions:} Uses the Farey algorithm to approximate decimals
#'     as fractions (e.g., `?/?`, `# ?/?`).
#'   \item \bold{Conditionals:} Evaluates bracketed conditions like `[<1000]`
#'     to select the appropriate formatting section.
#'   \item \bold{Literals:} Handles escaped characters (`\\`), quoted
#'     text (`"text"`), and the text placeholder (`@`).
#' }
#'
#' Rounding is based on the R function `round()` and does not match spreadsheet
#' software behavior.
#'
#' @return A character vector of formatted strings.
#'
#' @examples
#' # Numeric formatting
#' apply_numfmt(1234.5678, "#,##0.00") # "1,234.57"
#'
#' # Date and Time
#' apply_numfmt("2025-01-05", "dddd, mmm dd") # "Sunday, Jan 05"
#' apply_numfmt("13:45:30", "hh:mm AM/PM")    # "01:45 PM"
#'
#' # Durations
#' apply_numfmt("1900-01-12 08:17:47", "[h]:mm:ss") # "296:17:47"
#'
#' # Fractions
#' apply_numfmt(1.75, "# ?/?") # "1 3/4"
#'
#' @export
apply_numfmt <- function(value, format_code) {

  if (inherits(value, "hms") || inherits(value, "difftime")) {
    value <- convert_datetime(value)
  }

  # Standardize input lengths
  max_len <- max(length(value), length(format_code))
  value <- rep_len(value, max_len)
  format_code <- rep_len(format_code, max_len)

  results <- character(max_len)

  unique_fmts <- unique(format_code)

  for (fmt_raw in unique_fmts) {
    idx <- which(format_code == fmt_raw)
    v_subset <- value[idx]

    # --- Pre-processing ---
    f <- replaceXMLEntities(fmt_raw)
    sections <- split_ooxml_sections(f)
    n_sec <- length(sections)

    # Pre-clean metadata for each section
    sections_cleaned <- lapply(sections, function(target_fmt) {
      # Currency Extraction
      curr_matches <- regmatches(target_fmt, gregexpr("\\[\\$(.*?)-.*?\\]", target_fmt))[[1]]
      if (length(curr_matches) > 0) {
        for (m in curr_matches) {
          symbol <- gsub("^\\[\\$|\\-.*$", "", m)
          target_fmt <- gsub(m, symbol, target_fmt, fixed = TRUE)
        }
      }

      target_fmt <- gsub("\\* ", " ", target_fmt)
      target_fmt <- gsub("\\*.", "", target_fmt)
      target_fmt <- gsub("(?<!\\\\)_.", "", target_fmt, perl = TRUE)

      # Strip Conditions ([<1500]) and Colors ([Red]) but PROTECT Durations ([h], [m], [s])
      target_fmt <- gsub("\\[(?!(?:h|m|s|H|M|S))[^\\]]+\\]", "", target_fmt, perl = TRUE)
      target_fmt
    })

    # --- Value Processing ---
    results[idx] <- vapply(v_subset, function(v) {
      if (is.na(v)) return(NA_character_)

      # Type Detection
      is_numeric <- is.numeric(v)
      is_dt_obj <- inherits(v, "Date") || inherits(v, "POSIXct")
      is_dt_str <- !is_numeric && !is_dt_obj && is.character(v) && (
        grepl("^\\d{4}-\\d{2}-\\d{2}", v) || grepl("^\\d{1,2}:\\d{2}:\\d{2}", v)
      )

      target_fmt <- ""
      condition_matched <- FALSE
      sec_idx <- 1

      # 1. Section Selection Logic
      if (is_numeric && any(grepl("^\\[[<>=!]+[0-9.]+\\]", sections))) {
        for (i in seq_along(sections)) {
          sec <- sections[[i]]
          if (grepl("^\\[([<>=!]+[0-9.]+)\\]", sec)) {
            cond_str <- gsub("^\\[([<>=!]+[0-9.]+)\\].*", "\\1", sec)
            if (eval(parse(text = paste(v, cond_str)))) {
              sec_idx <- i
              condition_matched <- TRUE
              break
            }
          } else {
            sec_idx <- i
            condition_matched <- TRUE
            break
          }
        }
      }

      if (!condition_matched) {
        if (is_numeric) {
          if (v > 0 || (v == 0 && n_sec < 3)) {
            sec_idx <- 1
          } else if (v < 0) {
            if (n_sec >= 2) {
              sec_idx <- 2
              v <- abs(v)
            } else {
              sec_idx <- 1
            }
          } else {
            if (n_sec >= 3 && !grepl("[#0]", sections[[3]])) return(process_literals(sections_cleaned[[3]]))
            sec_idx <- if (n_sec >= 3) 3 else 1
          }
        } else if (is_dt_str || is_dt_obj) {
          sec_idx <- 1
        } else {
          if (n_sec >= 4) return(process_literals(sections_cleaned[[4]], v))
          if (grepl("@", sections[[1]])) return(process_literals(sections_cleaned[[1]], v))
          return(as.character(v))
        }
      }

      # Use the pre-cleaned section
      target_fmt <- sections_cleaned[[sec_idx]]
      orig_sec <- sections[[sec_idx]]

      # 3. Final Routing
      # A. Dates/Times
      if (is_dt_str || is_dt_obj || (grepl("[ymdhs\\[]", orig_sec, ignore.case = TRUE) && !grepl("[#0]", orig_sec))) {
        return(format_date_time(v, target_fmt))
      }

      # B. Fractions
      if (grepl("/[\\?\\d]", target_fmt) && is_numeric) {
        return(format_fraction(v, target_fmt))
      }

      # C. Percentages
      if (grepl("%", target_fmt)) {
        val_pct <- if (is_numeric) v * 100 else v
        res <- format_number(val_pct, gsub("%", "", target_fmt))
        return(paste0(res, "%"))
      }

      # D. Standard Numbers
      res <- format_number(v, target_fmt)
      if (is_numeric && v < 0 && !condition_matched && n_sec == 1 && !grepl("[\\(\\)-]", res)) {
        res <- paste0("-", res)
      }

      res
    }, character(1))
  }

  results
}
