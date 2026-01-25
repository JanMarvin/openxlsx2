#' Create a color object for workbook styling
#'
#' @description
#' `wb_color()` is a helper function used to define colors for fonts, fills,
#' borders, and other styling elements. It creates a `wbColour` object that
#' encapsulates color information in a format compatible with the OpenXML
#' specification.
#'
#' @details
#' Colors in spreadsheets can be defined in several ways. This function
#' standardizes those inputs into a single object.
#'
#' Hex Color Formats:
#' Hexadecimal colors represent the intensity of Red, Green, and Blue.
#' A major point of confusion is the Alpha (transparency) channel:
#' * ARGB (Default): Spreadsheet software expects the Alpha value before
#'   the RGB values (e.g., `FFFF0000` for solid red).
#' * RGBA: R functions like [grDevices::adjustcolor()] place the Alpha
#'   value after the RGB values (e.g., `FF0000FF` for solid red).
#' * Use the `format` argument to tell `wb_color()` how to interpret your hex
#'   string if it includes transparency.
#'
#' Theme Colors:
#' Instead of hard-coding a hex value, you can use `theme`. These are indices
#' (0-based) referencing the workbook's theme palette (e.g., "Text 1", "Accent 1").
#' Using theme colors allows the spreadsheet's appearance to change
#' automatically if the user changes the workbook theme.
#'
#' Tints:
#' The `tint` parameter modifies a color's lightness. A value of `-0.25`
#' darkens the color by 25%, while `0.5` lightens it by 50%.
#'
#' @param name A character string. Can be a standard R color name (e.g., `"red"`,
#'   `"steelblue"`) or a hex value.
#' @param hex A character string representing a hex color (e.g., `"FF0000"`).
#'   Leading "#" is optional.
#' @param theme A zero-based integer index referencing a value in the
#'   workbook's theme (usually 0-11).
#' @param tint A numeric value between -1.0 (darkest) and 1.0 (lightest)
#'   to modify the base color.
#' @param auto A logical value. If `TRUE`, the spreadsheet application
#'   determines the color automatically (usually for text contrast).
#' @param indexed An integer referencing a legacy indexed color value.
#' @param format The alpha channel format for hex strings: `"ARGB"` (default)
#'   or `"RGBA"`.
#'
#' @return A `wbColour` object (a named character vector).
#'
#' @examples
#' # Using standard R colors
#' font_col <- wb_color("darkblue")
#'
#' # Using Hex values with Alpha
#' bg_col <- wb_color(hex = "FFED7D31") # ARGB for orange
#'
#' # Using Theme colors with a tint (4th accent color, 40% lighter)
#' theme_col <- wb_color(theme = 4, tint = 0.4)
#'
#' # Converting R-style RGBA to spreadsheet-style ARGB
#' r_col <- adjustcolor("red", alpha.f = 0.5)
#' xlsx_col <- wb_color(hex = r_col, format = "RGBA")
#'
#' @seealso [wb_get_base_colors()], [grDevices::colors()]
#' @export
wb_color <- function(
    name = NULL,
    auto = NULL,
    indexed = NULL,
    hex = NULL,
    theme = NULL,
    tint = NULL,
    format = c("ARGB", "RGBA")
  ) {
  format <- match.arg(format)

  if (!is.null(name)) hex <-  validate_color(name, format = format)
  if (!is.null(hex))  hex <-  validate_color(hex, format = format)

  z <- c(
    auto    = as_xml_attr(auto),
    indexed = as_xml_attr(indexed),
    rgb     = as_xml_attr(hex),
    theme   = as_xml_attr(theme),
    tint    = as_xml_attr(tint)
  )

  z <- z[z != ""]

  # wbColour for historical reasons
  class(z) <- c("wbColour", "character")
  z
}

#' @export
#' @rdname wb_color
#' @usage NULL
wb_colour <- wb_color

is_wbColour <- function(x) inherits(x, "wbColour")
