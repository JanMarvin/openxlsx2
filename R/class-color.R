#' Helper to create a color
#'
#' Creates a `wbColour` object.
#'
#' The **format** of the hex color representation  can be either RGB, ARGB, or RGBA.
#' These hex formats differ only in a way how they encode the transparency value alpha,
#' ARGB expecting the alpha value before the RGB values (default in spreadsheets),
#' RGBA expects the alpha value after the RGB values (default in R),
#' and RGB is not encoding transparency at all.
#' If the colors some from functions such as `adjustcolor` that provide color in the
#' RGBA format, it is necessary to specify the `format = "RGBA"` when calling the
#' `wb_color()` function.
#'
#' @param name A name of a color known to R either as name or RGB/ARGB/RGBA value.
#' @param auto A boolean.
#' @param indexed An indexed color value. This color has to be provided by the workbook.
#' @param hex A rgb color a RGB/ARGB/RGBA hex value with or without leading "#".
#' @param theme A zero based index referencing a value in the theme.
#' @param tint A tint value applied. Range from -1 (dark) to 1 (light).
#' @param format A colour format, one of ARGB (default) or RGBA.
#' @seealso [wb_get_base_colors()] [grDevices::colors()]
#' @return a `wbColour` object
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
