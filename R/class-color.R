#' Helper to create a color
#'
#' Creates a `wbColour` object.
#' @param name A name of a color known to R either as name or RGB/ARGB value.
#' @param auto A boolean.
#' @param indexed An indexed color value. This color has to be provided by the workbook.
#' @param hex A rgb color either a ARGB hex value or RGB hex value With or without leading "#".
#' @param theme A zero based index referencing a value in the theme.
#' @param tint A tint value applied. Range from -1 (dark) to 1 (light).
#' @seealso [wb_get_base_colors()] [grDevices::colors()]
#' @return a `wbColour` object
#' @export
wb_color <- function(
    name = NULL,
    auto = NULL,
    indexed = NULL,
    hex = NULL,
    theme = NULL,
    tint = NULL
  ) {

  if (!is.null(name)) hex <-  validate_color(name)
  if (!is.null(hex))  hex <-  validate_color(hex)

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
