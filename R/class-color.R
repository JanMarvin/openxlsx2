
#' Create a new hyperlink object
#' @param name A name of a color known to R
#' @param auto A boolean.
#' @param indexed An indexed color values.
#' @param hex A rgb color as ARGB hex value "FF000000".
#' @param theme A zero based index referencing a value in the theme.
#' @param tint A tint value applied. Range from -1 (dark) to 1 (light).
#' @return a `wbColour` object
#' @rdname wbColour
#' @export
wb_colour <- function(
    name = NULL,
    auto = NULL,
    indexed = NULL,
    hex = NULL,
    theme = NULL,
    tint = NULL
  ) {

  if (!is.null(name)) hex <-  validate_colour(name)

  z <- c(
    auto    = auto,
    indexed = indexed,
    rgb     = hex,
    theme   = theme,
    tint    = tint
  )

  if (is.null(z))
    z <- c(name = "black")

  class(z) <- c("character", "wbColour")
  z
}

is_wbColour <- function(x) inherits(x, "wbColour")
