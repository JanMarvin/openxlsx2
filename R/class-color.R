#' Create a new wbColour object
#' @param name A name of a color known to R
#' @param auto A boolean.
#' @param indexed An indexed color values.
#' @param hex A rgb color as ARGB hex value "FF000000".
#' @param theme A zero based index referencing a value in the theme.
#' @param tint A tint value applied. Range from -1 (dark) to 1 (light).
#' @return a `wbColour` object
#' @rdname wbColour
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

  z <- c(
    auto    = auto,
    indexed = indexed,
    rgb     = hex,
    theme   = theme,
    tint    = tint
  )

  if (is.null(z))
    z <- c(name = "black")

  # wbColour for historical reasons
  class(z) <- c("character", "wbColour")
  z
}

#' @export
#' @rdname wbColour
#' @usage NULL
wb_colour <- wb_color

is_wbColour <- function(x) inherits(x, "wbColour")

#' takes colour and returns color
#' @param ... ...
#' @returns void. assigns an object in the parent frame
#' @keywords internal
#' @noRd
standardize_color_names <- function(...) {

  got <- ...names()
  # can be Color or color
  got_color <- which(grepl("colour", tolower(got)))

  if (length(got_color)) {
    for (got_col in got_color) {
      color <- got[got_col]
      name_color <- stringi::stri_replace_all_fixed(color, "olour", "olor", )
      value_color <- ...elt(got_col)
      assign(name_color, value_color, parent.frame())
    }
  }
}
