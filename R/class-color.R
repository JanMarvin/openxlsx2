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

#' @export
#' @rdname wbColour
#' @usage NULL
wb_color <- wb_colour

is_wbColour <- function(x) inherits(x, "wbColour")

#' takes color and returns colour
#' @param ... ...
#' @returns named colour argument
#' @keywords internal
#' @noRd
standardise_color_names <- function(...) {

  got <- ...names()
  # can be Colour or colour
  got_colour <- which(grepl("color", tolower(got)))

  if (length(got_colour)) {
    for (got_col in got_colour) {
      colour <- got[got_col]
      name_color <- stringi::stri_replace_all_fixed(colour, "olor", "olour", )
      value_color <- ...elt(got_col)
      assign(name_color, value_color, parent.frame())
    }
  }
}

#' takes colour and returns color
#' @param ... ...
#' @returns named color argument
#' @keywords internal
#' @noRd
standardize_colour_names <- function(...) {

  got <- ...names()
  # can be Colour or colour
  got_colour <- which(grepl("colour", tolower(got)))

  if (length(got_colour)) {
    for (got_col in got_colour) {
      colour <- got[got_col]
      name_color <- stringi::stri_replace_all_fixed(colour, "olour", "olor", )
      value_color <- ...elt(got_col)
      assign(name_color, value_color, parent.frame())
    }
  }
}
