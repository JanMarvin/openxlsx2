#' R6 class for a Workbook Color
#'
#' A color
#'
#' @export
wbColor <- R6::R6Class(
  "wbColor",

  public = list(

    #' @field auto auto
    auto = NULL,

    #' @field indexed indexed
    indexed = NULL,

    #' @field rgb rgb
    rgb = NULL,

    #' @field theme theme
    theme = NULL,

    #' @field tint tint
    tint = NULL,

    #' @description
    #' Creates a new `wbColor` object
    #' @param name A name of a color known to R
    #' @param auto A boolean.
    #' @param indexed An indexed color values.
    #' @param rgb A rgb color as ARGB hex value "FF000000".
    #' @param theme A zero based index referencing a value in the theme.
    #' @param tint A tint value applied. Range from -1 (dark) to 1 (light).
    #' @return a `wbColor` object
    initialize = function(auto = NULL, indexed = NULL, rgb = NULL, theme = NULL, tint = NULL) {
      self$auto    <- auto
      self$indexed <- indexed
      self$rgb     <- rgb
      self$theme   <- theme
      self$tint    <- tint

      invisible(self)
    },

    #' @description
    #' get as vector
    #' @return A named character vector
    get = function() {
      c(
        auto    = self$auto,
        indexed = self$indexed,
        rgb     = self$rgb,
        theme   = self$theme,
        tint    = self$tint
      )
    },

    #' @description
    #' Convert to xml
    #' @return A character vector of xml
    to_xml = function() {
      xml_node_create(
        "color",
        self$get()
      )
    }
  )
)

#' Create a new hyperlink object
#' @param name a color name as character string
#' @rdname wbColor
#' @export
wb_color <- function(name = NULL, auto = NULL, indexed = NULL, rgb = NULL, theme = NULL, tint = NULL) {
  if (!is.null(name)) rgb <-  validate_colour(name)
  wbColor$new(auto = auto, indexed = indexed, rgb = rgb, theme = theme, tint = tint)
}
