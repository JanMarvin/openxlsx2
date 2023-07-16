#' R6 class for a Workbook Hyperlink
#'
#' A hyperlink
#'
#' @noRd
wbHyperlink <- R6::R6Class(
  "wbHyperlink",

  public = list(

    #' @field ref ref
    ref = NULL,

    #' @field target target
    target = NULL,

    #' @field location location
    location = NULL,

    #' @field display display
    display = NULL,

    #' @field is_external is_external
    is_external = NULL,

    #' @description
    #' Creates a new `wbHyperlink` object
    #' @param ref ref
    #' @param target target
    #' @param location location
    #' @param display display
    #' @param is_external is_external
    #' @return a `wbHyperlink` object
    initialize = function(ref, target, location, display = NULL, is_external = TRUE) {
      self$ref         <- ref
      self$target      <- target
      self$location    <- location
      self$display     <- display
      self$is_external <- is_external

      invisible(self)
    },

    #' @description
    #' Convert to xml
    #' @param id ???
    #' @return A character vector of xml
    to_xml = function(id) {
      paste_c(
        "<hyperlink",
        sprintf('ref="%s"', self$ref),                     # rf
        if (self$is_external) sprintf('r:id="rId%s"', id), # rid
        sprintf('display="%s"', self$display),             # disp
        sprintf('location="%s"', self$location),           # loc
        "/>",
        sep = " "
      )
    },

    # TODO is this needed?  If target is TRUE then use this instead?

    #' @description
    #' Convert to target xml
    #' @param id ???
    #' @returns A character vector of html if `is_external` is `TRUE`, otherwise `NULL`
    to_target_xml = function(id) {
      if (self$is_external) {
        sprintf('<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="%s" TargetMode="External"/>', id, self$target)
      }
    }
  )
)

#' Create a new hyperlink object
#' @rdname Hyperlink
wb_hyperlink <- function() {
  wbHyperlink$new(ref = character(), target = character(), location = character())
}


xml_to_hyperlink <- function(xml) {
  # xml_to_hyperlink() is used once in wb_load()

  # TODO allow wbHyperlink$new(xml = xml)

  # xml <- c('<hyperlink ref="A1" r:id="rId1" location="Authority"/>',
  # '<hyperlink ref="B1" r:id="rId2"/>',
  # '<hyperlink ref="A1" location="Sheet2!A1" display="Sheet2!A1"/>')

  if (length(xml) == 0) {
    return(xml)
  }

  targets <- names(xml) %||% rep(NA, length(xml))
  xml <- unname(xml)

  # TODO a, names, and vals could be moved within the larger lapply()
  a <- unapply(xml, function(i) regmatches(i, gregexpr('[a-zA-Z]+=".*?"', i)), .recurse = FALSE)
  names <- lapply(a, function(i) regmatches(i, regexpr('[a-zA-Z]+(?=\\=".*?")', i, perl = TRUE)))
  vals <- lapply(a, function(i) {
    res <- regmatches(i, regexpr('(?<=").*?(?=")', i, perl = TRUE))
    res
  })

  lapply(seq_along(xml), function(i) {
    tmp_vals <- vals[[i]]
    tmp_nms <- names[[i]]
    names(tmp_vals) <- tmp_nms

    ## ref
    ref <- tmp_vals[["ref"]]

    ## location
    location <- if ("location" %in% tmp_nms) tmp_vals[["location"]]
    display <- if ("display" %in% tmp_nms) tmp_vals[["display"]]

    ## target/external
    if (is.na(targets[i])) {
      target <- NULL
      is_external <- FALSE
    } else {
      is_external <- TRUE
      target <- targets[i]
    }

    wbHyperlink$new(
      ref         = ref,
      target      = target,
      location    = location,
      display     = display,
      is_external = is_external
    )
  })
}
