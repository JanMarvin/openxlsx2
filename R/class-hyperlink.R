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
#' @noRd
wb_hyperlink <- function() {
  wbHyperlink$new(ref = character(), target = character(), location = character())
}

#' Helper function to create `wbHyperlink` objects from a workbook
#'
#' This function is used only in wb_to_df(show_hyperlinks = TRUE) to construct ref, target and location
#' @param wb a workbook
#' @param sheet a sheet index
#' @noRd
wb_to_hyperlink <- function(wb, sheet = 1) {

  xml     <- wb$worksheets[[sheet]]$hyperlinks
  relship <- wb$worksheets_rels[[sheet]]

  # prepare relships data frame
  relships <- rbindlist(xml_attr(relship, "Relationship"))
  relships <- relships[basename(relships$Type) == "hyperlink", ]

  # prepare hyperlinks data frame
  hlinks <- rbindlist(xml_attr(xml, "hyperlink"))

  # merge both
  hl_df <- merge(hlinks, relships, by.x = "r:id", by.y = "Id", all.x = TRUE, all.y = FALSE)

  lapply(seq_len(nrow(hl_df)), function(i) {

    x <- hl_df[i, ]

    ref         <- x$ref
    target      <- x$Target
    location    <- x$location
    display     <- x$display
    is_external <- x$TargetMode == "External"

    wbHyperlink$new(
      ref         = ref,
      target      = target,
      location    = location,
      display     = display,
      is_external = is_external
    )
  })
}
