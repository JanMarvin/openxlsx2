
Hyperlink <- setRefClass(
  "Hyperlink",
  fields = c(
    "ref",
    "target",
    "location",
    "display",
    "is_external"
  ),

  methods = list(
    # initiated the object
    initialize = function(ref, target, location, display = NULL, is_external = TRUE) {
      .self$ref         <- ref
      .self$target      <- target
      .self$location    <- location
      .self$display     <- display
      .self$is_external <- is_external

      invisible(.self)
    },

    # converts the Hyperlink to XML
    to_xml = function(id) {
      paste_c(
        "<hyperlink",
        sprintf('ref="%s"', .self$ref),                     # rf
        if (.self$is_external) sprintf('r:id="rId%s"', id), # rid
        sprintf('display="%s"', .self$display),             # disp
        sprintf('location="%s"', .self$location),           # loc
        "/>",
        sep = " "
      )
    },

    # converts the Hyperlink to target xml
    to_target_xml = function(id) {
      if (.self$is_external) {
        return(sprintf('<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="%s" TargetMode="External"/>', id, .self$target))
      } else {
        return(NULL)
      }
    }
  )
)

new_hyperlink <- function() {
  Hyperlink$new(ref = character(), target = character(), location = character())
}

xml_to_hyperlink <- function(xml) {

  # TODO allow Hyperlink$new(xml = xml)

  # xml <- c('<hyperlink ref="A1" r:id="rId1" location="Authority"/>',
  # '<hyperlink ref="B1" r:id="rId2"/>',
  # '<hyperlink ref="A1" location="Sheet2!A1" display="Sheet2!A1"/>')

  if (length(xml) == 0) {
    return(xml)
  }

  targets <- names(xml) %||% rep(NA, length(xml))
  xml <- unname(xml)

  a <- unlist(lapply(xml, function(i) regmatches(i, gregexpr('[a-zA-Z]+=".*?"', i))), recursive = FALSE)
  names <- lapply(a, function(i) regmatches(i, regexpr('[a-zA-Z]+(?=\\=".*?")', i, perl = TRUE)))
  vals <- lapply(a, function(i) {
    res <- regmatches(i, regexpr('(?<=").*?(?=")', i, perl = TRUE))
    Encoding(res) <- "UTF-8"
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

    Hyperlink$new(
      ref         = ref,
      target      = target,
      location    = location,
      display     = display,
      is_external = is_external
    )
  })
}

new_hyperlink <- function() {
  Hyperlink$new(ref = character(), target = character(), location = character())
}
