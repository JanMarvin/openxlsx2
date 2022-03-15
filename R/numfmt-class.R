
#' number format class
#'
#' @param x An object
#' @returns A single `character` for the accepted class except for the
#'   `data.frame` method which returns a single-row `data.frame` with the `x` as
#'   column names and results of each
#'
#' @keywords internal
#' @noRd
numfmt_class <- function(x) {
  UseMethod("numfmt_class")
}

# These have to be "exported" to be accessed but the main function isn't

#' @export
numfmt_class.data.frame <- function(x) {
  # check if we are inside a data.frame
  if (getOption("openxlsx2.numfmt.data.frame", FALSE)) {
    stop("nested data.frames are not supported")
  }

  # set the option that we are currently inspecting a data.frame
  options(openxlsx2.numfmt.data.frame = TRUE)
  on.exit(options(openxlsx2.numfmt.data.frame = FALSE), add = TRUE)

  # when numfmt_class.data.frame is called again, we'll hit the failure above
  quick_df(as.list(vapply(x, numfmt_class, NA_character_)))
}

#' @export
numfmt_class.default <- function(x) {
  cl <- paste(class(x), sep = ", ")
  warning("class(s) not supported: ", cl, "\n  using 'character'")
  "character"
}

#' @export
numfmt_class.factor      <- function(x) { "factor"     }
#' @export
numfmt_class.Date        <- function(x) { "date"       }
#' @export
numfmt_class.POSIXct     <- function(x) { "posix"      }
#' @export
numfmt_class.logical     <- function(x) { "logical"    }
#' @export
numfmt_class.character   <- function(x) { "character"  }
#' @export
numfmt_class.integer     <- function(x) { "integer"    }
#' @export
numfmt_class.numeric     <- function(x) { "numeric"    }
#' @export
numfmt_class.currency    <- function(x) { "currency"   }
#' @export
numfmt_class.accounting  <- function(x) { "accounting" }
#' @export
numfmt_class.percentage  <- function(x) { "percentage" }
#' @export
numfmt_class.scientific  <- function(x) { "scientific" }
#' @export
numfmt_class.comma       <- function(x) { "comma"      }
#' @export
numfmt_class.formula     <- function(x) { "formula"    }
#' @export
numfmt_class.hyperlink   <- function(x) { "hyperlink"  }
#' @export
numfmt_class.wbHyperlink <- function(x) { "hyperlink"  }
