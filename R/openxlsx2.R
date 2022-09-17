#' xlsx reading, writing and editing.
#'
#' openxlsx simplifies the the process of writing and styling Excel xlsx files from R
#' and removes the dependency on Java.
#'
#' @name openxlsx2
#' @docType package
#' @useDynLib openxlsx2, .registration=TRUE
#'
#' @import Rcpp
#' @import R6
#' @importFrom grDevices bmp col2rgb colours dev.copy dev.list dev.off jpeg png rgb tiff
#' @importFrom magrittr %>%
#' @importFrom stringi stri_c stri_isempty stri_join stri_match stri_match_all_regex stri_order stri_opts_collator stri_rand_strings stri_read_lines stri_replace_all_fixed stri_split_fixed stri_split_regex stri_sub stri_unescape_unicode
#' @importFrom utils download.file head menu read.csv unzip
#' @importFrom zip zip
#'
#' @seealso
#' \itemize{
#'    \item{`vignette("Introduction", package = "openxlsx2")`}
#'    \item{`vignette("formatting", package = "openxlsx2")`}
#'    \item{[write_data()]}
#'    \item{[write_datatable()]}
#'    \item{[write_xlsx()]}
#'    \item{[read_xlsx()]}
#'   }
#' for examples
#'
#' @details
#' The openxlsx package uses global options to simplify formatting:
#'
#' \itemize{
#'    \item{`options("openxlsx2.borderColour" = "black")`}
#'    \item{`options("openxlsx2.borderStyle" = "thin")`}
#'    \item{`options("openxlsx2.dateFormat" = "mm/dd/yyyy")`}
#'    \item{`options("openxlsx2.datetimeFormat" = "yyyy-mm-dd hh:mm:ss")`}
#'    \item{`options("openxlsx2.numFmt" = NULL)`}
#'    \item{`options("openxlsx2.paperSize" = 9)`} ## A4
#'    \item{`options("openxlsx2.orientation" = "portrait")`} ## page orientation
#' }
#'  See the Formatting vignette for examples.
#'
#'
#'
#'
#' Additional options
#'
#' \itemize{
#' \item{`options("openxlsx2.compressionLevel" = "9")`} ## set zip compression level, default is "1".
#' }
#'
NULL

# matches enum celltype
openxlsx2_celltype <- c(
  short_date     = 0,
  long_date      = 1,
  numeric        = 2,
  logical        = 3,
  character      = 4,
  formula        = 5,
  accounting     = 6,
  percentage     = 7,
  scientific     = 8,
  comma          = 9,
  hyperlink      = 10,
  array_formula  = 11,
  factor         = 12
)
