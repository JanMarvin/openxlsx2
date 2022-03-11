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
#' @importFrom grDevices bmp col2rgb colours dev.copy dev.list dev.off jpeg png rgb tiff
#' @importFrom magrittr %>%
#' @importFrom stats na.omit pchisq
#' @importFrom stringi stri_c stri_conv stri_isempty stri_join stri_match stri_replace_all_fixed stri_split_fixed stri_split_regex stri_sub
#' @importFrom utils download.file head menu unzip
#' @importFrom zip zipr
#'
#' @seealso
#' \itemize{
#'    \item{`vignette("Introduction", package = "openxlsx2")`}
#'    \item{`vignette("formatting", package = "openxlsx2")`}
#'    \item{[writeData()]}
#'    \item{[writeDataTable()]}
#'    \item{[write.xlsx()]}
#'    \item{[read.xlsx()]}
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
