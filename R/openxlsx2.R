#' xlsx reading, writing and editing.
#'
#' openxlsx simplifies the the process of writing and styling Excel xlsx files from R
#' and removes the dependency on Java.
#'
#' @name openxlsx2
#' @docType package
#' @useDynLib openxlsx2, .registration=TRUE
#' @importFrom zip zipr
#' @importFrom utils download.file head menu unzip
#' @import stringi
#'
#' @seealso
#' \itemize{
#'    \item{\code{vignette("Introduction", package = "openxlsx2")}}
#'    \item{\code{vignette("formatting", package = "openxlsx2")}}
#'    \item{\code{\link{writeData}}}
#'    \item{\code{\link{writeDataTable}}}
#'    \item{\code{\link{write.xlsx}}}
#'    \item{\code{\link{read.xlsx}}}
#'   }
#' for examples
#'
#' @details
#' The openxlsx package uses global options to simplify formatting:
#'
#' \itemize{
#'    \item{\code{options("openxlsx2.borderColour" = "black")}}
#'    \item{\code{options("openxlsx2.borderStyle" = "thin")}}
#'    \item{\code{options("openxlsx2.dateFormat" = "mm/dd/yyyy")}}
#'    \item{\code{options("openxlsx2.datetimeFormat" = "yyyy-mm-dd hh:mm:ss")}}
#'    \item{\code{options("openxlsx2.numFmt" = NULL)}}
#'    \item{\code{options("openxlsx2.paperSize" = 9)}} ## A4
#'    \item{\code{options("openxlsx2.orientation" = "portrait")}} ## page orientation
#' }
#'  See the Formatting vignette for examples.
#'
#'
#'
#'
#' Additional options
#'
#' \itemize{
#' \item{\code{options("openxlsx2.compressionLevel" = "9")}} ## set zip compression level, default is "1".
#' }
#'
NULL
