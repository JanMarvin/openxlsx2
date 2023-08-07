#' xlsx reading, writing and editing.
#'
#' This R package is a modern reinterpretation of the widely used popular
#' `openxlsx` package. Similar to its predecessor, it simplifies the creation of xlsx
#' files by providing a clean interface for writing, designing and editing worksheets.
#' Based on a powerful XML library and focusing on modern programming flows in pipes
#' or chains, `openxlsx2` allows to break many new ground.
#'
#' @name openxlsx2-package
#' @docType package
#' @useDynLib openxlsx2, .registration=TRUE
#'
#' @import Rcpp
#' @import R6
#' @importFrom grDevices bmp col2rgb colors dev.copy dev.list dev.off jpeg png rgb tiff
#' @importFrom magrittr %>%
#' @importFrom stringi stri_c stri_encode stri_isempty stri_join stri_match
#'   stri_match_all_regex stri_order stri_opts_collator stri_rand_strings
#'   stri_read_lines stri_replace_all_fixed stri_split_fixed stri_split_regex
#'   stri_sub stri_unescape_unicode stri_unique
#' @importFrom utils download.file head menu read.csv unzip
#' @importFrom zip zip
#'
#' @seealso
#' * `browseVignettes("openxlsx2")`
#' * [wb_add_data()]
#' * [wb_add_data_table()]
#' * [wb_to_df()]
#' * [read_xlsx()]
#' * [write_xlsx()]
#' * <https://janmarvin.github.io/openxlsx2/>
#' for examples
#'
#' @details
#' By default, openxlsx2 uses the American English word for color (written with 'o' instead of the British English 'ou').
#' However, both spellings are supported.
#' So where the documentation uses a 'color', the function should also accept a 'colour'.
#' However, this is not indicated by the autocompletion.
#'
#' ## Authors and contributions
#'
#' For a full list of all authors that have made this package possible and for whom we are grateful, please see:
#'
#' ``` r
#' system.file("AUTHORS", package = "openxlsx2")
#' ```
#'
#' If you feel like you should be included on this list, please let us know.
#' If you have something to contribute, you are welcome.
#' If something is not working as expected, open issues or if you have solved an issue, open a pull request.
#' Please be respectful and be aware that we are volunteers doing this for fun in our unpaid free time.
#' We will work on problems when we have time or need.
#'
#' ## License
#'
#' This package is licensed under the MIT license and
#' is based on [`openxlsx`](https://github.com/ycphs/openxlsx) (by Alexander Walker and Philipp Schauberger; COPYRIGHT 2014-2022)
#' and [`pugixml`](https://github.com/zeux/pugixml) (by Arseny Kapoulkine; COPYRIGHT 2006-2022). Both released under the MIT license.
#' @keywords internal
"_PACKAGE"

#' openxlsx2 Options
#'
#' The openxlsx2 package uses global options to simplify formatting:
#'
#' * `options("openxlsx2.borderColor" = "black")`
#' * `options("openxlsx2.borderStyle" = "thin")`
#' * `options("openxlsx2.dateFormat" = "mm/dd/yyyy")`
#' * `options("openxlsx2.datetimeFormat" = "yyyy-mm-dd hh:mm:ss")`
#' * `options("openxlsx2.numFmt" = NULL)`
#' * `options("openxlsx2.paperSize" = 9)` ## A4
#' * `options("openxlsx2.orientation" = "portrait")` ## page orientation
#' * `options("openxlsx2.sheet.default_name" = "Sheet")`
#' * `options("openxlsx2.rightToLeft" = NULL)`
#'
#' @name openxlsx2_options
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
  factor         = 12,
  string_nums    = 13,
  cm_formula     = 14,
  hms_time       = 15
)

#' Deprecated Functions in Package *openxlsx2*
#'
#' These functions are provided for compatibility with older versions of `openxlsx2`, and may be defunct as soon as the next release.
#' @details
#' * [convertToExcelDate()] -> [convert_to_excel_date()]
#' * [wb_grid_lines()] -> [wb_set_grid_lines()]
#' @seealso [.Deprecated]
#' @name openxlsx2-deprecated
NULL
