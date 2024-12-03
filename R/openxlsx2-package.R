#' xlsx reading, writing and editing.
#'
#' This R package is a modern reinterpretation of the widely used popular
#' `openxlsx` package. Similar to its predecessor, it simplifies the creation of xlsx
#' files by providing a clean interface for writing, designing and editing worksheets.
#' Based on a powerful XML library and focusing on modern programming flows in pipes
#' or chains, `openxlsx2` allows to break many new ground.
#'
#' @useDynLib openxlsx2, .registration=TRUE
#'
#' @import Rcpp
#' @import R6
#' @importFrom grDevices bmp col2rgb colors dev.copy dev.list dev.off jpeg palette png rgb tiff
#' @importFrom magrittr %>%
#' @importFrom stringi stri_encode stri_extract_all_regex stri_extract_first_regex
#'   stri_join stri_match_first_regex stri_order stri_opts_collator stri_pad_left
#'   stri_rand_strings stri_read_lines stri_replace_all_fixed stri_replace_all_regex
#'   stri_split_fixed stri_sub stri_unescape_unicode stri_unique
#' @importFrom utils download.file head menu read.csv unzip
#' @importFrom zip zip
#'
#' @seealso
#' * `browseVignettes("openxlsx2")`
#' * <https://janmarvin.github.io/openxlsx2/>
#' * <https://janmarvin.github.io/ox2-book/>
#' for examples
#'
#' @details
#' The `openxlsx2` package provides comprehensive functionality for interacting
#' with Office Open XML spreadsheet files. Users can read data using [read_xlsx()]
#' and write data to spreadsheets via [write_xlsx()], with options to specify
#' sheet names and cell ranges for targeted operations. Beyond basic read/write
#' capabilities, `openxlsx2` facilitates extensive workbook ([wb_workbook()])
#' manipulations, including:
#' * Loading a workbook into R with [wb_load()] and saving it with [wb_save()]
#' * Adding/removing and modifying worksheets and data with [wb_add_worksheet()],
#'  [wb_remove_worksheet()], and [wb_add_data()].
#' * Enhancing spreadsheets with comments ([wb_add_comment()]),
#'  images ([wb_add_image()]), plots ([wb_add_plot()]), charts ([wb_add_mschart()]),
#'  and pivot tables ([wb_add_pivot_table()]).
#' Customizing cell styles using fonts ([wb_add_font()]),
#' number formats ([wb_add_numfmt()]), backgrounds ([wb_add_fill()]),
#' and alignments ([wb_add_cell_style()]). Inserting custom text strings with
#' [fmt_txt()] and creating comprehensive table styles with [create_tablestyle()].
#'
#' ## Interaction
#' Interaction with `openxlsx2` objects can occur through two primary methods:
#'
#' *Wrapper Function Method*: Utilizes the `wb` family of functions that support
#'  piping to streamline operations.
#' ``` r
#' wb <- wb_workbook(creator = "My name here") %>%
#'   wb_add_worksheet(sheet = "Expenditure", grid_lines = FALSE) %>%
#'   wb_add_data(x = USPersonalExpenditure, row_names = TRUE)
#' ```
#' *Chaining Method*: Directly modifies the object through a series of chained
#'  function calls.
#' ``` r
#' wb <- wb_workbook(creator = "My name here")$
#'   add_worksheet(sheet = "Expenditure", grid_lines = FALSE)$
#'   add_data(x = USPersonalExpenditure, row_names = TRUE)
#' ```
#'
#' While wrapper functions require explicit assignment of their output to reflect
#' changes, chained functions inherently modify the input object. Both approaches
#' are equally supported, offering flexibility to suit user preferences. The
#' documentation mainly highlights the use of wrapper functions. To find information,
#' users should look up the wb function name e.g. `?wb_add_data_table` rather than
#' searching for `?wbWorkbook`.
#'
#' Function arguments follow the snake_case convention, but for backward compatibility,
#' camelCase is also supported at the moment. The API aims to maintain consistency
#' in its arguments, with a special focus on `sheet` ([wb_get_sheet_names()]) and
#' `dims` ([wb_dims]), which are of particular importance to users.
#'
#' ## Locale
#'
#' By default, `openxlsx2` uses the American English word for color (written with
#' 'o' instead of the British English 'ou'). However, both spellings are supported.
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
#' and [`pugixml`](https://github.com/zeux/pugixml) (by Arseny Kapoulkine; COPYRIGHT 2006-2023). Both released under the MIT license.
#' @keywords internal
#' @examples
#' # read xlsx or xlsm files
#' path <- system.file("extdata/openxlsx2_example.xlsx", package = "openxlsx2")
#' read_xlsx(path)
#'
#' # or import workbooks
#' wb <- wb_load(path)
#'
#' # read a data frame
#' wb_to_df(wb)
#'
#' # and save
#' temp <- temp_xlsx()
#' if (interactive()) wb_save(wb, temp)
#'
#' ## or create one yourself
#' wb <- wb_workbook()
#' # add a worksheet
#' wb$add_worksheet("sheet")
#' # add some data
#' wb$add_data("sheet", cars)
#' # open it in your default spreadsheet software
#' if (interactive()) wb$open()
"_PACKAGE"

#' Options consulted by openxlsx2
#'
#' @description
#' The openxlsx2 package allows the user to set global options to simplify formatting:
#'
#' If the built-in defaults don't suit you, set one or more of these options.
#' Typically, this is done in the `.Rprofile` startup file
#'
#' * `options("openxlsx2.borderColor" = "black")`
#' * `options("openxlsx2.borderStyle" = "thin")`
#' * `options("openxlsx2.dateFormat" = "mm/dd/yyyy")`
#' * `options("openxlsx2.datetimeFormat" = "yyyy-mm-dd hh:mm:ss")`
#' * `options("openxlsx2.maxWidth" = NULL)` (Maximum width allowed in Excel is 250)
#' * `options("openxlsx2.minWidth" = NULL)`
#' * `options("openxlsx2.numFmt" = NULL)`
#' * `options("openxlsx2.paperSize" = 9)` corresponds to a A4 paper size
#' * `options("openxlsx2.orientation" = "portrait")` page orientation
#' * `options("openxlsx2.sheet.default_name" = "Sheet")`
#' * `options("openxlsx2.rightToLeft" = NULL)`
#' * `options("openxlsx2.soon_deprecated" = FALSE)` Set to `TRUE` if you want a
#'    warning if using some functions deprecated recently in openxlsx2
#' * `options("openxlsx2.creator")` A default name for the creator of new
#'   `wbWorkbook` object with [wb_workbook()] or new comments with [wb_add_comment()]
#' * `options("openxlsx2.thread_id")` the default person id when adding a threaded comment
#'   to a cell with [wb_add_thread()]
#' * `options("openxlsx2.accountingFormat" = 4)`
#' * `options("openxlsx2.currencyFormat" = 4)`
#' * `options("openxlsx2.commaFormat" = 3)`
#' * `options("openxlsx2.percentageFormat" = 10)`
#' * `options("openxlsx2.scientificFormat" = 48)`
#' * `options("openxlsx2.string_nums" = TRUE)` numerics in character columns
#'    will be converted. `"1"` will be written as `1`
#' * `options("openxlsx2.na.strings" = "#N/A")` consulted by `write_xlsx()`,
#'   `wb_add_data()` and `wb_add_data_table()`.
#' * `options("openxlsx2.compression_level" = 6)` compression level for the output file. Increasing compression and time consumed from 1-9.
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
  hms_time       = 15,
  currency       = 16,
  list           = 17
)

#' Deprecated functions in package *openxlsx2*
#'
#' @description
#' These functions are provided for compatibility with older versions of `openxlsx2`,
#' and may be defunct as soon as the next release. This guide helps you update your
#' code to the latest standards.
#'
#' As of openxlsx2 v1.0, API change should be minimal.
#'
#' # Internal functions
#'
#' These functions are used internally by openxlsx2. It is no longer advertised
#' to use them in scripts. They originate from openxlsx, but do not fit openxlsx2's API.
#'
#' You should be able to modify
#' * [delete_data()] -> [wb_clean_sheet()]
#' * [write_data()] -> [wb_add_data()]
#' * [write_datatable()] -> [wb_add_data_table()]
#' * [write_comment()] -> [wb_add_comment()]
#' * [remove_comment()] -> [wb_remove_comment()]
#' * [write_formula()] -> [wb_add_formula()]
#'
#' You should be able to change those with minimal changes
#'
#' # Deprecated functions
#'
#' First of all, you can set an option that will add warnings when using deprecated
#' functions.
#'
#' ```
#' options("openxlsx2.soon_deprecated" = TRUE)
#' ```
#'
#' # Argument changes
#'
#' For consistency, arguments were renamed to snake_case for the 0.8 release.
#' It is now recommended to use `dims` (the cell range) in favor of `row`, `col`, `start_row`, `start_col`
#'
#' See [wb_dims()] as it provides many options on how to provide cell range
#'
#' # Functions with a new name
#'
#' These functions were renamed for consistency.
#' * [convertToExcelDate()] -> [convert_to_excel_date()]
#' * [wb_grid_lines()] -> [wb_set_grid_lines()]
#' * [create_comment()] -> [wb_comment()]
#'
#'
#' # Deprecated usage
#'
#' * `wb_get_named_regions()` will no longer allow providing a file.
#'
#' ```
#' ## Before
#' wb_get_named_regions(file)
#'
#' ## Now
#' wb <- wb_load(file)
#' wb_get_named_regions(wb)
#' # also possible
#' wb_load(file)$get_named_regions()`
#' ```
#'
#' @seealso [.Deprecated]
#' @name openxlsx2-deprecated
NULL
