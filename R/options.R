#' `openxlsx` options
#'
#' A list of standard options from `openxlsx`
#'
#' @description A named `list` of options that are set during `.onAttach()` and
#'   are available for global controls of other functions.
#'
#' \details{
#'   \item{`bandedCols`}{Table style colun bands}
#'   \item{`bandedRows`}{Table style row bands}
#'   \item{`firstColumn`}{Table style, first column}
#'   \item{`lastColumn`}{Table style, last column}
#'   \item{`withFilter`}{Include filters on tables}
#'   \item{`gridLines`}{Show grid lines}
#'   \item{`tableStyle`}{Table style}
#'
#'   \item{`borders`}{Cell borders (e.g., `"none"`, `"sorrounding"`, `"all"`)}
#'   \item{`borderColour`}{Style border colour}
#'   \item{`bordersStyle`}{Cell border style}
#'
#'   \item{`numFormat`}{Number format}
#'   \item{`dateFormat`}{Date format style}
#'   \item{`datetimeFormat`}{Date time format style}
#'
#'   \item{`header`}{Document headers; prioritized last}
#'   \item{`footer`}{Document footers; prioritized last}
#'   \item{`firstHeader`}{Document header; first page only; prioritized first}
#'   \item{`firstFooter`}{Document footer; first page only; prioritized first}
#'   \item{`evenFooter`}{Document footers; even pages}
#'   \item{`evenHeader`}{Document headers; even pages}
#'   \item{`oddFooter`}{Document footers; odd pages}
#'   \item{`oddHeader`}{Document headers; odd pages}
#'
#'   \item{`dpi`}{Dots per inch (see `hdpi` and `vdpi` for other controls)}
#'   \item{`hdpi`}{Horizontal dots per inch; prioritized over `dpi`}
#'   \item{`vdpi`}{Vertical dots per inch; prioritized over `dpi`}
#'
#'   \item{`creator`}{Sets workbook creator (used in internal `get_creator()`; see [wbWorkbook] for more details)}
#'   \item{`compressionLevel`}{Passed to [zip::zip()]}
#'   \item{`na`}{String to convert to `NA`}
#'   \item{`keepNA`}{Convert `NA` to `"#N/A"` (or string in `na`)}
#'
#'   \item{`orientation`}{paper orientation}
#'   \item{`paperSize`}{paper size, in inches}
#'   \item{`tabColour`}{tab colour}
#' }
#'
#' @export
op.openxlsx <- list(
  openxlsx.bandedCols       = FALSE,
  openxlsx.bandedRows       = TRUE,
  openxlsx.firstColumn      = FALSE,
  openxlsx.lastColumn       = FALSE,
  openxlsx.withFilter       = NULL,
  openxlsx.gridLines        = NA,
  openxlsx.tableStyle       = "TableStyleLight9",

  openxlsx.borders          = "none",
  openxlsx.borderColour     = "black",
  openxlsx.borderStyle      = "thin",

  openxlsx.numFmt           = "GENERAL",
  openxlsx.dateFormat       = "mm/dd/yyyy",
  openxlsx.datetimeFormat   = "yyyy-mm-dd hh:mm:ss",

  openxlsx.header           = NULL, # consider list(left = "",  center = "", right = "")
  openxlsx.footer           = NULL, # consider list(left = "",  center = "", right = "")
  # allows NA to turn off?
  openxlsx.firstFooter      = NULL,
  openxlsx.firstHeader      = NULL,
  openxlsx.evenFooter       = NULL, # consider list(left = "",  center = "", right = "")
  openxlsx.evenHeader       = NULL, # consider list(left = "",  center = "", right = "")
  openxlsx.oddFooter        = NULL,
  openxlsx.oddHeader        = NULL,

  # openxlsx.headerStyle      = NULL, # shouldn't this be a style object?

  openxlsx.dpi              = 300,
  openxlsx.hdpi             = 300,
  openxlsx.vdpi             = 300,

  openxlsx.creator          = NULL,
  openxlsx.compressionLevel = 9,
  openxlsx.na               = NULL, # default to "#N/A" ?
  openxlsx.keepNA           = FALSE,

  # openxlsx.maxWidth         = 250,
  # openxlsx.minWidth         = 3,
  openxlsx.orientation      = "portrait",
  openxlsx.paperSize        = 9,
  openxlsx.tabColour        = NULL
)
