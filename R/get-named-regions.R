
#' @name NamedRegions
#' @title Get create or remove named regions
#' @description Return a vector of named regions in a xlsx file or
#' Workbook object
#' @param x An xlsx file or Workbook object
#' @export
#' @seealso [wb_create_named_region()] [wb_delete_named_region()]
#' @examples
#' ## create named regions
#' wb <- wb_workbook()
#' wb$add_worksheet("Sheet 1")
#'
#' ## specify region
#' write_data(wb, sheet = 1, x = iris, startCol = 1, startRow = 1)
#' wb_create_named_region(
#'   wb = wb,
#'   sheet = 1,
#'   name = "iris",
#'   rows = seq_len(nrow(iris) + 1),
#'   cols = seq_along(iris)
#' )
#'
#'
#' ## using write_data 'name' argument to create a named region
#' write_data(wb, sheet = 1, x = iris, name = "iris2", startCol = 10)
#' \dontrun{
#' out_file <- tempfile(fileext = ".xlsx")
#' wb_save(wb, out_file, overwrite = TRUE)
#'
#' ## see named regions
#' getNamedRegions(wb) ## From Workbook object
#' getNamedRegions(out_file) ## From xlsx file
#'
#' ## read named regions
#' df <- read_xlsx(wb, namedRegion = "iris")
#' head(df)
#'
#' df <- read_xlsx(out_file, namedRegion = "iris2")
#' head(df)
#' }
getNamedRegions <- function(x) {
  UseMethod("getNamedRegions", x)
}


#' @export
getNamedRegions.default <- function(x) {
  wb <- wb_load(x)
  dn <- wb$workbook$definedNames

  if (length(dn) == 0) {
    return(NULL)
  }

  get_named_regions_from_string(wb = wb, dn = dn)
}


#' @export
getNamedRegions.wbWorkbook <- function(x) {
  dn <- x$workbook$definedNames

  if (length(dn) == 0) {
    return(NULL)
  }

  get_named_regions_from_string(wb = x, dn = dn)
}
