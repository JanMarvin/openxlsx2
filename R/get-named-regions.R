
#' @name getNamedRegions
#' @title Get named regions
#' @description Return a vector of named regions in a xlsx file or
#' Workbook object
#' @param x An xlsx file or Workbook object
#' @export
#' @seealso [createNamedRegion()]
#' @examples
#' ## create named regions
#' wb <- wb_workbook()
#' wb$addWorksheet("Sheet 1")
#'
#' ## specify region
#' writeData(wb, sheet = 1, x = iris, startCol = 1, startRow = 1)
#' createNamedRegion(
#'   wb = wb,
#'   sheet = 1,
#'   name = "iris",
#'   rows = seq_len(nrow(iris) + 1),
#'   cols = seq_along(iris)
#' )
#'
#'
#' ## using writeData 'name' argument to create a named region
#' writeData(wb, sheet = 1, x = iris, name = "iris2", startCol = 10)
#' \dontrun{
#' out_file <- tempfile(fileext = ".xlsx")
#' wb_save(wb, out_file, overwrite = TRUE)
#'
#' ## see named regions
#' getNamedRegions(wb) ## From Workbook object
#' getNamedRegions(out_file) ## From xlsx file
#'
#' ## read named regions
#' df <- read.xlsx(wb, namedRegion = "iris")
#' head(df)
#'
#' df <- read.xlsx(out_file, namedRegion = "iris2")
#' head(df)
#' }
getNamedRegions <- function(x) {
  UseMethod("getNamedRegions", x)
}

#' @export
getNamedRegions.default <- function(x) {
  if (!file.exists(x)) {
    stop(sprintf("File '%s' does not exist.", x))
  }

  xmlDir <- tempfile("named_regions_tmp")
  # don't unlink on exit
  on.exit(unlink(xmlDir, recursive = TRUE), add = TRUE)

  xmlFiles <- unzip(x, exdir = xmlDir)

  workbook <- grep("workbook.xml$", xmlFiles, perl = TRUE, value = TRUE)
  workbook <- read_xml(workbook)

  dn <- xml_node(xml = workbook, "workbook", "definedNames", "definedName")

  if (length(dn) == 0) {
    return(NULL)
  }

  get_named_regions_from_string(dn = dn)
}


#' @export
getNamedRegions.wbWorkbook <- function(x) {
  dn <- x$workbook$definedNames

  if (length(dn) == 0) {
    return(NULL)
  }

  get_named_regions_from_string(dn = dn)
}
