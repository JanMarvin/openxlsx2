
#' @name getNamedRegions
#' @title Get named regions
#' @description Return a vector of named regions in a xlsx file or
#' Workbook object
#' @param x An xlsx file or Workbook object
#' @export
#' @seealso \code{\link{createNamedRegion}}
#' @examples
#' ## create named regions
#' wb <- createWorkbook()
#' addWorksheet(wb, "Sheet 1")
#'
#' ## specify region
#' writeData(wb, sheet = 1, x = iris, startCol = 1, startRow = 1)
#' createNamedRegion(
#'   wb = wb,
#'   sheet = 1,
#'   name = "iris",
#'   rows = 1:(nrow(iris) + 1),
#'   cols = 1:ncol(iris)
#' )
#'
#'
#' ## using writeData 'name' argument to create a named region
#' writeData(wb, sheet = 1, x = iris, name = "iris2", startCol = 10)
#' \dontrun{
#' out_file <- tempfile(fileext = ".xlsx")
#' saveWorkbook(wb, out_file, overwrite = TRUE)
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

  xmlDir <- file.path(tempdir(), "named_regions_tmp")
  xmlFiles <- unzip(x, exdir = xmlDir)

  workbook <- xmlFiles[grepl("workbook.xml$", xmlFiles, perl = TRUE)]
  workbook <- read_xml(workbook)

  dn <- xml_node(xml = workbook, "workbook", "definedNames", "definedName")
  if (length(dn) == 0) {
    return(NULL)
  }

  dn_names <- get_named_regions_from_string(dn = dn)

  unlink(xmlDir, recursive = TRUE, force = TRUE)

  return(dn_names)
}


#' @export
getNamedRegions.Workbook <- function(x) {
  dn <- x$workbook$definedNames
  if (length(dn) == 0) {
    return(NULL)
  }

  dn_names <- get_named_regions_from_string(dn = dn)

  return(dn_names)
}
