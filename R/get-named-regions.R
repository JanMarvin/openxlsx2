#' convert `<definedName />` xml to R named region data frame
#' @param wb a workbook
#' @returns a data frame in named_region format
#' @noRd
get_named_regions_from_definedName <- function(wb) {

  dn <- wb$workbook$definedNames

  dn <- cbind(
    rbindlist(xml_attr(dn, "definedName")),
    value =  xml_value(dn, "definedName")
  )

  if (!is.null(dn$value)) {
    dn_pos <- dn$value
    dn_pos <- gsub("[$']", "", dn_pos)
    # for ws_page_setup we can have multiple defined names for column and row
    # separated by a colon. This keeps only the first and drops the second.
    # This will allow saving, but changes get_named_regions()
    dn_pos <- vapply(strsplit(dn_pos, ","), FUN = function(x) x[1], NA_character_)

    has_bang <- grepl("!", dn_pos, fixed = TRUE)
    dn$sheets <- ifelse(has_bang, gsub("^(.*)!.*$", "\\1", dn_pos), "")
    dn$coords <- ifelse(has_bang, gsub("^.*!(.*)$", "\\1", dn_pos), "")
  }

  dn$id <- seq_len(nrow(dn))

  if (!is.null(dn$localSheetId)) {
    dn$local <- as.integer(dn$localSheetId != "")
  } else {
    dn$local <- 0
  }
  dn$sheet <- vapply(dn$sheets, function(x) ifelse(x != "", wb_validate_sheet(wb, x), NA_integer_), NA_integer_)

  dn[order(dn[, "local"], dn[, "name"], dn[, "sheet"]),]
}

#' get named region from ´wb$tables`
#' @param a workbook
#' @returns a data frame in named_region format
#' @noRd
get_named_regions_from_table <- function(wb) {
  data.frame(
    #localSheetId is not always available
    name = wb$tables$tab_nam,
    value = "table",
    sheets = wb_get_sheet_name(wb, wb$tables$tab_sheet),
    coords = wb$tables$tab_ref,
    id = NA_integer_,
    local = NA_integer_,
    sheet = wb$tables$tab_sheet,
    stringsAsFactors = FALSE
  )
}

#' @name NamedRegions
#' @title Get create or remove named regions
#' @description Return a vector of named regions in a xlsx file or
#' Workbook object
#' @param x An xlsx file or Workbook object
#' @export
#' @seealso [wb_add_named_region()] [wb_remove_named_region()]
#' @examples
#' ## create named regions
#' wb <- wb_workbook()
#' wb$add_worksheet("Sheet 1")
#'
#' ## specify region
#' wb$add_data(sheet = 1, x = iris, startCol = 1, startRow = 1)
#' wb$add_named_region(
#'   sheet = 1,
#'   name = "iris",
#'   rows = seq_len(nrow(iris) + 1),
#'   cols = seq_along(iris)
#' )
#'
#'
#' ## using write_data 'name' argument to create a named region
#' wb$add_data(sheet = 1, x = iris, name = "iris2", startCol = 10)
#' 
#' out_file <- temp_xlsx()
#' wb$save(out_file, overwrite = TRUE)
#'
#' ## see named regions
#' get_named_regions(wb) ## From Workbook object
#' get_named_regions(out_file) ## From xlsx file
#'
#' ## read named regions
#' df <- read_xlsx(wb, namedRegion = "iris")
#' head(df)
#'
#' df <- read_xlsx(out_file, namedRegion = "iris2")
#' head(df)
get_named_regions <- function(x) {
  if (inherits(x, "wbWorkbook")) {
    wb <- x
  } else {
    wb <- wb_load(x)
  }

  z <- NULL

  if (length(wb$workbook$definedNames)) {
    z <- get_named_regions_from_definedName(wb)
  }

  if (!is.null(wb$tables)) {
    tb <- get_named_regions_from_table(wb)
    z <- merge(z, tb, all = TRUE, sort = FALSE)
  }

  z
}
