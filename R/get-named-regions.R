#' convert `<definedName />` xml to R named region data frame
#' @param wb a workbook
#' @returns a data frame in named_region format
#' @noRd
get_nr_from_definedName <- function(wb) {

  dn <- wb$workbook$definedNames

  dn <- cbind(
    rbindlist(xml_attr(dn, "definedName")),
    value =  xml_value(dn, "definedName")
  )

  if (!is.null(dn$value)) {
    dn_pos <- dn$value
    dn_pos <- gsub("[$']", "", dn_pos)
    # for wb_page_setup we can have multiple defined names for column and row
    # separated by a colon. This keeps only the first and drops the second.
    # This will allow saving, but changes wb_get_named_regions()
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

  dn[order(dn[, "local"], dn[, "name"], dn[, "sheet"]), ]
}

#' get named region from ´wb$tables`
#' @param wb a workbook
#' @returns a data frame in named_region format
#' @noRd
wb_get_named_regions_tab <- function(wb) {
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

#' Get create or remove named regions
#'
#' Return a vector of named regions in a xlsx file or `wbWorkbook` object
#' @param x An xlsx file or `wbWorkbook` object
#' @param tables add tables too if `TRUE`
#' @export
wb_get_named_regions <- function(x, tables = FALSE) {
  if (inherits(x, "wbWorkbook")) {
    wb <- x
  } else {
    wb <- wb_load(x)
  }

  z <- NULL

  if (length(wb$workbook$definedNames)) {
    z <- get_nr_from_definedName(wb)
  }

  if (tables && !is.null(wb$tables)) {
    tb <- wb_get_named_regions_tab(wb)

    if (is.null(z)) {
      z <- tb
    } else {
      z <- merge(z, tb, all = TRUE, sort = FALSE)
    }

  }

  z
}
