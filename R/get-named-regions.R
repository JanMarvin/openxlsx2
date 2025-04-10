#' convert `<definedName />` xml to R named region data frame
#' @param wb a workbook
#' @returns a data frame in named_region format
#' @noRd
get_nr_from_definedName <- function(wb) {

  dn <- wb$workbook$definedNames

  dn <- cbind(
    rbindlist(xml_attr(dn, "definedName")),
    value =  xml_value(dn, "definedName"),
    stringsAsFactors = FALSE
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
get_named_regions_tab <- function(wb) {
  tabs <- wb$tables[wb$tables$tab_act == 1, ]
  data.frame(
    #localSheetId is not always available
    name = tabs$tab_name,
    value = "table",
    sheets = unname(wb$get_sheet_names()[tabs$tab_sheet]),
    coords = tabs$tab_ref,
    id = NA_integer_,
    local = NA_integer_,
    sheet = tabs$tab_sheet,
    stringsAsFactors = FALSE
  )
}

#' @rdname named_region-wb
#' @param tables Should included both data tables and named regions in the result?
#' @param x Deprecated. Use `wb`. For Excel input use [wb_load()] to first load
#'   the xlsx file as a workbook.
#' @seealso [wb_get_tables()]
#' @returns A data frame with the all named regions in `wb`. Or `NULL`, if none are found.
#' @export
#' @examples
#'
#' # Extract named regions from a file
#' out_file <- temp_xlsx()
#' wb_save(wb, out_file, overwrite = TRUE)
#'
#' # Load the file as a workbook first, then get named regions.
#' wb1 <- wb_load(out_file)
#' wb1$get_named_regions()
#'
wb_get_named_regions <- function(wb, tables = FALSE, x = NULL) {
  # TODO merge this doc with wb_add_named_region
  if (!is.null(x)) {
    # Will only show up if the user named `x`
    .Deprecated("wb", old = "x", msg = "Use `wb` instead in `wb_get_named_regions()`")

    if (!missing(wb)) {
      # if a user tries to provide both x and wb.
      stop("x is a deprecated argument. Use wb instead. can't be supplied. Use `wb` only.")
    }

    wb <- x
  }

  if (!inherits(wb, "wbWorkbook")) {
    if (getOption("openxlsx2.soon_deprecated", default = FALSE)) {
      warning(
        "Using `wb_get_named_regions()` on an xlsx file is deprecated.\n",
        "Use `wb_load(file)$get_named_regions()` instead.",
        call. = FALSE
        )
    }
    wb <- wb_load(wb)
  }

  assert_workbook(wb)
  wb$get_named_regions(tables = tables)
}
