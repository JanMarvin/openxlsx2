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
get_named_regions_tab <- function(wb) {
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

#' Get named regions in a workbook
#'
#' @returns A vector of named regions in `x`.
#' @param wb A `wbWorkbook` object
#' @param tables Should data tables be included in the result?
#' @param x deprecated. Use `wb`. For Excel input use [wb_load()] to first load
#'   the xlsx file as a workbook.
#' @seealso [wb_add_named_region()], [wb_get_tables()]
#' @returns A data frame with the all named regions in `wb`. Or `NULL`, if none are found.
#' @export
#' @examples
#' wb <- wb_workbook()
#' wb$add_worksheet("Sheet 1")
#'
#' ## specify region
#' wb$add_data(x = iris, start_col = 1, start_row = 1)
#' wb$add_named_region(
#'   name = "iris",
#'   dims = wb_dims(x = iris)
#' )$add_data(sheet = 1, x = iris, name = "iris2", start_col = 10)
#' ## From Workbook object
#' wb_get_named_regions(wb)
#' # Use this info to extract the data frame
#' df <- wb$to_df(named_region = "iris2")
#' head(df)
#'
#' # Extract tables and named regions
#' wb$add_worksheet()$add_data_table(x = iris)
#'
#' wb$get_named_regions(tables = TRUE)
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
