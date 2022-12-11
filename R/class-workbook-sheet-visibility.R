workbook_get_sheet_visibility <- function(self, private) {
  state <- rep("visible", length(self$workbook$sheets))
  state[grepl("hidden", self$workbook$sheets)] <- "hidden"
  state[grepl("veryHidden", self$workbook$sheets, ignore.case = TRUE)] <- "veryHidden"
  state
}

workbook_set_sheet_visibility = function(
    self,
    private,
    sheet = current_sheet(),
    value
) {
  if (length(value) != length(sheet)) {
    stop("`value` and `sheet` must be the same length")
  }

  sheet <- private$get_sheet_index(sheet)

  value <- tolower(as.character(value))
  value[value %in% "true"] <- "visible"
  value[value %in% "false"] <- "hidden"
  value[value %in% "veryhidden"] <- "veryHidden"

  exState0 <- reg_match0(self$workbook$sheets[sheet], '(?<=state=")[^"]+')
  exState <- tolower(exState0)
  exState[exState %in% "true"] <- "visible"
  exState[exState %in% "hidden"] <- "hidden"
  exState[exState %in% "false"] <- "hidden"
  exState[exState %in% "veryhidden"] <- "veryHidden"


  inds <- which(value != exState)

  if (length(inds) == 0) {
    return(invisible(self))
  }

  for (i in seq_along(self$worksheets)) {
    self$workbook$sheets[sheet[i]] <- gsub(exState0[i], value[i], self$workbook$sheets[sheet[i]], fixed = TRUE)
  }

  if (!any(self$get_sheet_visibility() %in% c("true", "visible"))) {
    warning("A workbook must have atleast 1 visible worksheet.  Setting first for visible")
    self$set_sheet_visibility(1, TRUE)
  }

  invisible(self)
}
