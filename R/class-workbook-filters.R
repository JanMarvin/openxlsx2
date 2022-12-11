workbook_add_filter <- function(
    self,
    private,
    sheet = current_sheet(),
    rows,
    cols
) {
  sheet <- private$get_sheet_index(sheet)

  if (length(rows) != 1) {
    stop("row must be a numeric of length 1.")
  }

  if (!is.numeric(cols)) {
    cols <- col2int(cols)
  }

  self$worksheets[[sheet]]$autoFilter <- sprintf(
    '<autoFilter ref="%s"/>',
    paste(get_cell_refs(data.frame("x" = c(rows, rows), "y" = c(min(cols), max(cols)))), collapse = ":")
  )

  invisible(self)
}

workbook_remove_filter <- function(self, private, sheet = current_sheet()) {
  for (s in private$get_sheet_index(sheet)) {
    self$worksheets[[s]]$autoFilter <- character()
  }

  invisible(self)
}
