wb_add_comment_impl <- function(
    self,
    private,
    sheet   = current_sheet(),
    col,
    row,
    dims    = rowcol_to_dims(row, col),
    comment
) {
  if (!missing(dims)) {
    xy <- unlist(dims_to_rowcol(dims))
    col <- xy[[1]]
    row <- as.integer(xy[[2]])
  }

  write_comment(
    wb = self,
    sheet = sheet,
    col = col,
    row = row,
    comment = comment
  ) # has no use: xy

  invisible(self)
}

wb_remove_comment_impl <- function(
    self,
    private,
    sheet      = current_sheet(),
    col,
    row,
    dims       = rowcol_to_dims(row, col),
    gridExpand = TRUE
) {
  if (!missing(dims)) {
    xy <- unlist(dims_to_rowcol(dims))
    col <- xy[[1]]
    row <- as.integer(xy[[2]])
    # with gridExpand this is always true
    gridExpand <- TRUE
  }
  remove_comment(wb = self, sheet = sheet, col = col, row = row, gridExpand = TRUE)
  invisible(self)
}
