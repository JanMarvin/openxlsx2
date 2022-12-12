workbook_add_data <- function(
    self,
    private,
    sheet           = current_sheet(),
    x,
    startCol        = 1,
    startRow        = 1,
    dims            = rowcol_to_dims(startRow, startCol),
    array           = FALSE,
    xy              = NULL,
    colNames        = TRUE,
    rowNames        = FALSE,
    withFilter      = FALSE,
    name            = NULL,
    sep             = ", ",
    applyCellStyle  = TRUE,
    removeCellStyle = FALSE,
    na.strings
) {
  # TODO copy over the actual write_data()
  write_data(
    wb              = self,
    sheet           = sheet,
    x               = x,
    startCol        = startCol,
    startRow        = startRow,
    dims            = dims,
    array           = array,
    xy              = xy,
    colNames        = colNames,
    rowNames        = rowNames,
    withFilter      = withFilter,
    name            = name,
    sep             = sep,
    applyCellStyle  = applyCellStyle,
    removeCellStyle = removeCellStyle,
    na.strings      = na.strings
  )

  invisible(self)
}

workbook_add_data_table <- function(
    self,
    private,
    sheet       = current_sheet(),
    x,
    startCol    = 1,
    startRow    = 1,
    dims        = rowcol_to_dims(startRow, startCol),
    xy          = NULL,
    colNames    = TRUE,
    rowNames    = FALSE,
    tableStyle  = "TableStyleLight9",
    tableName   = NULL,
    withFilter  = TRUE,
    sep         = ", ",
    firstColumn = FALSE,
    lastColumn  = FALSE,
    bandedRows  = TRUE,
    bandedCols  = FALSE,
    applyCellStyle = TRUE,
    removeCellStyle = FALSE,
    na.strings
) {
  write_datatable(
    wb              = self,
    sheet           = sheet,
    x               = x,
    startCol        = startCol,
    startRow        = startRow,
    dims            = dims,
    xy              = xy,
    colNames        = colNames,
    rowNames        = rowNames,
    tableStyle      = tableStyle,
    tableName       = tableName,
    withFilter      = withFilter,
    sep             = sep,
    firstColumn     = firstColumn,
    lastColumn      = lastColumn,
    bandedRows      = bandedRows,
    bandedCols      = bandedCols,
    applyCellStyle  = applyCellStyle,
    removeCellStyle = removeCellStyle,
    na.strings      = na.strings
  )

  invisible(self)
}


workbook_add_formula <- function(
    self,
    private,
    sheet    = current_sheet(),
    x,
    startCol = 1,
    startRow = 1,
    dims     = rowcol_to_dims(startRow, startCol),
    array    = FALSE,
    xy       = NULL,
    applyCellStyle = TRUE,
    removeCellStyle = FALSE
) {
  write_formula(
    wb       = self,
    sheet    = sheet,
    x        = x,
    startCol = startCol,
    startRow = startRow,
    dims     = dims,
    array    = array,
    xy       = xy,
    applyCellStyle = applyCellStyle,
    removeCellStyle = removeCellStyle
  )
  invisible(self)
}
