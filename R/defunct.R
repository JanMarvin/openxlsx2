#' Internal comment functions
#'
#' Users are advised to use [wb_add_comment()] and [wb_remove_comment()].
#' `write_comment()` and `remove_comment()` are now defunct. Use the
#' replacement functions.
#' @name comment_internal
NULL

#' @rdname comment_internal
#' @inheritParams wb_add_comment
#' @param comment An object created by [create_comment()]
#' @param row,col Row and column of the cell
#' @param color optional background color
#' @param file optional background image (file extension must be png or jpeg)
#' @keywords internal
#' @export
#' @inherit wb_add_comment examples
write_comment <- function(
    wb,
    sheet,
    col     = NULL,
    row     = NULL,
    comment,
    dims    = rowcol_to_dim(row, col),
    color   = NULL,
    file    = NULL
) {
  .Defunct("wb_add_comment()", package = "openxlsx2", msg = "write_comment() is defunct")
}

#' @rdname comment_internal
#' @param gridExpand If `TRUE`, all data in rectangle min(rows):max(rows) X min(cols):max(cols)
#' will be removed.
#' @keywords internal
#' @export
remove_comment <- function(
    wb,
    sheet,
    col        = NULL,
    row        = NULL,
    gridExpand = TRUE,
    dims       = NULL
) {
  .Defunct("wb_remove_comment()", package = "openxlsx2", msg = "remove_comment() is defunct")
}

#' Convert to Excel data
#'
#' Use [convert_to_excel_date()].
#' @usage NULL
#' @inheritParams convert_to_excel_date
#' @keywords internal
#' @export
convertToExcelDate <- function(df, date1904 = FALSE) {
  .Defunct("convert_to_excel_date()", package = "openxlsx2", msg = "convertToExcelDate() is defunct")
}

#' Write an object to a worksheet
#'
#' Use [wb_add_data()] or [write_xlsx()] in new code.
#'
#' @inheritParams wb_add_data
#' @return invisible(0)
#' @export
#' @keywords internal
write_data <-  function(
    wb,
    sheet,
    x,
    dims              = wb_dims(start_row, start_col),
    start_col         = 1,
    start_row         = 1,
    array             = FALSE,
    col_names         = TRUE,
    row_names         = FALSE,
    with_filter       = FALSE,
    sep               = ", ",
    name              = NULL,
    apply_cell_style  = TRUE,
    remove_cell_style = FALSE,
    na.strings        = na_strings(),
    inline_strings    = TRUE,
    enforce           = FALSE,
    ...
) {
  .Defunct("wb_add_data()", package = "openxlsx2", msg = "write_data() is defunct")
}

#' Write to a worksheet as an Excel table
#'
#' Write to a worksheet and format as an Excel table. Use [wb_add_data_table()] in new code.
#' This function is defunct.
#' @inheritParams wb_add_data_table
#' @inherit wb_add_data_table details
#' @export
#' @keywords internal
write_datatable <- function(
    wb,
    sheet,
    x,
    dims              = wb_dims(start_row, start_col),
    start_col         = 1,
    start_row         = 1,
    col_names         = TRUE,
    row_names         = FALSE,
    table_style       = "TableStyleLight9",
    table_name        = NULL,
    with_filter       = TRUE,
    sep               = ", ",
    first_column      = FALSE,
    last_column       = FALSE,
    banded_rows       = TRUE,
    banded_cols       = FALSE,
    apply_cell_style  = TRUE,
    remove_cell_style = FALSE,
    na.strings        = na_strings(),
    inline_strings    = TRUE,
    total_row         = FALSE,
    ...
) {
  .Defunct("wb_add_data_table()", package = "openxlsx2", msg = "write_datatable() is defunct")
}

#' Write a character vector as an Excel Formula
#'
#' Write a a character vector containing Excel formula to a worksheet.
#' Use [wb_add_formula()] or `add_formula()` in new code. This function is
#' defunct.
#'
#' @inheritParams wb_add_formula
#' @export
#' @keywords internal
write_formula <- function(
    wb,
    sheet,
    x,
    dims              = wb_dims(start_row, start_col),
    start_col         = 1,
    start_row         = 1,
    array             = FALSE,
    cm                = FALSE,
    apply_cell_style  = TRUE,
    remove_cell_style = FALSE,
    enforce           = FALSE,
    ...
) {
  .Defunct("wb_add_formula()", package = "openxlsx2", msg = "write_formula() is defunct")
}
