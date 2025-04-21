#' Internal comment functions
#'
#' Users are advised to use [wb_add_comment()] and [wb_remove_comment()].
#' `write_comment()` and `remove_comment()` are now deprecated. openxlsx2 will stop
#' exporting it at some point in the future. Use the replacement functions.
#' @name comment_internal
NULL

#' Create a comment
#'
#' Use [wb_comment()] in new code. See [openxlsx2-deprecated]
#'
#' @inheritParams wb_comment
#' @param author A string, by default, will use "user"
#' @param visible Default: `TRUE`. Is the comment visible by default?
#' @keywords internal
#' @returns a `wbComment` object
#' @export
create_comment <- function(text,
                           author = Sys.info()[["user"]],
                           style = NULL,
                           visible = TRUE,
                           width = 2,
                           height = 4) {
  .Deprecated("wb_comment()", old = "create_comment()")
  wb_comment(text = text, author = author, style = style, visible = visible, width = width[1], height = height[1])
}

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
  .Deprecated("wb_add_comment()", package = "openxlsx2", old = "write_comment()")
  do_write_comment(
    wb,
    sheet,
    col,
    row,
    comment,
    dims,
    color,
    file
  )
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
  .Deprecated("wb_remove_comment()", package = "openxlsx2", old = "remove_comment()")
  do_remove_comment(
    wb,
    sheet,
    col,
    row,
    gridExpand,
    dims
  )
}

#' Convert to Excel data
#'
#' Use [convert_to_excel_date()].
#' @usage NULL
#' @inheritParams convert_to_excel_date
#' @keywords internal
#' @export
convertToExcelDate <- function(df, date1904 = FALSE) {
  stop("convertToExcelDate() is defunct and will be removed in new version. Use convert_to_excel_date().", call. = FALSE)
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
  .Deprecated("wb_add_data()", package = "openxlsx2", old = "write_data()")
  do_write_data(
    wb = wb,
    sheet = sheet,
    x = x,
    dims = dims,
    start_col         = start_col,
    start_row         = start_col,
    array             = array,
    col_names         = col_names,
    row_names         = row_names,
    with_filter       = with_filter,
    sep               = sep,
    name              = name,
    apply_cell_style  = apply_cell_style,
    remove_cell_style = remove_cell_style,
    na.strings        = na.strings,
    inline_strings    = inline_strings,
    enforce           = enforce,
    ...
  )
}

#' Write to a worksheet as an Excel table
#'
#' Write to a worksheet and format as an Excel table. Use [wb_add_data_table()] in new code.
#' This function is deprecated and may not be exported in the future.
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
  .Deprecated("wb_add_data_table()", package = "openxlsx2", old = "write_datatable()")
  do_write_datatable(
    wb = wb,
    sheet = sheet,
    x = x,
    dims = dims,
    start_col         = start_col,
    start_row         = start_row,
    col_names         = col_names,
    row_names         = row_names,
    table_style       = table_style,
    table_name        = table_name,
    with_filter       = with_filter,
    sep               = sep,
    first_column      = first_column,
    last_column       = last_column,
    banded_rows       = banded_rows,
    banded_cols       = banded_cols,
    apply_cell_style  = apply_cell_style,
    remove_cell_style = remove_cell_style,
    na.strings        = na.strings,
    inline_strings    = inline_strings,
    total_row         = total_row,
    ...

  )
}

#' Write a character vector as an Excel Formula
#'
#' Write a a character vector containing Excel formula to a worksheet.
#' Use [wb_add_formula()] or `add_formula()` in new code. This function is
#' deprecated and may be defunct.
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
  .Deprecated("wb_add_formula()", package = "openxlsx2", old = "write_formula()")

  do_write_formula(
    wb = wb,
    sheet = sheet,
    x = x,
    dims = dims,
    start_col = start_col,
    start_row = start_row,
    array = array,
    cm = cm,
    apply_cell_style = apply_cell_style,
    remove_cell_style = remove_cell_style,
    enforce = enforce,
    ...
  )
}
