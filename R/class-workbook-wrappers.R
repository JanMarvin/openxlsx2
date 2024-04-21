
#' Create a new Workbook object
#'
#' Initialize a [wbWorkbook] object. You can set workbook properties as well.
#'
#' `theme` can be one of
#' "Atlas", "Badge", "Berlin", "Celestial", "Crop", "Depth", "Droplet",
#' "Facet", "Feathered", "Gallery", "Headlines", "Integral", "Ion",
#' "Ion Boardroom", "LibreOffice", "Madison", "Main Event", "Mesh",
#' "Office 2007 - 2010 Theme", "Office 2013 - 2022 Theme", "Office Theme",
#' "Old Office Theme", "Organic", "Parallax", "Parcel", "Retrospect",
#' "Savon", "Slice", "Vapor Trail", "View", "Wisp", "Wood Type"
#'
#' @param creator Creator of the workbook (your name). Defaults to login username or `options("openxlsx2.creator")` if set.
#' @param title,subject,category,keywords,comments,manager,company Workbook property, a string.
#' @param datetime_created The time of the workbook is created
#' @param theme Optional theme identified by string or number.
#'   See **Details** for options.
#' @param ... additional arguments
#' @return A `wbWorkbook` object
#'
#' @export
#' @family workbook wrappers
#'
#' @examples
#' ## Create a new workbook
#' wb <- wb_workbook()
#'
#' ## Set Workbook properties
#' wb <- wb_workbook(
#'   creator  = "Me",
#'   title    = "Expense Report",
#'   subject  = "Expense Report - 2022 Q1",
#'   category = "sales"
#' )
wb_workbook <- function(
  creator          = NULL,
  title            = NULL,
  subject          = NULL,
  category         = NULL,
  datetime_created = Sys.time(),
  theme            = NULL,
  keywords         = NULL,
  comments         = NULL,
  manager          = NULL,
  company          = NULL,
  ...
) {
  wbWorkbook$new(
    creator          = creator,
    title            = title,
    subject          = subject,
    category         = category,
    datetime_created = datetime_created,
    theme            = theme,
    keywords         = keywords,
    comments         = comments,
    manager          = manager,
    company          = company,
    ...              = ...
  )
}


#' Save a workbook to file
#'
#' @param wb A `wbWorkbook` object to write to file
#' @param file A path to save the workbook to
#' @param overwrite If `FALSE`, will not overwrite when `file` already exists.
#' @param path Deprecated argument. Please use `file` in new code.
#'
#' @export
#' @family workbook wrappers
#'
#' @returns the `wbWorkbook` object, invisibly
#'
#' @examples
#' ## Create a new workbook and add a worksheet
#' wb <- wb_workbook("Creator of workbook")
#' wb$add_worksheet(sheet = "My first worksheet")
#'
#' ## Save workbook to working directory
#' \donttest{
#' wb_save(wb, file = temp_xlsx(), overwrite = TRUE)
#' }
wb_save <- function(wb, file = NULL, overwrite = TRUE, path = NULL) {
  assert_workbook(wb)
  wb$clone()$save(file = file, overwrite = overwrite, path = path)
}

# add data ----------------------------------------------------------------


#' Add data to a worksheet
#'
#' Add data to worksheet with optional styling.
#'
#' @param wb A Workbook object containing a worksheet.
#' @param sheet The worksheet to write to. Can be the worksheet index or name.
#' @param x Object to be written. For classes supported look at the examples.
#' @param dims Spreadsheet cell range that will determine `start_col` and `start_row`: "A1", "A1:B2", "A:B"
#' @param start_col A vector specifying the starting column to write `x` to.
#' @param start_row A vector specifying the starting row to write `x` to.
#' @param array A bool if the function written is of type array
#' @param col_names If `TRUE`, column names of `x` are written.
#' @param row_names If `TRUE`, the row names of `x` are written.
#' @param with_filter If `TRUE`, add filters to the column name row.
#'   NOTE: can only have one filter per worksheet.
#' @param name The name of a named region if specified.
#' @param sep Only applies to list columns. The separator used to collapse list
#'   columns to a character vector e.g. `sapply(x$list_column, paste, collapse = sep)`.
#' @param apply_cell_style Should we write cell styles to the workbook
#' @param remove_cell_style keep the cell style?
#' @param na.strings Value used for replacing `NA` values from `x`. Default
#'   looks if `options(openxlsx2.na.strings)` is set. Otherwise [na_strings()]
#'   uses the special `#N/A` value within the workbook.
#' @param inline_strings write characters as inline strings
#' @param enforce enforce that selected dims is filled. For this to work, `dims` must match `x`
#' @param ... additional arguments
#' @export
#' @details Formulae written using [wb_add_formula()] to a Workbook object will
#' not get picked up by `read_xlsx()`. This is because only the formula is written
#' and left to Excel to evaluate the formula when the file is opened in Excel.
#' The string `"_openxlsx_NA"` is reserved for `openxlsx2`.
#' If the data frame contains this string, the output will be broken.
#'
#' Supported classes are data frames, matrices and vectors of various types and
#' everything that can be converted into a data frame with `as.data.frame()`.
#' Everything else that the user wants to write should either be converted into
#' a vector or data frame or written in vector or data frame segments. This
#' includes base classes such as `table`, which were coerced internally in the
#' predecessor of this package.
#'
#' Even vectors and data frames can consist of different classes. Many base
#' classes are covered, though not all and far from all third-party classes.
#' When data of an unknown class is written, it is handled with `as.character()`.
#' It is not possible to write character nodes beginning with `<r>` or `<r/>`. Both
#' are reserved for internal functions. If you need these. You have to wrap
#' the input string in `fmt_txt()`.
#'
#' The columns of `x` with class Date/POSIXt, currency, accounting, hyperlink,
#' percentage are automatically styled as dates, currency, accounting,
#' hyperlinks, percentages respectively.
#'
#' Functions [wb_add_data()] and [wb_add_data_table()] behave quite similar. The
#' distinction is that the latter creates a table in the worksheet that can be
#' used for different kind of formulas and can be sorted independently, though
#' is less flexible than basic cell regions.
#' @family workbook wrappers
#' @family worksheet content functions
#' @return A `wbWorkbook`, invisibly.
#' @examples
#' ## See formatting vignette for further examples.
#'
#' ## Options for default styling (These are the defaults)
#' options("openxlsx2.dateFormat" = "mm/dd/yyyy")
#' options("openxlsx2.datetimeFormat" = "yyyy-mm-dd hh:mm:ss")
#' options("openxlsx2.numFmt" = NULL)
#'
#' #############################################################################
#' ## Create Workbook object and add worksheets
#' wb <- wb_workbook()
#'
#' ## Add worksheets
#' wb$add_worksheet("Cars")
#' wb$add_worksheet("Formula")
#'
#' x <- mtcars[1:6, ]
#' wb$add_data("Cars", x, start_col = 2, start_row = 3, row_names = TRUE)
#'
#' #############################################################################
#' ## Hyperlinks
#' ## - vectors/columns with class 'hyperlink' are written as hyperlinks'
#'
#' v <- rep("https://CRAN.R-project.org/", 4)
#' names(v) <- paste0("Hyperlink", 1:4) # Optional: names will be used as display text
#' class(v) <- "hyperlink"
#' wb$add_data("Cars", x = v, dims = "B32")
#'
#' #############################################################################
#' ## Formulas
#' ## - vectors/columns with class 'formula' are written as formulas'
#'
#' df <- data.frame(
#'   x = 1:3, y = 1:3,
#'   z = paste(paste0("A", 1:3 + 1L), paste0("B", 1:3 + 1L), sep = "+"),
#'   stringsAsFactors = FALSE
#' )
#'
#' class(df$z) <- c(class(df$z), "formula")
#'
#' wb$add_data(sheet = "Formula", x = df)
#'
#' #############################################################################
#' # update cell range and add mtcars
#' xlsxFile <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
#' wb2 <- wb_load(xlsxFile)
#'
#' # read dataset with inlinestr
#' wb_to_df(wb2)
#' wb2 <- wb2 %>% wb_add_data(sheet = 1, mtcars, dims = wb_dims(4, 4))
#' wb_to_df(wb2)
wb_add_data <- function(
    wb,
    sheet             = current_sheet(),
    x,
    dims              = wb_dims(start_row, start_col),
    start_col         = 1,
    start_row         = 1,
    array             = FALSE,
    col_names         = TRUE,
    row_names         = FALSE,
    with_filter       = FALSE,
    name              = NULL,
    sep               = ", ",
    apply_cell_style  = TRUE,
    remove_cell_style = FALSE,
    na.strings        = na_strings(),
    inline_strings    = TRUE,
    enforce           = FALSE,
    ...
) {
  assert_workbook(wb)
  wb$clone(deep = TRUE)$add_data(
    sheet             = sheet,
    x                 = x,
    dims              = dims,
    start_col         = start_col,
    start_row         = start_row,
    array             = array,
    col_names         = col_names,
    row_names         = row_names,
    with_filter       = with_filter,
    name              = name,
    sep               = sep,
    apply_cell_style  = apply_cell_style,
    remove_cell_style = remove_cell_style,
    na.strings        = na.strings,
    inline_strings    = inline_strings,
    enforce           = enforce,
    ...               = ...
  )
}

#' Add a data table to a worksheet
#'
#' Add data to a worksheet and format as an Excel table.
#'
#' @inherit wb_add_data details
#' @inheritParams wb_add_data
#' @param x A data frame
#' @param table_style Any table style name or "none" (see `vignette("openxlsx2_style_manual")`)
#' @param table_name Name of table in workbook. The table name must be unique.
#' @param with_filter If `TRUE`, columns with have filters in the first row.
#' @param sep Only applies to list columns. The separator used to collapse list
#'   columns to a character vector e.g.
#'   `sapply(x$list_column, paste, collapse = sep)`.
#' \cr\cr
#' \cr**The below options correspond to Excel table options:**
#' \cr
#' \if{html}{\figure{tableoptions.png}{options: width="40\%" alt="Figure: table_options.png"}}
#' \if{latex}{\figure{tableoptions.pdf}{options: width=7cm}}
#' @param first_column logical. If `TRUE`, the first column is bold.
#' @param last_column logical. If `TRUE`, the last column is bold.
#' @param banded_rows logical. If `TRUE`, rows are color banded.
#' @param banded_cols logical. If `TRUE`, the columns are color banded.
#' @param total_row logical. With the default `FALSE` no total row is added.
#' @param ... additional arguments
#'
#' @details # Modify total row argument
#' It is possible to further tweak the total row. In addition to the default
#' `FALSE` possible values are `TRUE` (the xlsx file will create column sums
#' each variable).
#'
#' In addition it is possible to tweak this further using a character string
#' with one of the following functions for each variable: `"average"`,
#' `"count"`, `"countNums"`, `"max"`, `"min"`, `"stdDev"`, `"sum"`, `"var"`.
#' It is possible to leave the cell empty `"none"` or to create a text input
#' using a named character with name `text` like: `c(text = "Total")`.
#' It's also possible to pass other spreadsheet software functions if they
#' return a single value and hence `"SUM"` would work too.
#'
#' @family worksheet content functions
#' @family workbook wrappers
#' @examples
#' wb <- wb_workbook()$add_worksheet()$
#'   add_data_table(
#'     x = as.data.frame(USPersonalExpenditure),
#'     row_names = TRUE,
#'     total_row = c(text = "Total", "none", "sum", "sum", "sum", "SUM")
#'   )
#' @export
wb_add_data_table <- function(
    wb,
    sheet             = current_sheet(),
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
    total_row        = FALSE,
    ...
) {
  assert_workbook(wb)
  wb$clone()$add_data_table(
    sheet             = sheet,
    x                 = x,
    dims              = dims,
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
    ...               = ...
  )
}

#' Add a pivot table to a worksheet
#'
#' The data must be specified using [wb_data()] to ensure the function works.
#' The sheet will be empty unless it is opened in spreadsheet software. Find
#' more details in the [section about pivot tables](https://janmarvin.github.io/ox2-book/chapters/openxlsx2_pivot_tables.html)
#' in the openxlsx2 book.
#'
#' @details
#' The pivot table is not actually written to the worksheet, therefore the cell
#' region has to remain empty. What is written to the workbook is something
#' like a recipe how the spreadsheet software has to construct the pivot table
#' when opening the file.
#'
#' It is possible to add slicers to the pivot table. For this the pivot
#' table has to be named and the variable used as slicer, must be part
#' of the selected pivot table names (`cols`, `rows`, `filter`, or
#' `slicer`). If these criteria are matched, a slicer can be added
#' using [wb_add_slicer()].
#'
#' Be aware that you should always test on a copy if a `param` argument works
#' with a pivot table. Not only to check if the desired effect appears, but
#' first and foremost if the file loads. Wildly mixing params might brick the
#' output file and cause spreadsheet software to crash.
#'
#' `fun` can be any of `AVERAGE`, `COUNT`, `COUNTA`, `MAX`, `MIN`,
#' `PRODUCT`, `STDEV`, `STDEVP`, `SUM`, `VAR`, `VARP`.
#'
#' `show_data_as` can be any of `normal`, `difference`, `percent`, `percentDiff`,
#' `runTotal`, `percentOfRow`, `percentOfCol`, `percentOfTotal`, `index`.
#'
#' It is possible to calculate data fields if the formula is assigned as a
#' variable name for the field to calculate. This would look like this:
#' `data = c("am", "disp/cyl" = "New")`
#'
#' Possible `params` arguments are listed below. Pivot tables accepts more
#' parameters, but they were either not tested or misbehaved (probably because
#' we misunderstood how the parameter should be used).
#'
#' Boolean arguments:
#' * apply_alignment_formats
#' * apply_number_formats
#' * apply_border_formats
#' * apply_font_formats
#' * apply_pattern_formats
#' * apply_width_height_formats
#' * no_style
#' * compact
#' * outline
#' * compact_data
#' * row_grand_totals
#' * col_grand_totals
#'
#' Table styles accepting character strings:
#' * auto_format_id: style id as character in the range of 4096 to 4117
#' * table_style: a predefined (pivot) table style `"TableStyleMedium23"`
#' * show_data_as: accepts character strings as listed above
#'
#' Miscellaneous:
#' * numfmt: accepts vectors of the form `c(formatCode = "0.0%")`
#' * choose: select variables in the form of a named logical vector like
#'  `c(agegp = 'x > "25-34"')` for the `esoph` dataset.
#' * sort_item: named list of index or character vectors
#'
#' @param wb A Workbook object containing a #' worksheet.
#' @param x A `data.frame` that inherits the [`wb_data`][wb_data()] class.
#' @param sheet A worksheet containing a #'
#' @param dims The worksheet cell where the pivot table is placed
#' @param filter The column name(s) of `x` used for filter.
#' @param rows The column name(s) of `x` used as rows
#' @param cols The column names(s) of `x` used as cols
#' @param data The column name(s) of `x` used as data
#' @param fun A vector of functions to be used with `data`. See **Details** for the list of available options.
#' @param params A list of parameters to modify pivot table creation. See **Details** for available options.
#' @param pivot_table An optional name for the pivot table
#' @param slicer Any additional column name(s) of `x` used as slicer
#' @seealso [wb_data()]
#' @examples
#' wb <- wb_workbook() %>% wb_add_worksheet() %>% wb_add_data(x = mtcars)
#'
#' df <- wb_data(wb, sheet = 1)
#'
#' wb <- wb %>%
#'   # default pivot table
#'   wb_add_pivot_table(df, dims = "A3",
#'     filter = "am", rows = "cyl", cols = "gear", data = "disp"
#'   ) %>%
#'   # with parameters
#'   wb_add_pivot_table(df,
#'     filter = "am", rows = "cyl", cols = "gear", data = "disp",
#'     params = list(no_style = TRUE, numfmt = c(formatCode = "##0.0"))
#'   )
#' @family workbook wrappers
#' @family worksheet content functions
#' @export
wb_add_pivot_table <- function(
    wb,
    x,
    sheet = next_sheet(),
    dims = "A3",
    filter,
    rows,
    cols,
    data,
    fun,
    params,
    pivot_table,
    slicer
) {
  assert_workbook(wb)
  if (missing(filter))      filter <- substitute()
  if (missing(rows))        rows   <- substitute()
  if (missing(cols))        cols   <- substitute()
  if (missing(data))        data   <- substitute()
  if (missing(fun))         fun    <- substitute()
  if (missing(params))      params <- substitute()
  if (missing(pivot_table)) pivot_table <- substitute()
  if (missing(slicer))      slicer <- substitute()

  wb$clone()$add_pivot_table(
    x           = x,
    sheet       = sheet,
    dims        = dims,
    filter      = filter,
    rows        = rows,
    cols        = cols,
    data        = data,
    fun         = fun,
    params      = params,
    pivot_table = pivot_table,
    slicer      = slicer
  )

}

#' Add a slicer to a pivot table
#'
#' Add a slicer to a previously created pivot table. This function is still experimental and might be changed/improved in upcoming releases.
#'
#' @details
#' This assumes that the slicer variable initialization has happened before. Unfortunately, it is unlikely that we can guarantee this for loaded workbooks, and we *strictly* discourage users from attempting this. If the variable has not been initialized properly, this may cause the spreadsheet software to crash.
#'
#' Possible `params` arguments are listed below.
#' * edit_as: "twoCell" to place the slicer into the cells
#' * style: "SlicerStyleLight2"
#' * column_count: integer used as column count
#' * caption: string used for a caption
#' * sort_order: "descending" / "ascending"
#' * choose: select variables in the form of a named logical vector like
#'  `c(agegp = 'x > "25-34"')` for the `esoph` dataset.
#'
#' @param wb A Workbook object containing a #' worksheet.
#' @param x A `data.frame` that inherits the [`wb_data`][wb_data()] class.
#' @param sheet A worksheet containing a #'
#' @param dims The worksheet cell where the pivot table is placed
#' @param pivot_table The name of a pivot table on the selected sheet
#' @param slicer A variable used as slicer for the pivot table
#' @param params A list of parameters to modify pivot table creation. See **Details** for available options.
#' @family workbook wrappers
#' @family worksheet content functions
#' @examples
#' wb <- wb_workbook() %>%
#'   wb_add_worksheet() %>% wb_add_data(x = mtcars)
#'
#' df <- wb_data(wb, sheet = 1)
#'
#' wb <- wb %>%
#'   wb_add_pivot_table(
#'     df, dims = "A3", slicer = "vs", rows = "cyl", cols = "gear", data = "disp",
#'     pivot_table = "mtcars"
#'   ) %>%
#'   wb_add_slicer(x = df, slicer = "vs", pivot_table = "mtcars")
#' @export
wb_add_slicer <- function(
    wb,
    x,
    dims        = "A1",
    sheet       = current_sheet(),
    pivot_table,
    slicer,
    params
) {
  assert_workbook(wb)
  if (missing(params)) params <- substitute()

  wb$clone()$add_slicer(
    x           = x,
    sheet       = sheet,
    dims        = dims,
    pivot_table = pivot_table,
    slicer      = slicer,
    params      = params
  )

}

#' Add a formula to a cell range in a worksheet
#'
#' This function can be used to add a formula to a worksheet.
#' In `wb_add_formula()`, you can provide the formula as a character vector.
#'
#' @details
#' Currently, the local translations of formulas are not supported.
#' Only the English functions work.
#'
#' The examples below show a small list of possible formulas:
#'
#' * SUM(B2:B4)
#' * AVERAGE(B2:B4)
#' * MIN(B2:B4)
#' * MAX(B2:B4)
#' * ...
#'
#' It is possible to pass vectors to `x`. If `x` is an array formula, it will
#' take `dims` as a reference. For some formulas, the result will span multiple
#' cells (see the `MMULT()` example below). For this type of formula, the
#' output range must be known a priori and passed to `dims`, otherwise only the
#' value of the first cell will be returned. This type of formula, whose result
#' extends over several cells, is only possible with single strings. If a vector
#' is passed, it is only possible to return individual cells.
#'
#' @param wb A Workbook object containing a worksheet.
#' @param sheet The worksheet to write to. (either as index or name)
#' @param x A formula as character vector.
#' @param dims Spreadsheet dimensions that will determine where `x` spans: "A1", "A1:B2", "A:B"
#' @param start_col A vector specifying the starting column to write to.
#' @param start_row A vector specifying the starting row to write to.
#' @param array A bool if the function written is of type array
#' @param cm A special kind of array function that hides the curly braces in the cell.
#'   Add this, if you see "@" inserted into your formulas.
#' @param apply_cell_style Should we write cell styles to the workbook?
#' @param remove_cell_style Should we keep the cell style?
#' @param enforce enforce dims
#' @param ... additional arguments
#' @return The workbook, invisibly.
#' @family workbook wrappers
#' @family worksheet content functions
#' @export
#' @examples
#' wb <- wb_workbook()$add_worksheet()
#' wb$add_data(dims = wb_dims(rows = 1, cols = 1:3), x = c(4, 5, 8))
#'
#' # calculate the sum of elements.
#' wb$add_formula(dims = "D1", x = "SUM(A1:C1)")
#'
#' # array formula with result spanning over multiple cells
#' mm <- matrix(1:4, 2, 2)
#'
#' wb$add_worksheet()$
#'  add_data(x = mm, dims = "A1:B2", col_names = FALSE)$
#'  add_data(x = mm, dims = "A4:B5", col_names = FALSE)$
#'  add_formula(x = "MMULT(A1:B2, A4:B5)", dims = "A7:B8", array = TRUE)
#'
wb_add_formula <- function(
    wb,
    sheet             = current_sheet(),
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
  assert_workbook(wb)
  wb$clone()$add_formula(
    sheet             = sheet,
    x                 = x,
    dims              = dims,
    start_col         = start_col,
    start_row         = start_row,
    array             = array,
    cm                = cm,
    apply_cell_style  = apply_cell_style,
    remove_cell_style = remove_cell_style,
    enforce           = enforce,
    ...               = ...
  )
}

#' Update a data table position in a worksheet
#'
#' Update the position of a data table, possibly written using [wb_add_data_table()]
#' @param wb A workbook
#' @param sheet A worksheet
#' @param dims Cell range used for new data table.
#' @param tabname A table name
#'
#' @details
#' Be aware that this function does not alter any filter.
#' Excluding or adding rows does not make rows appear nor will it hide them.
#' @examples
#' wb <- wb_workbook()$add_worksheet()$add_data_table(x = mtcars)
#' wb$update_table(tabname = "Table1", dims = "A1:J4")
#' @export
wb_update_table <- function(wb, sheet = current_sheet(), dims = "A1", tabname) {
  assert_workbook(wb)
  wb$clone()$update_table(sheet = sheet, dims = dims, tabname = tabname)
}

#' Copy cells around within a worksheet
#'
#' @param wb A workbook
#' @param sheet a worksheet
#' @param dims A cell where to place the copy
#' @param data A [`wb_data`][wb_data()] object containing cells to copy
#' @param as_value Should a copy of the value be written?
#' @param as_ref Should references to the cell be written?
#' @param transpose Should the data be written transposed?
#' @param ... additional arguments passed to add_data() if used with `as_value`
#' @examples
#' wb <- wb_workbook()$
#' add_worksheet()$
#'   add_data(x = mtcars)$
#'   add_fill(dims = "A1:F1", color = wb_color("yellow"))
#'
#' dat <- wb_data(wb, dims = "A1:D4", col_names = FALSE)
#' # 1:1 copy to M2
#' wb$
#'   clone_worksheet(old = 1, new = "Clone1")$
#'   copy_cells(data = dat, dims = "M2")
#' @family workbook wrappers
#' @seealso [wb_data()]
#' @export
#' @returns the `wbWorkbook` invisibly
wb_copy_cells <- function(
    wb,
    sheet     = current_sheet(),
    dims      = "A1",
    data,
    as_value  = FALSE,
    as_ref    = FALSE,
    transpose = FALSE,
    ...
) {
  assert_workbook(wb)
  wb$
    clone(deep = TRUE)$
    copy_cells(
      sheet     = sheet,
      dims      = dims,
      data      = data,
      as_value  = as_value,
      as_ref    = as_ref,
      transpose = transpose,
      ...       = ...
    )
}

# merge cells ------------------------------------------------------------------

#'
#' Merge cells within a worksheet
#'
#' Worksheet cell merging
#'
#' @details
#' If using the deprecated arguments `rows` and `cols` with a merged region must be rectangular,
#' only min and max of `cols` and `rows` are used.
#'
#' @param wb A Workbook object
#' @param sheet A name or index of a worksheet
#' @param dims worksheet cells
#' @param solve logical if intersecting merges should be solved
#' @param ... additional arguments
#'
#' @examples
#' # Create a new workbook
#' wb <- wb_workbook()$add_worksheet()
#'
#' # Merge cells: Row 2 column C to F (3:6)
#' wb <- wb_merge_cells(wb, dims = "C3:F6")
#'
#' # Merge cells:Rows 10 to 20 columns A to J (1:10)
#' wb <- wb_merge_cells(wb, dims = wb_dims(rows = 10:20, cols = 1:10))
#'
#' wb$add_worksheet()
#'
#' ## Intersecting merges
#' wb <- wb_merge_cells(wb, dims = wb_dims(cols = 1:10, rows = 1))
#' wb <- wb_merge_cells(wb, dims = wb_dims(cols = 5:10, rows = 2))
#' wb <- wb_merge_cells(wb, dims = wb_dims(cols = 1:10, rows = 12))
#' try(wb_merge_cells(wb, dims = "A1:A10"))
#'
#' ## remove merged cells
#' # removes any intersecting merges
#' wb <- wb_unmerge_cells(wb, dims = wb_dims(cols = 1, rows = 1))
#' wb <- wb_merge_cells(wb, dims = "A1:A10")
#'
#' # or let us decide how to solve this
#' wb <- wb_merge_cells(wb, dims = "A1:A10", solve = TRUE)
#'
#' @family workbook wrappers
#' @family worksheet content functions
#' @export
wb_merge_cells <- function(wb, sheet = current_sheet(), dims = NULL, solve = FALSE, ...) {
  assert_workbook(wb)
  wb$clone(deep = TRUE)$merge_cells(sheet = sheet, dims = dims, solve = solve, ... = ...)
}

#' @export
#' @rdname wb_merge_cells
wb_unmerge_cells <- function(wb, sheet = current_sheet(), dims = NULL, ...) {
  assert_workbook(wb)
  wb$clone()$unmerge_cells(sheet = sheet, dims = dims, ... = ...)
}


# sheets ------------------------------------------------------------------

#' Add a chartsheet to a workbook
#'
#' A chartsheet is a special type of sheet that handles charts output. You must
#' add a chart to the sheet. Otherwise, this will break the workbook.
#'
#' @param wb A Workbook object to attach the new chartsheet
#' @param sheet A name for the new chartsheet
#' @inheritParams wb_add_worksheet
#' @family workbook wrappers
#' @seealso [wb_add_mschart()]
#' @export
wb_add_chartsheet <- function(
  wb,
  sheet     = next_sheet(),
  tab_color = NULL,
  zoom      = 100,
  visible   = c("true", "false", "hidden", "visible", "veryhidden"),
  ...
) {
  assert_workbook(wb)
  wb$clone()$add_chartsheet(
    sheet       = sheet,
    tab_color   = tab_color,
    zoom        = zoom,
    visible     = visible,
    ...         = ...
  )
}

#' Add a worksheet to a workbook
#'
#' Add a worksheet to a [wbWorkbook] is the first step to build a workbook.
#' With the function, you can also set the sheet view with `zoom`, set headers
#' and footers as well as other features. See the function arguments.
#'
#' @details
#' Headers and footers can contain special tags
#' * **&\[Page\]** Page number
#' * **&\[Pages\]** Number of pages
#' * **&\[Date\]** Current date
#' * **&\[Time\]** Current time
#' * **&\[Path\]** File path
#' * **&\[File\]** File name
#' * **&\[Tab\]** Worksheet name
#'
#' @param wb A `wbWorkbook` object to attach the new worksheet
#' @param sheet A name for the new worksheet
#' @param grid_lines A logical. If `FALSE`, the worksheet grid lines will be
#'   hidden.
#' @param row_col_headers A logical. If `FALSE`, the worksheet colname and rowname will be
#'   hidden.
#' @param tab_color Color of the sheet tab. A  [wb_color()],  a valid color (belonging to
#'   `grDevices::colors()`) or a valid hex color beginning with "#".
#' @param zoom The sheet zoom level, a numeric between 10 and 400 as a
#'   percentage. (A zoom value smaller than 10 will default to 10.)
#' @param header,odd_header,even_header,first_header,footer,odd_footer,even_footer,first_footer
#'   Character vector of length 3 corresponding to positions left, center,
#'   right.  `header` and `footer` are used to default additional arguments.
#'   Setting `even`, `odd`, or `first`, overrides `header`/`footer`. Use `NA` to
#'   skip a position.
#' @param visible If `FALSE`, sheet is hidden else visible.
#' @param has_drawing If `TRUE` prepare a drawing output (TODO does this work?)
#' @param paper_size An integer corresponding to a paper size. See [wb_page_setup()] for
#'   details.
#' @param orientation One of "portrait" or "landscape"
#' @param hdpi,vdpi Horizontal and vertical DPI. Can be set with `options("openxlsx2.dpi" = X)`,
#'   `options("openxlsx2.hdpi" = X)` or `options("openxlsx2.vdpi" = X)`
#' @param ... Additional arguments
#' @return The `wbWorkbook` object, invisibly.
#'
#' @export
#' @family workbook wrappers
#' @examples
#' ## Create a new workbook
#' wb <- wb_workbook()
#'
#' ## Add a worksheet
#' wb$add_worksheet("Sheet 1")
#' ## No grid lines
#' wb$add_worksheet("Sheet 2", grid_lines = FALSE)
#' ## A red tab color
#' wb$add_worksheet("Sheet 3", tab_color = wb_color("red"))
#' ## All options combined with a zoom of 40%
#' wb$add_worksheet("Sheet 4", grid_lines = FALSE, tab_color = wb_color(hex = "#4F81BD"), zoom = 40)
#'
# TODO maybe leave the special cases to wb_set_header_footer?
#' ## Headers and Footers
#' wb$add_worksheet("Sheet 5",
#'   header = c("ODD HEAD LEFT", "ODD HEAD CENTER", "ODD HEAD RIGHT"),
#'   footer = c("ODD FOOT RIGHT", "ODD FOOT CENTER", "ODD FOOT RIGHT"),
#'   even_header = c("EVEN HEAD LEFT", "EVEN HEAD CENTER", "EVEN HEAD RIGHT"),
#'   even_footer = c("EVEN FOOT RIGHT", "EVEN FOOT CENTER", "EVEN FOOT RIGHT"),
#'   first_header = c("TOP", "OF FIRST", "PAGE"),
#'   first_footer = c("BOTTOM", "OF FIRST", "PAGE")
#' )
#'
#' wb$add_worksheet("Sheet 6",
#'   header = c("&[Date]", "ALL HEAD CENTER 2", "&[Page] / &[Pages]"),
#'   footer = c("&[Path]&[File]", NA, "&[Tab]"),
#'   first_header = c(NA, "Center Header of First Page", NA),
#'   first_footer = c(NA, "Center Footer of First Page", NA)
#' )
#'
#' wb$add_worksheet("Sheet 7",
#'   header = c("ALL HEAD LEFT 2", "ALL HEAD CENTER 2", "ALL HEAD RIGHT 2"),
#'   footer = c("ALL FOOT RIGHT 2", "ALL FOOT CENTER 2", "ALL FOOT RIGHT 2")
#' )
#'
#' wb$add_worksheet("Sheet 8",
#'   first_header = c("FIRST ONLY L", NA, "FIRST ONLY R"),
#'   first_footer = c("FIRST ONLY L", NA, "FIRST ONLY R")
#' )
#'
#' ## Need data on worksheet to see all headers and footers
#' wb$add_data(sheet = 5, 1:400)
#' wb$add_data(sheet = 6, 1:400)
#' wb$add_data(sheet = 7, 1:400)
#' wb$add_data(sheet = 8, 1:400)
wb_add_worksheet <- function(
  wb,
  sheet           = next_sheet(),
  grid_lines      = TRUE,
  row_col_headers = TRUE,
  tab_color       = NULL,
  zoom            = 100,
  header          = NULL,
  footer          = NULL,
  odd_header      = header,
  odd_footer      = footer,
  even_header     = header,
  even_footer     = footer,
  first_header    = header,
  first_footer    = footer,
  visible         = c("true", "false", "hidden", "visible", "veryhidden"),
  has_drawing     = FALSE,
  paper_size      = getOption("openxlsx2.paperSize", default = 9),
  orientation     = getOption("openxlsx2.orientation", default = "portrait"),
  hdpi            = getOption("openxlsx2.hdpi", default = getOption("openxlsx2.dpi", default = 300)),
  vdpi            = getOption("openxlsx2.vdpi", default = getOption("openxlsx2.dpi", default = 300)),
  ...
) {

  assert_workbook(wb)
  wb$clone()$add_worksheet(
    sheet           = sheet,
    grid_lines      = grid_lines,
    row_col_headers = row_col_headers,
    tab_color       = tab_color,
    zoom            = zoom,
    odd_header      = headerFooterSub(odd_header),
    odd_footer      = headerFooterSub(odd_footer),
    even_header     = headerFooterSub(even_header),
    even_footer     = headerFooterSub(even_footer),
    first_header    = headerFooterSub(first_header),
    first_footer    = headerFooterSub(first_footer),
    visible         = visible,
    paper_size      = paper_size,
    orientation     = orientation,
    vdpi            = vdpi,
    hdpi            = hdpi,
    ...             = ...
  )
}


#' Create copies of a worksheet within a workbook
#'
#' @description
#' Create a copy of a worksheet in the same `wbWorkbook` object.
#'
#' Cloning is possible only to a limited extent. References to sheet names in
#' formulas, charts, pivot tables, etc. may not be updated. Some elements like
#' named ranges and slicers cannot be cloned yet.
#'
#' Cloning from another workbook is still an experimental feature and might not
#' work reliably. Cloning data, media, charts and tables should work. Slicers
#' and pivot tables as well as everything everything relying on dxfs styles
#' (e.g. custom table styles and conditional formatting) is currently not
#' implemented.
#' Formula references are not updated to reflect interactions between workbooks.
#'
#' @param wb A `wbWorkbook` object
#' @param old Name of existing worksheet to copy
#' @param new Name of the new worksheet to create
#' @param from (optional) Workbook to clone old from
#' @return The `wbWorkbook` object, invisibly.
#'
#' @export
#' @family workbook wrappers
#'
#' @examples
#' # Create a new workbook
#' wb <- wb_workbook()
#'
#' # Add worksheets
#' wb$add_worksheet("Sheet 1")
#' wb$clone_worksheet("Sheet 1", new = "Sheet 2")
#' # Take advantage of waiver functions
#' wb$clone_worksheet(old = "Sheet 1")
#'
#' ## cloning from another workbook
#'
#' # create a workbook
#' wb <- wb_workbook()$
#' add_worksheet("NOT_SUM")$
#'   add_data(x = head(iris))$
#'   add_fill(dims = "A1:B2", color = wb_color("yellow"))$
#'   add_border(dims = "B2:C3")
#'
#' # we will clone this styled chart into another workbook
#' fl <- system.file("extdata", "oxlsx2_sheet.xlsx", package = "openxlsx2")
#' wb_from <- wb_load(fl)
#'
#' # clone styles and shared strings
#' wb$clone_worksheet(old = "SUM", new = "SUM", from = wb_from)
#'
wb_clone_worksheet <- function(wb, old = current_sheet(), new = next_sheet(), from = NULL) {
  assert_workbook(wb)
  wb$clone()$clone_worksheet(old = old, new = new, from = from)
}

# worksheets --------------------------------------------------------------

#' Freeze pane of a worksheet
#'
#' Add a Freeze pane in a worksheet.
#'
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param first_active_row Top row of active region
#' @param first_active_col Furthest left column of active region
#' @param first_row If `TRUE`, freezes the first row (equivalent to `first_active_row = 2`)
#' @param first_col If `TRUE`, freezes the first column (equivalent to `first_active_col = 2`)
#' @param ... additional arguments
#'
#' @export
#' @family workbook wrappers
#' @family worksheet content functions
#' @examples
#' ## Create a new workbook
#' wb <- wb_workbook("Kenshin")
#'
#' ## Add some worksheets
#' wb$add_worksheet("Sheet 1")
#' wb$add_worksheet("Sheet 2")
#' wb$add_worksheet("Sheet 3")
#' wb$add_worksheet("Sheet 4")
#'
#' ## Freeze Panes
#' wb$freeze_pane("Sheet 1", first_active_row = 5, first_active_col = 3)
#' wb$freeze_pane("Sheet 2", first_col = TRUE) ## shortcut to first_active_col = 2
#' wb$freeze_pane(3, first_row = TRUE) ## shortcut to first_active_row = 2
#' wb$freeze_pane(4, first_active_row = 1, first_active_col = "D")
wb_freeze_pane <- function(
  wb,
  sheet            = current_sheet(),
  first_active_row = NULL,
  first_active_col = NULL,
  first_row        = FALSE,
  first_col        = FALSE,
  ...
) {
  assert_workbook(wb)
  wb$clone()$freeze_pane(
    sheet            = sheet,
    first_active_row = first_active_row,
    first_active_col = first_active_col,
    first_row        = first_row,
    first_col        = first_col
  )
}


# heights and columns -----------------------------------------------------


#' Modify row heights of a worksheet
#'
#' Set / remove custom worksheet row heights
#'
#' @param wb A [wbWorkbook] object
#' @param sheet A name or index of a worksheet. (A vector is accepted for `remove_row_heights()`)
#' @param rows Indices of rows to set / remove (if any) custom height.
#' @param heights Heights to set `rows` to specified in a spreadsheet column height units.
#' @param hidden Option to hide rows. A logical vector of length 1 or length of `rows`
#' @name row_heights-wb
#' @family workbook wrappers
#' @family worksheet content functions
#'
#' @examples
#' ## Create a new workbook
#' wb <- wb_workbook()
#'
#' ## Add a worksheet
#' wb$add_worksheet("Sheet 1")
#'
#' ## set row heights
#' wb <- wb_set_row_heights(
#'   wb, 1,
#'   rows = c(1, 4, 22, 2, 19),
#'   heights = c(24, 28, 32, 42, 33)
#' )
#'
#' ## overwrite row 1 height
#' wb <- wb_set_row_heights(wb, 1, rows = 1, heights = 40)
#' ## remove any custom row heights in row 1
#' wb$remove_row_heights(sheet = 1, rows = 1)
NULL
#' @rdname row_heights-wb
#' @export
wb_set_row_heights <- function(wb, sheet = current_sheet(), rows, heights = NULL, hidden = FALSE) {
  assert_workbook(wb)
  assert_class(heights, c("numeric", "integer"), or_null = TRUE, arg_nm = "heights")
  wb$clone()$set_row_heights(sheet = sheet, rows, heights, hidden)
}
#' @rdname row_heights-wb
#' @export
wb_remove_row_heights <- function(wb, sheet = current_sheet(), rows) {
  assert_workbook(wb)
  wb$clone()$remove_row_heights(sheet = sheet, rows = rows)
}

#' Modify column widths of a worksheet
#'
#' Remove / set worksheet column widths to specified width or "auto".
#'
#' @details
#' The global min and max column width for "auto" columns is set by (default values show):
#' * `options("openxlsx2.minWidth" = 3)`
#' * `options("openxlsx2.maxWidth" = 250)` Maximum width allowed in Excel
#'
#' NOTE: The calculation of column widths can be slow for large worksheets.
#'
#' NOTE: The `hidden` parameter may conflict with the one set in [wb_group_cols()];
#' changing one will update the other.
#'
#' NOTE: The default column width varies by spreadsheet software, operating system,
#' and DPI settings used. Setting `widths` to specific value also is no guarantee
#' that the output will have consistent column widths.
#'
#' For automatic text wrapping of columns use
#' [wb_add_cell_style(wrap_text = TRUE)][wb_add_cell_style()]
#'
#' @param wb A `wbWorkbook` object.
#' @param sheet A name or index of a worksheet, a vector in the case of `remove_`
#' @param cols Indices of cols to set/remove column widths.
#' @param widths Width to set `cols` to specified column width or `"auto"` for
#'   automatic sizing. `widths` is recycled to the length of `cols`. openxlsx2
#'   sets the default width is 8.43, as this is the standard in some spreadsheet
#'   software. See **Details** for general information on column widths.
#' @param hidden Logical vector recycled to the length of `cols`.
#'   If `TRUE`, the columns are hidden.
#'
#' @family workbook wrappers
#' @family worksheet content functions
#'
#' @examples
#' ## Create a new workbook
#' wb <- wb_workbook()
#'
#' ## Add a worksheet
#' wb$add_worksheet("Sheet 1")
#'
#' ## set col widths
#' wb$set_col_widths(cols = c(1, 4, 6, 7, 9), widths = c(16, 15, 12, 18, 33))
#'
#' ## auto columns
#' wb$add_worksheet("Sheet 2")
#' wb$add_data(sheet = 2, x = iris)
#' wb$set_col_widths(sheet = 2, cols = 1:5, widths = "auto")
#'
#' ## removing column widths
#' ## Create a new workbook
#' wb <- wb_load(file = system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2"))
#'
#' ## remove column widths in columns 1 to 20
#' wb_remove_col_widths(wb, 1, cols = 1:20)
#' @name col_widths-wb
NULL

#' @rdname col_widths-wb
#' @export
wb_set_col_widths <- function(wb, sheet = current_sheet(), cols, widths = 8.43, hidden = FALSE) {
  assert_workbook(wb)
  wb$clone()$set_col_widths(
    sheet  = sheet,
    cols   = cols,
    widths = widths,
    # TODO allow either 1 or length(cols)
    hidden = hidden
  )
}
#' @rdname col_widths-wb
#' @export
wb_remove_col_widths <- function(wb, sheet = current_sheet(), cols) {
  assert_workbook(wb)
  wb$clone()$remove_col_widths(sheet = sheet, cols = cols)
}


# images ------------------------------------------------------------------


#' Insert the current plot into a worksheet
#'
#' The current plot is saved to a temporary image file using
#' [grDevices::dev.copy()] This file is then written to the workbook using
#' [wb_add_image()].
#'
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param dims Worksheet dimension, single cell ("A1") or cell range ("A1:D4")
#' @param row_offset,col_offset Offset for column and row
#' @param width Width of figure. Defaults to `6` in.
#' @param height Height of figure . Defaults to `4` in.
#' @param file_type File type of image
#' @param units Units of width and height. Can be `"in"`, `"cm"` or `"px"`
#' @param dpi Image resolution
#' @param ... additional arguments
#' @seealso [wb_add_chart_xml()] [wb_add_drawing()] [wb_add_image()] [wb_add_mschart()]
#' @export
#' @examples
#' if (requireNamespace("ggplot2") && interactive()) {
#' ## Create a new workbook
#' wb <- wb_workbook()
#'
#' ## Add a worksheet
#' wb$add_worksheet("Sheet 1", grid_lines = FALSE)
#'
#' ## create plot objects
#' require(ggplot2)
#' p1 <- ggplot(mtcars, aes(x = mpg, fill = as.factor(gear))) +
#'   ggtitle("Distribution of Gas Mileage") +
#'   geom_density(alpha = 0.5)
#' p2 <- ggplot(Orange, aes(x = age, y = circumference, color = Tree)) +
#'   geom_point() + geom_line()
#'
#' ## Insert currently displayed plot to sheet 1, row 1, column 1
#' print(p1) # plot needs to be showing
#' wb$add_plot(1, width = 5, height = 3.5, file_type = "png", units = "in")
#'
#' ## Insert plot 2
#' print(p2)
#' wb$add_plot(1, dims = "J2", width = 16, height = 10, file_type = "png", units = "cm")
#'
#' }
wb_add_plot <- function(
    wb,
    sheet      = current_sheet(),
    dims       = "A1",
    width      = 6,
    height     = 4,
    row_offset = 0,
    col_offset = 0,
    file_type  = "png",
    units      = "in",
    dpi        = 300,
    ...
) {
  assert_workbook(wb)
  wb$clone()$add_plot(
    sheet      = sheet,
    dims       = dims,
    width      = width,
    height     = height,
    row_offset = row_offset,
    col_offset = col_offset,
    file_type  = file_type,
    units      = units,
    dpi        = dpi,
    ...        = ...
  )
}

#' Add drawings to a worksheet
#'
#' Add drawings to a worksheet. This requires the `rvg` package.
#' @param wb A `wbWorkbook`
#' @param sheet A sheet in the workbook
#' @param dims The dimension where the drawing is added.
#' @param xml the drawing xml as character or file
#' @param col_offset,row_offset offsets for column and row
#' @param ... additional arguments
#' @examples
#' if (requireNamespace("rvg") && interactive()) {
#'
#' ## rvg example
#' require(rvg)
#' tmp <- tempfile(fileext = ".xml")
#' dml_xlsx(file =  tmp)
#' plot(1,1)
#' dev.off()
#'
#' wb <- wb_workbook()$
#'   add_worksheet()$
#'   add_drawing(xml = tmp)$
#'   add_drawing(xml = tmp, dims = NULL)
#' }
#' @seealso [wb_add_chart_xml()] [wb_add_image()] [wb_add_mschart()] [wb_add_plot()]
#' @export
wb_add_drawing <- function(
  wb,
  sheet      = current_sheet(),
  dims       = "A1",
  xml,
  col_offset = 0,
  row_offset = 0,
  ...
) {
  assert_workbook(wb)
  wb$clone()$add_drawing(
    sheet      = sheet,
    dims       = dims,
    xml        = xml,
    col_offset = col_offset,
    row_offset = row_offset,
    ...        = ...
  )
}

#' Add mschart object to a worksheet
#'
#' @param wb a workbook
#' @param sheet the sheet on which the graph will appear
#' @param dims the dimensions where the sheet will appear
#' @param graph mschart object
#' @param col_offset,row_offset offsets for column and row
#' @param ... additional arguments
#' @examples
#' if (requireNamespace("mschart")) {
#' require(mschart)
#'
#' ## Add mschart to worksheet (adds data and chart)
#' scatter <- ms_scatterchart(data = iris, x = "Sepal.Length", y = "Sepal.Width", group = "Species")
#' scatter <- chart_settings(scatter, scatterstyle = "marker")
#'
#' wb <- wb_workbook() %>%
#'  wb_add_worksheet() %>%
#'  wb_add_mschart(dims = "F4:L20", graph = scatter)
#'
#' ## Add mschart to worksheet and use available data
#' wb <- wb_workbook() %>%
#'   wb_add_worksheet() %>%
#'   wb_add_data(x = mtcars, dims = "B2")
#'
#' # create wb_data object
#' dat <- wb_data(wb, 1, dims = "B2:E6")
#'
#' # call ms_scatterplot
#' data_plot <- ms_scatterchart(
#'   data = dat,
#'   x = "mpg",
#'   y = c("disp", "hp"),
#'   labels = c("disp", "hp")
#' )
#'
#' # add the scatterplot to the data
#' wb <- wb %>%
#'   wb_add_mschart(dims = "F4:L20", graph = data_plot)
#' }
#' @seealso [wb_data()] [wb_add_chart_xml()] [wb_add_image] [wb_add_mschart()] [wb_add_plot]
#' @export
wb_add_mschart <- function(
    wb,
    sheet      = current_sheet(),
    dims       = NULL,
    graph,
    col_offset = 0,
    row_offset = 0,
    ...
) {
  assert_workbook(wb)
  wb$clone()$add_mschart(
    sheet      = sheet,
    dims       = dims,
    graph      = graph,
    col_offset = col_offset,
    row_offset = row_offset,
    ...        = ...
  )
}

#' Remove a worksheet from a workbook
#' @param wb A wbWorkbook object
#' @param sheet The sheet name or index to remove
#' @returns The `wbWorkbook` object, invisibly.
#' @export
#' @examples
#' ## load a workbook
#' wb <- wb_load(file = system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2"))
#'
#' ## Remove sheet 2
#' wb <- wb_remove_worksheet(wb, 2)
wb_remove_worksheet <- function(wb, sheet = current_sheet()) {
  assert_workbook(wb)
  wb$clone()$remove_worksheet(sheet = sheet)
}


# base font ---------------------------------------------------------------

#' Set the default font in a workbook
#'
#' Modify / get the default font for the workbook. This will alter the latin
#' major and minor font in the workbooks theme.
#'
#' The font name is not validated in anyway. Spreadsheet software replaces
#' unknown font names with system defaults.
#'
#' The default base font is Aptos Narrow, black, size 11. If `font_name` differs
#' from the name in [wb_get_base_font()], the theme is updated to use the newly
#' selected font name.
#'
#' @param wb A workbook object
#' @param font_size Font size
#' @param font_color Font color
#' @param font_name Name of a font
#' @param ... Additional arguments
#' @family workbook styling functions
#' @family workbook wrappers
#' @name base_font-wb
#' @examples
#' ## create a workbook
#' wb <- wb_workbook(theme = "Office 2013 - 2022 Theme")
#' wb$add_worksheet("S1")
#' ## modify base font to size 10 Aptos Narrow in red
#' wb$set_base_font(font_size = 10, font_color = wb_color("red"), font_name = "Aptos Narrow")
#'
#' wb$add_data(x = iris)
#'
#' ## font color does not affect tables
#' wb$add_data_table(x = iris, dims = wb_dims(from_col = 10))
#'
#' ## get the base font
#' wb_get_base_font(wb)
#'
NULL
#' @export
#' @rdname base_font-wb
wb_set_base_font <- function(
  wb,
  font_size  = 11,
  font_color = wb_color(theme = "1"),
  font_name  = "Aptos Narrow",
  ...
) {
  assert_workbook(wb)
  wb$clone()$set_base_font(
    font_size   = font_size,
    font_color  = font_color,
    font_name   = font_name,
    ...         = ...
  )
}
#' @export
#' @rdname base_font-wb
wb_get_base_font <- function(wb) {
  # TODO all of these class checks need to be cleaned up
  assert_workbook(wb)
  wb$get_base_font()
}

#' Set the default colors in a workbook
#'
#' Modify / get the default colors of the workbook.
#'
#' @details Theme must be any of the following:
#' "Aspect", "Blue", "Blue II", "Blue Green", "Blue Warm", "Greyscale",
#' "Green", "Green Yellow", "Marquee", "Median", "Office", "Office 2007 - 2010",
#' "Office 2013 - 2022", "Orange", "Orange Red", "Paper", "Red",
#' "Red Orange", "Red Violet", "Slipstream", "Violet", "Violet II",
#' "Yellow", "Yellow Orange"
#'
#' @name wb_base_colors
#' @param wb A workbook object
#' @param theme a predefined color theme
#' @param ... optional parameters
#' @param xml Logical if xml string should be returned
#' @param plot Logical if a barplot of the colors should be returned
#' @family workbook styling functions
#' @family workbook wrappers
#' @examples
#' wb <- wb_workbook()
#' wb$get_base_colors()
#' wb$set_base_colors(theme = 3)
#' wb$set_base_colors(theme = "Violet II")
#' wb$get_base_colours()
NULL
#' @export
#' @rdname wb_base_colors
wb_set_base_colors <- function(
  wb,
  theme = "Office",
  ...
) {
  assert_workbook(wb)
  wb$clone()$set_base_colors(
    theme = theme,
    ...   = ...
  )
}
#' @export
#' @rdname wb_base_colors
wb_get_base_colors <- function(wb, xml = FALSE, plot = TRUE) {
  assert_workbook(wb)
  wb$get_base_colors(xml = xml, plot = plot)
}
#' @export
#' @rdname wb_base_colors
#' @usage NULL
wb_set_base_colours <- wb_set_base_colors
#' @export
#' @rdname wb_base_colors
#' @usage NULL
wb_get_base_colours <- wb_get_base_colors


#' Set the workbook position, size and filter
#'
#' @param wb A [wbWorkbook] object
#' @param active_tab activeTab
#' @param auto_filter_date_grouping autoFilterDateGrouping
#' @param first_sheet The first sheet to be displayed
#' @param minimized minimized
#' @param show_horizontal_scroll showHorizontalScroll
#' @param show_sheet_tabs showSheetTabs
#' @param show_vertical_scroll showVerticalScroll
#' @param tab_ratio tabRatio
#' @param visibility visibility
#' @param window_height windowHeight
#' @param window_width windowWidth
#' @param x_window xWindow
#' @param y_window yWindow
#' @param ... additional arguments
#' @return The Workbook object
#' @export
wb_set_bookview <- function(
    wb,
    active_tab                = NULL,
    auto_filter_date_grouping = NULL,
    first_sheet               = NULL,
    minimized                 = NULL,
    show_horizontal_scroll    = NULL,
    show_sheet_tabs           = NULL,
    show_vertical_scroll      = NULL,
    tab_ratio                 = NULL,
    visibility                = NULL,
    window_height             = NULL,
    window_width              = NULL,
    x_window                  = NULL,
    y_window                  = NULL,
    ...
) {
  assert_workbook(wb)
  wb$clone()$set_bookview(
    active_tab                = active_tab,
    auto_filter_date_grouping = auto_filter_date_grouping,
    first_sheet               = first_sheet,
    minimized                 = minimized,
    show_horizontal_scroll    = show_horizontal_scroll,
    show_sheet_tabs           = show_sheet_tabs,
    show_vertical_scroll      = show_vertical_scroll,
    tab_ratio                 = tab_ratio,
    visibility                = visibility,
    window_height             = window_height,
    window_width              = window_width,
    x_window                  = x_window,
    y_window                  = y_window,
    ...                       = ...
  )
}

#' Set headers and footers of a worksheet
#'
#' Set document headers and footers. You can also do this when adding a worksheet
#' with [wb_add_worksheet()] with the `header`, `footer` arguments and friends.
#' These will show up when printing an xlsx file.
#'
#' Headers and footers can contain special tags
#' * **&\[Page\]** Page number
#' * **&\[Pages\]** Number of pages
#' * **&\[Date\]** Current date
#' * **&\[Time\]** Current time
#' * **&\[Path\]** File path
#' * **&\[File\]** File name
#' * **&\[Tab\]** Worksheet name
#'
#' @param wb A Workbook object
#' @param sheet A name or index of a worksheet
#' @param header,even_header,first_header,footer,even_footer,first_footer
#'   Character vector of length 3 corresponding to positions left, center,
#'   right.  `header` and `footer` are used to default additional arguments.
#'   Setting `even`, `odd`, or `first`, overrides `header`/`footer`. Use `NA` to
#'   skip a position.
# #' @inheritParams wb_add_worksheet
#' @param ... additional arguments
#' @export
#' @examples
#' wb <- wb_workbook()
#'
#' # Add example data
#' wb$add_worksheet("S1")$add_data(x = 1:400)
#' wb$add_worksheet("S2")$add_data(x = 1:400)
#' wb$add_worksheet("S3")$add_data(x = 3:400)
#' wb$add_worksheet("S4")$add_data(x = 3:400)
#'
#' wb$set_header_footer(
#'   sheet = "S1",
#'   header = c("ODD HEAD LEFT", "ODD HEAD CENTER", "ODD HEAD RIGHT"),
#'   footer = c("ODD FOOT RIGHT", "ODD FOOT CENTER", "ODD FOOT RIGHT"),
#'   even_header = c("EVEN HEAD LEFT", "EVEN HEAD CENTER", "EVEN HEAD RIGHT"),
#'   even_footer = c("EVEN FOOT RIGHT", "EVEN FOOT CENTER", "EVEN FOOT RIGHT"),
#'   first_header = c("TOP", "OF FIRST", "PAGE"),
#'   first_footer = c("BOTTOM", "OF FIRST", "PAGE")
#' )
#'
#' wb$set_header_footer(
#'   sheet = 2,
#'   header = c("&[Date]", "ALL HEAD CENTER 2", "&[Page] / &[Pages]"),
#'   footer = c("&[Path]&[File]", NA, "&[Tab]"),
#'   first_header = c(NA, "Center Header of First Page", NA),
#'   first_footer = c(NA, "Center Footer of First Page", NA)
#' )
#'
#' wb$set_header_footer(
#'   sheet = 3,
#'   header = c("ALL HEAD LEFT 2", "ALL HEAD CENTER 2", "ALL HEAD RIGHT 2"),
#'   footer = c("ALL FOOT RIGHT 2", "ALL FOOT CENTER 2", "ALL FOOT RIGHT 2")
#' )
#'
#' wb$set_header_footer(
#'   sheet = 4,
#'   first_header = c("FIRST ONLY L", NA, "FIRST ONLY R"),
#'   first_footer = c("FIRST ONLY L", NA, "FIRST ONLY R")
#' )
wb_set_header_footer <- function(
    wb,
    sheet        = current_sheet(),
    header       = NULL,
    footer       = NULL,
    even_header  = NULL,
    even_footer  = NULL,
    first_header = NULL,
    first_footer = NULL,
    ...
) {
  assert_workbook(wb)
  wb$clone()$set_header_footer(
    sheet        = sheet,
    header       = header,
    footer       = footer,
    even_header  = even_header,
    even_footer  = even_footer,
    first_header = first_header,
    first_footer = first_footer,
    ...          = ...
  )
}



#' Set page margins, orientation and print scaling of a worksheet
#'
#' Set page margins, orientation and print scaling.
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param orientation Page orientation. One of "portrait" or "landscape"
#' @param scale Print scaling. Numeric value between 10 and 400
#' @param left,right,top,bottom  Page margin in inches
#' @param header,footer  Margin in inches
#' @param fit_to_width,fit_to_height An integer that tells the spreadsheet software on how many pages the scaling should fit. This does not actually scale the sheet.
#' @param paper_size See details. Default value is 9 (A4 paper).
#' @param print_title_rows,print_title_cols Rows / columns to repeat at top of page when printing. Integer vector.
#' @param summary_row Location of summary rows in groupings. One of "Above" or "Below".
#' @param summary_col Location of summary columns in groupings. One of "Right" or "Left".
#' @param ... additional arguments
#' @export
#' @details
#'  When adding fitting to width and height manual adjustment of the scaling factor is required. Setting `fit_to_width` and `fit_to_height` only tells spreadsheet software that the scaling was applied, but not which scaling was applied.
#'
#' `paper_size` is an integer corresponding to:
#'
#'  | size   | "paper type"                                                                         |
#'  |--------|--------------------------------------------------------------------------------------|
#'  | 1      | Letter paper (8.5 in. by 11 in.)                                                     |
#'  | 2      |  Letter small paper (8.5 in. by 11 in.)                                              |
#'  | 3      |  Tabloid paper (11 in. by 17 in.)                                                    |
#'  | 4      |  Ledger paper (17 in. by 11 in.)                                                     |
#'  | 5      |  Legal paper (8.5 in. by 14 in.)                                                     |
#'  | 6      |  Statement paper (5.5 in. by 8.5 in.)                                                |
#'  | 7      |  Executive paper (7.25 in. by 10.5 in.)                                              |
#'  | 8      |  A3 paper (297 mm by 420 mm)                                                         |
#'  | 9      |  A4 paper (210 mm by 297 mm)                                                         |
#'  | 10     |  A4 small paper (210 mm by 297 mm)                                                   |
#'  | 11     |  A5 paper (148 mm by 210 mm)                                                         |
#'  | 12     |  B4 paper (250 mm by 353 mm)                                                         |
#'  | 13     |  B5 paper (176 mm by 250 mm)                                                         |
#'  | 14     |  Folio paper (8.5 in. by 13 in.)                                                     |
#'  | 15     |  Quarto paper (215 mm by 275 mm)                                                     |
#'  | 16     |  Standard paper (10 in. by 14 in.)                                                   |
#'  | 17     |  Standard paper (11 in. by 17 in.)                                                   |
#'  | 18     |  Note paper (8.5 in. by 11 in.)                                                      |
#'  | 19     |  #9 envelope (3.875 in. by 8.875 in.)                                                |
#'  | 20     |  #10 envelope (4.125 in. by 9.5 in.)                                                 |
#'  | 21     |  #11 envelope (4.5 in. by 10.375 in.)                                                |
#'  | 22     |  #12 envelope (4.75 in. by 11 in.)                                                   |
#'  | 23     |  #14 envelope (5 in. by 11.5 in.)                                                    |
#'  | 24     |  C paper (17 in. by 22 in.)                                                          |
#'  | 25     |  D paper (22 in. by 34 in.)                                                          |
#'  | 26     |  E paper (34 in. by 44 in.)                                                          |
#'  | 27     |  DL envelope (110 mm by 220 mm)                                                      |
#'  | 28     |  C5 envelope (162 mm by 229 mm)                                                      |
#'  | 29     |  C3 envelope (324 mm by 458 mm)                                                      |
#'  | 30     |  C4 envelope (229 mm by 324 mm)                                                      |
#'  | 31     |  C6 envelope (114 mm by 162 mm)                                                      |
#'  | 32     |  C65 envelope (114 mm by 229 mm)                                                     |
#'  | 33     |  B4 envelope (250 mm by 353 mm)                                                      |
#'  | 34     |  B5 envelope (176 mm by 250 mm)                                                      |
#'  | 35     |  B6 envelope (176 mm by 125 mm)                                                      |
#'  | 36     |  Italy envelope (110 mm by 230 mm)                                                   |
#'  | 37     |  Monarch envelope (3.875 in. by 7.5 in.)                                             |
#'  | 38     |  6 3/4 envelope (3.625 in. by 6.5 in.)                                               |
#'  | 39     |  US standard fanfold (14.875 in. by 11 in.)                                          |
#'  | 40     |  German standard fanfold (8.5 in. by 12 in.)                                         |
#'  | 41     |  German legal fanfold (8.5 in. by 13 in.)                                            |
#'  | 42     |  ISO B4 (250 mm by 353 mm)                                                           |
#'  | 43     |  Japanese double postcard (200 mm by 148 mm)                                         |
#'  | 44     |  Standard paper (9 in. by 11 in.)                                                    |
#'  | 45     |  Standard paper (10 in. by 11 in.)                                                   |
#'  | 46     |  Standard paper (15 in. by 11 in.)                                                   |
#'  | 47     |  Invite envelope (220 mm by 220 mm)                                                  |
#'  | 50     |  Letter extra paper (9.275 in. by 12 in.)                                            |
#'  | 51     |  Legal extra paper (9.275 in. by 15 in.)                                             |
#'  | 52     |  Tabloid extra paper (11.69 in. by 18 in.)                                           |
#'  | 53     |  A4 extra paper (236 mm by 322 mm)                                                   |
#'  | 54     |  Letter transverse paper (8.275 in. by 11 in.)                                       |
#'  | 55     |  A4 transverse paper (210 mm by 297 mm)                                              |
#'  | 56     |  Letter extra transverse paper (9.275 in. by 12 in.)                                 |
#'  | 57     |  SuperA/SuperA/A4 paper (227 mm by 356 mm)                                           |
#'  | 58     |  SuperB/SuperB/A3 paper (305 mm by 487 mm)                                           |
#'  | 59     |  Letter plus paper (8.5 in. by 12.69 in.)                                            |
#'  | 60     |  A4 plus paper (210 mm by 330 mm)                                                    |
#'  | 61     |  A5 transverse paper (148 mm by 210 mm)                                              |
#'  | 62     |  JIS B5 transverse paper (182 mm by 257 mm)                                          |
#'  | 63     |  A3 extra paper (322 mm by 445 mm)                                                   |
#'  | 64     |  A5 extra paper (174 mm by 235 mm)                                                   |
#'  | 65     |  ISO B5 extra paper (201 mm by 276 mm)                                               |
#'  | 66     |  A2 paper (420 mm by 594 mm)                                                         |
#'  | 67     |  A3 transverse paper (297 mm by 420 mm)                                              |
#'  | 68     |  A3 extra transverse paper (322 mm by 445 mm)                                        |
#'  | 69     | Japanese Double Postcard (200 mm x 148 mm) 70=A6(105mm x 148mm)                      |
#'  | 71     | Japanese Envelope Kaku #2                                                            |
#'  | 72     | Japanese Envelope Kaku #3                                                            |
#'  | 73     | Japanese Envelope Chou #3                                                            |
#'  | 74     | Japanese Envelope Chou #4                                                            |
#'  | 75     | Letter Rotated (11in x 8 1/2 11 in)                                                  |
#'  | 76     | A3 Rotated (420 mm x 297 mm)                                                         |
#'  | 77     | A4 Rotated (297 mm x 210 mm)                                                         |
#'  | 78     | A5 Rotated (210 mm x 148 mm)                                                         |
#'  | 79     | B4 (JIS) Rotated (364 mm x 257 mm)                                                   |
#'  | 80     | B5 (JIS) Rotated (257 mm x 182 mm)                                                   |
#'  | 81     | Japanese Postcard Rotated (148 mm x 100 mm)                                          |
#'  | 82     | Double Japanese Postcard Rotated (148 mm x 200 mm) 83 = A6 Rotated (148 mm x 105 mm) |
#'  | 84     | Japanese Envelope Kaku #2 Rotated                                                    |
#'  | 85     | Japanese Envelope Kaku #3 Rotated                                                    |
#'  | 86     | Japanese Envelope Chou #3 Rotated                                                    |
#'  | 87     | Japanese Envelope Chou #4 Rotated 88=B6(JIS)(128mm x 182mm)                          |
#'  | 89     | B6 (JIS) Rotated (182 mm x 128 mm)                                                   |
#'  | 90     | (12 in x 11 in)                                                                      |
#'  | 91     | Japanese Envelope You #4                                                             |
#'  | 92     | Japanese Envelope You #4 Rotated 93=PRC16K(146mm x 215mm) 94=PRC32K(97mm x 151mm)    |
#'  | 95     | PRC 32K(Big) (97 mm x 151 mm)                                                        |
#'  | 96     | PRC Envelope #1 (102 mm x 165 mm)                                                    |
#'  | 97     | PRC Envelope #2 (102 mm x 176 mm)                                                    |
#'  | 98     | PRC Envelope #3 (125 mm x 176 mm)                                                    |
#'  | 99     | PRC Envelope #4 (110 mm x 208 mm)                                                    |
#'  | 100    | PRC Envelope #5 (110 mm x 220 mm)                                                    |
#'  | 101    | PRC Envelope #6 (120 mm x 230 mm)                                                    |
#'  | 102    | PRC Envelope #7 (160 mm x 230 mm)                                                    |
#'  | 103    | PRC Envelope #8 (120 mm x 309 mm)                                                    |
#'  | 104    | PRC Envelope #9 (229 mm x 324 mm)                                                    |
#'  | 105    | PRC Envelope #10 (324 mm x 458 mm)                                                   |
#'  | 106    | PRC 16K Rotated                                                                      |
#'  | 107    | PRC 32K Rotated                                                                      |
#'  | 108    | PRC 32K(Big) Rotated                                                                 |
#'  | 109    | PRC Envelope #1 Rotated (165 mm x 102 mm)                                            |
#'  | 110    | PRC Envelope #2 Rotated (176 mm x 102 mm)                                            |
#'  | 111    | PRC Envelope #3 Rotated (176 mm x 125 mm)                                            |
#'  | 112    | PRC Envelope #4 Rotated (208 mm x 110 mm)                                            |
#'  | 113    | PRC Envelope #5 Rotated (220 mm x 110 mm)                                            |
#'  | 114    | PRC Envelope #6 Rotated (230 mm x 120 mm)                                            |
#'  | 115    | PRC Envelope #7 Rotated (230 mm x 160 mm)                                            |
#'  | 116    | PRC Envelope #8 Rotated (309 mm x 120 mm)                                            |
#'  | 117    | PRC Envelope #9 Rotated (324 mm x 229 mm)                                            |
#'  | 118    | PRC Envelope #10 Rotated (458 mm x 324 mm)                                           |
#'
#' @examples
#' wb <- wb_workbook()
#' wb$add_worksheet("S1")
#' wb$add_worksheet("S2")
#' wb$add_data_table(1, x = iris[1:30, ])
#' wb$add_data_table(2, x = iris[1:30, ], dims = c("C5"))
#'
#' ## landscape page scaled to 50%
#' wb$page_setup(sheet = 1, orientation = "landscape", scale = 50)
#'
#' ## portrait page scales to 300% with 0.5in left and right margins
#' wb$page_setup(sheet = 2, orientation = "portrait", scale = 300, left = 0.5, right = 0.5)
#'
#'
#' ## print titles
#' wb$add_worksheet("print_title_rows")
#' wb$add_worksheet("print_title_cols")
#'
#' wb$add_data("print_title_rows", rbind(iris, iris, iris, iris))
#' wb$add_data("print_title_cols", x = rbind(mtcars, mtcars, mtcars), row_names = TRUE)
#'
#' wb$page_setup(sheet = "print_title_rows", print_title_rows = 1) ## first row
#' wb$page_setup(sheet = "print_title_cols", print_title_cols = 1, print_title_rows = 1)
wb_page_setup <- function(
    wb,
    sheet            = current_sheet(),
    orientation      = NULL,
    scale            = 100,
    left             = 0.7,
    right            = 0.7,
    top              = 0.75,
    bottom           = 0.75,
    header           = 0.3,
    footer           = 0.3,
    fit_to_width     = FALSE,
    fit_to_height    = FALSE,
    paper_size       = NULL,
    print_title_rows = NULL,
    print_title_cols = NULL,
    summary_row      = NULL,
    summary_col      = NULL,
    ...
) {
  assert_workbook(wb)
  wb$clone()$page_setup(
    sheet            = sheet,
    orientation      = orientation,
    scale            = scale,
    left             = left,
    right            = right,
    top              = top,
    bottom           = bottom,
    header           = header,
    footer           = footer,
    fit_to_width     = fit_to_width,
    fit_to_height    = fit_to_height,
    paper_size       = paper_size,
    print_title_rows = print_title_rows,
    print_title_cols = print_title_cols,
    summary_row      = summary_row,
    summary_col      = summary_col,
    ...              = ...
  )
}


# protect ----------------------------------------------------------------------

#' Protect a worksheet from modifications
#'
#' Protect or unprotect a worksheet from modifications by the user in the graphical user interface. Replaces an existing protection. Certain features require applying unlocking of initialized cells in the worksheet and across columns and/or rows.
#'
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param protect Whether to protect or unprotect the sheet (default=TRUE)
#' @param password (optional) password required to unprotect the worksheet
#' @param properties A character vector of properties to lock.  Can be one or
#'   more of the following: `"selectLockedCells"`, `"selectUnlockedCells"`,
#'   `"formatCells"`, `"formatColumns"`, `"formatRows"`, `"insertColumns"`,
#'   `"insertRows"`, `"insertHyperlinks"`, `"deleteColumns"`, `"deleteRows"`,
#'   `"sort"`, `"autoFilter"`, `"pivotTables"`, `"objects"`, `"scenarios"`
#' @export
#' @examples
#' wb <- wb_workbook()
#' wb$add_worksheet("S1")
#' wb$add_data_table(1, x = iris[1:30, ])
#'
#' wb$protect_worksheet(
#'   "S1",
#'   protect = TRUE,
#'   properties = c("formatCells", "formatColumns", "insertColumns", "deleteColumns")
#' )
#'
#' # Formatting cells / columns is allowed , but inserting / deleting columns is protected:
#' wb$protect_worksheet(
#'   "S1",
#'   protect = TRUE,
#'    c(formatCells = FALSE, formatColumns = FALSE,
#'                  insertColumns = TRUE, deleteColumns = TRUE)
#' )
#'
#' # Remove the protection
#' wb$protect_worksheet("S1", protect = FALSE)
wb_protect_worksheet <- function(
    wb,
    sheet      = current_sheet(),
    protect    = TRUE,
    password   = NULL,
    properties = NULL
) {

  assert_workbook(wb)
  wb$clone(deep = TRUE)$protect_worksheet(
    sheet      = sheet,
    protect    = protect,
    password   = password,
    properties = properties
  )
}


#' Protect a workbook from modifications
#'
#' Protect or unprotect a workbook from modifications by the user in the
#' graphical user interface. Replaces an existing protection.
#'
#' @param wb A Workbook object
#' @param protect Whether to protect or unprotect the sheet (default `TRUE`)
#' @param password (optional) password required to unprotect the workbook
#' @param lock_structure Whether the workbook structure should be locked
#' @param lock_windows Whether the window position of the spreadsheet should be
#'   locked
#' @param type Lock type (see **Details**)
#' @param file_sharing Whether to enable a popup requesting the unlock password
#'   is prompted
#' @param username The username for the `file_sharing` popup
#' @param read_only_recommended Whether or not a post unlock message appears
#'   stating that the workbook is recommended to be opened in read-only mode.
#' @param ... additional arguments
#'
#' @details
#' Lock types:
#'
#' * `1`  xlsx with password (default)
#' * `2` xlsx recommends read-only
#' * `4` xlsx enforces read-only
#' * `8` xlsx is locked for annotation
#'
#' @export
#' @examples
#' wb <- wb_workbook()
#' wb$add_worksheet("S1")
#' wb_protect(wb, protect = TRUE, password = "Password", lock_structure = TRUE)
#'
#' # Remove the protection
#' wb_protect(wb, protect = FALSE)
#'
#' wb <- wb_protect(
#'   wb,
#'   protect = TRUE,
#'   password = "Password",
#'   lock_structure = TRUE,
#'   type = 2L,
#'   file_sharing = TRUE,
#'   username = "Test",
#'   read_only_recommended = TRUE
#' )
#'
wb_protect <- function(
    wb,
    protect               = TRUE,
    password              = NULL,
    lock_structure        = FALSE,
    lock_windows          = FALSE,
    type                  = 1,
    file_sharing          = FALSE,
    username              = unname(Sys.info()["user"]),
    read_only_recommended = FALSE,
    ...
) {
  assert_workbook(wb)
  wb$clone()$protect(
    protect               = protect,
    password              = password,
    lock_structure        = lock_structure,
    lock_windows          = lock_windows,
    type                  = type,
    file_sharing          = file_sharing,
    username              = username,
    read_only_recommended = read_only_recommended,
    ...                   = ...
  )
}


# grid lines --------------------------------------------------------------

#' Modify grid lines visibility in a worksheet
#'
#' Set worksheet grid lines to show or hide.
#' You can also add / remove grid lines when creating a worksheet with
#' [`wb_add_worksheet(grid_lines = FALSE)`][wb_add_worksheet()]
#'
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param show A logical. If `FALSE`, grid lines are hidden.
#' @param print A logical. If `FALSE`, grid lines are not printed.
#' @export
#' @examples
#' wb <- wb_workbook()$add_worksheet()$add_worksheet()
#' wb$get_sheet_names() ## list worksheets in workbook
#' wb$set_grid_lines(1, show = FALSE)
#' wb$set_grid_lines("Sheet 2", show = FALSE)
wb_set_grid_lines <- function(wb, sheet = current_sheet(), show = FALSE, print = show) {
  assert_workbook(wb)
  wb$clone()$set_grid_lines(sheet = sheet, show = show, print = print)
}

#' @rdname wb_set_grid_lines
#' @export
wb_grid_lines <- function(wb, sheet = current_sheet(), show = FALSE, print = show) {
  assert_workbook(wb)
  .Deprecated(old = "wb_grid_lines", new = "wb_set_grid_lines", package = "openxlsx2")
  wb$clone()$set_grid_lines(sheet = sheet, show = show, print = print)
}

# TODO hide gridlines?

# worksheet order ---------------------------------------------------------

#' Order worksheets in a workbook
#'
#' Get/set order of worksheets in a Workbook object
#'
#' @param wb A `wbWorkbook` object
#'
#' @details This function does not reorder the worksheets within the workbook
#'   object, it simply shuffles the order when writing to file.
#' @export
#' @examples
#' ## setup a workbook with 3 worksheets
#' wb <- wb_workbook()
#' wb$add_worksheet("Sheet 1", grid_lines = FALSE)
#' wb$add_data_table(x = iris)
#'
#' wb$add_worksheet("mtcars (Sheet 2)", grid_lines = FALSE)
#' wb$add_data(x = mtcars)
#'
#' wb$add_worksheet("Sheet 3", grid_lines = FALSE)
#' wb$add_data(x = Formaldehyde)
#'
#' wb_get_order(wb)
#' wb$get_sheet_na
#' wb$set_order(c(1, 3, 2)) # switch position of sheets 2 & 3
#' wb$add_data(2, 'This is still the "mtcars" worksheet', start_col = 15)
#' wb_get_order(wb)
#' wb$get_sheet_names() ## ordering within workbook is not changed
#' wb$set_order(3:1)
#' @name wb_order
wb_get_order <- function(wb) {
  assert_workbook(wb)
  wb$sheetOrder
}

#' @rdname wb_order
#' @param sheets Sheet order
#' @export
wb_set_order <- function(wb, sheets) {
  assert_workbook(wb)
  wb$clone()$set_order(sheets = sheets)
}


# named region ------------------------------------------------------------


#' Modify named regions in a worksheet
#'
#' Create / delete a named region. You can also specify a named region by using
#' the `name` argument in `wb_add_data(x = iris, name = "my-region")`.
#' It is important to note that named regions are not case-sensitive and must be unique.
#'
#' You can use the [wb_dims()] helper to specify the cell range of the named region
#'
#' @param wb A Workbook object
#' @param sheet A name or index of a worksheet
#' @param dims Worksheet cell range of the region ("A1:D4").
#' @param name Name for region. A character vector of length 1. Note that region
#'   names must be case-insensitive unique.
#' @param overwrite Boolean. Overwrite if exists? Default to `FALSE`.
#' @param local_sheet If `TRUE` the named region will be local for this sheet
#' @param comment description text for named region
#' @param hidden Should the named region be hidden?
#' @param custom_menu,description,is_function,function_group_id,help,local_name,publish_to_server,status_bar,vb_procedure,workbook_parameter,xml Unknown XML feature
#' @param ... additional arguments
#' @returns A workbook, invisibly.
#' @family worksheet content functions
#' @examples
#' ## create named regions
#' wb <- wb_workbook()
#' wb$add_worksheet("Sheet 1")
#'
#' ## specify region
#' wb$add_data(x = iris, start_col = 1, start_row = 1)
#' wb$add_named_region(
#'   name = "iris",
#'   dims = wb_dims(x = iris)
#' )
#'
#' ## using add_data 'name' argument
#' wb$add_data(sheet = 1, x = iris, name = "iris2", start_col = 10)
#'
#' ## delete one
#' wb$remove_named_region(name = "iris2")
#' wb$get_named_regions()
#' ## read named regions
#' df <- wb_to_df(wb, named_region = "iris")
#' head(df)
#'
#' @name named_region-wb
NULL

#' @rdname named_region-wb
#' @export
wb_add_named_region <- function(
  wb,
  sheet             = current_sheet(),
  dims              = "A1",
  name,
  local_sheet        = FALSE,
  overwrite          = FALSE,
  comment            = NULL,
  hidden             = NULL,
  custom_menu        = NULL,
  description        = NULL,
  is_function        = NULL,
  function_group_id  = NULL,
  help               = NULL,
  local_name         = NULL,
  publish_to_server  = NULL,
  status_bar         = NULL,
  vb_procedure       = NULL,
  workbook_parameter = NULL,
  xml                = NULL,
  ...
) {
  assert_workbook(wb)
  wb$clone()$add_named_region(
    sheet              = sheet,
    dims               = dims,
    name               = name,
    local_sheet        = local_sheet,
    overwrite          = overwrite,
    comment            = comment,
    custom_menu        = custom_menu,
    description        = description,
    is_function        = is_function,
    function_group_id  = function_group_id,
    help               = help,
    hidden             = hidden,
    local_name         = local_name,
    publish_to_server  = publish_to_server,
    status_bar         = status_bar,
    vb_procedure       = vb_procedure,
    workbook_parameter = workbook_parameter,
    xml                = xml,
    ...                = ...
  )
}

#' @rdname named_region-wb
#' @export
wb_remove_named_region <- function(wb, sheet = current_sheet(), name = NULL) {
  assert_workbook(wb)
  wb$clone()$remove_named_region(sheet = sheet, name = name)
}

# filters -----------------------------------------------------------------

#' Add/remove column filters in a worksheet
#'
#' Add or remove excel column filters to a worksheet
#'
#' Adds filters to worksheet columns, same as `with_filter = TRUE` in [wb_add_data()]
#' [wb_add_data_table()] automatically adds filters to first row of a table.
#'
#' NOTE Can only have a single filter per worksheet unless using tables.
#'
#' @param wb A workbook object
#' @param sheet A worksheet name or index.
#'   In `wb_remove_filter()`, you may supply a vector of worksheets.
#' @param cols columns to add filter to.
#' @param rows A row number.
#' @seealso [wb_add_data()], [wb_add_data_table()]
#' @examples
#' wb <- wb_workbook()
#' wb$add_worksheet("Sheet 1")
#' wb$add_worksheet("Sheet 2")
#' wb$add_worksheet("Sheet 3")
#'
#' wb$add_data(1, iris)
#' wb$add_filter(1, row = 1, cols = seq_along(iris))
#'
#' ## Equivalently
#' wb$add_data(2, x = iris, with_filter = TRUE)
#'
#' ## Similarly
#' wb$add_data_table(3, iris)
#' wb <- wb_workbook()
#' wb$add_worksheet("Sheet 1")
#' wb$add_worksheet("Sheet 2")
#' wb$add_worksheet("Sheet 3")
#'
#' wb$add_data(1, iris)
#' wb_add_filter(wb, 1, row = 1, cols = seq_along(iris))
#'
#' ## Equivalently
#' wb$add_data(2, x = iris, with_filter = TRUE)
#'
#' ## Similarly
#' wb$add_data_table(3, iris)
#'
#' ## remove filters
#' wb_remove_filter(wb, 1:2) ## remove filters
#' wb_remove_filter(wb, 3) ## Does not affect tables!
#' @name filter-wb
#' @family worksheet content functions
NULL
#' @rdname filter-wb
#' @export
wb_add_filter <- function(wb, sheet = current_sheet(), rows, cols) {
  assert_workbook(wb)
  wb$clone()$add_filter(sheet = sheet, rows = rows, cols = cols)
}
#' @rdname filter-wb
#' @export
wb_remove_filter <- function(wb, sheet = current_sheet()) {
  assert_workbook(wb)
  wb$clone()$remove_filter(sheet = sheet)
}


# validations ------------------------------------------------------------------

#' Add data validation to cells in a worksheet
#'
#' Add Excel data validation to cells
#'
#' @param wb A Workbook object
#' @param sheet A name or index of a worksheet
#' @param dims A cell dimension ("A1" or "A1:B2")
#' @param type One of 'whole', 'decimal', 'date', 'time', 'textLength', 'list'
#'   (see examples)
#' @param operator One of 'between', 'notBetween', 'equal',
#'  'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'
#' @param value a vector of length 1 or 2 depending on operator (see examples)
#' @param allow_blank logical
#' @param show_input_msg logical
#' @param show_error_msg logical
#' @param error_style The icon shown and the options how to deal with such inputs.
#'   Default "stop" (cancel), else "information" (prompt popup) or
#'   "warning" (prompt accept or change input)
#' @param error_title The error title
#' @param error The error text
#' @param prompt_title The prompt title
#' @param prompt The prompt text
#' @param ... additional arguments
#' @export
#' @examples
#' wb <- wb_workbook()
#' wb$add_worksheet("Sheet 1")
#' wb$add_worksheet("Sheet 2")
#'
#' wb$add_data_table(1, x = iris[1:30, ])
#' wb$add_data_validation(1,
#'   dims = "A2:C31", type = "whole",
#'   operator = "between", value = c(1, 9)
#' )
#' wb$add_data_validation(1,
#'   dims = "E2:E31", type = "textLength",
#'   operator = "between", value = c(4, 6)
#' )
#'
#' ## Date and Time cell validation
#' df <- data.frame(
#'   "d" = as.Date("2016-01-01") + -5:5,
#'   "t" = as.POSIXct("2016-01-01") + -5:5 * 10000
#' )
#' wb$add_data_table(2, x = df)
#' wb$add_data_validation(2, dims = "A2:A12", type = "date",
#'   operator = "greaterThanOrEqual", value = as.Date("2016-01-01")
#' )
#' wb$add_data_validation(2,
#'   dims = "B2:B12", type = "time",
#'   operator = "between", value = df$t[c(4, 8)]
#' )
#'
#'
#' ######################################################################
#' ## If type == 'list'
#' # operator argument is ignored.
#'
#' wb <- wb_workbook()
#' wb$add_worksheet("Sheet 1")
#' wb$add_worksheet("Sheet 2")
#'
#' wb$add_data_table(sheet = 1, x = iris[1:30, ])
#' wb$add_data(sheet = 2, x = sample(iris$Sepal.Length, 10))
#'
#' wb$add_data_validation(1, dims = "A2:A31", type = "list", value = "'Sheet 2'!$A$1:$A$10")
wb_add_data_validation <- function(
    wb,
    sheet          = current_sheet(),
    dims           = "A1",
    type,
    operator,
    value,
    allow_blank    = TRUE,
    show_input_msg = TRUE,
    show_error_msg = TRUE,
    error_style    = NULL,
    error_title    = NULL,
    error          = NULL,
    prompt_title   = NULL,
    prompt         = NULL,
    ...
) {
  assert_workbook(wb)
  wb$clone(deep = TRUE)$add_data_validation(
    sheet          = sheet,
    dims           = dims,
    type           = type,
    operator       = operator,
    value          = value,
    allowBlank     = allow_blank,
    show_input_msg = show_input_msg,
    show_error_msg = show_error_msg,
    error_style    = error_style,
    error_title    = error_title,
    error          = error,
    prompt_title   = prompt_title,
    prompt         = prompt,
    ...            = ...
  )
}


# visibility --------------------------------------------------------------

#' Get/set worksheet visible state in a workbook
#'
#' Get and set worksheet visible state. This allows to hide worksheets from the workbook.
#' The visibility of a worksheet can either be  "visible", "hidden", or "veryHidden".
#' You can set this when creating a worksheet with `wb_add_worksheet(visible = FALSE)`
#'
#' @return
#' * `wb_set_sheet_visibility`: The Workbook object, invisibly.
#' * `wb_get_sheet_visibility()`: A character vector of the worksheet visibility value
#' @examples
#'
#' wb <- wb_workbook()
#' wb$add_worksheet(sheet = "S1", visible = FALSE)
#' wb$add_worksheet(sheet = "S2", visible = TRUE)
#' wb$add_worksheet(sheet = "S3", visible = FALSE)
#'
#' wb$get_sheet_visibility()
#' wb$set_sheet_visibility(1, TRUE)         ## show sheet 1
#' wb$set_sheet_visibility(2, FALSE)        ## hide sheet 2
#' wb$set_sheet_visibility(3, "hidden")     ## hide sheet 3
#' wb$set_sheet_visibility(3, "veryHidden") ## hide sheet 3 from UI
#' @name sheet_visibility-wb
NULL

#' @rdname sheet_visibility-wb
#' @param wb A `wbWorkbook` object
#' @export
wb_get_sheet_visibility <- function(wb) {
  assert_workbook(wb)
  wb$get_sheet_visibility()
}

#' @rdname sheet_visibility-wb
#' @param sheet Worksheet identifier
#' @param value a logical/character vector the same length as sheet,
#'   if providing a character vector, you can provide any of "hidden", "visible", or "veryHidden"
#' @export
wb_set_sheet_visibility <- function(wb, sheet = current_sheet(), value) {
  assert_workbook(wb)
  wb$clone()$set_sheet_visibility(sheet = sheet, value = value)
}


#' Add a page break to a worksheet
#'
#' Insert page breaks into a worksheet
#'
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param row,col Either a row number of column number.  One must be `NULL`
#' @export
#' @seealso [wb_add_worksheet()]
#' @examples
#' wb <- wb_workbook()
#' wb$add_worksheet("Sheet 1")
#' wb$add_data(sheet = 1, x = iris)
#'
#' wb$add_page_break(sheet = 1, row = 10)
#' wb$add_page_break(sheet = 1, row = 20)
#' wb$add_page_break(sheet = 1, col = 2)
#'
#' ## In Excel: View tab -> Page Break Preview
wb_add_page_break <- function(wb, sheet = current_sheet(), row = NULL, col = NULL) {
  assert_workbook(wb)
  wb$clone(deep = TRUE)$add_page_break(sheet = sheet, row = row, col = col)
}


#' List Excel tables in a worksheet
#'
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @return A character vector of table names on the specified sheet
#' @examples
#'
#' wb <- wb_workbook()
#' wb$add_worksheet(sheet = "Sheet 1")
#' wb$add_data_table(x = iris)
#' wb$add_data_table(x = mtcars, table_name = "mtcars", start_col = 10)
#'
#' wb$get_tables(sheet = "Sheet 1")
#' @export
wb_get_tables <- function(wb, sheet = current_sheet()) {
  assert_workbook(wb)
  wb$clone()$get_tables(sheet = sheet)
}



#' Remove a data table from a worksheet
#'
#' Remove Excel tables in a workbook using its name.
#'
#' @param wb A Workbook object
#' @param sheet A name or index of a worksheet
#' @param table Name of table to remove. Use [wb_get_tables()] to view the
#'   tables present in the worksheet.
#' @param remove_data Default `TRUE`. If `FALSE`, will only remove the data table attributes
#'   but will keep the data in the worksheet.
#' @return The `wbWorkbook`, invisibly
#' @examples
#' wb <- wb_workbook()
#' wb$add_worksheet(sheet = "Sheet 1")
#' wb$add_worksheet(sheet = "Sheet 2")
#' wb$add_data_table(sheet = "Sheet 1", x = iris, table_name = "iris")
#' wb$add_data_table(sheet = 1, x = mtcars, table_name = "mtcars", start_col = 10)
#'
#' ## delete worksheet removes table objects
#' wb <- wb_remove_worksheet(wb, sheet = 1)
#'
#' wb$add_data_table(sheet = 1, x = iris, table_name = "iris")
#' wb$add_data_table(sheet = 1, x = mtcars, table_name = "mtcars", start_col = 10)
#'
#' ## wb_remove_tables() deletes table object and all data
#' wb_get_tables(wb, sheet = 1)
#' wb$remove_tables(sheet = 1, table = "iris")
#' wb$add_data_table(sheet = 1, x = iris, table_name = "iris")
#'
#' wb_get_tables(wb, sheet = 1)
#' wb$remove_tables(sheet = 1, table = "iris")
#' @export
wb_remove_tables <- function(wb, sheet = current_sheet(), table, remove_data = TRUE) {
  assert_workbook(wb)
  wb$clone()$remove_tables(sheet = sheet, table = table, remove_data = remove_data)
}


# grouping ----------------------------------------------------------------

#' Group rows and columns in a worksheet
#'
#' Group a selection of rows or cols
#'
#' @details If row was previously hidden, it will now be shown.
#'
#' @param wb A `wbWorkbook` object
#' @param sheet A name or index of a worksheet
#' @param rows,cols Indices of rows and columns to group
#' @param collapsed If `TRUE` the grouped columns are collapsed
#' @param levels levels
#' @family workbook wrappers
#' @family worksheet content functions
#' @examples
#' # create matrix
#' t1 <- AirPassengers
#' t2 <- do.call(cbind, split(t1, cycle(t1)))
#' dimnames(t2) <- dimnames(.preformat.ts(t1))
#'
#' wb <- wb_workbook()
#' wb$add_worksheet("AirPass")
#' wb$add_data("AirPass", t2, row_names = TRUE)
#'
#' # groups will always end on/show the last row. in the example 1950, 1955, and 1960
#' wb <- wb_group_rows(wb, "AirPass", 2:3, collapsed = TRUE) # group years < 1950
#' wb <- wb_group_rows(wb, "AirPass", 4:8, collapsed = TRUE) # group years 1951-1955
#' wb <- wb_group_rows(wb, "AirPass", 9:13)                  # group years 1956-1960
#'
#' wb <- wb_group_cols(wb, "AirPass", 2:4, collapsed = TRUE)
#' wb <- wb_group_cols(wb, "AirPass", 5:7, collapsed = TRUE)
#' wb <- wb_group_cols(wb, "AirPass", 8:10, collapsed = TRUE)
#' wb <- wb_group_cols(wb, "AirPass", 11:13)
#'
#' @name grouping-wb
#' @family workbook wrappers
NULL

#' @export
#' @rdname grouping-wb
wb_group_cols <- function(wb, sheet = current_sheet(), cols, collapsed = FALSE, levels = NULL) {
  assert_workbook(wb)
  wb$clone()$group_cols(
    sheet     = sheet,
    cols      = cols,
    collapsed = collapsed,
    levels    = levels
  )
}

#' @export
#' @rdname grouping-wb
wb_ungroup_cols <- function(wb, sheet = current_sheet(), cols) {
  assert_workbook(wb)
  wb$clone()$ungroup_cols(sheet = sheet, cols = cols)
}


#' @export
#' @rdname grouping-wb
#' @examples
#' ### create grouping levels
#' grp_rows <- list(
#'   "1" = seq(2, 3),
#'   "2" = seq(4, 8),
#'   "3" = seq(9, 13)
#' )
#'
#' grp_cols <- list(
#'   "1" = seq(2, 4),
#'   "2" = seq(5, 7),
#'   "3" = seq(8, 10),
#'   "4" = seq(11, 13)
#' )
#'
#' wb <- wb_workbook()
#' wb$add_worksheet("AirPass")
#' wb$add_data("AirPass", t2, row_names = TRUE)
#'
#' wb$group_cols("AirPass", cols = grp_cols)
#' wb$group_rows("AirPass", rows = grp_rows)
wb_group_rows <- function(wb, sheet = current_sheet(), rows, collapsed = FALSE, levels = NULL) {
  assert_workbook(wb)
  wb$clone()$group_rows(
    sheet     = sheet,
    rows      = rows,
    collapsed = collapsed,
    levels    = levels
  )
}

#' @export
#' @rdname grouping-wb
wb_ungroup_rows <- function(wb, sheet = current_sheet(), rows) {
  assert_workbook(wb)
  wb$clone()$ungroup_rows(sheet = sheet, rows = rows)
}


# creators ----------------------------------------------------------------

#' Modify workbook properties
#'
#' This function is useful for workbooks that are loaded. It can be used to set the
#' workbook `title`, `subject` and `category` field. Use [wb_workbook()]
#' to easily set these properties with a new workbook.
#'
#' To set properties, the following XML core properties are used.
#' - title = dc:title
#' - subject = dc:subject
#' - creator = dc:creator
#' - keywords = cp:keywords
#' - comments = dc:description
#' - modifier = cp:lastModifiedBy
#' - datetime_created = dcterms:created
#' - datetime_modified = dcterms:modified
#' - category = cp:category
#'
#' In addition, manager and company are used.
#' @name properties-wb
#' @param wb A Workbook object
#' @param modifier A character string indicating who was the last person to modify the workbook
#' @param custom A named vector of custom properties added to the workbook
#' @seealso [wb_workbook()]
#' @inheritParams wb_workbook
#' @return A wbWorkbook object, invisibly.
#' @export
#'
#' @examples
#' file <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
#' wb <- wb_load(file)
#' wb$get_properties()
#'
#' # Add a title to properties
#' wb$set_properties(title = "my title")
#' wb$get_properties()
wb_get_properties <- function(wb) {
  assert_workbook(wb)
  wb$get_properties()
}

#' @rdname properties-wb
#' @export
wb_set_properties <- function(wb, creator = NULL, title = NULL, subject = NULL, category = NULL, datetime_created = Sys.time(), modifier = NULL, keywords = NULL, comments = NULL, manager = NULL, company = NULL, custom = NULL) {
  assert_workbook(wb)
  wb$clone()$set_properties(
    creator           = creator,
    title             = title,
    subject           = subject,
    category          = category,
    datetime_created  = datetime_created,
    modifier          = modifier,
    keywords          = keywords,
    comments          = comments,
    manager           = manager,
    company           = company,
    custom            = custom
  )
}

#' wb get and apply MIP section
#'
#' Read sensitivity labels from files and apply them to workbooks
#'
#' @details
#' The MIP section is a special user-defined XML section that is used to create
#' sensitivity labels in workbooks. It consists of a series of XML property
#' nodes that define the sensitivity label. This XML string cannot be created
#' and it is necessary to first load a workbook with a suitable sensitivity
#' label. Once the workbook is loaded, the string `fmips <- wb_get_mips(wb)`
#' can be extracted. This xml string can later be assigned to an
#' `options("openxlsx2.mips_xml_string" = fmips)` option.
#'
#' The sensitivity label can then be assigned with `wb_add_mips(wb)`. If no xml
#' string is passed, the MIP section is taken from the option. This should make
#' it easier for users to read the section from a specific workbook, save it to
#' a file or string and copy it to an option via the .Rprofile.
#'
#' @param wb a workbook
#' @param xml a mips string obtained from [wb_get_mips()] or a global option "openxlsx2.mips_xml_string"
#' @returns the workbook invisible ([wb_add_mips()]) or the xml string ([wb_get_mips()])
#' @export
wb_add_mips <- function(wb, xml = NULL) {
  assert_workbook(wb)
  wb$clone()$set_properties(custom = xml)
}

#' @param single_xml option to define if the string should be exported as single string. helpful if storing as option is desired.
#' @param quiet option to print a MIP section name. This is not always a human readable string.
#' @rdname wb_add_mips
#' @export
wb_get_mips <- function(wb, single_xml = TRUE, quiet = TRUE) {
  assert_workbook(wb)
  wb$get_mips(single_xml = single_xml, quiet = quiet)
}

#' Modify creators of a workbook
#'
#' Modify and get workbook creators
#'
#' @param wb A `wbWorkbook` object
#' @examples
#'
#' # workbook made with default creator (see [wbWorkbook])
#' wb <- wb_workbook()
#' wb_get_creators(wb)
#'
#' # add a new creator (assuming "test" isn't your default creator)
#' wb <- wb_add_creators(wb, "test")
#' wb_get_creators(wb)
#'
#' # remove the creator (should be the same as before)
#' wb <- wb_remove_creators(wb, "test")
#' wb_get_creators(wb)
#'
#' @return
#' * `wb_set_creators()`, `wb_add_creators()`, and `wb_remove_creators()` return
#'    the `wbWorkbook` object
#' * `wb_get_creators()` returns a `character` vector of creators
#'
#' @name creators-wb
#' @family workbook wrappers
NULL

#' @rdname creators-wb
#' @export
#' @param creators A character vector of names
wb_add_creators <- function(wb, creators) {
  assert_workbook(wb)
  wb$clone()$add_creators(creators)
}

#' @rdname creators-wb
#' @export
wb_set_creators <- function(wb, creators) {
  assert_workbook(wb)
  wb$clone()$set_creators(creators)
}

#' @rdname creators-wb
#' @export
wb_remove_creators <- function(wb, creators) {
  assert_workbook(wb)
  wb$clone()$remove_creators(creators)
}

#' @rdname creators-wb
#' @export
wb_get_creators <- function(wb) {
  assert_workbook(wb)
  strsplit(wb$get_properties()[["creator"]], ";")[[1]]
}


# names -------------------------------------------------------------------

#' Get / Set worksheet names for a workbook
#'
#' Gets / Sets the worksheet names for a [wbWorkbook] object.
#'
#' This only changes the sheet name as shown in spreadsheet software
#' and will not alter it elsewhere. Not in formulas, chart references,
#' named regions, pivot tables or anywhere else.
#'
#' @param wb A [wbWorkbook] object
#' @param old The name (or index) of the old sheet name. If `NULL` will assume
#'   all worksheets are to be renamed.
#' @param new The name of the new sheet
#' @name sheet_names-wb
#' @returns
#'   * `set_`: The `wbWorkbook` object.
#'   * `get_`: A named character vector of sheet names in order. The
#'   names represent the original value of the worksheet prior to any character
#'   substitutions.
#'
NULL
#' @rdname sheet_names-wb
#' @export
wb_set_sheet_names <- function(wb, old = NULL, new) {
  assert_workbook(wb)
  wb$clone()$set_sheet_names(old = old, new = new)
}
#' @rdname sheet_names-wb
#' @param escape Should the xml special characters be escaped?
#' @export
wb_get_sheet_names <- function(wb, escape = FALSE) {
  assert_workbook(wb)
  wb$get_sheet_names(escape = escape)
}

# others? -----------------------------------------------------------------

#' Modify author in the metadata of a workbook
#'
#' Just a wrapper of `wb$set_last_modified_by()`
#'
#' @param wb A workbook object
#' @param name A string object with the name of the LastModifiedBy-User
#' @param ... additional arguments
#' @family workbook wrappers
#' @export
#' @examples
#' wb <- wb_workbook()
#' wb_set_last_modified_by(wb, "test")
wb_set_last_modified_by <- function(wb, name, ...) {
  if (missing(name)) name <- substitute()
  assert_workbook(wb)
  wb$clone()$set_last_modified_by(name, ...)
}

#' Insert an image into a worksheet
#'
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param dims Dimensions where to plot. Default absolute anchor, single cell (eg. "A1")
#'   oneCellAnchor, cell range (eg. "A1:D4") twoCellAnchor
#' @param file An image file. Valid file types are:` "jpeg"`, `"png"`, `"bmp"`
#' @param width Width of figure.
#' @param height Height of figure.
#' @param row_offset offset vector for one or two cell anchor within cell (row)
#' @param col_offset offset vector for one or two cell anchor within cell (column)
#' @param units Units of width and height. Can be `"in"`, `"cm"` or `"px"`
#' @param dpi Image resolution used for conversion between units.
#' @param ... additional arguments
#' @seealso [wb_add_chart_xml()] [wb_add_drawing()] [wb_add_mschart()] [wb_add_plot()]
#' @export
#' @examples
#' img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")
#'
#' wb <- wb_workbook()$
#'   add_worksheet()$
#'   add_image("Sheet 1", dims = "C5", file = img, width = 6, height = 5)$
#'   add_worksheet()$
#'   add_image(dims = "B2", file = img)$
#'   add_worksheet()$
#'   add_image(dims = "G3", file = img, width = 15, height = 12, units = "cm")
wb_add_image <- function(
  wb,
  sheet      = current_sheet(),
  dims       = "A1",
  file,
  width      = 6,
  height     = 3,
  row_offset = 0,
  col_offset = 0,
  units      = "in",
  dpi        = 300,
  ...
) {
  assert_workbook(wb)
  wb$clone()$add_image(
    sheet      = sheet,
    dims       = dims,
    file       = file,
    width      = width,
    height     = height,
    row_offset = row_offset,
    col_offset = col_offset,
    units      = units,
    dpi        = dpi,
    ...
  )
}


#' Add a chart XML to a worksheet
#'
#' @param wb a workbook
#' @param sheet the sheet on which the graph will appear
#' @param dims the dimensions where the sheet will appear
#' @param xml chart xml
#' @param col_offset,row_offset positioning
#' @param ... additional arguments
#' @seealso [wb_add_drawing()] [wb_add_image()] [wb_add_mschart()] [wb_add_plot()]
#' @export
wb_add_chart_xml <- function(
  wb,
  sheet      = current_sheet(),
  dims       = NULL,
  xml,
  col_offset = 0,
  row_offset = 0,
  ...
) {
  assert_workbook(wb)
  wb$clone()$add_chart_xml(
    sheet      = sheet,
    xml        = xml,
    dims       = dims,
    col_offset = col_offset,
    row_offset = row_offset,
    ...        = ...
  )
}


#' Remove all values in a worksheet
#'
#' Remove content of a worksheet completely, or a region if specifying `dims`.
#'
#' @param wb A Workbook object
#' @param sheet sheet to clean
#' @param dims spreadsheet dimensions (optional)
#' @param numbers remove all numbers
#' @param characters remove all characters
#' @param styles remove all styles
#' @param merged_cells remove all merged_cells
#' @return A Workbook object
#' @export
wb_clean_sheet <- function(
    wb,
    sheet        = current_sheet(),
    dims         = NULL,
    numbers      = TRUE,
    characters   = TRUE,
    styles       = TRUE,
    merged_cells = TRUE
) {
  assert_workbook(wb)
  wb$clone(deep = TRUE)$clean_sheet(
    sheet        = sheet,
    dims         = dims,
    numbers      = numbers,
    characters   = characters,
    styles       = styles,
    merged_cells = merged_cells
  )
}

#' Preview a workbook in a spreadsheet software
#'
#' You can also use the shorter `wb$open()` as a replacement.
#' To open xlsx files, see [xl_open()].
#'
#' @param wb a [wbWorkbook] object
#' @export
wb_open <- function(wb) {
  assert_workbook(wb)
  wb$open()
}

#' Set the default style in a workbook
#'
#' wb wrapper to add style to workbook
#'
#' @param wb A workbook
#' @param style style xml character, created by a `create_*()` function.
#' @param style_name style name used optional argument
#' @return The `wbWorkbook` object, invisibly.
#' @seealso
#' * [create_border()]
#' * [create_cell_style()]
#' * [create_dxfs_style()]
#' * [create_fill()]
#' * [create_font()]
#' * [create_numfmt()]
#' @family workbook styling functions
#' @examples
#' yellow_f <- wb_color(hex = "FF9C6500")
#' yellow_b <- wb_color(hex = "FFFFEB9C")
#'
#' yellow <- create_dxfs_style(font_color = yellow_f, bg_fill = yellow_b)
#' wb <- wb_workbook() %>% wb_add_style(yellow)
#' @export
wb_add_style <- function(wb, style = NULL, style_name = NULL) {
  assert_workbook(wb)
  # deparse this name, otherwise it will remain "style"
  if (is.null(style_name)) style_name <- deparse(substitute(style))
  wb$clone()$add_style(style, style_name)
}

#' Apply styling to a cell region
#'
#' Setting a style across only impacts cells that are not yet part of a workbook. The effect is similar to setting the cell style for all cells in a row independently, though much quicker and less memory consuming.
#'
#' @name wb_cell_style
#' @param wb A `wbWorkbook` object
#' @param sheet sheet
#' @param dims A cell range in the worksheet
#' @family styles
#' @examples
#' # set a style in b1
#' wb <- wb_workbook()$add_worksheet()$
#'   add_numfmt(dims = "B1", numfmt = "#,0")
#'
#' # get style from b1 to assign it to a1
#' numfmt <- wb$get_cell_style(dims = "B1")
#'
#' # assign style to a1
#' wb$set_cell_style(dims = "A1", style = numfmt)
#' @return A Workbook object
#' @export
wb_get_cell_style <- function(wb, sheet = current_sheet(), dims) {
  assert_workbook(wb)
  wb$clone()$get_cell_style(sheet, dims)
}

#' @rdname wb_cell_style
#' @param style A style or a cell with a certain style
#' @export
wb_set_cell_style <- function(wb, sheet = current_sheet(), dims, style) {
  assert_workbook(wb)
  # needs deep clone for nested calls as in styles vignette copy cell styles
  wb$clone(deep = TRUE)$set_cell_style(sheet, dims, style)
}

#' @rdname wb_cell_style
#' @param cols The columns the style will be applied to, either "A:D" or 1:4
#' @param rows The rows the style will be applied to
#' @examples
#' wb <- wb_workbook() %>%
#'   wb_add_worksheet() %>%
#'   wb_add_fill(dims = "C3", color = wb_color("yellow")) %>%
#'   wb_set_cell_style_across(style = "C3", cols = "C:D", rows = 3:4)
#' @export
wb_set_cell_style_across <- function(wb, sheet = current_sheet(), style, cols = NULL, rows = NULL) {
  assert_workbook(wb)
  wb$clone(deep = TRUE)$set_cell_style_across(sheet = sheet, style = style, cols = cols, rows = rows)
}

#' Modify borders in a cell region of a worksheet
#'
#' wb wrapper to create borders for cell regions.
#' @param wb A `wbWorkbook`
#' @param sheet A worksheet
#' @param dims Cell range in the worksheet e.g. "A1", "A1:A5", "A1:H5"
#' @param bottom_color,left_color,right_color,top_color,inner_hcolor,inner_vcolor
#'   a color, either something openxml knows or some RGB color
#' @param left_border,right_border,top_border,bottom_border,inner_hgrid,inner_vgrid
#'   the border style, if `NULL` no border is drawn.
#'   See [create_border()] for possible border styles
#' @param ... additional arguments
#' @seealso [create_border()]
#' @examples
#' wb <- wb_workbook() %>% wb_add_worksheet("S1") %>%  wb_add_data("S1", mtcars)
#' wb <- wb_add_border(wb, 1, dims = "A1:K1",
#'  left_border = NULL, right_border = NULL,
#'  top_border = NULL, bottom_border = "double")
#' wb <- wb_add_border(wb, 1, dims = "A5",
#'  left_border = "dotted", right_border = "dotted",
#'  top_border = "hair", bottom_border = "thick")
#' wb <- wb_add_border(wb, 1, dims = "C2:C5")
#' wb <- wb_add_border(wb, 1, dims = "G2:H3")
#'
#' wb <- wb_add_border(wb, 1, dims = "G12:H13",
#'  left_color = wb_color(hex = "FF9400D3"), right_color = wb_color(hex = "FF4B0082"),
#'  top_color = wb_color(hex = "FF0000FF"), bottom_color = wb_color(hex = "FF00FF00"))
#' wb <- wb_add_border(wb, 1, dims = "A20:C23")
#' wb <- wb_add_border(wb, 1, dims = "B12:D14",
#'  left_color = wb_color(hex = "FFFFFF00"), right_color = wb_color(hex = "FFFF7F00"),
#'  bottom_color = wb_color(hex = "FFFF0000"))
#' wb <- wb_add_border(wb, 1, dims = "D28:E28")
#'
#' # With chaining
#'
#' wb <- wb_workbook()
#' wb$add_worksheet("S1")$add_data("S1", mtcars)
#' wb$add_border(1, dims = "A1:K1",
#'  left_border = NULL, right_border = NULL,
#'  top_border = NULL, bottom_border = "double")
#' wb$add_border(1, dims = "A5",
#'  left_border = "dotted", right_border = "dotted",
#'  top_border = "hair", bottom_border = "thick")
#' wb$add_border(1, dims = "C2:C5")
#' wb$add_border(1, dims = "G2:H3")
#' wb$add_border(1, dims = "G12:H13",
#'  left_color = wb_color(hex = "FF9400D3"), right_color = wb_color(hex = "FF4B0082"),
#'  top_color = wb_color(hex = "FF0000FF"), bottom_color = wb_color(hex = "FF00FF00"))
#' wb$add_border(1, dims = "A20:C23")
#' wb$add_border(1, dims = "B12:D14",
#'  left_color = wb_color(hex = "FFFFFF00"), right_color = wb_color(hex = "FFFF7F00"),
#'  bottom_color = wb_color(hex = "FFFF0000"))
#' wb$add_border(1, dims = "D28:E28")
#' # if (interactive()) wb$open()
#'
#' wb <- wb_workbook()
#' wb$add_worksheet("S1")$add_data("S1", mtcars)
#' wb$add_border(1, dims = "A2:K33", inner_vgrid = "thin", inner_vcolor = c(rgb="FF808080"))
#' @family styles
#' @export
wb_add_border <- function(
    wb,
    sheet          = current_sheet(),
    dims           = "A1",
    bottom_color   = wb_color(hex = "FF000000"),
    left_color     = wb_color(hex = "FF000000"),
    right_color    = wb_color(hex = "FF000000"),
    top_color      = wb_color(hex = "FF000000"),
    bottom_border  = "thin",
    left_border    = "thin",
    right_border   = "thin",
    top_border     = "thin",
    inner_hgrid    = NULL,
    inner_hcolor   = NULL,
    inner_vgrid    = NULL,
    inner_vcolor   = NULL,
    ...
) {
  assert_workbook(wb)
  wb$clone()$add_border(
    sheet         = sheet,
    dims          = dims,
    bottom_color  = bottom_color,
    left_color    = left_color,
    right_color   = right_color,
    top_color     = top_color,
    bottom_border = bottom_border,
    left_border   = left_border,
    right_border  = right_border,
    top_border    = top_border,
    inner_hgrid   = inner_hgrid,
    inner_hcolor  = inner_hcolor,
    inner_vgrid   = inner_vgrid,
    inner_vcolor  = inner_vcolor,
    ...           = ...
  )

}

#' Modify the background fill color in a cell region
#'
#' Add fill to a cell region.
#'
#' @param wb a workbook
#' @param sheet the worksheet
#' @param dims the cell range
#' @param color the colors to apply, e.g. yellow: wb_color(hex = "FFFFFF00")
#' @param pattern various default "none" but others are possible:
#'  "solid", "mediumGray", "darkGray", "lightGray", "darkHorizontal",
#'  "darkVertical", "darkDown", "darkUp", "darkGrid", "darkTrellis",
#'  "lightHorizontal", "lightVertical", "lightDown", "lightUp", "lightGrid",
#'  "lightTrellis", "gray125", "gray0625"
#' @param gradient_fill a gradient fill xml pattern.
#' @param every_nth_col which col should be filled
#' @param every_nth_row which row should be filled
#' @param ... ...
#' @examples
#' wb <- wb_workbook() %>% wb_add_worksheet("S1") %>% wb_add_data("S1", mtcars)
#' wb <- wb %>% wb_add_fill("S1", dims = "D5:J23", color = wb_color(hex = "FFFFFF00"))
#' wb <- wb %>% wb_add_fill("S1", dims = "B22:D27", color = wb_color(hex = "FF00FF00"))
#'
#' wb <- wb %>%  wb_add_worksheet("S2") %>% wb_add_data("S2", mtcars)
#'
#' gradient_fill1 <- '<gradientFill degree="90">
#' <stop position="0"><color rgb="FF92D050"/></stop>
#' <stop position="1"><color rgb="FF0070C0"/></stop>
#' </gradientFill>'
#' wb <- wb %>% wb_add_fill("S2", dims = "A2:K5", gradient_fill = gradient_fill1)
#'
#' gradient_fill2 <- '<gradientFill type="path" left="0.2" right="0.8" top="0.2" bottom="0.8">
#' <stop position="0"><color theme="0"/></stop>
#' <stop position="1"><color theme="4"/></stop>
#' </gradientFill>'
#' wb <- wb %>% wb_add_fill("S2", dims = "A7:K10", gradient_fill = gradient_fill2)
#' @return The `wbWorkbook` object, invisibly
#' @family styles
#' @export
wb_add_fill <- function(
    wb,
    sheet         = current_sheet(),
    dims          = "A1",
    color         = wb_color(hex = "FFFFFF00"),
    pattern       = "solid",
    gradient_fill = "",
    every_nth_col = 1,
    every_nth_row = 1,
    ...
) {
  assert_workbook(wb)
  wb$clone()$add_fill(
    sheet         = sheet,
    dims          = dims,
    color         = color,
    pattern       = pattern,
    gradient_fill = gradient_fill,
    every_nth_col = every_nth_col,
    every_nth_row = every_nth_row,
    ...           = ...
  )
}

#' Modify font in a cell region
#'
#' Modify the font in a cell region with more precision
#' You can specify the font in a cell with other cell styling functions,
#' but `wb_add_font()` gives you more control.
#'
#' `wb_add_font()` provides all the options openxml accepts for a font node,
#' not all have to be set. Usually `name`, `size` and `color` should be what the user wants.
#' @param wb A Workbook object
#' @param sheet the worksheet
#' @param dims the cell range
#' @param name Font name: default "Aptos Narrow"
#' @param color An object created by [wb_color()]
#' @param size Font size: default "11",
#' @param bold bold, "single" or "double", default: ""
#' @param italic italic
#' @param outline outline
#' @param strike strike
#' @param underline underline
#' @param family font family
#' @param charset charset
#' @param condense condense
#' @param scheme font scheme
#' @param shadow shadow
#' @param extend extend
#' @param vert_align vertical alignment
#' @param ... ...
#' @examples
#'  wb <- wb_workbook() %>% wb_add_worksheet("S1") %>% wb_add_data("S1", mtcars)
#'  wb %>% wb_add_font("S1", "A1:K1", name = "Arial", color = wb_color(theme = "4"))
#' # With chaining
#'  wb <- wb_workbook()$add_worksheet("S1")$add_data("S1", mtcars)
#'  wb$add_font("S1", "A1:K1", name = "Arial", color = wb_color(theme = "4"))
#' @return A `wbWorkbook`, invisibly
#' @family styles
#' @export
wb_add_font <- function(
      wb,
      sheet      = current_sheet(),
      dims       = "A1",
      name       = "Aptos Narrow",
      color      = wb_color(hex = "FF000000"),
      size       = "11",
      bold       = "",
      italic     = "",
      outline    = "",
      strike     = "",
      underline  = "",
      # fine tuning
      charset    = "",
      condense   = "",
      extend     = "",
      family     = "",
      scheme     = "",
      shadow     = "",
      vert_align = "",
      ...
) {
  assert_workbook(wb)
  wb$clone()$add_font(
    sheet      = sheet,
    dims       = dims,
    name       = name,
    color      = color,
    size       = size,
    bold       = bold,
    italic     = italic,
    outline    = outline,
    strike     = strike,
    underline  = underline,
    # fine tuning
    charset    = charset,
    condense   = condense,
    extend     = extend,
    family     = family,
    scheme     = scheme,
    shadow     = shadow,
    vert_align = vert_align,
    ...        = ...
  )
}

#' Modify number formatting in a cell region
#'
#' Add number formatting to a cell region. You can use a number format created
#' by [create_numfmt()].
#'
#' The list of number formats ID is located in the **Details** section of [create_cell_style()].
#' @param wb A Workbook
#' @param sheet the worksheet
#' @param dims the cell range
#' @param numfmt either an id or a character
#' @examples
#' wb <- wb_workbook() %>% wb_add_worksheet("S1") %>% wb_add_data("S1", mtcars)
#' wb %>% wb_add_numfmt("S1", dims = "F1:F33", numfmt = "#.0")
#' # Chaining
#' wb <- wb_workbook()$add_worksheet("S1")$add_data("S1", mtcars)
#' wb$add_numfmt("S1", "A1:A33", numfmt = 1)
#' @return The `wbWorkbook` object, invisibly.
#' @family styles
#' @export

wb_add_numfmt <- function(
    wb,
    sheet = current_sheet(),
    dims  = "A1",
    numfmt
) {
  assert_workbook(wb)
  wb$clone()$add_numfmt(
    sheet  = sheet,
    dims   = dims,
    numfmt = numfmt
  )
}

#' Modify the style in a cell region
#'
#' Add cell style to a cell region
#' @param wb a workbook
#' @param sheet the worksheet
#' @param dims the cell range
#' @param ext_lst extension list something like `<extLst>...</extLst>`
#' @param hidden logical cell is hidden
#' @param horizontal align content horizontal ('left', 'center', 'right')
#' @param indent logical indent content
#' @param justify_last_line logical justify last line
#' @param locked logical cell is locked
#' @param pivot_button unknown
#' @param quote_prefix unknown
#' @param reading_order reading order left to right
#' @param relative_indent relative indentation
#' @param shrink_to_fit logical shrink to fit
#' @param text_rotation degrees of text rotation
#' @param vertical vertical alignment of content ('top', 'center', 'bottom')
#' @param wrap_text wrap text in cell
## alignments
#' @param apply_alignment logical apply alignment
#' @param apply_border logical apply border
#' @param apply_fill logical apply fill
#' @param apply_font logical apply font
#' @param apply_number_format logical apply number format
#' @param apply_protection logical apply protection
## ids
#' @param border_id border ID to apply
#' @param fill_id fill ID to apply
#' @param font_id font ID to apply
#' @param num_fmt_id number format ID to apply
#' @param xf_id xf ID to apply
#' @param ... additional arguments
#' @examples
#' wb <- wb_workbook() %>%
#'   wb_add_worksheet("S1") %>%
#'   wb_add_data("S1", x = mtcars)
#'
#' wb %>%
#'   wb_add_cell_style(
#'     dims = "A1:K1",
#'     text_rotation = "45",
#'     horizontal = "center",
#'     vertical = "center",
#'     wrap_text = "1"
#' )
#' # Chaining
#' wb <- wb_workbook()$add_worksheet("S1")$add_data(x = mtcars)
#' wb$add_cell_style(dims = "A1:K1",
#'                   text_rotation = "45",
#'                   horizontal = "center",
#'                   vertical = "center",
#'                   wrap_text = "1")
#' @return The `wbWorkbook` object, invisibly
#' @family styles
#' @export
wb_add_cell_style <- function(
    wb,
    sheet               = current_sheet(),
    dims                = "A1",
    apply_alignment     = NULL,
    apply_border        = NULL,
    apply_fill          = NULL,
    apply_font          = NULL,
    apply_number_format = NULL,
    apply_protection    = NULL,
    border_id           = NULL,
    ext_lst             = NULL,
    fill_id             = NULL,
    font_id             = NULL,
    hidden              = NULL,
    horizontal          = NULL,
    indent              = NULL,
    justify_last_line   = NULL,
    locked              = NULL,
    num_fmt_id          = NULL,
    pivot_button        = NULL,
    quote_prefix        = NULL,
    reading_order       = NULL,
    relative_indent     = NULL,
    shrink_to_fit       = NULL,
    text_rotation       = NULL,
    vertical            = NULL,
    wrap_text           = NULL,
    xf_id               = NULL,
    ...
) {
  assert_workbook(wb)
  wb$clone()$add_cell_style(
    sheet               = sheet,
    dims                = dims,
    apply_alignment     = apply_alignment,
    apply_border        = apply_border,
    apply_fill          = apply_fill,
    apply_font          = apply_font,
    apply_number_format = apply_number_format,
    apply_protection    = apply_protection,
    border_id           = border_id,
    ext_lst             = ext_lst,
    fill_id             = fill_id,
    font_id             = font_id,
    hidden              = hidden,
    horizontal          = horizontal,
    indent              = indent,
    justify_last_line   = justify_last_line,
    locked              = locked,
    num_fmt_id          = num_fmt_id,
    pivot_button        = pivot_button,
    quote_prefix        = quote_prefix,
    reading_order       = reading_order,
    relative_indent     = relative_indent,
    shrink_to_fit       = shrink_to_fit,
    text_rotation       = text_rotation,
    vertical            = vertical,
    wrap_text           = wrap_text,
    xfId                = xf_id,
    ...                 = ...
  )
}

#' Apply styling to a cell region with a named style
#'
#' Set the styling to a named style for a cell region. Use [wb_add_cell_style()]
#' to style a cell region with custom parameters.
#' A named style is the one in spreadsheet software, like "Normal", "Warning".
#' @param wb A `wbWorkbook` object
#' @param sheet A worksheet
#' @param dims A cell range
#' @param name The named style name.
#' @family styles
#' @param font_name,font_size optional else the default of the theme
#' @return The `wbWorkbook`, invisibly
#' @export
wb_add_named_style <- function(
    wb,
    sheet = current_sheet(),
    dims = "A1",
    name = "Normal",
    font_name = NULL,
    font_size = NULL
) {
  assert_workbook(wb)
  assert_class(name, "character")
  wb$clone()$add_named_style(
    sheet = sheet,
    dims = dims,
    name = name,
    font_name = font_name,
    font_size = font_size
  )
}

#' Set a dxfs styling for the workbook
#'
#' These styles are used with conditional formatting and custom table styles.
#'
#' @param wb A Workbook object.
#' @param name the style name
#' @param font_name the font name
#' @param font_size the font size
#' @param font_color the font color (a `wb_color()` object)
#' @param num_fmt the number format
#' @param border logical if borders are applied
#' @param border_color the border color
#' @param border_style the border style
#' @param bg_fill any background fill
#' @param gradient_fill any gradient fill
#' @param text_bold logical if text is bold
#' @param text_italic logical if text is italic
#' @param text_underline logical if text is underlined
#' @param ... additional arguments passed to [create_dxfs_style()]
#' @family workbook styling functions
#' @return The Workbook object, invisibly
#' @examples
#' wb <- wb_workbook() %>%
#'   wb_add_worksheet() %>%
#'   wb_add_dxfs_style(
#'    name = "nay",
#'    font_color = wb_color(hex = "FF9C0006"),
#'    bg_fill = wb_color(hex = "FFFFC7CE")
#'   )
#' @export
wb_add_dxfs_style <- function(
  wb,
  name,
  font_name      = NULL,
  font_size      = NULL,
  font_color     = NULL,
  num_fmt        = NULL,
  border         = NULL,
  border_color   = wb_color(getOption("openxlsx2.borderColor", "black")),
  border_style   = getOption("openxlsx2.borderStyle", "thin"),
  bg_fill        = NULL,
  gradient_fill  = NULL,
  text_bold      = NULL,
  text_italic    = NULL,
  text_underline = NULL,
  ...
) {

  assert_workbook(wb)
  wb$clone()$add_dxfs_style(
    name           = name,
    font_name      = font_name,
    font_size      = font_size,
    font_color     = font_color,
    num_fmt        = num_fmt,
    border         = border,
    border_color   = border_color,
    border_style   = border_style,
    bg_fill        = bg_fill,
    gradient_fill  = gradient_fill,
    text_bold      = text_bold,
    text_italic    = text_italic,
    text_underline = text_underline,
    ...            = ...
  )

}

#' Add comment to worksheet
#'
#' @details
#' If applying a `comment` with a string, it will use [wb_comment()] default values. If additional background colors are applied, RGB colors should be provided, either as hex code or with builtin R colors. The alpha channel is ignored.
#'
#' @param wb A workbook object
#' @param sheet A worksheet of the workbook
#' @param dims Optional row and column as spreadsheet dimension, e.g. "A1"
#' @param comment A comment to apply to `dims` created by [wb_comment()], a string or a [fmt_txt()] object
#' @param ... additional arguments
#' @returns The Workbook object, invisibly.
#' @seealso [wb_comment()], [wb_add_thread()]
#' @keywords comments
#' @examples
#' wb <- wb_workbook()
#' wb$add_worksheet("Sheet 1")
#' # add a comment without author
#' c1 <- wb_comment(text = "this is a comment", author = "")
#' wb$add_comment(dims = "B10", comment = c1)
#' #' # Remove comment
#' wb$remove_comment(sheet = "Sheet 1", dims = "B10")
#' # Write another comment with author information
#' c2 <- wb_comment(text = "this is another comment", author = "Marco Polo", visible = TRUE)
#' wb$add_comment(sheet = 1, dims = "C10", comment = c2)
#' # Works with formatted text also.
#' formatted_text <- fmt_txt("bar", underline = TRUE)
#' wb$add_comment(dims = "B5", comment = formatted_text)
#' # With background color
#' wb$add_comment(dims = "B7", comment = formatted_text, color = wb_color("green"))
#' # With background image. File extension must be png or jpeg, not jpg?
#' tmp <- tempfile(fileext = ".png")
#' png(file = tmp, bg = "transparent")
#' plot(1:10)
#' rect(1, 5, 3, 7, col = "white")
#' dev.off()
#'
#' c1 <- wb_comment(text = "this is a comment", author = "", visible = TRUE)
#' wb$add_comment(dims = "B12", comment = c1, file = tmp)
#' @export
wb_add_comment <- function(
    wb,
    sheet   = current_sheet(),
    dims    = "A1",
    comment,
    ...
  ) {

  assert_workbook(wb)

  if (is.character(comment)) {
    comment <- wb_comment(text = comment, author = getOption("openxlsx2.creator"))
  }

  assert_comment(comment)

  wb$clone()$add_comment(
    sheet   = sheet,
    dims    = dims,
    comment = comment,
    ...     = ...
  )
}

#' @rdname wb_add_comment
#' @export
wb_get_comment <- function(
    wb,
    sheet = current_sheet(),
    dims  = NULL
) {

  assert_workbook(wb)

  wb$clone(deep = TRUE)$get_comment(
    sheet       = sheet,
    dims        = dims
  )
}

#' @rdname wb_add_comment
#' @export
wb_remove_comment <- function(
    wb,
    sheet      = current_sheet(),
    dims       = "A1",
    ...
  ) {

  assert_workbook(wb)

  wb$clone(deep = TRUE)$remove_comment(
    sheet       = sheet,
    dims        = dims,
    ...         = ...
  )
}

#' Helper for adding threaded comments
#'
#' Adds a person to a workbook, so that they can be the author of threaded
#' comments in a workbook with [wb_add_thread()]
#'
#' @name person-wb
#' @param wb a Workbook
#' @param name the name of the person to display.
#' @param id (optional) the display id
#' @param user_id (optional) the user id
#' @param provider_id (optional) the provider id
#' @seealso [wb_add_thread()]
#' @keywords comments
#' @export
wb_add_person <- function(
    wb,
    name        = NULL,
    id          = NULL,
    user_id     = NULL,
    provider_id = "None"
) {
  assert_workbook(wb)
  wb$clone()$add_person(
    name        = name,
    id          = id,
    user_id     = user_id,
    provider_id = provider_id
  )
}

#' @rdname person-wb
#' @export
wb_get_person <- function(wb, name = NULL) {
  assert_workbook(wb)
  wb$get_person(name)
}

#' Add threaded comments to a cell in a worksheet
#'
#' These functions allow adding thread comments to spreadsheets.
#' This is not yet supported by all spreadsheet software. A threaded comment must
#' be tied to a person created by [wb_add_person()].
#'
#' If a threaded comment is added, it needs a person attached to it.
#' The default is to create a person with provider id `"None"`.
#' Other providers are possible with specific values for `id` and `user_id`.
#' If you require the following, create a workbook via spreadsheet software load
#' it and get the values with [wb_get_person()]
#'
#' @param wb A workbook
#' @param sheet A worksheet
#' @param dims A cell
#' @param comment The text to add, a character vector.
#' @param person_id the person Id this should be added. The default is
#'   `getOption("openxlsx2.thread_id")` if set.
#' @param reply Is the comment a reply? (default `FALSE`)
#' @param resolve Should the comment be resolved? (default `FALSE`)
#' @seealso [wb_add_comment()] [`person-wb`]
#' @family worksheet content functions
#' @examples
#' wb <- wb_workbook()$add_worksheet()
#' # Add a person to the workbook.
#' wb$add_person(name = "someone who likes to edit workbooks")
#'
#' pid <- wb$get_person(name = "someone who likes to edit workbooks")$id
#'
#' # write a comment to a thread, reply to one and solve some
#' wb <- wb %>%
#'   wb_add_thread(dims = "A1", comment = "wow it works!", person_id = pid) %>%
#'   wb_add_thread(dims = "A2", comment = "indeed", person_id = pid, resolve = TRUE) %>%
#'   wb_add_thread(dims = "A1", comment = "so cool", person_id = pid, reply = TRUE)
#' @export
wb_add_thread <- function(
    wb,
    sheet     = current_sheet(),
    dims      = "A1",
    comment   = NULL,
    person_id,
    reply     = FALSE,
    resolve   = FALSE
) {
  # wb_add_thread now uses wb_comment internally. No change detected.
  # means that add_thread does not look at visibility. (I think it's fine.)
  if (missing(person_id)) {
    person_id <- substitute()
  }

  assert_workbook(wb)
  wb$clone()$add_thread(
    sheet     = sheet,
    dims      = dims,
    comment   = comment,
    person_id = person_id,
    reply     = reply,
    resolve   = resolve
  )
}

#' @rdname wb_add_thread
#' @export
wb_get_thread <- function(
    wb,
    sheet = current_sheet(),
    dims  = NULL
) {

  assert_workbook(wb)

  wb$clone(deep = TRUE)$get_thread(
    sheet       = sheet,
    dims        = dims
  )
}

#' Add a checkbox, radio button or drop menu to a cell in a worksheet
#'
#' You can add Form Control to a cell. The three supported types are a Checkbox,
#' a Radio button, or a Drop menu.
#'
#' @param wb A Workbook object
#' @param sheet A worksheet of the workbook
#' @param dims A single cell as spreadsheet dimension, e.g. "A1".
#' @param type A type "Checkbox" (the default), "Radio" a radio button or "Drop" a drop down menu
#' @param text A text to be shown next to the Checkbox or radio button (optional)
#' @param link A cell range to link to
#' @param range A cell range used as input
#' @param checked A logical indicating if the Checkbox or Radio button is checked
#' @returns The `wbWorkbook` object, invisibly.
#' @examples
#' wb <- wb_workbook() %>% wb_add_worksheet() %>%
#'   wb_add_form_control()
#' # Add
#' wb$add_form_control(dims = "C5", type = "Radio", checked = TRUE)
#' @export
wb_add_form_control <- function(
    wb,
    sheet   = current_sheet(),
    dims    = "A1",
    type    = c("Checkbox", "Radio", "Drop"),
    text    = NULL,
    link    = NULL,
    range   = NULL,
    checked = FALSE
) {
  assert_workbook(wb)
  wb$clone()$add_form_control(
      sheet   = sheet,
      dims    = dims,
      type    = type,
      text    = text,
      link    = link,
      range   = range,
      checked = checked
  )

}

#' Add conditional formatting to cells in a worksheet
#'
#' Add conditional formatting to cells.
#' You can find more details in `vignette("conditional-formatting")`.
#'
#' openxml uses the alpha channel first then RGB, whereas the usual default is RGBA.
#' @param wb A Workbook object
#' @param sheet A name or index of a worksheet
#' @param dims A cell or cell range like "A1" or "A1:B2"
#' @param rule The condition under which to apply the formatting. See **Examples**.
#' @param style A style to apply to those cells that satisfy the rule.
#'   Default is `font_color = "FF9C0006"` and `bg_fill = "FFFFC7CE"`
#' @param type The type of conditional formatting rule to apply. One of `"expression"`, `"colorScale"` or others mentioned in **Details**.
#' @param params A list of additional parameters passed.  See **Details** for more.
#' @param ... additional arguments
#' @family worksheet content functions
#' @details
#' Conditional formatting `type` accept different parameters. Unless noted,
#' unlisted parameters are ignored.
#' \describe{
#'   \item{`expression`}{
#'     `[style]`\cr A `Style` object\cr\cr
#'     `[rule]`\cr An Excel expression (as a character). Valid operators are: `<`, `<=`, `>`, `>=`, `==`, `!=`
#'   }
#'   \item{`colorScale`}{
#'     `[style]`\cr A `character` vector of valid colors with length `2` or `3`\cr\cr
#'     `[rule]`\cr `NULL` or a `character` vector of valid colors of equal length to `styles`
#'   }
#'   \item{`dataBar`}{
#'     `[style]`\cr A `character` vector of valid colors with length `2` or `3`\cr\cr
#'     `[rule]`\cr A `numeric` vector specifying the range of the databar colors. Must be equal length to `style`\cr\cr
#'     `[params$showValue]`\cr If `FALSE` the cell value is hidden. Default `TRUE`\cr\cr
#'     `[params$gradient]`\cr If `FALSE` color gradient is removed. Default `TRUE`\cr\cr
#'     `[params$border]`\cr If `FALSE` the border around the database is hidden. Default `TRUE`
#'   }
#'   \item{`duplicatedValues` / `uniqueValues` / `containsErrors`}{
#'     `[style]`\cr A `Style` object
#'   }
#'   \item{`contains`}{
#'     `[style]`\cr A `Style` object\cr\cr
#'     `[rule]`\cr The text to look for within cells
#'   }
#'   \item{`between`}{
#'     `[style]`\cr A `Style` object.\cr\cr
#'     `[rule]`\cr A `numeric` vector of length `2` specifying lower and upper bound (Inclusive)
#'   }
#'   \item{`topN`}{
#'     `[style]`\cr A `Style` object\cr\cr
#'     `[params$rank]`\cr A `numeric` vector of length `1` indicating number of highest values. Default `5L`\cr\cr
#'     `[params$percent]` If `TRUE`, uses percentage
#'   }
#'   \item{`bottomN`}{
#'     `[style]`\cr A `Style` object\cr\cr
#'     `[params$rank]`\cr A `numeric` vector of length `1` indicating number of lowest values. Default `5L`\cr\cr
#'     `[params$percent]`\cr If `TRUE`, uses percentage
#'   }
#'  \item{`iconSet`}{
#'     `[params$showValue]`\cr If `FALSE`, the cell value is hidden. Default `TRUE`\cr\cr
#'     `[params$reverse]`\cr If `TRUE`, the order is reversed. Default `FALSE`\cr\cr
#'     `[params$percent]`\cr If `TRUE`, uses percentage\cr\cr
#'     `[params$iconSet]`\cr Uses one of the implemented icon sets. Values must match the length of the icons
#'      in the set 3Arrows, 3ArrowsGray, 3Flags, 3Signs, 3Symbols, 3Symbols2, 3TrafficLights1, 3TrafficLights2,
#'      4Arrows, 4ArrowsGray, 4Rating, 4RedToBlack, 4TrafficLights, 5Arrows, 5ArrowsGray, 5Quarters, 5Rating. The
#'      default is 3TrafficLights1.
#'  }
#' }
#'
#' @examples
#' wb <- wb_workbook()
#' wb$add_worksheet("a")
#' wb$add_data(x = 1:4, col_names = FALSE)
#' wb$add_conditional_formatting(dims = wb_dims(cols = "A", rows = 1:4), rule = ">2")
#' @export
wb_add_conditional_formatting <- function(
    wb,
    sheet  = current_sheet(),
    dims   = NULL,
    rule   = NULL,
    style  = NULL,
    type   = c(
      "expression", "colorScale",
      "dataBar", "iconSet",
      "duplicatedValues", "uniqueValues",
      "containsErrors", "notContainsErrors",
      "containsBlanks", "notContainsBlanks",
      "containsText", "notContainsText",
      "beginsWith", "endsWith",
      "between", "topN", "bottomN"
    ),
    params = list(
      showValue = TRUE,
      gradient  = TRUE,
      border    = TRUE,
      percent   = FALSE,
      rank      = 5L
    ),
    ...
) {
  assert_workbook(wb)
  wb$clone()$add_conditional_formatting(
    sheet  = sheet,
    dims   = dims,
    rule   = rule,
    style  = style,
    type   = type,
    params = params,
    ...    = ...
  )
}

#' Apply styling from a sheet to another within a workbook
#'
#' This function can be used to apply styling from a cell range, and apply it
#' to another cell range.
#' @param wb A workbook
#' @param from sheet we select the style from
#' @param to sheet to apply the style to
#' @export
wb_clone_sheet_style <- function(wb, from = current_sheet(), to) {
  assert_workbook(wb)
  wb$clone()$clone_sheet_style(from, to)
}

#' Add sparklines to a worksheet
#'
#' @param wb A `wbWorkbook`
#' @param sheet sheet to add the sparklines to
#' @param sparklines sparklines object created with [create_sparklines()]
#' @seealso [create_sparklines()]
#' @examples
#'  sl <- create_sparklines("Sheet 1", dims = "A3:K3", sqref = "L3")
#'  wb <- wb_workbook() %>%
#'    wb_add_worksheet() %>%
#'    wb_add_data(x = mtcars) %>%
#'    wb_add_sparklines(sparklines = sl)
#' @export
wb_add_sparklines <- function(wb, sheet = current_sheet(), sparklines) {
  assert_workbook(wb)
  wb$clone(deep = TRUE)$add_sparklines(sheet, sparklines)
}

#' Ignore error types on worksheet
#'
#' This function allows to hide / ignore certain types of errors shown in a worksheet.
#' @param wb A workbook
#' @param sheet A sheet name or index.
#' @param dims Cell range to ignore the error
#' @param calculated_column calculatedColumn
#' @param empty_cell_reference emptyCellReference
#' @param eval_error evalError
#' @param formula formula
#' @param formula_range formulaRange
#' @param list_data_validation listDataValidation
#' @param number_stored_as_text If `TRUE`, will not display the error if numbers are stored as text.
#' @param two_digit_text_year twoDigitTextYear
#' @param unlocked_formula unlockedFormula
#' @param ... additional arguments
#' @return The `wbWorkbook` object, invisibly.
#' @export
wb_add_ignore_error <- function(
    wb,
    sheet                 = current_sheet(),
    dims                  = "A1",
    calculated_column     = FALSE,
    empty_cell_reference  = FALSE,
    eval_error            = FALSE,
    formula               = FALSE,
    formula_range         = FALSE,
    list_data_validation  = FALSE,
    number_stored_as_text = FALSE,
    two_digit_text_year   = FALSE,
    unlocked_formula      = FALSE,
    ...
) {
  assert_workbook(wb)
  wb$clone()$add_ignore_error(
    sheet                 = sheet,
    dims                  = dims,
    calculated_column     = calculated_column,
    empty_cell_reference  = empty_cell_reference,
    eval_error            = eval_error,
    formula               = formula,
    formula_range         = formula_range,
    list_data_validation  = list_data_validation,
    number_stored_as_text = number_stored_as_text,
    two_digit_text_year   = two_digit_text_year,
    unlocked_formula      = unlocked_formula,
    ...                   = ...
    )
}

#' Modify the default view of a worksheet
#'
#' This helps set a worksheet's appearance, such as the zoom, whether to show grid lines
#'
#' @param wb A Workbook object
#' @param sheet sheet
#' @param color_id,default_grid_color Integer: A color, default is 64
#' @param right_to_left Logical: if `TRUE` column ordering is right  to left
#' @param show_formulas Logical: if `TRUE` cell formulas are shown
#' @param show_grid_lines Logical: if `TRUE` the worksheet grid is shown
#' @param show_outline_symbols Logical: if `TRUE` outline symbols are shown
#' @param show_row_col_headers Logical: if `TRUE` row and column headers are shown
#' @param show_ruler Logical: if `TRUE` a ruler is shown in page layout view
#' @param show_white_space Logical: if `TRUE` margins are shown in page layout view
#' @param show_zeros Logical: if `FALSE` cells containing zero are shown blank if `show_formulas = FALSE`
#' @param tab_selected Integer: zero vector indicating the selected tab
#' @param top_left_cell Cell: the cell shown in the top left corner / or top right with rightToLeft
#' @param view View: "normal", "pageBreakPreview" or "pageLayout"
#' @param window_protection Logical: if `TRUE` the panes are protected
#' @param workbook_view_id integer: Pointing to some other view inside the workbook
#' @param zoom_scale,zoom_scale_normal,zoom_scale_page_layout_view,zoom_scale_sheet_layout_view
#'   Integer: the zoom scale should be between 10 and 400. These are values for current, normal etc.
#' @param ... additional arguments
#' @examples
#' wb <- wb_workbook()$add_worksheet()
#'
#' wb$set_sheetview(
#'   zoom_scale = 75,
#'   right_to_left = FALSE,
#'   show_formulas = TRUE,
#'   show_grid_lines = TRUE,
#'   show_outline_symbols = FALSE,
#'   show_row_col_headers = TRUE,
#'   show_ruler = TRUE,
#'   show_white_space = FALSE,
#'   tab_selected = 1,
#'   top_left_cell = "B1",
#'   view = "normal",
#'   window_protection = TRUE
#' )
#' @return The `wbWorkbook` object, invisibly
#' @export
wb_set_sheetview <- function(
    wb,
    sheet                        = current_sheet(),
    color_id                     = NULL,
    default_grid_color           = NULL,
    right_to_left                = NULL,
    show_formulas                = NULL,
    show_grid_lines              = NULL,
    show_outline_symbols         = NULL,
    show_row_col_headers         = NULL,
    show_ruler                   = NULL,
    show_white_space             = NULL,
    show_zeros                   = NULL,
    tab_selected                 = NULL,
    top_left_cell                = NULL,
    view                         = NULL,
    window_protection            = NULL,
    workbook_view_id             = NULL,
    zoom_scale                   = NULL,
    zoom_scale_normal            = NULL,
    zoom_scale_page_layout_view  = NULL,
    zoom_scale_sheet_layout_view = NULL,
    ...
) {
  assert_workbook(wb)
  wb$clone()$set_sheetview(
    sheet                        = sheet,
    color_id                     = color_id,
    default_grid_color           = default_grid_color,
    right_to_left                = right_to_left,
    show_formulas                = show_formulas,
    show_grid_lines              = show_grid_lines,
    show_outline_symbols         = show_outline_symbols,
    show_row_col_headers         = show_row_col_headers,
    show_ruler                   = show_ruler,
    show_white_space             = show_white_space,
    show_zeros                   = show_zeros,
    tab_selected                 = tab_selected,
    top_left_cell                = top_left_cell,
    view                         = view,
    window_protection            = window_protection,
    workbook_view_id             = workbook_view_id,
    zoom_scale                   = zoom_scale,
    zoom_scale_normal            = zoom_scale_normal,
    zoom_scale_page_layout_view  = zoom_scale_page_layout_view,
    zoom_scale_sheet_layout_view = zoom_scale_sheet_layout_view,
    ...                          = ...
  )
}

#' Modify the state of active and selected sheets in a workbook
#'
#' @description
#' Get and set table of sheets and their state as selected and active in a workbook
#'
#' Multiple sheets can be selected, but only a single one can be active (visible).
#' The visible sheet, must not necessarily be a selected sheet.
#'
#' @param wb a workbook
#' @returns a data frame with tabSelected and names
#' @export
#' @examples
#' wb <- wb_load(file = system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2"))
#' # testing is the selected sheet
#' wb_get_selected(wb)
#' # change the selected sheet to Sheet2
#' wb <- wb_set_selected(wb, "Sheet2")
#' # get the active sheet
#' wb_get_active_sheet(wb)
#' # change the selected sheet to Sheet2
#' wb <- wb_set_active_sheet(wb, sheet = "Sheet2")
#' @name active_sheet-wb
wb_get_active_sheet <- function(wb) {
  assert_workbook(wb)
  wb$get_active_sheet()
}

#' @rdname active_sheet-wb
#' @param sheet a sheet name of the workbook
#' @export
wb_set_active_sheet <- function(wb, sheet) {
  # active tab requires a c index
  assert_workbook(wb)
  wb$clone()$set_active_sheet(sheet = sheet)
}

#' @rdname active_sheet-wb
#' @export
wb_get_selected <- function(wb) {
  assert_workbook(wb)
  wb$get_selected()
}

#' @rdname active_sheet-wb
#' @export
wb_set_selected <- function(wb, sheet) {
  assert_workbook(wb)
  wb$clone()$set_selected(sheet = sheet)
}
