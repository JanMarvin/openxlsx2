# Workbook class

This is the class used by `openxlsx2` to modify workbooks from R. You
can load an existing workbook with
[`wb_load()`](https://janmarvin.github.io/openxlsx2/reference/wb_load.md)
and create a new one with
[`wb_workbook()`](https://janmarvin.github.io/openxlsx2/reference/wb_workbook.md).

After that, you can modify the `wbWorkbook` object through two primary
methods:

*Wrapper Function Method*: Utilizes the `wb` family of functions that
support piping to streamline operations.

    wb <- wb_workbook(creator = "My name here")
    wb <- wb_add_worksheet(wb, sheet = "Expenditure", grid_lines = FALSE)
    wb <- wb_add_data(wb, x = USPersonalExpenditure, row_names = TRUE)

*Chaining Method*: Directly modifies the object through a series of
chained function calls.

    wb <- wb_workbook(creator = "My name here")$
      add_worksheet(sheet = "Expenditure", grid_lines = FALSE)$
      add_data(x = USPersonalExpenditure, row_names = TRUE)

While wrapper functions require explicit assignment of their output to
reflect changes, chained functions inherently modify the input object.
Both approaches are equally supported, offering flexibility to suit user
preferences. The documentation mainly highlights the use of wrapper
functions.

    # Import workbooks
    path <- system.file("extdata/openxlsx2_example.xlsx", package = "openxlsx2")
    wb <- wb_load(path)

    ## or create one yourself
    wb <- wb_workbook()
    # add a worksheet
    wb$add_worksheet("sheet")
    # add some data
    wb$add_data("sheet", cars)
    # Add data in a different location
    wb <- wb_add_data(wb, x = cars, dims = wb_dims(from_dims = "D4"))
    # open it in your default spreadsheet software
    if (interactive()) wb$open()

Note that the documentation is more complete in each of the wrapper
functions. (i.e.
[`?wb_add_data`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data.md)
rather than `?wbWorkbook`).

## Public fields

- `sheet_names`:

  The names of the sheets

- `calcChain`:

  calcChain

- `charts`:

  charts

- `is_chartsheet`:

  A logical vector identifying if a sheet is a chartsheet.

- `customXml`:

  customXml

- `connections`:

  connections

- `ctrlProps`:

  ctrlProps

- `Content_Types`:

  Content_Types

- `app`:

  app

- `core`:

  The XML core

- `custom`:

  custom

- `drawings`:

  drawings

- `drawings_rels`:

  drawings_rels

- `docMetadata`:

  doc_meta_data

- `activeX`:

  activeX

- `embeddings`:

  embeddings

- `externalLinks`:

  externalLinks

- `externalLinksRels`:

  externalLinksRels

- `featurePropertyBag`:

  featurePropertyBag

- `python`:

  python

- `webextensions`:

  webextensions

- `headFoot`:

  The header and footer

- `media`:

  media

- `metadata`:

  contains cell/value metadata imported on load from xl/metadata.xml

- `persons`:

  Persons of the workbook. to be used with
  [`wb_add_thread()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_thread.md)

- `pivotTables`:

  pivotTables

- `pivotTables.xml.rels`:

  pivotTables.xml.rels

- `pivotDefinitions`:

  pivotDefinitions

- `pivotRecords`:

  pivotRecords

- `pivotDefinitionsRels`:

  pivotDefinitionsRels

- `queryTables`:

  queryTables

- `richData`:

  richData

- `slicers`:

  slicers

- `slicerCaches`:

  slicerCaches

- `sharedStrings`:

  sharedStrings

- `styles_mgr`:

  styles_mgr

- `tables`:

  tables

- `tables.xml.rels`:

  tables.xml.rels

- `tableSingleCells`:

  tableSingleCells

- `theme`:

  theme

- `vbaProject`:

  vbaProject

- `vml`:

  vml

- `vml_rels`:

  vml_rels

- `comments`:

  Comments (notes) present in the workbook.

- `threadComments`:

  Threaded comments

- `timelines`:

  timelines

- `timelineCaches`:

  timelineCaches

- `workbook`:

  workbook

- `workbook.xml.rels`:

  workbook.xml.rels

- `worksheets`:

  worksheets

- `worksheets_rels`:

  worksheets_rels

- `sheetOrder`:

  The sheet order. Controls ordering for worksheets and worksheet names.

- `path`:

  path

- `tmpDir`:

  tmpDir

- `namedSheetViews`:

  namedSheetViews

- `xmlMaps`:

  xmlMaps

## Methods

### Public methods

- [`wbWorkbook$new()`](#method-wbWorkbook-new)

- [`wbWorkbook$append()`](#method-wbWorkbook-append)

- [`wbWorkbook$append_sheets()`](#method-wbWorkbook-append_sheets)

- [`wbWorkbook$validate_sheet()`](#method-wbWorkbook-validate_sheet)

- [`wbWorkbook$add_chartsheet()`](#method-wbWorkbook-add_chartsheet)

- [`wbWorkbook$add_worksheet()`](#method-wbWorkbook-add_worksheet)

- [`wbWorkbook$clone_worksheet()`](#method-wbWorkbook-clone_worksheet)

- [`wbWorkbook$add_data()`](#method-wbWorkbook-add_data)

- [`wbWorkbook$add_data_table()`](#method-wbWorkbook-add_data_table)

- [`wbWorkbook$add_pivot_table()`](#method-wbWorkbook-add_pivot_table)

- [`wbWorkbook$add_slicer()`](#method-wbWorkbook-add_slicer)

- [`wbWorkbook$remove_slicer()`](#method-wbWorkbook-remove_slicer)

- [`wbWorkbook$add_timeline()`](#method-wbWorkbook-add_timeline)

- [`wbWorkbook$remove_timeline()`](#method-wbWorkbook-remove_timeline)

- [`wbWorkbook$add_formula()`](#method-wbWorkbook-add_formula)

- [`wbWorkbook$add_hyperlink()`](#method-wbWorkbook-add_hyperlink)

- [`wbWorkbook$remove_hyperlink()`](#method-wbWorkbook-remove_hyperlink)

- [`wbWorkbook$add_style()`](#method-wbWorkbook-add_style)

- [`wbWorkbook$to_df()`](#method-wbWorkbook-to_df)

- [`wbWorkbook$load()`](#method-wbWorkbook-load)

- [`wbWorkbook$save()`](#method-wbWorkbook-save)

- [`wbWorkbook$open()`](#method-wbWorkbook-open)

- [`wbWorkbook$buildTable()`](#method-wbWorkbook-buildTable)

- [`wbWorkbook$update_table()`](#method-wbWorkbook-update_table)

- [`wbWorkbook$copy_cells()`](#method-wbWorkbook-copy_cells)

- [`wbWorkbook$get_base_font()`](#method-wbWorkbook-get_base_font)

- [`wbWorkbook$set_base_font()`](#method-wbWorkbook-set_base_font)

- [`wbWorkbook$get_base_colors()`](#method-wbWorkbook-get_base_colors)

- [`wbWorkbook$get_base_colours()`](#method-wbWorkbook-get_base_colours)

- [`wbWorkbook$set_base_colors()`](#method-wbWorkbook-set_base_colors)

- [`wbWorkbook$set_base_colours()`](#method-wbWorkbook-set_base_colours)

- [`wbWorkbook$get_bookview()`](#method-wbWorkbook-get_bookview)

- [`wbWorkbook$remove_bookview()`](#method-wbWorkbook-remove_bookview)

- [`wbWorkbook$set_bookview()`](#method-wbWorkbook-set_bookview)

- [`wbWorkbook$get_sheet_names()`](#method-wbWorkbook-get_sheet_names)

- [`wbWorkbook$set_sheet_names()`](#method-wbWorkbook-set_sheet_names)

- [`wbWorkbook$set_row_heights()`](#method-wbWorkbook-set_row_heights)

- [`wbWorkbook$remove_row_heights()`](#method-wbWorkbook-remove_row_heights)

- [`wbWorkbook$createCols()`](#method-wbWorkbook-createCols)

- [`wbWorkbook$group_cols()`](#method-wbWorkbook-group_cols)

- [`wbWorkbook$ungroup_cols()`](#method-wbWorkbook-ungroup_cols)

- [`wbWorkbook$remove_col_widths()`](#method-wbWorkbook-remove_col_widths)

- [`wbWorkbook$set_col_widths()`](#method-wbWorkbook-set_col_widths)

- [`wbWorkbook$group_rows()`](#method-wbWorkbook-group_rows)

- [`wbWorkbook$ungroup_rows()`](#method-wbWorkbook-ungroup_rows)

- [`wbWorkbook$remove_worksheet()`](#method-wbWorkbook-remove_worksheet)

- [`wbWorkbook$add_data_validation()`](#method-wbWorkbook-add_data_validation)

- [`wbWorkbook$merge_cells()`](#method-wbWorkbook-merge_cells)

- [`wbWorkbook$unmerge_cells()`](#method-wbWorkbook-unmerge_cells)

- [`wbWorkbook$freeze_pane()`](#method-wbWorkbook-freeze_pane)

- [`wbWorkbook$add_comment()`](#method-wbWorkbook-add_comment)

- [`wbWorkbook$get_comment()`](#method-wbWorkbook-get_comment)

- [`wbWorkbook$remove_comment()`](#method-wbWorkbook-remove_comment)

- [`wbWorkbook$add_thread()`](#method-wbWorkbook-add_thread)

- [`wbWorkbook$get_thread()`](#method-wbWorkbook-get_thread)

- [`wbWorkbook$add_conditional_formatting()`](#method-wbWorkbook-add_conditional_formatting)

- [`wbWorkbook$remove_conditional_formatting()`](#method-wbWorkbook-remove_conditional_formatting)

- [`wbWorkbook$add_image()`](#method-wbWorkbook-add_image)

- [`wbWorkbook$add_plot()`](#method-wbWorkbook-add_plot)

- [`wbWorkbook$add_drawing()`](#method-wbWorkbook-add_drawing)

- [`wbWorkbook$add_chart_xml()`](#method-wbWorkbook-add_chart_xml)

- [`wbWorkbook$add_mschart()`](#method-wbWorkbook-add_mschart)

- [`wbWorkbook$add_form_control()`](#method-wbWorkbook-add_form_control)

- [`wbWorkbook$print()`](#method-wbWorkbook-print)

- [`wbWorkbook$protect()`](#method-wbWorkbook-protect)

- [`wbWorkbook$protect_worksheet()`](#method-wbWorkbook-protect_worksheet)

- [`wbWorkbook$get_properties()`](#method-wbWorkbook-get_properties)

- [`wbWorkbook$set_properties()`](#method-wbWorkbook-set_properties)

- [`wbWorkbook$add_mips()`](#method-wbWorkbook-add_mips)

- [`wbWorkbook$get_mips()`](#method-wbWorkbook-get_mips)

- [`wbWorkbook$set_creators()`](#method-wbWorkbook-set_creators)

- [`wbWorkbook$add_creators()`](#method-wbWorkbook-add_creators)

- [`wbWorkbook$remove_creators()`](#method-wbWorkbook-remove_creators)

- [`wbWorkbook$set_last_modified_by()`](#method-wbWorkbook-set_last_modified_by)

- [`wbWorkbook$set_page_setup()`](#method-wbWorkbook-set_page_setup)

- [`wbWorkbook$page_setup()`](#method-wbWorkbook-page_setup)

- [`wbWorkbook$set_header_footer()`](#method-wbWorkbook-set_header_footer)

- [`wbWorkbook$get_tables()`](#method-wbWorkbook-get_tables)

- [`wbWorkbook$remove_tables()`](#method-wbWorkbook-remove_tables)

- [`wbWorkbook$add_filter()`](#method-wbWorkbook-add_filter)

- [`wbWorkbook$remove_filter()`](#method-wbWorkbook-remove_filter)

- [`wbWorkbook$set_grid_lines()`](#method-wbWorkbook-set_grid_lines)

- [`wbWorkbook$grid_lines()`](#method-wbWorkbook-grid_lines)

- [`wbWorkbook$add_named_region()`](#method-wbWorkbook-add_named_region)

- [`wbWorkbook$get_named_regions()`](#method-wbWorkbook-get_named_regions)

- [`wbWorkbook$remove_named_region()`](#method-wbWorkbook-remove_named_region)

- [`wbWorkbook$set_order()`](#method-wbWorkbook-set_order)

- [`wbWorkbook$get_sheet_visibility()`](#method-wbWorkbook-get_sheet_visibility)

- [`wbWorkbook$set_sheet_visibility()`](#method-wbWorkbook-set_sheet_visibility)

- [`wbWorkbook$add_page_break()`](#method-wbWorkbook-add_page_break)

- [`wbWorkbook$clean_sheet()`](#method-wbWorkbook-clean_sheet)

- [`wbWorkbook$add_border()`](#method-wbWorkbook-add_border)

- [`wbWorkbook$add_fill()`](#method-wbWorkbook-add_fill)

- [`wbWorkbook$add_font()`](#method-wbWorkbook-add_font)

- [`wbWorkbook$add_numfmt()`](#method-wbWorkbook-add_numfmt)

- [`wbWorkbook$add_cell_style()`](#method-wbWorkbook-add_cell_style)

- [`wbWorkbook$get_cell_style()`](#method-wbWorkbook-get_cell_style)

- [`wbWorkbook$set_cell_style()`](#method-wbWorkbook-set_cell_style)

- [`wbWorkbook$set_cell_style_across()`](#method-wbWorkbook-set_cell_style_across)

- [`wbWorkbook$add_named_style()`](#method-wbWorkbook-add_named_style)

- [`wbWorkbook$add_dxfs_style()`](#method-wbWorkbook-add_dxfs_style)

- [`wbWorkbook$clone_sheet_style()`](#method-wbWorkbook-clone_sheet_style)

- [`wbWorkbook$add_sparklines()`](#method-wbWorkbook-add_sparklines)

- [`wbWorkbook$add_ignore_error()`](#method-wbWorkbook-add_ignore_error)

- [`wbWorkbook$set_sheetview()`](#method-wbWorkbook-set_sheetview)

- [`wbWorkbook$add_person()`](#method-wbWorkbook-add_person)

- [`wbWorkbook$get_person()`](#method-wbWorkbook-get_person)

- [`wbWorkbook$get_active_sheet()`](#method-wbWorkbook-get_active_sheet)

- [`wbWorkbook$set_active_sheet()`](#method-wbWorkbook-set_active_sheet)

- [`wbWorkbook$get_selected()`](#method-wbWorkbook-get_selected)

- [`wbWorkbook$set_selected()`](#method-wbWorkbook-set_selected)

- [`wbWorkbook$clone()`](#method-wbWorkbook-clone)

------------------------------------------------------------------------

### Method `new()`

Creates a new `wbWorkbook` object

#### Usage

    wbWorkbook$new(
      creator = NULL,
      title = NULL,
      subject = NULL,
      category = NULL,
      datetime_created = Sys.time(),
      datetime_modified = NULL,
      theme = NULL,
      keywords = NULL,
      comments = NULL,
      manager = NULL,
      company = NULL,
      ...
    )

#### Arguments

- `creator`:

  character vector of creators. Duplicated are ignored.

- `title, subject, category, keywords, comments, manager, company`:

  workbook properties

- `datetime_created`:

  The datetime (as `POSIXt`) the workbook is created. Defaults to the
  current [`Sys.time()`](https://rdrr.io/r/base/Sys.time.html) when the
  workbook object is created, not when the
  [`wb_workbook()`](https://janmarvin.github.io/openxlsx2/reference/wb_workbook.md)
  files are saved.

- `datetime_modified`:

  The datetime (as `POSIXt`) that should be recorded as last
  modification date. Defaults to the creation date.

- `theme`:

  Optional theme identified by string or number

- `...`:

  additional arguments

#### Returns

a `wbWorkbook` object

------------------------------------------------------------------------

### Method [`append()`](https://rdrr.io/r/base/append.html)

Append a field. This method is intended for internal use

#### Usage

    wbWorkbook$append(field, value)

#### Arguments

- `field`:

  A valid field name

- `value`:

  A value for the field

------------------------------------------------------------------------

### Method `append_sheets()`

Append to `self$workbook$sheets` This method is intended for internal
use

#### Usage

    wbWorkbook$append_sheets(value)

#### Arguments

- `value`:

  A value for `self$workbook$sheets`

------------------------------------------------------------------------

### Method `validate_sheet()`

validate sheet

#### Usage

    wbWorkbook$validate_sheet(sheet)

#### Arguments

- `sheet`:

  A character sheet name or integer location

#### Returns

The integer position of the sheet

------------------------------------------------------------------------

### Method `add_chartsheet()`

Add a chart sheet to the workbook

#### Usage

    wbWorkbook$add_chartsheet(
      sheet = next_sheet(),
      tab_color = NULL,
      zoom = 100,
      visible = c("true", "false", "hidden", "visible", "veryhidden"),
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `tab_color`:

  tab_color

- `zoom`:

  zoom

- `visible`:

  visible

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object, invisibly

------------------------------------------------------------------------

### Method `add_worksheet()`

Add worksheet to the `wbWorkbook` object

#### Usage

    wbWorkbook$add_worksheet(
      sheet = next_sheet(),
      grid_lines = TRUE,
      row_col_headers = TRUE,
      tab_color = NULL,
      zoom = 100,
      header = NULL,
      footer = NULL,
      odd_header = header,
      odd_footer = footer,
      even_header = header,
      even_footer = footer,
      first_header = header,
      first_footer = footer,
      visible = c("true", "false", "hidden", "visible", "veryhidden"),
      has_drawing = FALSE,
      paper_size = getOption("openxlsx2.paperSize", default = 9),
      orientation = getOption("openxlsx2.orientation", default = "portrait"),
      hdpi = getOption("openxlsx2.hdpi", default = getOption("openxlsx2.dpi", default = 300)),
      vdpi = getOption("openxlsx2.vdpi", default = getOption("openxlsx2.dpi", default = 300)),
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `grid_lines`:

  gridLines

- `row_col_headers`:

  rowColHeaders

- `tab_color`:

  tabColor

- `zoom`:

  zoom

- `header`:

  header

- `footer`:

  footer

- `odd_header`:

  oddHeader

- `odd_footer`:

  oddFooter

- `even_header`:

  evenHeader

- `even_footer`:

  evenFooter

- `first_header`:

  firstHeader

- `first_footer`:

  firstFooter

- `visible`:

  visible

- `has_drawing`:

  hasDrawing

- `paper_size`:

  paperSize

- `orientation`:

  orientation

- `hdpi`:

  hdpi

- `vdpi`:

  vdpi

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object, invisibly

------------------------------------------------------------------------

### Method `clone_worksheet()`

Clone a workbooksheet to another workbook

#### Usage

    wbWorkbook$clone_worksheet(
      old = current_sheet(),
      new = next_sheet(),
      from = NULL
    )

#### Arguments

- `old`:

  name of worksheet to clone

- `new`:

  name of new worksheet to add

- `from`:

  name of new worksheet to add

------------------------------------------------------------------------

### Method `add_data()`

add data

#### Usage

    wbWorkbook$add_data(
      sheet = current_sheet(),
      x,
      dims = wb_dims(start_row, start_col),
      start_col = 1,
      start_row = 1,
      array = FALSE,
      col_names = TRUE,
      row_names = FALSE,
      with_filter = FALSE,
      name = NULL,
      sep = ", ",
      apply_cell_style = TRUE,
      remove_cell_style = FALSE,
      na = na_strings(),
      inline_strings = TRUE,
      enforce = FALSE,
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `x`:

  x

- `dims`:

  Cell range in a sheet

- `start_col`:

  startCol

- `start_row`:

  startRow

- `array`:

  array

- `col_names`:

  colNames

- `row_names`:

  rowNames

- `with_filter`:

  withFilter

- `name`:

  name

- `sep`:

  sep

- `apply_cell_style`:

  applyCellStyle

- `remove_cell_style`:

  if writing into existing cells, should the cell style be removed?

- `na`:

  Value used for replacing `NA` values from `x`. Default
  [`na_strings()`](https://janmarvin.github.io/openxlsx2/reference/waivers.md)
  uses the special `#N/A` value within the workbook.

- `inline_strings`:

  write characters as inline strings

- `enforce`:

  enforce that selected dims is filled. For this to work, `dims` must
  match `x`

- `...`:

  additional arguments

- `return`:

  The `wbWorkbook` object

------------------------------------------------------------------------

### Method `add_data_table()`

add a data table

#### Usage

    wbWorkbook$add_data_table(
      sheet = current_sheet(),
      x,
      dims = wb_dims(start_row, start_col),
      start_col = 1,
      start_row = 1,
      col_names = TRUE,
      row_names = FALSE,
      table_style = "TableStyleLight9",
      table_name = NULL,
      with_filter = TRUE,
      sep = ", ",
      first_column = FALSE,
      last_column = FALSE,
      banded_rows = TRUE,
      banded_cols = FALSE,
      apply_cell_style = TRUE,
      remove_cell_style = FALSE,
      na = na_strings(),
      inline_strings = TRUE,
      total_row = FALSE,
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `x`:

  x

- `dims`:

  Cell range in a sheet

- `start_col`:

  startCol

- `start_row`:

  startRow

- `col_names`:

  colNames

- `row_names`:

  rowNames

- `table_style`:

  tableStyle

- `table_name`:

  tableName

- `with_filter`:

  withFilter

- `sep`:

  sep

- `first_column`:

  firstColumn

- `last_column`:

  lastColumn

- `banded_rows`:

  bandedRows

- `banded_cols`:

  bandedCols

- `apply_cell_style`:

  applyCellStyle

- `remove_cell_style`:

  if writing into existing cells, should the cell style be removed?

- `na`:

  Value used for replacing `NA` values from `x`. Default
  [`na_strings()`](https://janmarvin.github.io/openxlsx2/reference/waivers.md)
  uses the special `#N/A` value within the workbook.

- `inline_strings`:

  write characters as inline strings

- `total_row`:

  write total rows to table

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `add_pivot_table()`

add pivot table

#### Usage

    wbWorkbook$add_pivot_table(
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
      slicer,
      timeline
    )

#### Arguments

- `x`:

  a wb_data object

- `sheet`:

  The name of the sheet

- `dims`:

  the worksheet cell where the pivot table is placed

- `filter`:

  a character object with names used to filter

- `rows`:

  a character object with names used as rows

- `cols`:

  a character object with names used as cols

- `data`:

  a character object with names used as data

- `fun`:

  a character object of functions to be used with the data

- `params`:

  a list of parameters to modify pivot table creation

- `pivot_table`:

  a character object with a name for the pivot table

- `slicer`:

  a character object with names used as slicer

- `timeline`:

  a character object with names used as timeline

#### Details

`fun` can be either of AVERAGE, COUNT, COUNTA, MAX, MIN, PRODUCT, STDEV,
STDEVP, SUM, VAR, VARP

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `add_slicer()`

add pivot table

#### Usage

    wbWorkbook$add_slicer(
      x,
      dims = "A1",
      sheet = current_sheet(),
      pivot_table,
      slicer,
      params
    )

#### Arguments

- `x`:

  a wb_data object

- `dims`:

  the worksheet cell where the pivot table is placed

- `sheet`:

  The name of the sheet

- `pivot_table`:

  the name of a pivot table on the selected sheet

- `slicer`:

  a variable used as slicer for the pivot table

- `params`:

  a list of parameters to modify pivot table creation

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `remove_slicer()`

add pivot table

#### Usage

    wbWorkbook$remove_slicer(sheet = current_sheet())

#### Arguments

- `sheet`:

  The name of the sheet

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `add_timeline()`

add pivot table

#### Usage

    wbWorkbook$add_timeline(
      x,
      dims = "A1",
      sheet = current_sheet(),
      pivot_table,
      timeline,
      params
    )

#### Arguments

- `x`:

  a wb_data object

- `dims`:

  the worksheet cell where the pivot table is placed

- `sheet`:

  The name of the sheet

- `pivot_table`:

  the name of a pivot table on the selected sheet

- `timeline`:

  a variable used as timeline for the pivot table

- `params`:

  a list of parameters to modify pivot table creation

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `remove_timeline()`

add pivot table

#### Usage

    wbWorkbook$remove_timeline(sheet = current_sheet())

#### Arguments

- `sheet`:

  The name of the sheet

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `add_formula()`

Add formula

#### Usage

    wbWorkbook$add_formula(
      sheet = current_sheet(),
      x,
      dims = wb_dims(start_row, start_col),
      start_col = 1,
      start_row = 1,
      array = FALSE,
      cm = FALSE,
      apply_cell_style = TRUE,
      remove_cell_style = FALSE,
      enforce = FALSE,
      shared = FALSE,
      name = NULL,
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `x`:

  x

- `dims`:

  Cell range in a sheet

- `start_col`:

  startCol

- `start_row`:

  startRow

- `array`:

  array

- `cm`:

  cm

- `apply_cell_style`:

  applyCellStyle

- `remove_cell_style`:

  if writing into existing cells, should the cell style be removed?

- `enforce`:

  enforce dims

- `shared`:

  shared formula

- `name`:

  name

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `add_hyperlink()`

Add hyperlink

#### Usage

    wbWorkbook$add_hyperlink(
      sheet = current_sheet(),
      dims = "A1",
      target = NULL,
      tooltip = NULL,
      is_external = TRUE,
      col_names = FALSE
    )

#### Arguments

- `sheet`:

  sheet

- `dims`:

  dims

- `target`:

  target

- `tooltip`:

  tooltip

- `is_external`:

  is_external

- `col_names`:

  col_names

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `remove_hyperlink()`

remove hyperlink

#### Usage

    wbWorkbook$remove_hyperlink(sheet = current_sheet(), dims = NULL)

#### Arguments

- `sheet`:

  sheet

- `dims`:

  dims

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `add_style()`

add style

#### Usage

    wbWorkbook$add_style(style = NULL, style_name = NULL)

#### Arguments

- `style`:

  style

- `style_name`:

  style_name

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `to_df()`

to_df

#### Usage

    wbWorkbook$to_df(
      sheet,
      start_row = NULL,
      start_col = NULL,
      row_names = FALSE,
      col_names = TRUE,
      skip_empty_rows = FALSE,
      skip_empty_cols = FALSE,
      skip_hidden_rows = FALSE,
      skip_hidden_cols = FALSE,
      rows = NULL,
      cols = NULL,
      detect_dates = TRUE,
      na = "#N/A",
      fill_merged_cells = FALSE,
      dims,
      show_formula = FALSE,
      convert = TRUE,
      types,
      named_region,
      keep_attributes = FALSE,
      check_names = FALSE,
      show_hyperlinks = FALSE,
      apply_numfmts = FALSE,
      ...
    )

#### Arguments

- `sheet`:

  Either sheet name or index. When missing the first sheet in the
  workbook is selected.

- `start_row`:

  first row to begin looking for data.

- `start_col`:

  first column to begin looking for data.

- `row_names`:

  If TRUE, the first col of data will be used as row names.

- `col_names`:

  If TRUE, the first row of data will be used as column names.

- `skip_empty_rows`:

  If TRUE, empty rows are skipped.

- `skip_empty_cols`:

  If TRUE, empty columns are skipped.

- `skip_hidden_rows`:

  If TRUE, hidden rows are skipped.

- `skip_hidden_cols`:

  If TRUE, hidden columns are skipped.

- `rows`:

  A numeric vector specifying which rows in the spreadsheet to read. If
  NULL, all rows are read.

- `cols`:

  A numeric vector specifying which columns in the spreadsheet to read.
  If NULL, all columns are read.

- `detect_dates`:

  If TRUE, attempt to recognize dates and perform conversion.

- `na`:

  Defines values to be treated as NA. Can be a character vector of
  strings or a named list: list(strings = ..., numbers = ...). Blank
  cells are always converted to `NA`.

- `fill_merged_cells`:

  If TRUE, the value in a merged cell is given to all cells within the
  merge.

- `dims`:

  Character string of type "A1:B2" as optional dimensions to be
  imported.

- `show_formula`:

  If TRUE, the underlying spreadsheet formulas are shown.

- `convert`:

  If TRUE, a conversion to dates and numerics is attempted.

- `types`:

  A named numeric indicating, the type of the data. 0: character, 1:
  numeric, 2: date, 3: posixt, 4:logical. Names must match the returned
  data

- `named_region`:

  Character string with a named_region (defined name or table). If no
  sheet is selected, the first appearance will be selected.

- `keep_attributes`:

  If TRUE additional attributes are returned. (These are used internally
  to define a cell type.)

- `check_names`:

  If TRUE then the names of the variables in the data frame are checked
  to ensure that they are syntactically valid variable names.

- `show_hyperlinks`:

  If `TRUE` instead of the displayed text, hyperlink targets are shown.

- `apply_numfmts`:

  If `TRUE` numeric formats are applied if detected.

- `...`:

  additional arguments

#### Returns

a data frame

------------------------------------------------------------------------

### Method [`load()`](https://rdrr.io/r/base/load.html)

load workbook

#### Usage

    wbWorkbook$load(file, sheet, data_only = FALSE, ...)

#### Arguments

- `file`:

  file

- `sheet`:

  The name of the sheet

- `data_only`:

  data_only

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object invisibly

------------------------------------------------------------------------

### Method [`save()`](https://rdrr.io/r/base/save.html)

Save the workbook

#### Usage

    wbWorkbook$save(file = self$path, overwrite = TRUE, path = NULL, flush = FALSE)

#### Arguments

- `file`:

  The path to save the workbook to

- `overwrite`:

  If `FALSE`, will not overwrite when `path` exists

- `path`:

  Deprecated argument previously used for file. Please use file in new
  code.

- `flush`:

  Experimental, streams the worksheet file to disk

#### Returns

The `wbWorkbook` object invisibly

------------------------------------------------------------------------

### Method [`open()`](https://rdrr.io/r/base/connections.html)

open wbWorkbook in spreadsheet software

#### Usage

    wbWorkbook$open(interactive = NA, flush = FALSE)

#### Arguments

- `interactive`:

  If `FALSE` will throw a warning and not open the path. This can be
  manually set to `TRUE`, otherwise when `NA` (default) uses the value
  returned from
  [`base::interactive()`](https://rdrr.io/r/base/interactive.html)

- `flush`:

  flush

#### Details

minor helper wrapping xl_open which does the entire same thing

#### Returns

The `wbWorkbook`, invisibly

------------------------------------------------------------------------

### Method `buildTable()`

Build table

#### Usage

    wbWorkbook$buildTable(
      sheet = current_sheet(),
      colNames,
      ref,
      showColNames,
      tableStyle,
      tableName,
      withFilter = TRUE,
      totalsRowCount = 0,
      totalLabel = FALSE,
      showFirstColumn = 0,
      showLastColumn = 0,
      showRowStripes = 1,
      showColumnStripes = 0
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `colNames`:

  colNames

- `ref`:

  ref

- `showColNames`:

  showColNames

- `tableStyle`:

  tableStyle

- `tableName`:

  tableName

- `withFilter`:

  withFilter

- `totalsRowCount`:

  totalsRowCount

- `totalLabel`:

  totalLabel

- `showFirstColumn`:

  showFirstColumn

- `showLastColumn`:

  showLastColumn

- `showRowStripes`:

  showRowStripes

- `showColumnStripes`:

  showColumnStripes

#### Returns

The `wbWorksheet` object, invisibly

------------------------------------------------------------------------

### Method `update_table()`

update a data_table

#### Usage

    wbWorkbook$update_table(sheet = current_sheet(), dims = "A1", tabname)

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  Cell range in a sheet

- `tabname`:

  a tablename

#### Returns

The `wbWorksheet` object, invisibly

------------------------------------------------------------------------

### Method `copy_cells()`

copy cells around in a workbook

#### Usage

    wbWorkbook$copy_cells(
      sheet = current_sheet(),
      dims = "A1",
      data,
      as_value = FALSE,
      as_ref = FALSE,
      transpose = FALSE,
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  Cell range in a sheet

- `data`:

  a wb_data object

- `as_value`:

  should a copy of the value be written

- `as_ref`:

  should references to the cell be written

- `transpose`:

  should the data be written transposed

- `...`:

  additional arguments passed to add_data() if used with `as_value`

#### Returns

The `wbWorksheet` object, invisibly

------------------------------------------------------------------------

### Method `get_base_font()`

Get the base font

#### Usage

    wbWorkbook$get_base_font()

#### Returns

A list of of the font

------------------------------------------------------------------------

### Method `set_base_font()`

Set the base font

#### Usage

    wbWorkbook$set_base_font(
      font_size = 11,
      font_color = wb_color(theme = "1"),
      font_name = "Aptos Narrow",
      ...
    )

#### Arguments

- `font_size`:

  fontSize

- `font_color`:

  font_color

- `font_name`:

  font_name

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `get_base_colors()`

Get the base color

#### Usage

    wbWorkbook$get_base_colors(xml = FALSE, plot = TRUE)

#### Arguments

- `xml`:

  xml

- `plot`:

  plot

------------------------------------------------------------------------

### Method `get_base_colours()`

Get the base colour

#### Usage

    wbWorkbook$get_base_colours(xml = FALSE, plot = TRUE)

#### Arguments

- `xml`:

  xml

- `plot`:

  plot

------------------------------------------------------------------------

### Method `set_base_colors()`

Set the base color

#### Usage

    wbWorkbook$set_base_colors(theme = "Office", ...)

#### Arguments

- `theme`:

  theme

- `...`:

  ...

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `set_base_colours()`

Set the base colour

#### Usage

    wbWorkbook$set_base_colours(theme = "Office", ...)

#### Arguments

- `theme`:

  theme

- `...`:

  ...

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `get_bookview()`

Get the book views

#### Usage

    wbWorkbook$get_bookview()

#### Returns

A dataframe with the bookview properties

------------------------------------------------------------------------

### Method `remove_bookview()`

Get the book views

#### Usage

    wbWorkbook$remove_bookview(view = NULL)

#### Arguments

- `view`:

  view

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `set_bookview()`

#### Usage

    wbWorkbook$set_bookview(
      active_tab = NULL,
      auto_filter_date_grouping = NULL,
      first_sheet = NULL,
      minimized = NULL,
      show_horizontal_scroll = NULL,
      show_sheet_tabs = NULL,
      show_vertical_scroll = NULL,
      tab_ratio = NULL,
      visibility = NULL,
      window_height = NULL,
      window_width = NULL,
      x_window = NULL,
      y_window = NULL,
      view = 1L,
      ...
    )

#### Arguments

- `active_tab`:

  activeTab

- `auto_filter_date_grouping`:

  autoFilterDateGrouping

- `first_sheet`:

  firstSheet

- `minimized`:

  minimized

- `show_horizontal_scroll`:

  showHorizontalScroll

- `show_sheet_tabs`:

  showSheetTabs

- `show_vertical_scroll`:

  showVerticalScroll

- `tab_ratio`:

  tabRatio

- `visibility`:

  visibility

- `window_height`:

  windowHeight

- `window_width`:

  windowWidth

- `x_window`:

  xWindow

- `y_window`:

  yWindow

- `view`:

  view

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `get_sheet_names()`

Get sheet names

#### Usage

    wbWorkbook$get_sheet_names(escape = FALSE)

#### Arguments

- `escape`:

  Logical if the xml special characters are escaped

#### Returns

A `named` `character` vector of sheet names in their order. The names
represent the original value of the worksheet prior to any character
substitutions.

------------------------------------------------------------------------

### Method `set_sheet_names()`

Sets a sheet name

#### Usage

    wbWorkbook$set_sheet_names(old = NULL, new)

#### Arguments

- `old`:

  Old sheet name

- `new`:

  New sheet name

#### Returns

The `wbWorkbook` object, invisibly

------------------------------------------------------------------------

### Method `set_row_heights()`

Sets a row height for a sheet

#### Usage

    wbWorkbook$set_row_heights(
      sheet = current_sheet(),
      rows,
      heights = NULL,
      hidden = FALSE,
      hide_blanks = NULL
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `rows`:

  rows

- `heights`:

  heights

- `hidden`:

  hidden

- `hide_blanks`:

  hide_blanks

#### Returns

The `wbWorkbook` object, invisibly

------------------------------------------------------------------------

### Method `remove_row_heights()`

Removes a row height for a sheet

#### Usage

    wbWorkbook$remove_row_heights(sheet = current_sheet(), rows)

#### Arguments

- `sheet`:

  The name of the sheet

- `rows`:

  rows

#### Returns

The `wbWorkbook` object, invisibly

------------------------------------------------------------------------

### Method `createCols()`

creates column object for worksheet

#### Usage

    wbWorkbook$createCols(sheet = current_sheet(), n, beg, end)

#### Arguments

- `sheet`:

  The name of the sheet

- `n`:

  n

- `beg`:

  beg

- `end`:

  end

------------------------------------------------------------------------

### Method `group_cols()`

Group cols

#### Usage

    wbWorkbook$group_cols(
      sheet = current_sheet(),
      cols,
      collapsed = FALSE,
      levels = NULL
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `cols`:

  cols

- `collapsed`:

  collapsed

- `levels`:

  levels

#### Returns

The `wbWorkbook` object, invisibly

------------------------------------------------------------------------

### Method `ungroup_cols()`

ungroup cols

#### Usage

    wbWorkbook$ungroup_cols(sheet = current_sheet(), cols)

#### Arguments

- `sheet`:

  The name of the sheet

- `cols`:

  columns

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `remove_col_widths()`

Remove row heights from a worksheet

#### Usage

    wbWorkbook$remove_col_widths(sheet = current_sheet(), cols)

#### Arguments

- `sheet`:

  A name or index of a worksheet

- `cols`:

  Indices of columns to remove custom width (if any) from.

#### Returns

The `wbWorkbook` object, invisibly

------------------------------------------------------------------------

### Method `set_col_widths()`

Set column widths

#### Usage

    wbWorkbook$set_col_widths(
      sheet = current_sheet(),
      cols,
      widths = 8.43,
      hidden = FALSE
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `cols`:

  cols

- `widths`:

  Width of columns

- `hidden`:

  A logical vector to determine which cols are hidden; values are
  repeated across length of `cols`

#### Returns

The `wbWorkbook` object, invisibly

------------------------------------------------------------------------

### Method `group_rows()`

Group rows

#### Usage

    wbWorkbook$group_rows(
      sheet = current_sheet(),
      rows,
      collapsed = FALSE,
      levels = NULL
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `rows`:

  rows

- `collapsed`:

  collapsed

- `levels`:

  levels

#### Returns

The `wbWorkbook` object, invisibly

------------------------------------------------------------------------

### Method `ungroup_rows()`

ungroup rows

#### Usage

    wbWorkbook$ungroup_rows(sheet = current_sheet(), rows)

#### Arguments

- `sheet`:

  The name of the sheet

- `rows`:

  rows

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `remove_worksheet()`

Remove a worksheet

#### Usage

    wbWorkbook$remove_worksheet(sheet = current_sheet())

#### Arguments

- `sheet`:

  The worksheet to delete

#### Returns

The `wbWorkbook` object, invisibly

------------------------------------------------------------------------

### Method `add_data_validation()`

Adds data validation

#### Usage

    wbWorkbook$add_data_validation(
      sheet = current_sheet(),
      dims = "A1",
      type,
      operator,
      value,
      allow_blank = TRUE,
      show_input_msg = TRUE,
      show_error_msg = TRUE,
      error_style = NULL,
      error_title = NULL,
      error = NULL,
      prompt_title = NULL,
      prompt = NULL,
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  Cell range in a sheet

- `type`:

  type

- `operator`:

  operator

- `value`:

  value

- `allow_blank`:

  allowBlank

- `show_input_msg`:

  showInputMsg

- `show_error_msg`:

  showErrorMsg

- `error_style`:

  The icon shown and the options how to deal with such inputs. Default
  "stop" (cancel), else "information" (prompt popup) or "warning"
  (prompt accept or change input)

- `error_title`:

  The error title

- `error`:

  The error text

- `prompt_title`:

  The prompt title

- `prompt`:

  The prompt text

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `merge_cells()`

Set cell merging for a sheet

#### Usage

    wbWorkbook$merge_cells(
      sheet = current_sheet(),
      dims = NULL,
      solve = FALSE,
      direction = NULL,
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  Cell range in a sheet

- `solve`:

  logical if intersecting cells should be solved

- `direction`:

  direction in which to split the cell merging. Allows "row" or "col".

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object, invisibly

------------------------------------------------------------------------

### Method `unmerge_cells()`

Removes cell merging for a sheet

#### Usage

    wbWorkbook$unmerge_cells(sheet = current_sheet(), dims = NULL, ...)

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  Cell range in a sheet

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object, invisibly

------------------------------------------------------------------------

### Method `freeze_pane()`

Set freeze panes for a sheet

#### Usage

    wbWorkbook$freeze_pane(
      sheet = current_sheet(),
      first_active_row = NULL,
      first_active_col = NULL,
      first_row = FALSE,
      first_col = FALSE,
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `first_active_row`:

  first_active_row

- `first_active_col`:

  first_active_col

- `first_row`:

  first_row

- `first_col`:

  first_col

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object, invisibly

------------------------------------------------------------------------

### Method `add_comment()`

Add comment

#### Usage

    wbWorkbook$add_comment(sheet = current_sheet(), dims = "A1", comment, ...)

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  row and column as spreadsheet dimension, e.g. "A1"

- `comment`:

  a comment to apply to the worksheet

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `get_comment()`

Get comments

#### Usage

    wbWorkbook$get_comment(sheet = current_sheet(), dims = NULL)

#### Arguments

- `sheet`:

  sheet

- `dims`:

  dims

#### Returns

A data frame containing comments

------------------------------------------------------------------------

### Method [`remove_comment()`](https://janmarvin.github.io/openxlsx2/reference/comment_internal.md)

Remove comment

#### Usage

    wbWorkbook$remove_comment(sheet = current_sheet(), dims = "A1", ...)

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  row and column as spreadsheet dimension, e.g. "A1"

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `add_thread()`

add threaded comment to worksheet

#### Usage

    wbWorkbook$add_thread(
      sheet = current_sheet(),
      dims = "A1",
      comment = NULL,
      person_id,
      reply = FALSE,
      resolve = FALSE
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  Cell range in a sheet

- `comment`:

  the comment to add

- `person_id`:

  the person Id this should be added for

- `reply`:

  logical if the comment is a reply

- `resolve`:

  logical if the comment should be marked as resolved

------------------------------------------------------------------------

### Method `get_thread()`

Get threads

#### Usage

    wbWorkbook$get_thread(sheet = current_sheet(), dims = NULL)

#### Arguments

- `sheet`:

  sheet

- `dims`:

  dims

#### Returns

A data frame containing threads

------------------------------------------------------------------------

### Method `add_conditional_formatting()`

Add conditional formatting

#### Usage

    wbWorkbook$add_conditional_formatting(
      sheet = current_sheet(),
      dims = NULL,
      rule = NULL,
      style = NULL,
      type = c("expression", "colorScale", "dataBar", "iconSet", "duplicatedValues",
        "uniqueValues", "containsErrors", "notContainsErrors", "containsBlanks",
        "notContainsBlanks", "containsText", "notContainsText", "beginsWith", "endsWith",
        "between", "topN", "bottomN"),
      params = list(showValue = TRUE, gradient = TRUE, border = TRUE, percent = FALSE, rank =
        5L, axisPosition = "automatic"),
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  Cell range in a sheet

- `rule`:

  rule

- `style`:

  style

- `type`:

  type

- `params`:

  Additional parameters

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `remove_conditional_formatting()`

Remove conditional formatting

#### Usage

    wbWorkbook$remove_conditional_formatting(
      sheet = current_sheet(),
      dims = NULL,
      first = FALSE,
      last = FALSE
    )

#### Arguments

- `sheet`:

  sheet

- `dims`:

  dims

- `first`:

  first

- `last`:

  last

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `add_image()`

Insert an image into a sheet

#### Usage

    wbWorkbook$add_image(
      sheet = current_sheet(),
      dims = "A1",
      file,
      width = 6,
      height = 3,
      row_offset = 0,
      col_offset = 0,
      units = "in",
      dpi = 300,
      address = NULL,
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  Cell range in a sheet

- `file`:

  file

- `width`:

  width

- `height`:

  height

- `row_offset, col_offset`:

  offsets

- `units`:

  units

- `dpi`:

  dpi

- `address`:

  address

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object, invisibly

------------------------------------------------------------------------

### Method `add_plot()`

Add plot. A wrapper for add_image()

#### Usage

    wbWorkbook$add_plot(
      sheet = current_sheet(),
      dims = "A1",
      width = 6,
      height = 4,
      row_offset = 0,
      col_offset = 0,
      file_type = "png",
      units = "in",
      dpi = 300,
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  Cell range in a sheet

- `width`:

  width

- `height`:

  height

- `row_offset, col_offset`:

  offsets

- `file_type`:

  fileType

- `units`:

  units

- `dpi`:

  dpi

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `add_drawing()`

Add xml drawing

#### Usage

    wbWorkbook$add_drawing(
      sheet = current_sheet(),
      dims = "A1",
      xml,
      col_offset = 0,
      row_offset = 0,
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  Cell range in a sheet

- `xml`:

  xml

- `col_offset, row_offset`:

  offsets for column and row

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `add_chart_xml()`

Add xml chart

#### Usage

    wbWorkbook$add_chart_xml(
      sheet = current_sheet(),
      dims = NULL,
      xml,
      col_offset = 0,
      row_offset = 0,
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  Cell range in a sheet

- `xml`:

  xml

- `col_offset, row_offset`:

  positioning parameters

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `add_mschart()`

Add mschart chart to the workbook

#### Usage

    wbWorkbook$add_mschart(
      sheet = current_sheet(),
      dims = NULL,
      graph,
      col_offset = 0,
      row_offset = 0,
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  the dimensions where the sheet will appear

- `graph`:

  mschart graph

- `col_offset, row_offset`:

  offsets for column and row

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `add_form_control()`

Add form control to workbook

#### Usage

    wbWorkbook$add_form_control(
      sheet = current_sheet(),
      dims = "A1",
      type = c("Checkbox", "Radio", "Drop"),
      text = NULL,
      link = NULL,
      range = NULL,
      checked = FALSE
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  Cell range in a sheet

- `type`:

  type

- `text`:

  text

- `link`:

  link

- `range`:

  range

- `checked`:

  checked

#### Returns

The `wbWorkbook` object, invisibly

------------------------------------------------------------------------

### Method [`print()`](https://rdrr.io/r/base/print.html)

Prints the `wbWorkbook` object

#### Usage

    wbWorkbook$print()

#### Returns

The `wbWorkbook` object, invisibly; called for its side-effects

------------------------------------------------------------------------

### Method `protect()`

Protect a workbook

#### Usage

    wbWorkbook$protect(
      protect = TRUE,
      password = NULL,
      lock_structure = FALSE,
      lock_windows = FALSE,
      type = 1,
      file_sharing = FALSE,
      username = unname(Sys.info()["user"]),
      read_only_recommended = FALSE,
      ...
    )

#### Arguments

- `protect`:

  protect

- `password`:

  password

- `lock_structure`:

  lock_structure

- `lock_windows`:

  lock_windows

- `type`:

  type

- `file_sharing`:

  file_sharing

- `username`:

  username

- `read_only_recommended`:

  read_only_recommended

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object, invisibly

------------------------------------------------------------------------

### Method `protect_worksheet()`

protect worksheet

#### Usage

    wbWorkbook$protect_worksheet(
      sheet = current_sheet(),
      protect = TRUE,
      password = NULL,
      properties = NULL
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `protect`:

  protect

- `password`:

  password

- `properties`:

  A character vector of properties to lock. Can be one or more of the
  following: `"selectLockedCells"`, `"selectUnlockedCells"`,
  `"formatCells"`, `"formatColumns"`, `"formatRows"`, `"insertColumns"`,
  `"insertRows"`, `"insertHyperlinks"`, `"deleteColumns"`,
  `"deleteRows"`, `"sort"`, `"autoFilter"`, `"pivotTables"`,
  `"objects"`, `"scenarios"`

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `get_properties()`

Get properties of a workbook

#### Usage

    wbWorkbook$get_properties()

------------------------------------------------------------------------

### Method `set_properties()`

Set a property of a workbook

#### Usage

    wbWorkbook$set_properties(
      creator = NULL,
      title = NULL,
      subject = NULL,
      category = NULL,
      datetime_created = NULL,
      datetime_modified = NULL,
      modifier = NULL,
      keywords = NULL,
      comments = NULL,
      manager = NULL,
      company = NULL,
      custom = NULL
    )

#### Arguments

- `creator`:

  character vector of creators. Duplicated are ignored.

- `title, subject, category, datetime_created, datetime_modified, modifier, keywords, comments, manager, company, custom`:

  A workbook property to set

------------------------------------------------------------------------

### Method `add_mips()`

add mips string

#### Usage

    wbWorkbook$add_mips(xml = NULL)

#### Arguments

- `xml`:

  A mips string added to self\$custom

------------------------------------------------------------------------

### Method `get_mips()`

get mips string

#### Usage

    wbWorkbook$get_mips(single_xml = TRUE, quiet = TRUE)

#### Arguments

- `single_xml`:

  single_xml

- `quiet`:

  quiet

------------------------------------------------------------------------

### Method `set_creators()`

Set creator(s)

#### Usage

    wbWorkbook$set_creators(creators)

#### Arguments

- `creators`:

  A character vector of creators to set. Duplicates are ignored.

------------------------------------------------------------------------

### Method `add_creators()`

Add creator(s)

#### Usage

    wbWorkbook$add_creators(creators)

#### Arguments

- `creators`:

  A character vector of creators to add. Duplicates are ignored.

------------------------------------------------------------------------

### Method `remove_creators()`

Remove creator(s)

#### Usage

    wbWorkbook$remove_creators(creators)

#### Arguments

- `creators`:

  A character vector of creators to remove. All duplicated are removed.

------------------------------------------------------------------------

### Method `set_last_modified_by()`

Change the last modified by

#### Usage

    wbWorkbook$set_last_modified_by(name, ...)

#### Arguments

- `name`:

  A new value

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object, invisibly

------------------------------------------------------------------------

### Method `set_page_setup()`

set_page_setup() this function is intended to supersede page_setup(),
but is not yet stable

#### Usage

    wbWorkbook$set_page_setup(
      sheet = current_sheet(),
      black_and_white = NULL,
      cell_comments = NULL,
      copies = NULL,
      draft = NULL,
      errors = NULL,
      first_page_number = NULL,
      id = NULL,
      page_order = NULL,
      paper_height = NULL,
      paper_width = NULL,
      hdpi = NULL,
      vdpi = NULL,
      use_first_page_number = NULL,
      use_printer_defaults = NULL,
      orientation = NULL,
      scale = NULL,
      left = 0.7,
      right = 0.7,
      top = 0.75,
      bottom = 0.75,
      header = 0.3,
      footer = 0.3,
      fit_to_width = FALSE,
      fit_to_height = FALSE,
      paper_size = NULL,
      print_title_rows = NULL,
      print_title_cols = NULL,
      summary_row = NULL,
      summary_col = NULL,
      tab_color = NULL,
      horizontal_centered = NULL,
      vertical_centered = NULL,
      print_headings = NULL,
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `black_and_white`:

  black_and_white

- `cell_comments`:

  cell_comment

- `copies`:

  copies

- `draft`:

  draft

- `errors`:

  errors

- `first_page_number`:

  first_page_number

- `id`:

  id

- `page_order`:

  page_order

- `paper_height, paper_width`:

  paper size

- `hdpi, vdpi`:

  horizontal and vertical dpi

- `use_first_page_number`:

  use_first_page_number

- `use_printer_defaults`:

  use_printer_defaults

- `orientation`:

  orientation

- `scale`:

  scale

- `left`:

  left

- `right`:

  right

- `top`:

  top

- `bottom`:

  bottom

- `header`:

  header

- `footer`:

  footer

- `fit_to_width`:

  fitToWidth

- `fit_to_height`:

  fitToHeight

- `paper_size`:

  paperSize

- `print_title_rows`:

  printTitleRows

- `print_title_cols`:

  printTitleCols

- `summary_row`:

  summaryRow

- `summary_col`:

  summaryCol

- `tab_color`:

  tabColor

- `horizontal_centered`:

  horizontal_centered

- `vertical_centered`:

  vertical_centered

- `print_headings`:

  print_headings

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object, invisibly

------------------------------------------------------------------------

### Method `page_setup()`

page_setup()

#### Usage

    wbWorkbook$page_setup(
      sheet = current_sheet(),
      orientation = NULL,
      scale = 100,
      left = 0.7,
      right = 0.7,
      top = 0.75,
      bottom = 0.75,
      header = 0.3,
      footer = 0.3,
      fit_to_width = FALSE,
      fit_to_height = FALSE,
      paper_size = NULL,
      print_title_rows = NULL,
      print_title_cols = NULL,
      summary_row = NULL,
      summary_col = NULL,
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `orientation`:

  orientation

- `scale`:

  scale

- `left`:

  left

- `right`:

  right

- `top`:

  top

- `bottom`:

  bottom

- `header`:

  header

- `footer`:

  footer

- `fit_to_width`:

  fitToWidth

- `fit_to_height`:

  fitToHeight

- `paper_size`:

  paperSize

- `print_title_rows`:

  printTitleRows

- `print_title_cols`:

  printTitleCols

- `summary_row`:

  summaryRow

- `summary_col`:

  summaryCol

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object, invisibly

------------------------------------------------------------------------

### Method `set_header_footer()`

Sets headers and footers

#### Usage

    wbWorkbook$set_header_footer(
      sheet = current_sheet(),
      header = NULL,
      footer = NULL,
      even_header = NULL,
      even_footer = NULL,
      first_header = NULL,
      first_footer = NULL,
      align_with_margins = NULL,
      scale_with_doc = NULL,
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `header`:

  header

- `footer`:

  footer

- `even_header`:

  evenHeader

- `even_footer`:

  evenFooter

- `first_header`:

  firstHeader

- `first_footer`:

  firstFooter

- `align_with_margins`:

  align_with_margins

- `scale_with_doc`:

  scale_with_doc

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object, invisibly

------------------------------------------------------------------------

### Method `get_tables()`

get tables

#### Usage

    wbWorkbook$get_tables(sheet = current_sheet())

#### Arguments

- `sheet`:

  The name of the sheet

#### Returns

The sheet tables. [`character()`](https://rdrr.io/r/base/character.html)
if empty

------------------------------------------------------------------------

### Method `remove_tables()`

remove tables

#### Usage

    wbWorkbook$remove_tables(sheet = current_sheet(), table, remove_data = TRUE)

#### Arguments

- `sheet`:

  The name of the sheet

- `table`:

  table

- `remove_data`:

  removes the data as well

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `add_filter()`

add filters

#### Usage

    wbWorkbook$add_filter(sheet = current_sheet(), rows, cols)

#### Arguments

- `sheet`:

  The name of the sheet

- `rows`:

  rows

- `cols`:

  cols

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `remove_filter()`

remove filters

#### Usage

    wbWorkbook$remove_filter(sheet = current_sheet())

#### Arguments

- `sheet`:

  The name of the sheet

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `set_grid_lines()`

grid lines

#### Usage

    wbWorkbook$set_grid_lines(sheet = current_sheet(), show = FALSE, print = show)

#### Arguments

- `sheet`:

  The name of the sheet

- `show`:

  show

- `print`:

  print

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `grid_lines()`

grid lines

#### Usage

    wbWorkbook$grid_lines(sheet = current_sheet(), show = FALSE, print = show)

#### Arguments

- `sheet`:

  The name of the sheet

- `show`:

  show

- `print`:

  print

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `add_named_region()`

add a named region

#### Usage

    wbWorkbook$add_named_region(
      sheet = current_sheet(),
      dims = "A1",
      name,
      local_sheet = FALSE,
      overwrite = FALSE,
      comment = NULL,
      hidden = NULL,
      custom_menu = NULL,
      description = NULL,
      is_function = NULL,
      function_group_id = NULL,
      help = NULL,
      local_name = NULL,
      publish_to_server = NULL,
      status_bar = NULL,
      vb_procedure = NULL,
      workbook_parameter = NULL,
      xml = NULL,
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  Cell range in a sheet

- `name`:

  name

- `local_sheet`:

  local_sheet

- `overwrite`:

  overwrite

- `comment`:

  comment

- `hidden`:

  hidden

- `custom_menu`:

  custom_menu

- `description`:

  description

- `is_function`:

  function

- `function_group_id`:

  function group id

- `help`:

  help

- `local_name`:

  localName

- `publish_to_server`:

  publish to server

- `status_bar`:

  status bar

- `vb_procedure`:

  vb procedure

- `workbook_parameter`:

  workbookParameter

- `xml`:

  xml

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `get_named_regions()`

get named regions in a workbook

#### Usage

    wbWorkbook$get_named_regions(tables = FALSE, x = NULL)

#### Arguments

- `tables`:

  Return tables as well?

- `x`:

  Not used.

#### Returns

A `data.frame` of named regions

------------------------------------------------------------------------

### Method `remove_named_region()`

remove a named region

#### Usage

    wbWorkbook$remove_named_region(sheet = current_sheet(), name = NULL)

#### Arguments

- `sheet`:

  The name of the sheet

- `name`:

  name

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `set_order()`

set worksheet order

#### Usage

    wbWorkbook$set_order(sheets)

#### Arguments

- `sheets`:

  sheets

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `get_sheet_visibility()`

Get sheet visibility

#### Usage

    wbWorkbook$get_sheet_visibility()

#### Returns

Returns sheet visibility

------------------------------------------------------------------------

### Method `set_sheet_visibility()`

Set sheet visibility

#### Usage

    wbWorkbook$set_sheet_visibility(sheet = current_sheet(), value)

#### Arguments

- `sheet`:

  The name of the sheet

- `value`:

  value

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `add_page_break()`

Add a page break

#### Usage

    wbWorkbook$add_page_break(sheet = current_sheet(), row = NULL, col = NULL)

#### Arguments

- `sheet`:

  The name of the sheet

- `row`:

  row

- `col`:

  col

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `clean_sheet()`

clean sheet (remove all values)

#### Usage

    wbWorkbook$clean_sheet(
      sheet = current_sheet(),
      dims = NULL,
      numbers = TRUE,
      characters = TRUE,
      styles = TRUE,
      merged_cells = TRUE,
      hyperlinks = TRUE
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  Cell range in a sheet

- `numbers`:

  remove all numbers

- `characters`:

  remove all characters

- `styles`:

  remove all styles

- `merged_cells`:

  remove all merged_cells

- `hyperlinks`:

  remove all hyperlinks

#### Returns

The `wbWorksheetObject`, invisibly

------------------------------------------------------------------------

### Method `add_border()`

create borders for cell region

#### Usage

    wbWorkbook$add_border(
      sheet = current_sheet(),
      dims = "A1",
      bottom_color = wb_color(hex = "FF000000"),
      left_color = wb_color(hex = "FF000000"),
      right_color = wb_color(hex = "FF000000"),
      top_color = wb_color(hex = "FF000000"),
      bottom_border = "thin",
      left_border = "thin",
      right_border = "thin",
      top_border = "thin",
      inner_hgrid = NULL,
      inner_hcolor = NULL,
      inner_vgrid = NULL,
      inner_vcolor = NULL,
      update = FALSE,
      diagonal_up = NULL,
      diagonal_down = NULL,
      diagonal_color = NULL,
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  dimensions on the worksheet e.g. "A1", "A1:A5", "A1:H5"

- `bottom_color, left_color, right_color, top_color, inner_hcolor, inner_vcolor`:

  a color, either something openxml knows or some RGB color

- `left_border, right_border, top_border, bottom_border, inner_hgrid, inner_vgrid`:

  the border style, if NULL no border is drawn. See create_border for
  possible border styles

- `update`:

  update

- `diagonal_up, diagonal_down, diagonal_color`:

  (optional) arguments for diagonal border lines

- `...`:

  additional arguments

#### Returns

The `wbWorkbook`, invisibly

------------------------------------------------------------------------

### Method `add_fill()`

provide simple fill function

#### Usage

    wbWorkbook$add_fill(
      sheet = current_sheet(),
      dims = "A1",
      color = wb_color(hex = "FFFFFF00"),
      pattern = "solid",
      gradient_fill = "",
      every_nth_col = 1,
      every_nth_row = 1,
      bg_color = NULL,
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  Cell range in a sheet

- `color`:

  the colors to apply, e.g. yellow: wb_color(hex = "FFFFFF00")

- `pattern`:

  various default "none" but others are possible: "solid", "mediumGray",
  "darkGray", "lightGray", "darkHorizontal", "darkVertical", "darkDown",
  "darkUp", "darkGrid", "darkTrellis", "lightHorizontal",
  "lightVertical", "lightDown", "lightUp", "lightGrid", "lightTrellis",
  "gray125", "gray0625"

- `gradient_fill`:

  a gradient fill xml pattern.

- `every_nth_col`:

  which col should be filled

- `every_nth_row`:

  which row should be filled

- `bg_color`:

  (optional) background
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/reference/wb_color.md)

- `...`:

  additional arguments

#### Returns

The `wbWorksheetObject`, invisibly

------------------------------------------------------------------------

### Method `add_font()`

provide simple font function

#### Usage

    wbWorkbook$add_font(
      sheet = current_sheet(),
      dims = "A1",
      name = "Aptos Narrow",
      color = wb_color(hex = "FF000000"),
      size = "11",
      bold = "",
      italic = "",
      outline = "",
      strike = "",
      underline = "",
      charset = "",
      condense = "",
      extend = "",
      family = "",
      scheme = "",
      shadow = "",
      vert_align = "",
      update = FALSE,
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  Cell range in a sheet

- `name`:

  font name: default "Aptos Narrow"

- `color`:

  rgb color: default "FF000000"

- `size`:

  font size: default "11",

- `bold`:

  bold

- `italic`:

  italic

- `outline`:

  outline

- `strike`:

  strike

- `underline`:

  underline

- `charset`:

  charset

- `condense`:

  condense

- `extend`:

  extend

- `family`:

  font family

- `scheme`:

  font scheme

- `shadow`:

  shadow

- `vert_align`:

  vertical alignment

- `update`:

  update

- `...`:

  additional arguments

#### Returns

The `wbWorkbook`, invisibly

------------------------------------------------------------------------

### Method `add_numfmt()`

provide simple number format function

#### Usage

    wbWorkbook$add_numfmt(sheet = current_sheet(), dims = "A1", numfmt)

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  Cell range in a sheet

- `numfmt`:

  number format id or a character of the format

#### Returns

The `wbWorksheetObject`, invisibly

------------------------------------------------------------------------

### Method `add_cell_style()`

provide simple cell style format function

#### Usage

    wbWorkbook$add_cell_style(
      sheet = current_sheet(),
      dims = "A1",
      apply_alignment = NULL,
      apply_border = NULL,
      apply_fill = NULL,
      apply_font = NULL,
      apply_number_format = NULL,
      apply_protection = NULL,
      border_id = NULL,
      ext_lst = NULL,
      fill_id = NULL,
      font_id = NULL,
      hidden = NULL,
      horizontal = NULL,
      indent = NULL,
      justify_last_line = NULL,
      locked = NULL,
      num_fmt_id = NULL,
      pivot_button = NULL,
      quote_prefix = NULL,
      reading_order = NULL,
      relative_indent = NULL,
      shrink_to_fit = NULL,
      text_rotation = NULL,
      vertical = NULL,
      wrap_text = NULL,
      xf_id = NULL,
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  Cell range in a sheet

- `apply_alignment`:

  logical apply alignment

- `apply_border`:

  logical apply border

- `apply_fill`:

  logical apply fill

- `apply_font`:

  logical apply font

- `apply_number_format`:

  logical apply number format

- `apply_protection`:

  logical apply protection

- `border_id`:

  border ID to apply

- `ext_lst`:

  extension list something like `<extLst>...</extLst>`

- `fill_id`:

  fill ID to apply

- `font_id`:

  font ID to apply

- `hidden`:

  logical cell is hidden

- `horizontal`:

  align content horizontal ('left', 'center', 'right')

- `indent`:

  logical indent content

- `justify_last_line`:

  logical justify last line

- `locked`:

  logical cell is locked

- `num_fmt_id`:

  number format ID to apply

- `pivot_button`:

  unknown

- `quote_prefix`:

  unknown

- `reading_order`:

  reading order left to right

- `relative_indent`:

  relative indentation

- `shrink_to_fit`:

  logical shrink to fit

- `text_rotation`:

  degrees of text rotation

- `vertical`:

  vertical alignment of content ('top', 'center', 'bottom')

- `wrap_text`:

  wrap text in cell

- `xf_id`:

  xf ID to apply

- `...`:

  additional arguments

#### Returns

The `wbWorkbook` object, invisibly

------------------------------------------------------------------------

### Method `get_cell_style()`

get sheet style

#### Usage

    wbWorkbook$get_cell_style(sheet = current_sheet(), dims)

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  Cell range in a sheet

#### Returns

a character vector of cell styles

------------------------------------------------------------------------

### Method `set_cell_style()`

set sheet style

#### Usage

    wbWorkbook$set_cell_style(sheet = current_sheet(), dims, style)

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  Cell range in a sheet

- `style`:

  style

#### Returns

The `wbWorksheetObject`, invisibly

------------------------------------------------------------------------

### Method `set_cell_style_across()`

set style across columns and/or rows

#### Usage

    wbWorkbook$set_cell_style_across(
      sheet = current_sheet(),
      style,
      cols = NULL,
      rows = NULL
    )

#### Arguments

- `sheet`:

  sheet

- `style`:

  style

- `cols`:

  cols

- `rows`:

  rows

#### Returns

The `wbWorkbook` object

------------------------------------------------------------------------

### Method `add_named_style()`

set sheet style

#### Usage

    wbWorkbook$add_named_style(
      sheet = current_sheet(),
      dims = "A1",
      name = "Normal",
      font_name = NULL,
      font_size = NULL
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  Cell range in a sheet

- `name`:

  name

- `font_name, font_size`:

  optional else the default of the theme

#### Returns

The `wbWorkbook`, invisibly

------------------------------------------------------------------------

### Method `add_dxfs_style()`

create dxfs style These styles are used with conditional formatting and
custom table styles

#### Usage

    wbWorkbook$add_dxfs_style(
      name,
      font_name = NULL,
      font_size = NULL,
      font_color = NULL,
      num_fmt = NULL,
      border = NULL,
      border_color = wb_color(getOption("openxlsx2.borderColor", "black")),
      border_style = getOption("openxlsx2.borderStyle", "thin"),
      bg_fill = NULL,
      gradient_fill = NULL,
      text_bold = NULL,
      text_italic = NULL,
      text_underline = NULL,
      ...
    )

#### Arguments

- `name`:

  the style name

- `font_name`:

  the font name

- `font_size`:

  the font size

- `font_color`:

  the font color (a
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/reference/wb_color.md)
  object)

- `num_fmt`:

  the number format

- `border`:

  logical if borders are applied

- `border_color`:

  the border color

- `border_style`:

  the border style

- `bg_fill`:

  any background fill

- `gradient_fill`:

  any gradient fill

- `text_bold`:

  logical if text is bold

- `text_italic`:

  logical if text is italic

- `text_underline`:

  logical if text is underlined

- `...`:

  additional arguments passed to
  [`create_dxfs_style()`](https://janmarvin.github.io/openxlsx2/reference/create_dxfs_style.md)

#### Returns

The `wbWorksheetObject`, invisibly

------------------------------------------------------------------------

### Method `clone_sheet_style()`

clone style from one sheet to another

#### Usage

    wbWorkbook$clone_sheet_style(from = current_sheet(), to)

#### Arguments

- `from`:

  the worksheet you are cloning

- `to`:

  the worksheet the style is applied to

------------------------------------------------------------------------

### Method `add_sparklines()`

apply sparkline to worksheet

#### Usage

    wbWorkbook$add_sparklines(sheet = current_sheet(), sparklines)

#### Arguments

- `sheet`:

  The name of the sheet

- `sparklines`:

  sparkline created by `create_sparkline()`

------------------------------------------------------------------------

### Method `add_ignore_error()`

Ignore error on worksheet

#### Usage

    wbWorkbook$add_ignore_error(
      sheet = current_sheet(),
      dims = "A1",
      calculated_column = FALSE,
      empty_cell_reference = FALSE,
      eval_error = FALSE,
      formula = FALSE,
      formula_range = FALSE,
      list_data_validation = FALSE,
      number_stored_as_text = FALSE,
      two_digit_text_year = FALSE,
      unlocked_formula = FALSE,
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `dims`:

  Cell range in a sheet

- `calculated_column`:

  calculatedColumn

- `empty_cell_reference`:

  emptyCellReference

- `eval_error`:

  evalError

- `formula`:

  formula

- `formula_range`:

  formulaRange

- `list_data_validation`:

  listDataValidation

- `number_stored_as_text`:

  numberStoredAsText

- `two_digit_text_year`:

  twoDigitTextYear

- `unlocked_formula`:

  unlockedFormula

- `...`:

  additional arguments

------------------------------------------------------------------------

### Method `set_sheetview()`

add sheetview

#### Usage

    wbWorkbook$set_sheetview(
      sheet = current_sheet(),
      color_id = NULL,
      default_grid_color = NULL,
      right_to_left = NULL,
      show_formulas = NULL,
      show_grid_lines = NULL,
      show_outline_symbols = NULL,
      show_row_col_headers = NULL,
      show_ruler = NULL,
      show_white_space = NULL,
      show_zeros = NULL,
      tab_selected = NULL,
      top_left_cell = NULL,
      view = NULL,
      window_protection = NULL,
      workbook_view_id = NULL,
      zoom_scale = NULL,
      zoom_scale_normal = NULL,
      zoom_scale_page_layout_view = NULL,
      zoom_scale_sheet_layout_view = NULL,
      ...
    )

#### Arguments

- `sheet`:

  The name of the sheet

- `color_id, default_grid_color`:

  Integer: A color, default is 64

- `right_to_left`:

  Logical: if TRUE column ordering is right to left

- `show_formulas`:

  Logical: if TRUE cell formulas are shown

- `show_grid_lines`:

  Logical: if TRUE the worksheet grid is shown

- `show_outline_symbols`:

  Logical: if TRUE outline symbols are shown

- `show_row_col_headers`:

  Logical: if TRUE row and column headers are shown

- `show_ruler`:

  Logical: if TRUE a ruler is shown in page layout view

- `show_white_space`:

  Logical: if TRUE margins are shown in page layout view

- `show_zeros`:

  Logical: if FALSE cells containing zero are shown blank if
  !showFormulas

- `tab_selected`:

  Integer: zero vector indicating the selected tab

- `top_left_cell`:

  Cell: the cell shown in the top left corner / or top right with
  rightToLeft

- `view`:

  View: "normal", "pageBreakPreview" or "pageLayout"

- `window_protection`:

  Logical: if TRUE the panes are protected

- `workbook_view_id`:

  integer: Pointing to some other view inside the workbook

- `zoom_scale, zoom_scale_normal, zoom_scale_page_layout_view, zoom_scale_sheet_layout_view`:

  Integer: the zoom scale should be between 10 and 400. These are values
  for current, normal etc.

- `...`:

  additional arguments

#### Returns

The `wbWorksheetObject`, invisibly

------------------------------------------------------------------------

### Method `add_person()`

add person to workbook

#### Usage

    wbWorkbook$add_person(
      name = NULL,
      id = NULL,
      user_id = NULL,
      provider_id = "None"
    )

#### Arguments

- `name`:

  name

- `id`:

  id

- `user_id`:

  user_id

- `provider_id`:

  provider_id

------------------------------------------------------------------------

### Method `get_person()`

description get person

#### Usage

    wbWorkbook$get_person(name = NULL)

#### Arguments

- `name`:

  name

------------------------------------------------------------------------

### Method `get_active_sheet()`

description get active sheet

#### Usage

    wbWorkbook$get_active_sheet()

------------------------------------------------------------------------

### Method `set_active_sheet()`

description set active sheet

#### Usage

    wbWorkbook$set_active_sheet(sheet = current_sheet())

#### Arguments

- `sheet`:

  The name of the sheet

------------------------------------------------------------------------

### Method `get_selected()`

description get selected sheets

#### Usage

    wbWorkbook$get_selected()

------------------------------------------------------------------------

### Method `set_selected()`

set selected sheet

#### Usage

    wbWorkbook$set_selected(sheet = current_sheet())

#### Arguments

- `sheet`:

  The name of the sheet

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    wbWorkbook$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
