# Package index

## Read tabular data

Read data out of an xlsx file as an R object

- [`wb_to_df()`](https://janmarvin.github.io/openxlsx2/reference/wb_to_df.md)
  [`read_xlsx()`](https://janmarvin.github.io/openxlsx2/reference/wb_to_df.md)
  [`wb_read()`](https://janmarvin.github.io/openxlsx2/reference/wb_to_df.md)
  : Create a data frame from a Workbook

## Write tabular data

Write tabular data frame(s) to file with options

- [`write_xlsx()`](https://janmarvin.github.io/openxlsx2/reference/write_xlsx.md)
  : Write data to an xlsx file

## Load a workbook

Load and interact with an existing workbook

- [`wb_load()`](https://janmarvin.github.io/openxlsx2/reference/wb_load.md)
  : Load an existing .xlsx, .xlsm or .xlsb file
- [`wbWorkbook`](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.md)
  : Workbook class

## Create a workbook

Add content to a help you interact to a new or existing workbook.

- [`wb_workbook()`](https://janmarvin.github.io/openxlsx2/reference/wb_workbook.md)
  : Create a new Workbook object

## Add content to a workbook

Add data or images to a new or existing workbook

### Add a sheet to a workbook

- [`wb_add_worksheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_worksheet.md)
  : Add a worksheet to a workbook

### Add to a worksheet

These functions help you write in worksheets. They invisibly return the
`wbWorkbook` object, except for `get_` functions, who return a character
vector, unless specified otherwise.

- [`wb_set_col_widths()`](https://janmarvin.github.io/openxlsx2/reference/col_widths-wb.md)
  [`wb_remove_col_widths()`](https://janmarvin.github.io/openxlsx2/reference/col_widths-wb.md)
  : Modify column widths of a worksheet
- [`wb_add_filter()`](https://janmarvin.github.io/openxlsx2/reference/filter-wb.md)
  [`wb_remove_filter()`](https://janmarvin.github.io/openxlsx2/reference/filter-wb.md)
  : Add/remove column filters in a worksheet
- [`wb_group_cols()`](https://janmarvin.github.io/openxlsx2/reference/grouping-wb.md)
  [`wb_ungroup_cols()`](https://janmarvin.github.io/openxlsx2/reference/grouping-wb.md)
  [`wb_group_rows()`](https://janmarvin.github.io/openxlsx2/reference/grouping-wb.md)
  [`wb_ungroup_rows()`](https://janmarvin.github.io/openxlsx2/reference/grouping-wb.md)
  : Group rows and columns in a worksheet
- [`wb_add_named_region()`](https://janmarvin.github.io/openxlsx2/reference/named_region-wb.md)
  [`wb_remove_named_region()`](https://janmarvin.github.io/openxlsx2/reference/named_region-wb.md)
  [`wb_get_named_regions()`](https://janmarvin.github.io/openxlsx2/reference/named_region-wb.md)
  : Modify named regions in a worksheet
- [`wb_set_row_heights()`](https://janmarvin.github.io/openxlsx2/reference/row_heights-wb.md)
  [`wb_remove_row_heights()`](https://janmarvin.github.io/openxlsx2/reference/row_heights-wb.md)
  : Modify row heights of a worksheet
- [`wb_add_conditional_formatting()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_conditional_formatting.md)
  [`wb_remove_conditional_formatting()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_conditional_formatting.md)
  : Add conditional formatting to cells in a worksheet
- [`wb_add_data()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data.md)
  : Add data to a worksheet
- [`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data_table.md)
  : Add a data table to a worksheet
- [`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_formula.md)
  : Add a formula to a cell range in a worksheet
- [`wb_add_hyperlink()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_hyperlink.md)
  [`wb_remove_hyperlink()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_hyperlink.md)
  : wb_add_hyperlink
- [`wb_add_pivot_table()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_pivot_table.md)
  : Add a pivot table to a worksheet
- [`wb_add_slicer()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_slicer.md)
  [`wb_remove_slicer()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_slicer.md)
  [`wb_add_timeline()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_slicer.md)
  [`wb_remove_timeline()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_slicer.md)
  : Add a slicer/timeline to a pivot table
- [`wb_add_thread()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_thread.md)
  [`wb_get_thread()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_thread.md)
  : Add threaded comments to a cell in a worksheet
- [`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/reference/wb_freeze_pane.md)
  : Freeze panes of a worksheet
- [`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/reference/wb_merge_cells.md)
  [`wb_unmerge_cells()`](https://janmarvin.github.io/openxlsx2/reference/wb_merge_cells.md)
  : Merge cells within a worksheet

### Add images and charts to a worksheet

Add images or Excel charts to a worksheet with the mschart package. See
[`vignette("openxlsx2_charts_manual")`](https://janmarvin.github.io/openxlsx2/articles/openxlsx2_charts_manual.md).

- [`wb_add_image()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_image.md)
  : Insert an image into a worksheet
- [`wb_add_plot()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_plot.md)
  : Insert the current R plot into a worksheet
- [`wb_add_chartsheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_chartsheet.md)
  : Add a chartsheet to a workbook
- [`wb_add_mschart()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_mschart.md)
  : Add an mschart object to a worksheet

## Style a workbook

Style a cell region, a worksheet or the entire workbook. See
[`vignette("openxlsx2_style_manual")`](https://janmarvin.github.io/openxlsx2/articles/openxlsx2_style_manual.md).

### Worksheet styling

Add styling to a cell region in a worksheet

- [`wb_add_border()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_border.md)
  : Modify borders in a cell region of a worksheet
- [`wb_add_cell_style()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_cell_style.md)
  : Modify the style in a cell region
- [`wb_add_fill()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_fill.md)
  : Modify the background fill color in a cell region
- [`wb_add_font()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_font.md)
  : Modify font properties in a cell region
- [`wb_add_named_style()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_named_style.md)
  : Apply styling to a cell region with a named style
- [`wb_add_numfmt()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_numfmt.md)
  : Modify number formatting in a cell region
- [`wb_get_cell_style()`](https://janmarvin.github.io/openxlsx2/reference/wb_cell_style.md)
  [`wb_set_cell_style()`](https://janmarvin.github.io/openxlsx2/reference/wb_cell_style.md)
  [`wb_set_cell_style_across()`](https://janmarvin.github.io/openxlsx2/reference/wb_cell_style.md)
  : Get or set cell style indices

### Workbook styling

These styling functions apply formatting to the workbook globally.

- [`wb_set_base_font()`](https://janmarvin.github.io/openxlsx2/reference/base_font-wb.md)
  [`wb_get_base_font()`](https://janmarvin.github.io/openxlsx2/reference/base_font-wb.md)
  : Set the default font in a workbook
- [`wb_add_dxfs_style()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_dxfs_style.md)
  : Set a dxfs style for the workbook
- [`wb_add_style()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_style.md)
  : Register a style in a workbook
- [`wb_set_base_colors()`](https://janmarvin.github.io/openxlsx2/reference/wb_base_colors.md)
  [`wb_get_base_colors()`](https://janmarvin.github.io/openxlsx2/reference/wb_base_colors.md)
  : Set the default colors in a workbook

## View and save a workbook

Visualize the content of a workbook in a spreadsheet software and save
it

- [`wb_open()`](https://janmarvin.github.io/openxlsx2/reference/wb_open.md)
  : Preview a workbook in spreadsheet software
- [`wb_save()`](https://janmarvin.github.io/openxlsx2/reference/wb_save.md)
  : Save a workbook to file
- [`xl_open()`](https://janmarvin.github.io/openxlsx2/reference/xl_open.md)
  : Open a file or workbook object in spreadsheet software

## Edit workbook defaults

These functions are used to modify setting of a workbook as a whole.

- [`wb_get_properties()`](https://janmarvin.github.io/openxlsx2/reference/properties-wb.md)
  [`wb_set_properties()`](https://janmarvin.github.io/openxlsx2/reference/properties-wb.md)
  : Modify workbook properties
- [`wb_set_last_modified_by()`](https://janmarvin.github.io/openxlsx2/reference/wb_set_last_modified_by.md)
  : Modify author in the metadata of a workbook
- [`wb_add_creators()`](https://janmarvin.github.io/openxlsx2/reference/creators-wb.md)
  [`wb_set_creators()`](https://janmarvin.github.io/openxlsx2/reference/creators-wb.md)
  [`wb_remove_creators()`](https://janmarvin.github.io/openxlsx2/reference/creators-wb.md)
  [`wb_get_creators()`](https://janmarvin.github.io/openxlsx2/reference/creators-wb.md)
  : Modify creators of a workbook
- [`wb_get_order()`](https://janmarvin.github.io/openxlsx2/reference/wb_order.md)
  [`wb_set_order()`](https://janmarvin.github.io/openxlsx2/reference/wb_order.md)
  : Order worksheets in a workbook
- [`wb_set_sheet_names()`](https://janmarvin.github.io/openxlsx2/reference/sheet_names-wb.md)
  [`wb_get_sheet_names()`](https://janmarvin.github.io/openxlsx2/reference/sheet_names-wb.md)
  : Get / Set worksheet names for a workbook
- [`wb_remove_worksheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_remove_worksheet.md)
  : Remove a worksheet from a workbook
- [`wb_get_active_sheet()`](https://janmarvin.github.io/openxlsx2/reference/active_sheet-wb.md)
  [`wb_set_active_sheet()`](https://janmarvin.github.io/openxlsx2/reference/active_sheet-wb.md)
  [`wb_get_selected()`](https://janmarvin.github.io/openxlsx2/reference/active_sheet-wb.md)
  [`wb_set_selected()`](https://janmarvin.github.io/openxlsx2/reference/active_sheet-wb.md)
  : Modify the state of active and selected sheets in a workbook
- [`wb_get_sheet_visibility()`](https://janmarvin.github.io/openxlsx2/reference/sheet_visibility-wb.md)
  [`wb_set_sheet_visibility()`](https://janmarvin.github.io/openxlsx2/reference/sheet_visibility-wb.md)
  : Get/set worksheet visible state in a workbook
- [`wb_get_bookview()`](https://janmarvin.github.io/openxlsx2/reference/wb_set_bookview.md)
  [`wb_remove_bookview()`](https://janmarvin.github.io/openxlsx2/reference/wb_set_bookview.md)
  [`wb_set_bookview()`](https://janmarvin.github.io/openxlsx2/reference/wb_set_bookview.md)
  : Get and Set the workbook position, size and filter
- [`wb_add_person()`](https://janmarvin.github.io/openxlsx2/reference/person-wb.md)
  [`wb_get_person()`](https://janmarvin.github.io/openxlsx2/reference/person-wb.md)
  : Helper for adding threaded comments
- [`wb_protect()`](https://janmarvin.github.io/openxlsx2/reference/wb_protect.md)
  : Protect a workbook from modifications
- [`wb_get_tables()`](https://janmarvin.github.io/openxlsx2/reference/wb_get_tables.md)
  : List tables in a worksheet
- [`wb_remove_tables()`](https://janmarvin.github.io/openxlsx2/reference/wb_remove_tables.md)
  : Remove a data table from a worksheet

## Workbook editing helpers

These functions are helpers that create an intermediate object. They are
helpful with editing a workbook. They are useful when you add content,
styling or you want to modify certain elements.

- [`wb_color()`](https://janmarvin.github.io/openxlsx2/reference/wb_color.md)
  : Create a color object for workbook styling

- [`wb_comment()`](https://janmarvin.github.io/openxlsx2/reference/wb_comment.md)
  : Helper to create a comment object

- [`wb_data()`](https://janmarvin.github.io/openxlsx2/reference/wb_data.md)
  [`` `[`( ``*`<wb_data>`*`)`](https://janmarvin.github.io/openxlsx2/reference/wb_data.md)
  :

  Add the `wb_data` attribute to a data frame in a worksheet

- [`wb_dims()`](https://janmarvin.github.io/openxlsx2/reference/wb_dims.md)
  :

  Helper to specify the `dims` argument

- [`create_border()`](https://janmarvin.github.io/openxlsx2/reference/create_border.md)
  : Create border format

- [`create_cell_style()`](https://janmarvin.github.io/openxlsx2/reference/create_cell_style.md)
  : Create cell style

- [`create_colors_xml()`](https://janmarvin.github.io/openxlsx2/reference/create_colors_xml.md)
  : Create custom color xml schemes

- [`create_dxfs_style()`](https://janmarvin.github.io/openxlsx2/reference/create_dxfs_style.md)
  : Create a custom formatting style

- [`create_fill()`](https://janmarvin.github.io/openxlsx2/reference/create_fill.md)
  : Create fill pattern

- [`create_font()`](https://janmarvin.github.io/openxlsx2/reference/create_font.md)
  : Create font format

- [`create_hyperlink()`](https://janmarvin.github.io/openxlsx2/reference/create_hyperlink.md)
  : Create spreadsheet hyperlink string

- [`create_numfmt()`](https://janmarvin.github.io/openxlsx2/reference/create_numfmt.md)
  : Create number format

- [`create_shape()`](https://janmarvin.github.io/openxlsx2/reference/create_shape.md)
  : Helper to create a shape

- [`create_sparklines()`](https://janmarvin.github.io/openxlsx2/reference/create_sparklines.md)
  : Create a sparklines object

- [`create_tablestyle()`](https://janmarvin.github.io/openxlsx2/reference/create_tablestyle.md)
  [`create_pivottablestyle()`](https://janmarvin.github.io/openxlsx2/reference/create_tablestyle.md)
  : Create custom (pivot) table styles

## Other functions

These functions can be used to achieve more specialized operations.

### Other tools to interact with worksheets

- [`wb_add_comment()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_comment.md)
  [`wb_get_comment()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_comment.md)
  [`wb_remove_comment()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_comment.md)
  : Add comment to worksheet
- [`wb_add_sparklines()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_sparklines.md)
  : Add sparklines to a worksheet
- [`wb_add_chart_xml()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_chart_xml.md)
  : Add a chart XML to a worksheet
- [`wb_add_drawing()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_drawing.md)
  : Add drawings to a worksheet
- [`wb_add_data_validation()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data_validation.md)
  : Add data validation to cells in a worksheet
- [`wb_add_form_control()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_form_control.md)
  : Add a checkbox, radio button or drop menu to a cell in a worksheet
- [`wb_add_ignore_error()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_ignore_error.md)
  : Ignore error types on a worksheet
- [`wb_add_page_break()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_page_break.md)
  : Add a page break to a worksheet
- [`wb_add_mips()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_mips.md)
  [`wb_get_mips()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_mips.md)
  : wb get and apply MIP section
- [`wb_clean_sheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_clean_sheet.md)
  : Clear content and formatting from a worksheet
- [`wb_clone_sheet_style()`](https://janmarvin.github.io/openxlsx2/reference/wb_clone_sheet_style.md)
  : Apply styling from a sheet to another within a workbook
- [`wb_clone_worksheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_clone_worksheet.md)
  : Create copies of a worksheet within a workbook
- [`wb_copy_cells()`](https://janmarvin.github.io/openxlsx2/reference/wb_copy_cells.md)
  : Copy cells around within a worksheet
- [`wb_set_page_setup()`](https://janmarvin.github.io/openxlsx2/reference/wb_page_setup.md)
  [`wb_page_setup()`](https://janmarvin.github.io/openxlsx2/reference/wb_page_setup.md)
  : Set page margins, orientation and print scaling of a worksheet
- [`wb_protect_worksheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_protect_worksheet.md)
  : Protect a worksheet from modifications
- [`wb_set_grid_lines()`](https://janmarvin.github.io/openxlsx2/reference/wb_set_grid_lines.md)
  [`wb_grid_lines()`](https://janmarvin.github.io/openxlsx2/reference/wb_set_grid_lines.md)
  : Modify grid lines visibility in a worksheet
- [`wb_set_header_footer()`](https://janmarvin.github.io/openxlsx2/reference/wb_set_header_footer.md)
  : Set headers and footers of a worksheet
- [`wb_set_sheetview()`](https://janmarvin.github.io/openxlsx2/reference/wb_set_sheetview.md)
  : Modify the default view of a worksheet
- [`wb_update_table()`](https://janmarvin.github.io/openxlsx2/reference/wb_update_table.md)
  : Update a data table position in a worksheet

### XML functions

A curated list of XML functions used in openxlsx2.

- [`read_xml()`](https://janmarvin.github.io/openxlsx2/reference/read_xml.md)
  : read xml file
- [`xml_node()`](https://janmarvin.github.io/openxlsx2/reference/pugixml.md)
  [`xml_node_name()`](https://janmarvin.github.io/openxlsx2/reference/pugixml.md)
  [`xml_value()`](https://janmarvin.github.io/openxlsx2/reference/pugixml.md)
  [`xml_attr()`](https://janmarvin.github.io/openxlsx2/reference/pugixml.md)
  : xml_node
- [`xml_add_child()`](https://janmarvin.github.io/openxlsx2/reference/xml_add_child.md)
  : append xml child to node
- [`xml_attr_mod()`](https://janmarvin.github.io/openxlsx2/reference/xml_attr_mod.md)
  : adds or updates attribute(s) in existing xml node
- [`xml_node_create()`](https://janmarvin.github.io/openxlsx2/reference/xml_node_create.md)
  : create xml_node from R objects
- [`xml_rm_child()`](https://janmarvin.github.io/openxlsx2/reference/xml_rm_child.md)
  : remove xml child to node
- [`as_xml()`](https://janmarvin.github.io/openxlsx2/reference/as_xml.md)
  : loads character string to pugixml and returns an externalptr
- [`print(`*`<pugi_xml>`*`)`](https://janmarvin.github.io/openxlsx2/reference/print.pugi_xml.md)
  : print pugi_xml

### Other helpers

- [`int2col()`](https://janmarvin.github.io/openxlsx2/reference/int2col.md)
  : Convert integers to spreadsheet column notation

- [`col2int()`](https://janmarvin.github.io/openxlsx2/reference/col2int.md)
  : Convert spreadsheet column notation to integers

- [`fmt_txt()`](https://janmarvin.github.io/openxlsx2/reference/fmt_txt.md)
  [`` `+`( ``*`<fmt_txt>`*`)`](https://janmarvin.github.io/openxlsx2/reference/fmt_txt.md)
  [`as.character(`*`<fmt_txt>`*`)`](https://janmarvin.github.io/openxlsx2/reference/fmt_txt.md)
  [`print(`*`<fmt_txt>`*`)`](https://janmarvin.github.io/openxlsx2/reference/fmt_txt.md)
  : format strings independent of the cell style.

- [`convert_date()`](https://janmarvin.github.io/openxlsx2/reference/convert_date.md)
  [`convert_datetime()`](https://janmarvin.github.io/openxlsx2/reference/convert_date.md)
  [`convert_hms()`](https://janmarvin.github.io/openxlsx2/reference/convert_date.md)
  : Convert from spreadsheet date, datetime or hms number to R Date type

- [`convert_to_excel_date()`](https://janmarvin.github.io/openxlsx2/reference/convert_to_excel_date.md)
  : convert back to an Excel Date

- [`apply_numfmt()`](https://janmarvin.github.io/openxlsx2/reference/apply_numfmt.md)
  : Format Values using OOXML (Spreadsheet) Number Format Codes

- [`temp_xlsx()`](https://janmarvin.github.io/openxlsx2/reference/temp_xlsx.md)
  : helper function to create temporary directory for testing purpose

- [`current_sheet()`](https://janmarvin.github.io/openxlsx2/reference/waivers.md)
  [`next_sheet()`](https://janmarvin.github.io/openxlsx2/reference/waivers.md)
  [`na_strings()`](https://janmarvin.github.io/openxlsx2/reference/waivers.md)
  :

  `openxlsx2` waivers

- [`dims_to_rowcol()`](https://janmarvin.github.io/openxlsx2/reference/dims_helper.md)
  [`validate_dims()`](https://janmarvin.github.io/openxlsx2/reference/dims_helper.md)
  [`rowcol_to_dims()`](https://janmarvin.github.io/openxlsx2/reference/dims_helper.md)
  :

  Helper functions to work with `dims`

- [`clean_worksheet_name()`](https://janmarvin.github.io/openxlsx2/reference/clean_worksheet_name.md)
  : Clean worksheet name

- [`styles_on_sheet()`](https://janmarvin.github.io/openxlsx2/reference/styles_on_sheet.md)
  : Get all styles on a sheet

## Package introduction

- [`openxlsx2`](https://janmarvin.github.io/openxlsx2/reference/openxlsx2-package.md)
  [`openxlsx2-package`](https://janmarvin.github.io/openxlsx2/reference/openxlsx2-package.md)
  : xlsx reading, writing and editing.

## Package options

- [`openxlsx2_options`](https://janmarvin.github.io/openxlsx2/reference/openxlsx2_options.md)
  : Options consulted by openxlsx2

## Deprecated

- [`openxlsx2-deprecated`](https://janmarvin.github.io/openxlsx2/reference/openxlsx2-deprecated.md)
  :

  Deprecated functions in package *openxlsx2*
