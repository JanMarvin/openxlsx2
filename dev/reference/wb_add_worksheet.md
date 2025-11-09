# Add a worksheet to a workbook

Add a worksheet to a
[wbWorkbook](https://janmarvin.github.io/openxlsx2/dev/reference/wbWorkbook.md)
is the first step to build a workbook. With the function, you can also
set the sheet view with `zoom`, set headers and footers as well as other
features. See the function arguments.

## Usage

``` r
wb_add_worksheet(
  wb,
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
```

## Arguments

- wb:

  A `wbWorkbook` object to attach the new worksheet

- sheet:

  A name for the new worksheet

- grid_lines:

  A logical. If `FALSE`, the worksheet grid lines will be hidden.

- row_col_headers:

  A logical. If `FALSE`, the worksheet colname and rowname will be
  hidden.

- tab_color:

  Color of the sheet tab. A
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_color.md),
  a valid color (belonging to
  [`grDevices::colors()`](https://rdrr.io/r/grDevices/colors.html)) or a
  valid hex color beginning with "#".

- zoom:

  The sheet zoom level, a numeric between 10 and 400 as a percentage. (A
  zoom value smaller than 10 will default to 10.)

- header, odd_header, even_header, first_header, footer, odd_footer,
  even_footer, first_footer:

  Character vector of length 3 corresponding to positions left, center,
  right. `header` and `footer` are used to default additional arguments.
  Setting `even`, `odd`, or `first`, overrides `header`/`footer`. Use
  `NA` to skip a position.

- visible:

  If `FALSE`, sheet is hidden else visible.

- has_drawing:

  If `TRUE` prepare a drawing output (TODO does this work?)

- paper_size:

  An integer corresponding to a paper size. See
  [`wb_page_setup()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_page_setup.md)
  for details.

- orientation:

  One of "portrait" or "landscape"

- hdpi, vdpi:

  Horizontal and vertical DPI. Can be set with
  `options("openxlsx2.dpi" = X)`, `options("openxlsx2.hdpi" = X)` or
  `options("openxlsx2.vdpi" = X)`

- ...:

  Additional arguments

## Value

The `wbWorkbook` object, invisibly.

## Details

Headers and footers can contain special tags

- **&\[Page\]** Page number

- **&\[Pages\]** Number of pages

- **&\[Date\]** Current date

- **&\[Time\]** Current time

- **&\[Path\]** File path

- **&\[File\]** File name

- **&\[Tab\]** Worksheet name

## See also

Other workbook wrappers:
[`base_font-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/base_font-wb.md),
[`col_widths-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/col_widths-wb.md),
[`creators-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/creators-wb.md),
[`grouping-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/grouping-wb.md),
[`row_heights-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/row_heights-wb.md),
[`wb_add_chartsheet()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_chartsheet.md),
[`wb_add_data()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data.md),
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data_table.md),
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_formula.md),
[`wb_add_hyperlink()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_hyperlink.md),
[`wb_add_pivot_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_pivot_table.md),
[`wb_add_slicer()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_slicer.md),
[`wb_base_colors`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_base_colors.md),
[`wb_clone_worksheet()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_clone_worksheet.md),
[`wb_copy_cells()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_copy_cells.md),
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_freeze_pane.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_merge_cells.md),
[`wb_save()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_save.md),
[`wb_set_last_modified_by()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_set_last_modified_by.md),
[`wb_workbook()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_workbook.md)

## Examples

``` r
## Create a new workbook
wb <- wb_workbook()

## Add a worksheet
wb$add_worksheet("Sheet 1")
## No grid lines
wb$add_worksheet("Sheet 2", grid_lines = FALSE)
## A red tab color
wb$add_worksheet("Sheet 3", tab_color = wb_color("red"))
## All options combined with a zoom of 40%
wb$add_worksheet("Sheet 4", grid_lines = FALSE, tab_color = wb_color(hex = "#4F81BD"), zoom = 40)

## Headers and Footers
wb$add_worksheet("Sheet 5",
  header = c("ODD HEAD LEFT", "ODD HEAD CENTER", "ODD HEAD RIGHT"),
  footer = c("ODD FOOT RIGHT", "ODD FOOT CENTER", "ODD FOOT RIGHT"),
  even_header = c("EVEN HEAD LEFT", "EVEN HEAD CENTER", "EVEN HEAD RIGHT"),
  even_footer = c("EVEN FOOT RIGHT", "EVEN FOOT CENTER", "EVEN FOOT RIGHT"),
  first_header = c("TOP", "OF FIRST", "PAGE"),
  first_footer = c("BOTTOM", "OF FIRST", "PAGE")
)

wb$add_worksheet("Sheet 6",
  header = c("&[Date]", "ALL HEAD CENTER 2", "&[Page] / &[Pages]"),
  footer = c("&[Path]&[File]", NA, "&[Tab]"),
  first_header = c(NA, "Center Header of First Page", NA),
  first_footer = c(NA, "Center Footer of First Page", NA)
)

wb$add_worksheet("Sheet 7",
  header = c("ALL HEAD LEFT 2", "ALL HEAD CENTER 2", "ALL HEAD RIGHT 2"),
  footer = c("ALL FOOT RIGHT 2", "ALL FOOT CENTER 2", "ALL FOOT RIGHT 2")
)

wb$add_worksheet("Sheet 8",
  first_header = c("FIRST ONLY L", NA, "FIRST ONLY R"),
  first_footer = c("FIRST ONLY L", NA, "FIRST ONLY R")
)

## Need data on worksheet to see all headers and footers
wb$add_data(sheet = 5, 1:400)
wb$add_data(sheet = 6, 1:400)
wb$add_data(sheet = 7, 1:400)
wb$add_data(sheet = 8, 1:400)
```
