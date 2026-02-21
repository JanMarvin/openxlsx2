# Add a worksheet to a workbook

The `wb_add_worksheet()` function is a fundamental step in workbook
construction, appending a new worksheet to a `wbWorkbook` object. It
provides extensive parameters for configuring the sheet's initial state,
including visibility, visual cues like grid lines, and metadata such as
tab colors and page setup properties.

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

  A
  [wbWorkbook](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.md)
  object to which the new worksheet will be attached.

- sheet:

  A character string for the worksheet name. Defaults to a sequentially
  generated name (e.g., "Sheet 1").

- grid_lines:

  Logical; if `FALSE`, the worksheet grid lines are hidden.

- row_col_headers:

  Logical; if `FALSE`, row numbers and column letters are hidden.

- tab_color:

  The color of the worksheet tab. Accepts a
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/reference/wb_color.md)
  object, a standard R color name, or a hex color code (e.g.,
  "#4F81BD").

- zoom:

  The sheet zoom level as a percentage; a numeric value between 10
  and 400. Values below 10 default to 10.

- header, footer:

  Default character vectors of length three for the left, center, and
  right sections of the header or footer.

- odd_header, odd_footer:

  Specific definitions for odd-numbered pages. Defaults to the values
  provided in `header` and `footer`.

- even_header, even_footer:

  Specific definitions for even-numbered pages. Defaults to the values
  provided in `header` and `footer`.

- first_header, first_footer:

  Specific definitions for the first page of the worksheet. Defaults to
  the values provided in `header` and `footer`.

- visible:

  The visibility state of the sheet. One of "visible", "hidden", or
  "veryHidden".

- has_drawing:

  *defunct*

- paper_size:

  An integer code representing a standard paper size. Refer to
  [`wb_page_setup()`](https://janmarvin.github.io/openxlsx2/reference/wb_page_setup.md)
  for a complete list of codes.

- orientation:

  The page orientation, either "portrait" or "landscape".

- hdpi, vdpi:

  The horizontal and vertical DPI (dots per inch) for printing and
  rendering. Can be set globally via `options("openxlsx2.hdpi")`.

- ...:

  Additional arguments passed to internal sheet configuration methods.

## Value

The
[wbWorkbook](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.md)
object, invisibly.

## Details

Worksheets can be configured with complex headers and footers that adapt
to document layout requirements. The function supports distinct
definitions for odd pages, even pages, and the first page of a document.
Headers and footers are defined as character vectors of length three,
representing the left, center, and right sections respectively.

Within these sections, special dynamic tags can be utilized to include
automatic metadata:

- `&[Page]`: The current page number

- `&[Pages]`: The total number of pages

- `&[Date]`: The current system date

- `&[Time]`: The current system time

- `&[Path]`: The file path of the workbook

- `&[File]`: The name of the file

- `&[Tab]`: The name of the worksheet

The function also initializes the sheet view. Parameters like `zoom` and
`grid_lines` determine how the sheet is presented upon opening the file
in spreadsheet software. For advanced page configuration, such as DPI
settings and paper sizes, the function integrates with the package-wide
options system but allows for per-sheet overrides.

## Notes

- As of recent versions, the `has_drawing` argument has been removed and
  is no longer part of the public API.

- If `zoom` is provided outside the 10–400 range, it is automatically
  clamped to the nearest boundary.

- The `sheet` name is validated against a set of illegal characters
  prohibited by spreadsheet software standards.

## See also

Other workbook wrappers:
[`base_font-wb`](https://janmarvin.github.io/openxlsx2/reference/base_font-wb.md),
[`col_widths-wb`](https://janmarvin.github.io/openxlsx2/reference/col_widths-wb.md),
[`creators-wb`](https://janmarvin.github.io/openxlsx2/reference/creators-wb.md),
[`grouping-wb`](https://janmarvin.github.io/openxlsx2/reference/grouping-wb.md),
[`row_heights-wb`](https://janmarvin.github.io/openxlsx2/reference/row_heights-wb.md),
[`wb_add_chartsheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_chartsheet.md),
[`wb_add_data()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data.md),
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data_table.md),
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_formula.md),
[`wb_add_hyperlink()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_hyperlink.md),
[`wb_add_pivot_table()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_pivot_table.md),
[`wb_add_slicer()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_slicer.md),
[`wb_base_colors`](https://janmarvin.github.io/openxlsx2/reference/wb_base_colors.md),
[`wb_clone_worksheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_clone_worksheet.md),
[`wb_copy_cells()`](https://janmarvin.github.io/openxlsx2/reference/wb_copy_cells.md),
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/reference/wb_freeze_pane.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/reference/wb_merge_cells.md),
[`wb_save()`](https://janmarvin.github.io/openxlsx2/reference/wb_save.md),
[`wb_set_last_modified_by()`](https://janmarvin.github.io/openxlsx2/reference/wb_set_last_modified_by.md),
[`wb_workbook()`](https://janmarvin.github.io/openxlsx2/reference/wb_workbook.md)

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
