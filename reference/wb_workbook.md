# Create a new Workbook object

This function initializes and returns a
[wbWorkbook](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.md)
object, which is the core structure for building and modifying openxml
files (`.xlsx`) in `openxlsx2`.

## Usage

``` r
wb_workbook(
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
```

## Arguments

- creator:

  Creator of the workbook (your name). Defaults to login username or
  `options("openxlsx2.creator")` if set.

- title, subject, category, keywords, comments, manager, company:

  Workbook property, a string.

- datetime_created:

  The time of the workbook is created

- datetime_modified:

  The time of the workbook was last modified

- theme:

  Optional theme identified by string or number. See **Details** for
  options.

- ...:

  additional arguments

## Value

A `wbWorkbook` object

## Details

You can define various metadata properties at creation, such as the
`creator`, `title`, `subject`, and timestamps. You can also specify a
workbook theme.

The returned `wb_workbook()` object is an
[`R6::R6Class()`](https://r6.r-lib.org/reference/R6Class.html) instance.
Once created, the standard workflow is to immediately add a worksheet
using
[`wb_add_worksheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_worksheet.md).
From there, you can populate the sheet with data
([`wb_add_data()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data.md)),
or formulas
([`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_formula.md)),
and apply styling or other elements.

`theme` can be one of "Atlas", "Badge", "Berlin", "Celestial", "Crop",
"Depth", "Droplet", "Facet", "Feathered", "Gallery", "Headlines",
"Integral", "Ion", "Ion Boardroom", "LibreOffice", "Madison", "Main
Event", "Mesh", "Office 2007 - 2010 Theme", "Office 2013 - 2022 Theme",
"Office Theme", "Old Office Theme", "Organic", "Parallax", "Parcel",
"Retrospect", "Savon", "Slice", "Vapor Trail", "View", "Wisp", "Wood
Type"

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
[`wb_add_worksheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_worksheet.md),
[`wb_base_colors`](https://janmarvin.github.io/openxlsx2/reference/wb_base_colors.md),
[`wb_clone_worksheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_clone_worksheet.md),
[`wb_copy_cells()`](https://janmarvin.github.io/openxlsx2/reference/wb_copy_cells.md),
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/reference/wb_freeze_pane.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/reference/wb_merge_cells.md),
[`wb_save()`](https://janmarvin.github.io/openxlsx2/reference/wb_save.md),
[`wb_set_last_modified_by()`](https://janmarvin.github.io/openxlsx2/reference/wb_set_last_modified_by.md)

## Examples

``` r
## Create a new workbook
wb <- wb_workbook()

## Set Workbook properties
wb <- wb_workbook(
  creator  = "Me",
  title    = "Expense Report",
  subject  = "Expense Report - 2022 Q1",
  category = "sales"
)

## Cloning a workbook
wb1 <- wb_workbook()
wb2 <- wb1$clone(deep = TRUE)
```
