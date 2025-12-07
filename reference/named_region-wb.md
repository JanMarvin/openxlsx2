# Modify named regions in a worksheet

Create / delete a named region. You can also specify a named region by
using the `name` argument in
`wb_add_data(x = iris, name = "my-region")`. It is important to note
that named regions are not case-sensitive and must be unique.

## Usage

``` r
wb_add_named_region(
  wb,
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

wb_remove_named_region(wb, sheet = current_sheet(), name = NULL)

wb_get_named_regions(wb, tables = FALSE, x = NULL)
```

## Arguments

- wb:

  A Workbook object

- sheet:

  A name or index of a worksheet

- dims:

  Worksheet cell range of the region ("A1:D4").

- name:

  Name for region. A character vector of length 1. Note that region
  names must be case-insensitive unique.

- local_sheet:

  If `TRUE` the named region will be local for this sheet

- overwrite:

  Boolean. Overwrite if exists? Default to `FALSE`.

- comment:

  description text for named region

- hidden:

  Should the named region be hidden?

- custom_menu, description, is_function, function_group_id, help,
  local_name, publish_to_server, status_bar, vb_procedure,
  workbook_parameter, xml:

  Unknown XML feature

- ...:

  additional arguments

- tables:

  Should included both data tables and named regions in the result?

- x:

  Deprecated. Use `wb`. For workbook input use
  [`wb_load()`](https://janmarvin.github.io/openxlsx2/reference/wb_load.md)
  to first load the xlsx file as a workbook.

## Value

A workbook, invisibly.

A data frame with the all named regions in `wb`. Or `NULL`, if none are
found.

## Details

You can use the
[`wb_dims()`](https://janmarvin.github.io/openxlsx2/reference/wb_dims.md)
helper to specify the cell range of the named region

## See also

[`wb_get_tables()`](https://janmarvin.github.io/openxlsx2/reference/wb_get_tables.md)

Other worksheet content functions:
[`col_widths-wb`](https://janmarvin.github.io/openxlsx2/reference/col_widths-wb.md),
[`filter-wb`](https://janmarvin.github.io/openxlsx2/reference/filter-wb.md),
[`grouping-wb`](https://janmarvin.github.io/openxlsx2/reference/grouping-wb.md),
[`row_heights-wb`](https://janmarvin.github.io/openxlsx2/reference/row_heights-wb.md),
[`wb_add_conditional_formatting()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_conditional_formatting.md),
[`wb_add_data()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data.md),
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data_table.md),
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_formula.md),
[`wb_add_hyperlink()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_hyperlink.md),
[`wb_add_pivot_table()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_pivot_table.md),
[`wb_add_slicer()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_slicer.md),
[`wb_add_thread()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_thread.md),
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/reference/wb_freeze_pane.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/reference/wb_merge_cells.md)

## Examples

``` r
## create named regions
wb <- wb_workbook()
wb$add_worksheet("Sheet 1")

## specify region
wb$add_data(x = iris, start_col = 1, start_row = 1)
wb$add_named_region(
  name = "iris",
  dims = wb_dims(x = iris)
)

## using add_data 'name' argument
wb$add_data(sheet = 1, x = iris, name = "iris2", start_col = 10)

## delete one
wb$remove_named_region(name = "iris2")
wb$get_named_regions()
#>   name                 value  sheets  coords id local sheet
#> 1 iris 'Sheet 1'!$A$1:$E$151 Sheet 1 A1:E151  1     0     1
## read named regions
df <- wb_to_df(wb, named_region = "iris")
head(df)
#>   Sepal.Length Sepal.Width Petal.Length Petal.Width Species
#> 2          5.1         3.5          1.4         0.2  setosa
#> 3          4.9         3.0          1.4         0.2  setosa
#> 4          4.7         3.2          1.3         0.2  setosa
#> 5          4.6         3.1          1.5         0.2  setosa
#> 6          5.0         3.6          1.4         0.2  setosa
#> 7          5.4         3.9          1.7         0.4  setosa


# Extract named regions from a file
out_file <- temp_xlsx()
wb_save(wb, out_file, overwrite = TRUE)

# Load the file as a workbook first, then get named regions.
wb1 <- wb_load(out_file)
wb1$get_named_regions()
#>   name                 value  sheets  coords id local sheet
#> 1 iris 'Sheet 1'!$A$1:$E$151 Sheet 1 A1:E151  1     0     1
```
