# Add a checkbox, radio button or drop menu to a cell in a worksheet

You can add Form Control to a cell. The three supported types are a
Checkbox, a Radio button, or a Drop menu.

## Usage

``` r
wb_add_form_control(
  wb,
  sheet = current_sheet(),
  dims = "A1",
  type = c("Checkbox", "Radio", "Drop"),
  text = NULL,
  link = NULL,
  range = NULL,
  checked = FALSE
)
```

## Arguments

- wb:

  A Workbook object

- sheet:

  A worksheet of the workbook

- dims:

  A single cell as spreadsheet dimension, e.g. "A1".

- type:

  A type "Checkbox" (the default), "Radio" a radio button or "Drop" a
  drop down menu

- text:

  A text to be shown next to the Checkbox or radio button (optional)

- link:

  A cell range to link to

- range:

  A cell range used as input

- checked:

  A logical indicating if the Checkbox or Radio button is checked

## Value

The `wbWorkbook` object, invisibly.

## Examples

``` r
wb <- wb_workbook()
wb <- wb_add_worksheet(wb)
wb <- wb_add_form_control(wb)
# Add
wb$add_form_control(dims = "C5", type = "Radio", checked = TRUE)
```
