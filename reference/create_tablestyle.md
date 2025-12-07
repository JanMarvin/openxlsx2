# Create custom (pivot) table styles

Create a custom (pivot) table style. These functions are for expert use
only. Use other styling functions instead.

## Usage

``` r
create_tablestyle(
  name,
  whole_table = NULL,
  header_row = NULL,
  total_row = NULL,
  first_column = NULL,
  last_column = NULL,
  first_row_stripe = NULL,
  second_row_stripe = NULL,
  first_column_stripe = NULL,
  second_column_stripe = NULL,
  first_header_cell = NULL,
  last_header_cell = NULL,
  first_total_cell = NULL,
  last_total_cell = NULL,
  ...
)

create_pivottablestyle(
  name,
  whole_table = NULL,
  header_row = NULL,
  grand_total_row = NULL,
  first_column = NULL,
  grand_total_column = NULL,
  first_row_stripe = NULL,
  second_row_stripe = NULL,
  first_column_stripe = NULL,
  second_column_stripe = NULL,
  first_header_cell = NULL,
  first_subtotal_column = NULL,
  second_subtotal_column = NULL,
  third_subtotal_column = NULL,
  first_subtotal_row = NULL,
  second_subtotal_row = NULL,
  third_subtotal_row = NULL,
  blank_row = NULL,
  first_column_subheading = NULL,
  second_column_subheading = NULL,
  third_column_subheading = NULL,
  first_row_subheading = NULL,
  second_row_subheading = NULL,
  third_row_subheading = NULL,
  page_field_labels = NULL,
  page_field_values = NULL,
  ...
)
```

## Arguments

- name:

  name

- whole_table:

  wholeTable

- header_row, total_row:

  ...Row

- first_column, last_column:

  ...Column

- first_row_stripe, second_row_stripe:

  ...RowStripe

- first_column_stripe, second_column_stripe:

  ...ColumnStripe

- first_header_cell, last_header_cell:

  ...HeaderCell

- first_total_cell, last_total_cell:

  ...TotalCell

- ...:

  additional arguments

- grand_total_row:

  totalRow

- grand_total_column:

  lastColumn

- first_subtotal_column, second_subtotal_column, third_subtotal_column:

  ...SubtotalColumn

- first_subtotal_row, second_subtotal_row, third_subtotal_row:

  ...SubtotalRow

- blank_row:

  blankRow

- first_column_subheading, second_column_subheading,
  third_column_subheading:

  ...ColumnSubheading

- first_row_subheading, second_row_subheading, third_row_subheading:

  ...RowSubheading

- page_field_labels:

  pageFieldLabels

- page_field_values:

  pageFieldValues

## See also

Other style creating functions:
[`create_border()`](https://janmarvin.github.io/openxlsx2/reference/create_border.md),
[`create_cell_style()`](https://janmarvin.github.io/openxlsx2/reference/create_cell_style.md),
[`create_colors_xml()`](https://janmarvin.github.io/openxlsx2/reference/create_colors_xml.md),
[`create_dxfs_style()`](https://janmarvin.github.io/openxlsx2/reference/create_dxfs_style.md),
[`create_fill()`](https://janmarvin.github.io/openxlsx2/reference/create_fill.md),
[`create_font()`](https://janmarvin.github.io/openxlsx2/reference/create_font.md),
[`create_numfmt()`](https://janmarvin.github.io/openxlsx2/reference/create_numfmt.md)
