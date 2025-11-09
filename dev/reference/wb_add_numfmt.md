# Modify number formatting in a cell region

Add number formatting to a cell region. You can use a number format
created by
[`create_numfmt()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_numfmt.md).

## Usage

``` r
wb_add_numfmt(wb, sheet = current_sheet(), dims = "A1", numfmt)
```

## Arguments

- wb:

  A Workbook

- sheet:

  the worksheet

- dims:

  the cell range

- numfmt:

  either an integer id for a builtin numeric font or a character string
  as described in the ***Details***

## Value

The `wbWorkbook` object, invisibly.

## Details

The list of number formats ID is located in the **Details** section of
[`create_cell_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_cell_style.md).

### General Number Formatting

- `"0"`: Displays numbers as integers without decimal places.

- `"0.00"`: Displays numbers with two decimal places (e.g., `123.45`).

- `"#,##0"`: Displays thousands separators without decimals (e.g.,
  `1,000`).

- `"#,##0.00"`: Displays thousands separators with two decimal places
  (e.g., `1,000.00`).

### Currency Formatting

- `"$#,##0.00"`: Formats numbers as currency with two decimal places
  (e.g., `$1,000.00`).

- `"[$$-409]#,##0.00"`: Localized currency format in U.S. dollars.

- `"¥#,##0"`: Custom currency format (e.g., for Japanese yen) without
  decimals.

- `"£#,##0.00"`: GBP currency format with two decimal places.

### Percentage Formatting

- `"0%"`: Displays numbers as percentages with no decimal places (e.g.,
  `50%`).

- `"0.00%"`: Displays numbers as percentages with two decimal places
  (e.g., `50.00%`).

### Scientific Formatting

- `"0.00E+00"`: Scientific notation with two decimal places (e.g.,
  `1.23E+03` for `1230`).

### Date and Time Formatting

- `"yyyy-mm-dd"`: Year-month-day format (e.g., `2023-10-31`).

- `"dd/mm/yyyy"`: Day/month/year format (e.g., `31/10/2023`).

- `"mmm d, yyyy"`: Month abbreviation with day and year (e.g.,
  `Oct 31, 2023`).

- `"h:mm AM/PM"`: Time with AM/PM format (e.g., `1:30 PM`).

- `"h:mm:ss"`: Time with seconds (e.g., `13:30:15` for `1:30:15 PM`).

- `"yyyy-mm-dd h:mm:ss"`: Full date and time format.

### Fraction Formatting

- `"# ?/?"`: Displays numbers as a fraction with a single digit
  denominator (e.g., `1/2`).

- `"# ??/??"`: Displays numbers as a fraction with a two-digit
  denominator (e.g., `1 12/25`).

### Custom Formatting

- `"_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)`: Custom currency
  format with parentheses for negative values and dashes for zero
  values.

- `"[Red]0.00;[Blue](0.00);0"`: Displays positive numbers in red,
  negatives in blue, and zeroes as plain.

- `"@"`: Text placeholder format (e.g., for cells with mixed text and
  numeric values).

### Formatting Symbols Reference

- `0`: Digit placeholder, displays a digit or zero.

- `#`: Digit placeholder, does not display extra zeroes.

- `.`: Decimal point.

- `,`: Thousands separator.

- `E+`, `E-`: Scientific notation.

- `_` (underscore): Adds a space equal to the width of the next
  character.

- `"text"`: Displays literal text within quotes.

- `*`: Repeat character to fill the cell width.

## See also

Other styles:
[`wb_add_border()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_border.md),
[`wb_add_cell_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_cell_style.md),
[`wb_add_fill()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_fill.md),
[`wb_add_font()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_font.md),
[`wb_add_named_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_named_style.md),
[`wb_cell_style`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_cell_style.md)

## Examples

``` r
wb <- wb_workbook()
wb <- wb_add_worksheet(wb, "S1")
wb <- wb_add_data(wb, "S1", mtcars)
wb <- wb_add_numfmt(wb, "S1", dims = "F1:F33", numfmt = "#.0")
# Chaining
wb <- wb_workbook()$add_worksheet("S1")$add_data("S1", mtcars)
wb$add_numfmt("S1", "A1:A33", numfmt = 1)
```
