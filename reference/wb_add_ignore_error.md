# Ignore error types on a worksheet

The `wb_add_ignore_error()` function allows you to suppress specific
types of background error checking warnings for a given cell range. This
is useful for preventing the display of green error indicators
(triangles) in cases where "errors" are intentional, such as numbers
being stored as text for formatting purposes.

## Usage

``` r
wb_add_ignore_error(
  wb,
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
```

## Arguments

- wb:

  A
  [wbWorkbook](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.md)
  object.

- sheet:

  The name or index of the worksheet. Defaults to the current sheet.

- dims:

  A character string defining the cell range (e.g., "A1:A100").

- calculated_column:

  Logical; if `TRUE`, ignores errors in calculated columns of a table.

- empty_cell_reference:

  Logical; if `TRUE`, ignores errors when a formula refers to an empty
  cell.

- eval_error:

  Logical; if `TRUE`, ignores errors resulting from formula evaluation
  (e.g., `#DIV/0!`, `#N/A`).

- formula:

  Logical; if `TRUE`, ignores formula consistency errors.

- formula_range:

  Logical; if `TRUE`, ignores errors where a formula omits cells in a
  region.

- list_data_validation:

  Logical; if `TRUE`, ignores errors related to list data validation
  mapping.

- number_stored_as_text:

  Logical; if `TRUE`, suppresses the error displayed when numeric values
  are stored as string/text types.

- two_digit_text_year:

  Logical; if `TRUE`, ignores warnings about dates containing two-digit
  years.

- unlocked_formula:

  Logical; if `TRUE`, ignores errors for formulas in cells that are not
  locked.

- ...:

  Additional arguments.

## Value

The
[wbWorkbook](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.md)
object, invisibly.

## Details

Spreadsheet software performs background validation on formulas and data
entries. When a cell triggers a rule, a visual indicator appears. This
function modifies the `<ignoredErrors>` section of the worksheet XML to
whitelist specific ranges against specific rules.

Most commonly, this is used with `number_stored_as_text = TRUE` when IDs
or codes (like "00123") must be preserved as character strings but
contain only numeric digits.

## Notes

- This function does not fix the underlying data; it only instructs the
  spreadsheet application not to flag the specific error type visually.

- If multiple error types need to be ignored for the same range, you can
  set multiple arguments to `TRUE` in a single call.
