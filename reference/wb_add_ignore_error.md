# Ignore error types on worksheet

This function allows to hide / ignore certain types of errors shown in a
worksheet.

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

  A workbook

- sheet:

  A sheet name or index.

- dims:

  Cell range to ignore the error

- calculated_column:

  calculatedColumn

- empty_cell_reference:

  emptyCellReference

- eval_error:

  evalError

- formula:

  formula

- formula_range:

  formulaRange

- list_data_validation:

  listDataValidation

- number_stored_as_text:

  If `TRUE`, will not display the error if numbers are stored as text.

- two_digit_text_year:

  twoDigitTextYear

- unlocked_formula:

  unlockedFormula

- ...:

  additional arguments

## Value

The `wbWorkbook` object, invisibly.
