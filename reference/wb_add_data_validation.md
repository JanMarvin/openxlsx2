# Add data validation to cells in a worksheet

Add spreadsheet data validation to cells

## Usage

``` r
wb_add_data_validation(
  wb,
  sheet = current_sheet(),
  dims = "A1",
  type,
  operator,
  value,
  allow_blank = TRUE,
  show_input_msg = TRUE,
  show_error_msg = TRUE,
  error_style = NULL,
  error_title = NULL,
  error = NULL,
  prompt_title = NULL,
  prompt = NULL,
  ...
)
```

## Arguments

- wb:

  A Workbook object

- sheet:

  A name or index of a worksheet

- dims:

  A cell dimension ("A1" or "A1:B2")

- type:

  One of 'whole', 'decimal', 'date', 'time', 'textLength', 'list' (see
  examples)

- operator:

  One of 'between', 'notBetween', 'equal', 'notEqual', 'greaterThan',
  'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'

- value:

  a vector of length 1 or 2 depending on operator (see examples)

- allow_blank:

  logical

- show_input_msg:

  logical

- show_error_msg:

  logical

- error_style:

  The icon shown and the options how to deal with such inputs. Default
  "stop" (cancel), else "information" (prompt popup) or "warning"
  (prompt accept or change input)

- error_title:

  The error title

- error:

  The error text

- prompt_title:

  The prompt title

- prompt:

  The prompt text

- ...:

  additional arguments

## Examples

``` r
wb <- wb_workbook()
wb$add_worksheet("Sheet 1")
wb$add_worksheet("Sheet 2")

wb$add_data_table(1, x = iris[1:30, ])
wb$add_data_validation(1,
  dims = "A2:C31", type = "whole",
  operator = "between", value = c(1, 9)
)
wb$add_data_validation(1,
  dims = "E2:E31", type = "textLength",
  operator = "between", value = c(4, 6)
)

## Date and Time cell validation
df <- data.frame(
  "d" = as.Date("2016-01-01") + -5:5,
  "t" = as.POSIXct("2016-01-01") + -5:5 * 10000
)
wb$add_data_table(2, x = df)
wb$add_data_validation(2, dims = "A2:A12", type = "date",
  operator = "greaterThanOrEqual", value = as.Date("2016-01-01")
)
wb$add_data_validation(2,
  dims = "B2:B12", type = "time",
  operator = "between", value = df$t[c(4, 8)]
)


######################################################################
## If type == 'list'
# operator argument is ignored.

wb <- wb_workbook()
wb$add_worksheet("Sheet 1")
wb$add_worksheet("Sheet 2")

wb$add_data_table(sheet = 1, x = iris[1:30, ])
wb$add_data(sheet = 2, x = sample(iris$Sepal.Length, 10))

wb$add_data_validation(1, dims = "A2:A31", type = "list", value = "'Sheet 2'!$A$1:$A$10")
```
