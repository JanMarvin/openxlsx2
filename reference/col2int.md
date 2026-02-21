# Convert spreadsheet column notation to integers

`col2int()` transforms spreadsheet-style column identifiers (e.g., "A",
"B", "AA") into their corresponding integer indices. This utility is
fundamental for programmatic data manipulation, where "A" is mapped to
1, "B" to 2, and "ZZ" to 702.

## Usage

``` r
col2int(x)
```

## Arguments

- x:

  A character vector of column labels, a numeric vector of indices, or a
  factor. Supports range notation like "A:Z".

## Value

An integer vector representing the column indices. Returns `NULL` if the
input `x` is `NULL`, or an empty integer vector if the length of `x` is
zero.

## Details

The function is designed to handle various input formats encountered
during spreadsheet data processing. In addition to single column labels,
it supports range notation using the colon operator (e.g., "A:C"). When
a range is detected, the function internally expands the notation into a
complete sequence of integers (e.g., 1, 2, 3). This behavior is
particularly useful when passing column selections to functions like
[`wb_to_df()`](https://janmarvin.github.io/openxlsx2/reference/wb_to_df.md)
or
[`wb_read()`](https://janmarvin.github.io/openxlsx2/reference/wb_to_df.md).

Input validation ensures that only atomic vectors are processed. If the
input is already numeric or a factor, the function ensures the values
fall within the valid spreadsheet column range before coercion to
integers. Note that the presence of `NA` values in the input will
trigger an error to maintain data integrity during index calculation.

## Notes

- Range expansion via `:` is performed iteratively until all sequences
  are resolved into individual integer components.

- In compliance with spreadsheet software standards, the function
  validates that indices do not exceed the maximum allowable column
  limit.

## See also

[`int2col()`](https://janmarvin.github.io/openxlsx2/reference/int2col.md)

## Examples

``` r
# Convert standard labels
col2int(c("A", "B", "Z"))
#> [1]  1  2 26

# Convert ranges to integer sequences
col2int("A:C")
#> [1] 1 2 3

# Mix individual columns and ranges
col2int(c("A", "C:E", "G"))
#> [1] 1 3 4 5 7

# Handle numeric inputs
col2int(c(1, 2, 26))
#> [1]  1  2 26
```
