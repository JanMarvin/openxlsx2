# Convert integers to spreadsheet column notation

`int2col()` performs the inverse operation of
[`col2int()`](https://janmarvin.github.io/openxlsx2/dev/reference/col2int.md),
transforming numeric column indices into their corresponding
spreadsheet-style character labels (e.g., 1 to "A", 28 to "AB"). This is
essential for converting calculated indices back into a format
compatible with spreadsheet cell referencing.

## Usage

``` r
int2col(x)
```

## Arguments

- x:

  A numeric vector representing the column indices to be converted.

## Value

A character vector of spreadsheet column labels. Returns `NULL` if the
input `x` is `NULL`.

## Details

The function accepts a numeric vector and maps each integer to its
positional representation in a base-26 derived system. This mapping
follows standard spreadsheet conventions where the sequence progresses
from "A" through "Z", followed by "AA", "AB", and so forth.

Validation is performed to ensure the input is numeric and finite. In
accordance with the Office Open XML specification used by most
spreadsheet software, the maximum supported column index is 16,384,
which corresponds to the column label "XFD". Inputs exceeding this range
may result in coordinates that are incompatible with standard
spreadsheet applications.

## Notes

- Non-integer numeric values will typically be coerced or truncated;
  however, infinite values will trigger an error to prevent invalid
  coordinate generation.

## See also

[`col2int()`](https://janmarvin.github.io/openxlsx2/dev/reference/col2int.md)

## Examples

``` r
# Convert a single index
int2col(27)
#> [1] "AA"

# Convert a sequence of indices
int2col(1:10)
#>  [1] "A" "B" "C" "D" "E" "F" "G" "H" "I" "J"

# Handle large column indices
int2col(c(702, 703, 16384))
#> [1] "ZZ"  "AAA" "XFD"
```
