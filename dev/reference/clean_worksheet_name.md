# Clean worksheet name

Cleans a worksheet name by removing legal characters.

## Usage

``` r
clean_worksheet_name(x, replacement = " ")
```

## Arguments

- x:

  A vector, coerced to `character`

- replacement:

  A single value to replace illegal characters by.

## Value

x with bad characters removed

## Details

Illegal characters are considered `\`, `/`, `?`, `*`, `:`, `[`, and `]`.
These must be intentionally removed from worksheet names prior to
creating a new worksheet.
