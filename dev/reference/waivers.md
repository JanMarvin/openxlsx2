# `openxlsx2` waivers

Waiver functions for `openxlsx2` functions.

- `current_sheet()` uses
  [`wb_get_active_sheet()`](https://janmarvin.github.io/openxlsx2/dev/reference/active_sheet-wb.md)
  by default if performing actions on a worksheet, for example when you
  add data.

- `next_sheet()` is used when you add a new worksheet, a new chartsheet
  or when you add a pivot table. It is defined as available sheets + 1L.

## Usage

``` r
current_sheet()

next_sheet()

na_strings()
```

## Value

An object of class `openxlsx2_waiver`
