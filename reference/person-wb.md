# Helper for adding threaded comments

Adds a person to a workbook, so that they can be the author of threaded
comments in a workbook with
[`wb_add_thread()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_thread.md)

## Usage

``` r
wb_add_person(wb, name = NULL, id = NULL, user_id = NULL, provider_id = "None")

wb_get_person(wb, name = NULL)
```

## Arguments

- wb:

  a Workbook

- name:

  the name of the person to display.

- id:

  (optional) the display id

- user_id:

  (optional) the user id

- provider_id:

  (optional) the provider id

## See also

[`wb_add_thread()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_thread.md)
