# Set headers and footers of a worksheet

Set document headers and footers. You can also do this when adding a
worksheet with
[`wb_add_worksheet()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_worksheet.md)
with the `header`, `footer` arguments and friends. These will show up
when printing an xlsx file.

## Usage

``` r
wb_set_header_footer(
  wb,
  sheet = current_sheet(),
  header = NULL,
  footer = NULL,
  even_header = NULL,
  even_footer = NULL,
  first_header = NULL,
  first_footer = NULL,
  align_with_margins = NULL,
  scale_with_doc = NULL,
  ...
)
```

## Arguments

- wb:

  A Workbook object

- sheet:

  A name or index of a worksheet

- header, even_header, first_header, footer, even_footer, first_footer:

  Character vector of length 3 corresponding to positions left, center,
  right. `header` and `footer` are used to default additional arguments.
  Setting `even`, `odd`, or `first`, overrides `header`/`footer`. Use
  `NA` to skip a position.

- align_with_margins:

  Align header/footer with margins

- scale_with_doc:

  Scale header/footer with document

- ...:

  additional arguments

## Details

Headers and footers can contain special tags

- **&\[Page\]** Page number

- **&\[Pages\]** Number of pages

- **&\[Date\]** Current date

- **&\[Time\]** Current time

- **&\[Path\]** File path

- **&\[File\]** File name

- **&\[Tab\]** Worksheet name

## Examples

``` r
wb <- wb_workbook()

# Add example data
wb$add_worksheet("S1")$add_data(x = 1:400)
wb$add_worksheet("S2")$add_data(x = 1:400)
wb$add_worksheet("S3")$add_data(x = 3:400)
wb$add_worksheet("S4")$add_data(x = 3:400)

wb$set_header_footer(
  sheet = "S1",
  header = c("ODD HEAD LEFT", "ODD HEAD CENTER", "ODD HEAD RIGHT"),
  footer = c("ODD FOOT RIGHT", "ODD FOOT CENTER", "ODD FOOT RIGHT"),
  even_header = c("EVEN HEAD LEFT", "EVEN HEAD CENTER", "EVEN HEAD RIGHT"),
  even_footer = c("EVEN FOOT RIGHT", "EVEN FOOT CENTER", "EVEN FOOT RIGHT"),
  first_header = c("TOP", "OF FIRST", "PAGE"),
  first_footer = c("BOTTOM", "OF FIRST", "PAGE")
)

wb$set_header_footer(
  sheet = 2,
  header = c("&[Date]", "ALL HEAD CENTER 2", "&[Page] / &[Pages]"),
  footer = c("&[Path]&[File]", NA, "&[Tab]"),
  first_header = c(NA, "Center Header of First Page", NA),
  first_footer = c(NA, "Center Footer of First Page", NA)
)

wb$set_header_footer(
  sheet = 3,
  header = c("ALL HEAD LEFT 2", "ALL HEAD CENTER 2", "ALL HEAD RIGHT 2"),
  footer = c("ALL FOOT RIGHT 2", "ALL FOOT CENTER 2", "ALL FOOT RIGHT 2")
)

wb$set_header_footer(
  sheet = 4,
  first_header = c("FIRST ONLY L", NA, "FIRST ONLY R"),
  first_footer = c("FIRST ONLY L", NA, "FIRST ONLY R")
)

# ---- Updating the header ----
## Variant a
## this will keep the odd and even header / footer from the original header /
## footerkeep the first header / footer and will set the first page header /
## footer and will use the original header / footer for the missing element
wb$set_header_footer(
  header = NA,
  footer = NA,
  even_header = NA,
  even_footer = NA,
  first_header = c("FIRST ONLY L", NA, "FIRST ONLY R"),
  first_footer = c("FIRST ONLY L", NA, "FIRST ONLY R")
)

## Variant b
## this will keep the first header / footer only and will use the missing
## element from the original header / footer
wb$set_header_footer(
  first_header = c("FIRST ONLY L", NA, "FIRST ONLY R"),
  first_footer = c("FIRST ONLY L", NA, "FIRST ONLY R")
)
```
