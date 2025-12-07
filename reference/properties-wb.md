# Modify workbook properties

This function is useful for workbooks that are loaded. It can be used to
set the workbook `title`, `subject` and `category` field. Use
[`wb_workbook()`](https://janmarvin.github.io/openxlsx2/reference/wb_workbook.md)
to easily set these properties with a new workbook.

## Usage

``` r
wb_get_properties(wb)

wb_set_properties(
  wb,
  creator = NULL,
  title = NULL,
  subject = NULL,
  category = NULL,
  datetime_created = NULL,
  datetime_modified = NULL,
  modifier = NULL,
  keywords = NULL,
  comments = NULL,
  manager = NULL,
  company = NULL,
  custom = NULL
)
```

## Arguments

- wb:

  A Workbook object

- creator:

  Creator of the workbook (your name). Defaults to login username or
  `options("openxlsx2.creator")` if set.

- title, subject, category, keywords, comments, manager, company:

  Workbook property, a string.

- datetime_created:

  The time of the workbook is created

- datetime_modified:

  The time of the workbook was last modified

- modifier:

  A character string indicating who was the last person to modify the
  workbook

- custom:

  A named vector of custom properties added to the workbook

## Value

A wbWorkbook object, invisibly.

## Details

To set properties, the following XML core properties are used.

- title = dc:title

- subject = dc:subject

- creator = dc:creator

- keywords = cp:keywords

- comments = dc:description

- modifier = cp:lastModifiedBy

- datetime_created = dcterms:created

- datetime_modified = dcterms:modified

- category = cp:category

In addition, manager and company are used.

## See also

[`wb_workbook()`](https://janmarvin.github.io/openxlsx2/reference/wb_workbook.md)

## Examples

``` r
file <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
wb <- wb_load(file)
wb$get_properties()
#>                creator               modifier       datetime_created 
#> "Jan Marvin Garbuszus" "Jan Marvin Garbuszus" "2023-05-29T07:43:12Z" 
#>      datetime_modified                company 
#> "2023-05-29T10:47:37Z"                     "" 

# Add a title to properties
wb$set_properties(title = "my title")
wb$get_properties()
#>                creator               modifier       datetime_created 
#> "Jan Marvin Garbuszus" "Jan Marvin Garbuszus" "2023-05-29T07:43:12Z" 
#>      datetime_modified                  title                company 
#> "2023-05-29T10:47:37Z"             "my title"                     "" 
```
