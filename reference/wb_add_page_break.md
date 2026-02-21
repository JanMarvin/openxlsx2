# Add a page break to a worksheet

The `wb_add_page_break()` function allows you to manually insert
horizontal or vertical page breaks into a worksheet. These breaks
determine where the spreadsheet software starts a new page when printing
or generating a PDF.

## Usage

``` r
wb_add_page_break(wb, sheet = current_sheet(), row = NULL, col = NULL)
```

## Arguments

- wb:

  A
  [wbWorkbook](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.md)
  object.

- sheet:

  The name or index of the worksheet. Defaults to the current sheet.

- row:

  Integer; the row number where the horizontal page break should be
  inserted.

- col:

  Integer or character; the column number or name (e.g., "B") where the
  vertical page break should be inserted.

## Value

The
[wbWorkbook](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.md)
object, invisibly.

## Details

Manual page breaks override the automatic breaks calculated by the
software based on margins and paper size.

- Row Breaks: When a `row` is specified, the horizontal break is placed
  *above* the specified row. For example, setting `row = 10` ensures
  that Row 10 starts on a new page.

- Column Breaks: When a `col` is specified, the vertical break is placed
  to the *left* of that column. For example, `col = "B"` (or `2`)
  ensures Column B is the first column on the next vertical page.

You must provide either a `row` or a `col` index, but not both in a
single call. To create a page intersection (both horizontal and
vertical), call the function twice.

## Notes

- Manual breaks are visible in "Page Break Preview" mode within most
  spreadsheet applications.

## See also

[`wb_add_worksheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_worksheet.md)

## Examples

``` r
wb <- wb_workbook()
wb$add_worksheet("Sheet 1")
wb$add_data(sheet = 1, x = iris)

wb$add_page_break(sheet = 1, row = 10)
wb$add_page_break(sheet = 1, row = 20)
wb$add_page_break(sheet = 1, col = 2)
```
