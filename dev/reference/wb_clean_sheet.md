# Clear content and formatting from a worksheet

The `wb_clean_sheet()` function removes data, formulas, and formatting
from a worksheet. It can be used to wipe an entire sheet clean or to
target specific cell regions (`dims`). This is particularly useful when
you want to reuse an existing sheet structure but replace the data or
reset the styling.

## Usage

``` r
wb_clean_sheet(
  wb,
  sheet = current_sheet(),
  dims = NULL,
  numbers = TRUE,
  characters = TRUE,
  styles = TRUE,
  merged_cells = TRUE,
  hyperlinks = TRUE
)
```

## Arguments

- wb:

  A
  [wbWorkbook](https://janmarvin.github.io/openxlsx2/dev/reference/wbWorkbook.md)
  object.

- sheet:

  The name or index of the worksheet to clean. Defaults to the current
  sheet.

- dims:

  Optional character string defining a cell range (e.g., "A1:G20"). If
  `NULL`, the entire worksheet is cleaned.

- numbers:

  Logical; if `TRUE`, removes all numeric values, booleans, and error
  codes.

- characters:

  Logical; if `TRUE`, removes all text strings (shared, inline, or
  formula-based strings).

- styles:

  Logical; if `TRUE`, removes all cell styles and resets formatting to
  default.

- merged_cells:

  Logical; if `TRUE`, unmerges all cells (or those within `dims`).

- hyperlinks:

  Logical; if `TRUE`, removes hyperlinks from the specified region.

## Value

The
[wbWorkbook](https://janmarvin.github.io/openxlsx2/dev/reference/wbWorkbook.md)
object, invisibly.

## Details

Unlike deleting a worksheet, cleaning a sheet preserves the sheet's
existence, name, and properties (like tab color or sheet views) while
emptying the cell-level data.

Selective Removal: By toggling the logical arguments, you can choose
exactly what to discard. For example, you can remove data but keep the
cell styles (borders, fills), or vice-versa.

- Numbers/Characters: Targeting these specifically allows you to clear
  constants while potentially leaving other elements intact.

- Styles: Resets cells to the workbook's default appearance.

- Merged Cells: Unmerges ranges; if `dims` is provided, only merges
  within that range are broken.

## Notes

- Currently, this function does not remove objects like images, charts,
  comments, or pivot tables.
