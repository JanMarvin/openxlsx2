# openxlsx2 (development version)

## New features

* Add function to remove conditional formatting from worksheet `wb_remove_conditional_formatting()` [1011](https://github.com/JanMarvin/openxlsx2/pull/1011)

* Silence a warning triggered by a folder called `"[trash]"`. [1012](https://github.com/JanMarvin/openxlsx2/pull/1012)

* Initial support for pivot table timelines. [1016](https://github.com/JanMarvin/openxlsx2/pull/1016)

* Add `wb_add_timeline()` and extend `wb_add_slicer()`. [1017](https://github.com/JanMarvin/openxlsx2/pull/1017)

## Fixes

* Fixed an issue with non consecutive dims, where columns or rows were silently dropped. [1015](https://github.com/JanMarvin/openxlsx2/pull/1015)

* Fixes to `wb_clone_worksheet()` cloning drawings and images should be restored. [1016](https://github.com/JanMarvin/openxlsx2/pull/1016)

* Fixed an issue where non consecutive columns with special types would overlap. If columns A and C were dates, column B would be formatted as date too. [1026](https://github.com/JanMarvin/openxlsx2/pull/1026)


***************************************************************************

# openxlsx2 1.6

## New features

* Helper to read sensitivity labels from files and apply them to workbooks. 
[983](https://github.com/JanMarvin/openxlsx2/pull/983)

* It is now possible to pass non consecutive dims like `"A1:B1,C2:D2"` to various style helpers like `wb_add_fill()`. In addition it is now possible to write a data set into a predefined dims region using `enforce = TRUE`. This handles either `","` or `";"` as cell separator. [993](https://github.com/JanMarvin/openxlsx2/pull/993)

```r
# force a dataset into a specific cell dimension
wb <- wb_workbook()$add_worksheet()
wb$add_data(dims = "I2:J2;A1:B2;G5:H6", x = matrix(1:8, 4, 2), enforce = TRUE)
```

## Fixes

* Allow writing data frames with zero rows. [987](https://github.com/JanMarvin/openxlsx2/pull/987)

* `wb_dims()` has been improved and is safer on 0-length inputs. In particular, it will error for a case where a `cols` doesn't exist in `x`  ([990](https://github.com/JanMarvin/openxlsx2/pull/990), @olivroy).

```r
# Previously created a wrong dims
wb_dims(x = mtcars, cols = "non-existent-col")
# Now errors
```

* `wb_set_col_widths()` is more strict about its arguments. If you provide `cols`, `widths`, or `hidden` don't have appropriate length, it will throw a warning. This may change to an error in the future, so it is recommended to use appropriate values. ([991](https://github.com/JanMarvin/openxlsx2/pull/991), @olivroy).


***************************************************************************

# openxlsx2 1.5

(This was updated post release.)

## New features

* It's now possible to pass array formula vectors to `wb_add_formula()`. [958](https://github.com/JanMarvin/openxlsx2/pull/958)

* `wb_add_data_table()` gained a new `total_row` argument. This allows to add a total row to spreadsheets including text and spreadsheet formulas. [959](https://github.com/JanMarvin/openxlsx2/pull/959)

* `wb_dims()` now accepts `from_dims` to specify a starting cell [960](https://github.com/JanMarvin/openxlsx2/pull/960).

* You can now set `options(openxlsx2.na.strings)` to a value  to have a default value for `na.strings` in `wb_add_data()`, `wb_add_data_table()`, and `write_xlsx()` [968](https://github.com/JanMarvin/openxlsx2/pull/968).

* The direction vectors are written is now controlled via `dims`. Previously it was required to transpose a vector to write it horizontally: `wb_add_data(x = t(letters), col_names = FALSE)`. Now the direction is defined by `dims`. The default is still to write vectors vertically, but for a horizontal vector it is possible to write `wb_add_data(x = letters, dims = "A1:Z1")`. This change impacts vectors, hyperlinks and formulas and basically everything that is not a two dimensional `x` object.

```r
# before (workaround needed)
wb$add_data(dims = wb_dims(rows = 1, cols = 1:3), x = t(c(4, 5, 8)), col_names = FALSE)
# now (listens to dims)
wb$add_data(dims = wb_dims(rows = 1, cols = 1:3), x = c(4, 5, 8))
```

## Fixes

* Export `wb_add_ignore_error()`. [955](https://github.com/JanMarvin/openxlsx2/pull/955)


***************************************************************************


# openxlsx2 1.4

## New features

* Experimental support for calculation of pivot table fields. [892](https://github.com/JanMarvin/openxlsx2/pull/892)

* Improve sparkline creation with new options and support snake case arguments. [920](https://github.com/JanMarvin/openxlsx2/pull/920)

* Experimental support to get and set base colors. [938](https://github.com/JanMarvin/openxlsx2/pull/938)

## Fixes

* Character strings with XML content were not written correctly: `a <br/> b` was converted to something neither we nor spreadsheet software was able to decipher. [895](https://github.com/JanMarvin/openxlsx2/pull/895)

* Restore `first_active_row`/`first_active_col` and `first_col`/`first_row` functionality in `write_xlsx()`. [903](https://github.com/JanMarvin/openxlsx2/pull/903)

* Further attempts to fix pivot table sorting. [912](https://github.com/JanMarvin/openxlsx2/pull/912) In addition improve handling of non distinct names in `wb_data()` and add `create_pivottablestyle()`. [914](https://github.com/JanMarvin/openxlsx2/pull/914)

* Document adding background color and images to comments and fix adding more than two images as background. [919](https://github.com/JanMarvin/openxlsx2/pull/919)

* Improvements to `wb_set_base_font()`. This modifies the workbook theme now, including panose values. [935](https://github.com/JanMarvin/openxlsx2/pull/935)

* Hyperlinks now use the color of the theme and the base size. [937](https://github.com/JanMarvin/openxlsx2/pull/937)

* `wb_add_data()` and `wb_add_data_table()` yield better error messages if attempting to add data to an empty workbook ([942](https://github.com/JanMarvin/openxlsx2/pull/942), @olivroy).

## Breaking changes

* Updating to themes. This includes updates to the default style `'Office Theme'` [899](https://github.com/JanMarvin/openxlsx2/pull/899)
  * This includes switching to the new default font `'Aptos Narrow'`
  * A new style `'Office 2013 - 2022 Theme'` was added

  Users that want to remain on the old style should use `wb_workbook(theme = 'Office 2013 - 2022 Theme')` or `wb_set_base_font(font_name = "Calibri")`.

## Maintenance

* Code simplifications ([924](https://github.com/JanMarvin/openxlsx2/pull/924), @olivroy), partial matching ([923](https://github.com/JanMarvin/openxlsx2/pull/923), @olivroy)

* Updates to test files and findings by various linters ([922](https://github.com/JanMarvin/openxlsx2/pull/922), @olivroy)

* Improve [wbWorkbook documentation](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.html) to help disambiguate the chaining and piping methods ([936](https://github.com/JanMarvin/openxlsx2/pull/936), @olivroy and [928](https://github.com/JanMarvin/openxlsx2/pull/928)).


***************************************************************************


# openxlsx2 1.3

## Documentation improvement

* Further tweaks to documentation and vignettes to make them more consistent.
  * `wb_add_pivot_table()` / `wb_add_slicer()`
  * `wb_load()`: `calc_chain` is no longer visible and the previous text that might have been misleading in regards of its use, has been replaced by a more detailed description of what are the consequences of keeping the calculation chain

## New features

* Allow further modifications of comments. The background can now be filled with a color or an image. [870](https://github.com/JanMarvin/openxlsx2/pull/870)

* Added `wb_set_cell_style_across()` to apply a cell style to selected columns and rows. This allows unlocking cells to make use column and row properties of `wb_protect_worksheet()` which require additional cell styles (see issue 871 for a more detailed explanation). `wb_set_cell_style()` now accepts a cell dimension in the `style` argument. [873](https://github.com/JanMarvin/openxlsx2/pull/873)

## Fixes

* `wb_add_ignore_error()` now returns a `wbWorkbook`. [865](https://github.com/JanMarvin/openxlsx2/pull/865)

* Deactivate the `is_hyperlink` check for non-dataframe objects in `wb_add_data()`. Internally, `vapply()` is applied to the input object, which is applied column-wise for a data frame and cell-wise for a matrix. This speeds up the writing of larger matrices considerably. [876](https://github.com/JanMarvin/openxlsx2/pull/876)

* Column style `currency` is now correctly applied to numeric vectors. Previously it was not handled. This applies the built in spreadsheet style for currency presumably linked to the spreadsheet software locale. [879](https://github.com/JanMarvin/openxlsx2/pull/879)

* `wb_to_df(col_names = FALSE)` no longer drops column names from logical vectors. Previously, column names were replaced by `NA`. Now the column name is returned as a cell value in a character column. [884](https://github.com/JanMarvin/openxlsx2/pull/884)


***************************************************************************


# openxlsx2 1.2

## New features

* Add new params to `wb_add_pivot_table()`. It is now possible to set the `show_data_as` value and set a tabular table design. [833](https://github.com/JanMarvin/openxlsx2/pull/833)

## Fixes

* Previously formulas written as data frames were not xml escaped. [834](https://github.com/JanMarvin/openxlsx2/pull/834)

* Improve drawing relationship id selection that could cause issues with unordered relationship ids in loaded workbooks. [838](https://github.com/JanMarvin/openxlsx2/pull/838)

* Improve copying cells in transpose mode and with hyperlinks. [850](https://github.com/JanMarvin/openxlsx2/pull/850)

* Options `openxlsx2.maxWidth` and `openxlsx2.minWidth` are now respected as documented when setting column widths with `wb_set_col_widths()`. [847](https://github.com/JanMarvin/openxlsx2/issues/847)


***************************************************************************


# openxlsx2 1.1

## New features

* Basic support for cloning worksheets across workbooks was added in `wb_clone_worksheet()`. This should work with tabular and formula data, images and charts. Support for more complex worksheets is known pending. [622](https://github.com/JanMarvin/openxlsx2/pull/622)

* Enable loading workbooks without sheets with `wb_load(sheet = NULL)`. This is helpful if the sheets contain a lot of data and only some general workbook data is of interest. [810](https://github.com/JanMarvin/openxlsx2/pull/810)

* It is no longer needed to manually create columns for `wb_group_cols()`. [781](https://github.com/JanMarvin/openxlsx2/pull/781)

* It is now possible to sort pivot tables created by `wb_add_pivot_table()`. [795](https://github.com/JanMarvin/openxlsx2/pull/795)

* Basic support for the `xlsb` file format. We parse the binary file format into pseudo-openxml files that we can import. Therefore, after importing, it is possible to interact with the file as if it had been provided as `xlsx` in the first place. This is of course slower than reading directly from the binary file. Our implementation is also still missing some features: some array formulas are still broken, conditional formatting and data validation are not implemented, nor are pivot tables and slicers. [688](https://github.com/JanMarvin/openxlsx2/pull/688)
  * Please note that `openxlsx2` is not a security tool and the `xlsb` parser was written with the intention of reading valid `xlsb` files on little endian systems.
  * Please do not raise issues about `xlsb` in terms of speed or completeness of features of the openxml standard. If you have access to other file conversion tools, such as spreadsheet software, they may provide a better solution for your needs.
  * Writing `xlsb` files is outside the scope of this project.

* New set of function `wb_get_properties()`/`wb_set_properties()` to view and modify workbook properties. [782](https://github.com/JanMarvin/openxlsx2/pull/782) This was subsequently improved to handle more workbook properties like `company` and `manager`. ([799](https://github.com/JanMarvin/openxlsx2/pull/799), @olivroy)

* Basic (experimental) support to add slicers to pivot tables created by `openxlsx2`. [822](https://github.com/JanMarvin/openxlsx2/pull/822)

## Fixes

* Removing the worksheet that is the active tab does no longer result in warnings in spreadsheet software. [792](https://github.com/JanMarvin/openxlsx2/pull/792)

* `update_table()` works now without if the table has no `autofilter`. [802](https://github.com/JanMarvin/openxlsx2/pull/802)


***************************************************************************


# openxlsx2 1.0

## Breaking changes

* `wb_get_worksheet()`, `wb_ws()`, `wb_get_sheet_name()` are no longer exported. [735](https://github.com/JanMarvin/openxlsx2/pull/735) They never worked as expected.
  * Use `wb_get_sheet_names()`, `wb_get_active_sheet()`, `wb_get_selected()` instead.

* `wbChartSheet` is now internal ([760](https://github.com/JanMarvin/openxlsx2/pull/760), @olivroy)

## Deprecated functions

These functions are no longer recommended. A [guide](https://janmarvin.github.io/openxlsx2/dev/reference/openxlsx2-deprecated.html) was created to help users
They will continue to work for some time, but changing to newer functions is recommended.

* `delete_data()` is deprecated in favor of `wb_remove_tables()` or `wb_clean_sheet()`
* `create_comment()` is deprecated in favor of `wb_comment()`. Note that while `wb_comment()` has some new defaults, the behavior of `create_comment()` has not changed. ([758](https://github.com/JanMarvin/openxlsx2/pull/758), @olivroy)

## New features

* `wb_comment()` is a new helper function to help create `wbComment` objects
  It is very similar to `create_comment()`, with the following differences:
  * `author` looks at `options("openxlsx2.creator")`; (`create_comment()` only used `sys.getenv("USERNAME")` / `Sys.getenv("USER")`)
  * `visible` defaults to `FALSE` to account for modern spreadsheet software behavior. (`create_comment()`, it is `TRUE`).
  * `width` and `height` must now be of length 1. (In `create_comment()`, the first element is taken, other are ignored.)

* `wb_get_sheet_names()` gains a `escape` argument to allow special XML characters to be escaped. [252](https://github.com/JanMarvin/openxlsx2/issues/252)

* `wb_color()` now accepts hex colors with leading sharp (e.g. "#FFFFF") [728](https://github.com/JanMarvin/openxlsx2/pull/728).

* `wb_merge_cells()` gains a `solve` argument. This allows to solve cell intersecting regions. [733](https://github.com/JanMarvin/openxlsx2/pull/733)

* `wb_add_comment(comment = "x")` no longer errors when a comment as a character vector no longer fails ([758](https://github.com/JanMarvin/openxlsx2/pull/758), @olivroy)


## Documentation improvement

* Tweaks to documentation and vignettes to make them more consistent.

## Fixes

* `wb_to_df()` now handles `date1904` detection. Previous results from this somewhat rare file type were using the default timezone origin of 1900-01-01. [737](https://github.com/JanMarvin/openxlsx2/pull/737)
* `wb_load()` handles more cases.
  * With workbooks with threaded comments [731](https://github.com/JanMarvin/openxlsx2/issues/731)
  * With workbooks with embeddings other than docx [732](https://github.com/JanMarvin/openxlsx2/pulls/732)
  * with workbooks with long hyperlinks [753](https://github.com/JanMarvin/openxlsx2/issues/753)

* `wb_load()` adds the `path` to the `wbWorkbook` object. [741](https://github.com/JanMarvin/openxlsx2/issues/741)
* `wb_set_header_footer()` now works with special characters [747](https://github.com/JanMarvin/openxlsx2/issues/747)

## Internal changes

* `wb_get_active_sheet()`, `wb_set_active_sheet()`, `wb_get_selected()` and `wb_set_selected()`, `wb_get_named_regions()` are now wrapper functions. [735](https://github.com/JanMarvin/openxlsx2/pull/735)


***************************************************************************


# openxlsx2 0.8

## API Change

* Function arguments are now defaulting to `snake_case`. For the time being, both arguments are accepted and `camelCase` will be switched to `snake_case` under the hood. Some examples are currently still displaying `camelCase` and maybe some `camelCase` function slipped through. [678](https://github.com/JanMarvin/openxlsx2/pull/678)

* `write_formula()`, `write_data()`, `write_datatable()`, `write_comment()` are no longer recommended,
  * Use `wb_add_formula()`, `wb_add_data()`, `wb_add_data_table()`, `wb_add_comment()` instead.

## Breaking changes

* Order of arguments in `wb_add_named_region()` changed, because previously overlooked `dims` argument was added.
* In various functions the order of `dims` was changed, to highlight it's importance

* Cleanups
  * remove deprecated functions

  * remove deprecated arguments
    * `xy` argument
    * arguments `col`, `row`, `cols`, `rows`. `start_col`, `start_row` and `gridExpand` were deprecated in favor of `dims`. Row and column vectors can be converted to `dims` using `wb_dims()`.
    * `xlsx_file` in favor of `file` in `wb_to_df()`

  * deprecating function
    * `convertToExcelDate()` for `convert_to_excel_date()`
    * `wb_grid_lines()` for `wb_set_grid_lines()`

  * make `get_cell_refs()`, `get_date_origin()`, `guess_col_type()`, and `write_file()`, `dataframe_to_dims()`, `dims_to_dataframe()`, `wb_get_sheet_name()` internal functions
  * make classes `styles_mgr()`, `wbSheetData`, `wbWorksheet`, `wbComment`, `wbHyperlink` internal

## New features

* `wb_dims()` was added as a more convenient replacement for `rowcol_to_dims()`.([691](https://github.com/JanMarvin/openxlsx2/pull/691) and [702](https://github.com/JanMarvin/openxlsx2/pull/702), @olivroy) The new function can take either numeric (for rows or columns) or character (column) vectors, in addition it is able to create dimensions for R objects that are coercible to data frame. This allows the following variants:
  * `wb_dims(1:5, letters)`
  * `wb_dims(1:5, 1:26)`
  * `wb_dims(x = matrix(1, 5, 26))`
  * `wb_dims(x = mtcars, from_col = "C", from_row = 2, row_names = TRUE)`
* Handling of thread comments is now possible via `wb_add_thread()`. This includes options to reply and resolve comments.

## Fixes

* Improve `show_formula`. Previously it was called to early in the function and skipped a few cases. [715](https://github.com/JanMarvin/openxlsx2/pull/715)

## Refactoring

* Cleanup / revisit documentation and vignettes ([682](https://github.com/JanMarvin/openxlsx2/pull/682), @olivroy)

* The [function index](https://janmarvin.github.io/openxlsx2/reference/) has been improved. ([717](https://github.com/JanMarvin/openxlsx2/pull/717), @olivroy)


***************************************************************************


# openxlsx2 0.7.1

## New features

* It is now possible to apply a specific theme to a workbook. [630](https://github.com/JanMarvin/openxlsx2/pull/630)

* Removed a few of the former example files and updated the code to use a new default example. This changes internal testing to only run locally if online and external files are required. This reduces the package footprint a little, because 1MB of xlsx files are now excluded. [632](https://github.com/JanMarvin/openxlsx2/pull/632)

* The handling of `fmt_txt()` objects has been improved. It now creates objects of class `fmt_txt` with their own `print()`, `+`, and `as.character()` methods. Such objects can now also be used as `text` in `create_comment()`. [636](https://github.com/JanMarvin/openxlsx2/pull/636)

* Improve support for inputs with `labels` attribute. If e.g. a `factor` label is numeric, we now try to write the label as number. This should impact the way other partially labeled variables are written. [639](https://github.com/JanMarvin/openxlsx2/pull/639)

* Added new wrapper function `wb_add_named_style()` this supports pre-defined theme aware cell styles like `Title` or `Note`. In addition loading of cell styles was improved and additional custom cell styles should be available as well. [628](https://github.com/JanMarvin/openxlsx2/pull/628)

* Provide additional options to write special characters in non-unicode environments. [641](https://github.com/JanMarvin/openxlsx2/pull/641)

* Add `wb_add_dxfs_style()` as single line wrapper to create dxf styles used in conditional formatting and custom table styles. [665](https://github.com/JanMarvin/openxlsx2/pull/665)

## Fixes

* On load `app.xml` is now assigned to `wb$app`. Previously it was loaded but not assigned. [629](https://github.com/JanMarvin/openxlsx2/pull/629)

* Previously if `wb_to_df()` was used with argument `cols`, columns that were missing were created at the end of the output frame. Now columns are returned ordered. [631](https://github.com/JanMarvin/openxlsx2/pull/631)

* Fix a bug in `wb_load()` that was modifying the cell range of conditional formatting. [647](https://github.com/JanMarvin/openxlsx2/pull/647)

## Breaking changes

* Order of arguments in `wb_add_conditional_formatting()` changed, because previously overlooked `dims` argument was added. [642](https://github.com/JanMarvin/openxlsx2/pull/642)

* New argument `gradientFill` was added to `create_dxfs_style()`. [651](https://github.com/JanMarvin/openxlsx2/pull/651)

* Special characters are now escaped in conditional formatting. Hence, previously manually escaped conditional formatting needs updates. [666](https://github.com/JanMarvin/openxlsx2/pull/666)


***************************************************************************


# openxlsx2 0.7

## New features

* The `dims` argument of `wb_add_formula()` can be used to create a array references. A new `cm` argument was added which might be useful, if formulas previously created addition `@` in spreadsheet software. Examples how to use formulas were added to a new vignette. [593](https://github.com/JanMarvin/openxlsx2/pull/593)

* Allow using custom data table styles. This fixes a few minor style inconsistencies. [594](https://github.com/JanMarvin/openxlsx2/pull/594)

* Allow reading and writing `hms` columns. [601](https://github.com/JanMarvin/openxlsx2/pull/601)

* Import `tableStyles` with `wb_load()` and improve `dxf` style creation. [603](https://github.com/JanMarvin/openxlsx2/pull/603)

* Add `fmt_txt()` to style character strings. [607](https://github.com/JanMarvin/openxlsx2/pull/607)

* Add new wrapper to ignore worksheet errors `wb_add_ignore_error()`. [617](https://github.com/JanMarvin/openxlsx2/pull/617)

* Add new wrapper to update table references `wb_update_table()`. [606](https://github.com/JanMarvin/openxlsx2/pull/606)

## Fixes

* Improve handling of non standard `OutDec` options. [611](https://github.com/JanMarvin/openxlsx2/pull/611)

* `openxlsx2` now does a better job of trying to return `character` values from classes that are foreign to it. This has been going on for quite some time, although previously we had a bug that treated such classes as `numeric`, resulting in corrupted xlsx files. [615](https://github.com/JanMarvin/openxlsx2/pull/615)

* We now return a few additional xml arguments from worksheets. [617](https://github.com/JanMarvin/openxlsx2/pull/617)


***************************************************************************


# openxlsx2 0.6.1

## New features

* Improve `col2int()` to accept column ranges like `col2int("A:Z")`. This should allow using column ranges in various places like `wb_merge_cells(cols = "B:D", ...)` or `wb_read(cols = c("A","C:D"))`. [575](https://github.com/JanMarvin/openxlsx2/pull/575)

* Add `dims` argument to `wb_add_image()` and `wb_add_plot()`. This can be used to place images starting at a cell or span a cell range. This deprecates `xy` in `wb_add_plot()`. This adds colOffset and rowOffset to `wb_add_drawing()` and `wb_add_mschart()` and `wb_add_chart_xml()`. [578](https://github.com/JanMarvin/openxlsx2/pull/578)

* Add `skipHiddenCols` and `skipHiddenRows` to `wb_to_df()`. In this way, hidden columns and rows are ignored, assuming that the person who has hidden them assumes that they are not important. [579](https://github.com/JanMarvin/openxlsx2/pull/579)

* When writing `tibble` use `to.data.frame()` just like in the `data.table` case. [582](https://github.com/JanMarvin/openxlsx2/pull/582)

* Add cleanup internal comment code in `write_comment()`. This should not impact the workbook wrapper code in `wb_add_comment()`. [586](https://github.com/JanMarvin/openxlsx2/pull/586)

* Added chain functions for `wb_to_df()` and `wb_load()`. [587](https://github.com/JanMarvin/openxlsx2/pull/587)


***************************************************************************


# openxlsx2 0.6

## New features

* Styles arguments now accept logical and numeric arguments where applicable. [558](https://github.com/JanMarvin/openxlsx2/pull/558)

* Adding `dims` argument to `wb_clean_sheet()`. This allows to clean only a selected range. [563](https://github.com/JanMarvin/openxlsx2/pull/563)

## Fixes

* `na.strings = NULL` is no longer ignored in `write_xlsx()` [552](https://github.com/JanMarvin/openxlsx2/issues/552)

* Explicit type conversion to date and datetime is finally available. [551](https://github.com/JanMarvin/openxlsx2/pull/551)

## Breaking changes

* `skipEmptyCols` and `skipEmptyRows` behavior in `wb_to_df()` related functions was switched to include empty columns that have a name. Previously we would exclude columns if they were empty, even if they had a name. [555](https://github.com/JanMarvin/openxlsx2/pull/555)

* Cleanups in [548](https://github.com/JanMarvin/openxlsx2/pull/548)
  * remove deprecated functions
    * `cloneSheetStyle()`
    * `get_cell_style()`
    * `set_cell_style()`
    * `wb_conditional_formatting()`

  * remove deprecated arguments
    * `xy` argument for `write_data_table()` interacting functions
    * `file` from `xl_open()`
    * `definedName` from `wb_to_df()` interacting functions

  * deprecating function
    * `get_named_regions()` for `wb_get_named_regions()`


***************************************************************************


# openxlsx2 0.5.1

## New features

* Allow hierarchical grouping. `wb_group_cols`/`wb_group_rows` now accept nested lists as grouping variable. [537](https://github.com/JanMarvin/openxlsx2/pull/537)

* It is now possible to add form control types `Checkbox`, `Radio` and `Drop` to a workbook using `wb_add_form_control()`. [533](https://github.com/JanMarvin/openxlsx2/pull/533)

* Improve `wb_to_df(fillMergedCells = TRUE)` to work better with dimensions. It is now possible to fill cells where the merged cells intersect with the selected dimensions. [541](https://github.com/JanMarvin/openxlsx2/pull/541)

* Speedup cell initialization. This is used in wb_style functions like `wb_add_numfmt()`. The previous loop was replaced with a faster implementation. [545](https://github.com/JanMarvin/openxlsx2/pull/545)

* Improve date detection in `wb_to_df()`. This improves date and posix detection with custom date formats. [547](https://github.com/JanMarvin/openxlsx2/pull/547)

* `na_strings()` is now used as the explicit default value for `na.strings` parameters in exported workbook functions. [473](https://github.com/JanMarvin/openxlsx2/issues/473)

* waiver functions (i.e., `next_worksheet()`, `current_worksheet()`, and `na_strings()`) are now exports [474](https://github.com/JanMarvin/openxlsx2/issues/474)

## Fixes

* Fixed a bug when loading input with multiple sheets where not every sheet contains a drawing/comment. Previously we assumed that every sheet had a comment and ordered them incorrectly. This caused confusion in spreadsheet software. [536](https://github.com/JanMarvin/openxlsx2/pull/536)

* Fixed a bug with files containing 10 or more external references. In this case we did not load the references in numeric order and instead as "1.xml", "10.xml", ..., "2.xml", ... This jumbled up the external references. [538](https://github.com/JanMarvin/openxlsx2/pull/538)


***************************************************************************


# openxlsx2 0.5

## New features

* Improve column and row grouping. It is now possible to group by list, so that you can create various levels of groupings. [486](https://github.com/JanMarvin/openxlsx2/pull/486)

* `writeData()` calls `force(x)` to evaluate the object before options are set ([#264](https://github.com/ycphs/openxlsx/issues/264))

* `tabColor` in `wb_add_worksheet()` now allows passing `wb_color()`. [500](https://github.com/JanMarvin/openxlsx2/pull/500)

* Add `wb_copy_cells()` a wrapper that allows copying cell ranges in a workbook as direct copy, as reference or as value. [515](https://github.com/JanMarvin/openxlsx2/pull/515)

* Experimental option: `openxlsx2.string_nums` to write string numerics differently. A string numeric is a numeric in a string like: `as.character(1.5)`. The option can be
  * 0 = the current default. Writes string numeric as string (the incorrect way according to spreadsheet software)
  * 1 = writes string numeric as numeric with a character flag (the correct way according to spreadsheet software)
  * 2 = convert all string numeric to numeric when writing

  This is experimental, because the impact is somewhat unknown. It might trigger unintended side effects. Feedback is requested.

* Enable writing strings as `sharedStrings` with argument `inline_strings = FALSE`. This creates a `sharedStrings` table in openxml that allows to reuse strings in the workbook efficiently and can reduce the file size if a workbook has many cells that are duplicates. [499](https://github.com/JanMarvin/openxlsx2/pull/499)

* Initial implementation of `wb_add_pivot_table()`. This allows adding native pivot tables to `openxlsx2` workbooks. The pivot table area will remain empty until the sheet is opened in spreadsheet software and evaluated successfully. This feature is newly developed and can cause unexpected side effects. Be aware that using it might currently break workbooks.

## Fixes

* Reading of files with frozen panes and more than one section node was restored. [495](https://github.com/JanMarvin/openxlsx2/pull/495)

* Fixed a copy and paste mistake in `add_border()` which used left borders for right borders. [496](https://github.com/JanMarvin/openxlsx2/pull/496)

* Improve XML unescaping. [497](https://github.com/JanMarvin/openxlsx2/pull/497)

* Fix reading and saving workbooks with multiple slicers per sheet. [505](https://github.com/JanMarvin/openxlsx2/pull/505)

* Fix tab selection always selecting the first sheet since #303. [506](https://github.com/JanMarvin/openxlsx2/pull/506)

## Breaking changes

* Do not export `write_data2()` anymore. This was used in development in the early stages of the package and should not be used directly anymore.

* Only documentation: `openxlsx2` defaults to American English 'color' from now on. Though, we fully support the previous 'colour'. Users will not have to adjust their code. Our documentation only lists `color`, but you can pass `colour` just the same way you used to. [501](https://github.com/JanMarvin/openxlsx2/pull/501) [502](https://github.com/JanMarvin/openxlsx2/pull/502)


***************************************************************************


# openxlsx2 0.4.1

## New features

* Provide new argument `calc_chain` to `wb_load()`. This is set to `FALSE` per default, to ignore the calculation chain if it is found. This change only reflects files imported from third party spreadsheet software and should not be visible to the user. [461](https://github.com/JanMarvin/openxlsx2/pull/461)

* Tweaks to `wb_data()`. Dims is now optional and will select data similar to `wb_to_df()`, similar it allows passing down other `wb_to_df()` arguments. Though, it probably is a good idea not be to creative passing down arguments, not all will result in a usable `wb_data` object. [462](https://github.com/JanMarvin/openxlsx2/pull/462)

* Add `hidden` argument and change the default for `heights` to `NULL` in `set_row_heights()`. This allows changing the row height and/or hiding selected rows. This does not yet provide a way to hide rows per default. [475](https://github.com/JanMarvin/openxlsx2/pull/475)

* Add `wb_add_chartsheet()` for chart sheet support. Along with internal cleanup around chart sheet code. [466](https://github.com/JanMarvin/openxlsx2/pull/466)

## Fixes

* Fix `wb_freeze_pane()`. This changes the load logic a bit. Previously we put everything into `sheetViews` (the frozen pane is part of this). Though `wb_freeze_pane()` assumes that `freezePane` is used. We now try to be smart and split sheetViews upon loading. [465](https://github.com/JanMarvin/openxlsx2/pull/465)

* Previously, adding mschart objects to sheets was only possible if (1) the worksheet already contained a drawing (if the workbook was loaded) or (2) to the last sheet of the workbook. This has now been fixed. Adding mschart objects to any worksheet in the workbook is now possible as intended. [458](https://github.com/JanMarvin/openxlsx2/pull/458)

* Really fix double xml escaping when saving. [467](https://github.com/JanMarvin/openxlsx2/pull/467)

* Improving the drawing logic. There are some workbooks with various drawings per sheet and previously there were combinations possible that were not reflecting this. [478](https://github.com/JanMarvin/openxlsx2/pull/478)


***************************************************************************


# openxlsx2 0.4

## New features

* Provide `rvg` support via `wb_add_drawing()`. This allows integrating `rvg` plots into xlsx files. [449](https://github.com/JanMarvin/openxlsx2/pull/449)

* Improve print options. Defaults to printing grid lines, if the worksheet contains grid lines. [440](https://github.com/JanMarvin/openxlsx2/pull/440)

* Support reading files with form control. [426](https://github.com/JanMarvin/openxlsx2/pull/426)

* Handle input files with chart extensions. [443](https://github.com/JanMarvin/openxlsx2/pull/443)

* Improve writing styles to workbook. Previously every cell was checked, this has been changed to check unique styles. [423](https://github.com/JanMarvin/openxlsx2/pull/423)

* Implement reading custom file properties. [418](https://github.com/JanMarvin/openxlsx2/pull/418)

* Improved `add_named_region()`. This function includes now various xml options. [386](https://github.com/JanMarvin/openxlsx2/pull/386)

* Add ... as argument to `read_xlsx()` and `wb_read()`. [381](https://github.com/JanMarvin/openxlsx2/pull/381)

* Allow reading files with xml namespace created by third party software. [405](https://github.com/JanMarvin/openxlsx2/pull/405)

## Fixes

* Update or remove calculation chain when overwriting formulas in a workbook. [438](https://github.com/JanMarvin/openxlsx2/pull/438)

* Fix double xml escaping when saving. [435](https://github.com/JanMarvin/openxlsx2/pull/435)

* Minor tweak for POSIXct dates and try to avoid the notorious 29Feb1900. [424](https://github.com/JanMarvin/openxlsx2/pull/424)

* Implement reading `customXml` folder for input files with connection. [419](https://github.com/JanMarvin/openxlsx2/pull/419)

* Fixed saving files with `<sheetPr/>` tag. Previously this was wrapped in a second `sheetPr` node. This issue occurs with xlsm files only. [417](https://github.com/JanMarvin/openxlsx2/pull/417)

* Fixed a case where embedded files were assigned incorrectly in worksheet relationships. This caused corrupted output. [403](https://github.com/JanMarvin/openxlsx2/pull/403)

## Breaking changes

* Remove `merge_` functions for styles. [450](https://github.com/JanMarvin/openxlsx2/issues/450)

* Previously if a loaded workbook contained formulas pointing to cells modified by `openxlsx2`, these formulas were not updated, once the workbook was opened in spreadsheet software. This is now enforced, unless the option `openxlsx2.disableFullCalcOnLoad` is set. In this case we would respect the original calculation properties of the  workbook.

* `wb_save()` no longer returns the `path` that the object was saved to, but instead the `wbWorkbook` object, invisibly.  This is consistent with the behavior of others wrappers. [378](https://github.com/JanMarvin/openxlsx2/issues/378)

* Remove never used `all.equal.wbWorkbook()`. The idea was nice, but it never developed into something useful.

* Remove never used `wb_chart_sheet()` function. [399](https://github.com/JanMarvin/openxlsx2/pull/399)

## Internal changes

* Provide `set_sheetview()` in sheets. Can be used to provide a `wbWorkbook` function and wrapper in the future. [399](https://github.com/JanMarvin/openxlsx2/pull/399)


***************************************************************************


# openxlsx2 0.3.1

## New features

* Functions adding data to a workbook now bring an `applyCellStyle` argument. If this is `TRUE` `openxlsx2` will apply a numeric style, if `FALSE` we will simply add the numeric value without additional styling and use the previous cell style. [365](https://github.com/JanMarvin/openxlsx2/pull/365)

* Reading from file or workbook with `showFormula` now returns all formulas found in the workbook. Previously only those with type `e` or `str` were returned. Now we will be able to see formulas like hyperlinks too. [352](https://github.com/JanMarvin/openxlsx2/issues/352)

## Internal changes

* Moved data validation list from x14 to data validation. This enables data validation lists without x14 extension [openxlsx 386](https://github.com/ycphs/openxlsx/issues/386). [347](https://github.com/JanMarvin/openxlsx2/pull/347)

* Removed `build_cell_merges()` and replaced it with a plain R solution. [390](https://github.com/JanMarvin/openxlsx2/pull/390)

## Fixes

* Improvements to setting column widths. Previously values set by `set_col_widths()` were a little off. This has now been improved. There are still corner cases where the column width set with `openxlsx2` does not match those shown in spreadsheet software. Notable differences can be seen with floating point values (e.g., `10L` works while `10.1` is slightly off) and with column width on Mac. [350](https://github.com/JanMarvin/openxlsx2/pull/350)

* Improve `rowNames` when writing data to worksheet. Previously the name for the rownames column defaulted to `1`. This has been changed. Now with data it defaults to an empty cell and with a data table it defaults to `_rowNames_`. [375](https://github.com/JanMarvin/openxlsx2/pull/375)

* Fix the workbook xml relationship file to not include a reference to shared strings per default. This solves [360](https://github.com/JanMarvin/openxlsx2/issues/360) for plain data files written from `openxlsx2`. [363](https://github.com/JanMarvin/openxlsx2/pull/363)

* Adding cell styles has been streamlined to increase consistency. This includes all style functions like `wb_add_font()` and covers all cases of hyperlinks. [365](https://github.com/JanMarvin/openxlsx2/pull/365)

* Fix cloning pivot charts. [361](https://github.com/JanMarvin/openxlsx2/issues/361)

* Fix loading and writing files with slicers. Loading would add a few empty slicer xml files to `Content_Types` and `workbook.xml.rels`. [361](https://github.com/JanMarvin/openxlsx2/issues/361)

* Align the logic for writing data to empty worksheets and updating/writing to worksheets with data. This removes `update_cell_loop()` and changes how `update_cell()` behaves. Not only does this remove duplicated code, it also brings great speed improvements (issue [356](https://github.com/JanMarvin/openxlsx2/issues/356)). [356](https://github.com/JanMarvin/openxlsx2/issues/356)

* It is now possible to use special characters in formulas without coding. Previously `&` had to be encoded like `&amp;` [251](https://github.com/JanMarvin/openxlsx2/issues/251)

## Breaking changes

* Previously deprecated `names.wbWorkbook()` and `names<-.wbWorkbook()` have been removed. [367](https://github.com/JanMarvin/openxlsx2/pull/367)

* Conditional style defaults for `create_dxfs_style()` have changed to be more permissive. Previously we shipped a default font, default font size and font color. This has been changed to better reflect a behavior the user expects. [343](https://github.com/JanMarvin/openxlsx2/issues/343)


***************************************************************************


# openxlsx2 0.3

## New features

* New argument `startCol` in read to data frame functions `wb_to_df()`, `wb_read()` and `read_xlsx()`. [330](https://github.com/JanMarvin/openxlsx2/issues/330)

* New function `wb_colour()` to ease working with color vectors used in `openxlsx2` styles. [292](https://github.com/JanMarvin/openxlsx2/issues/292)

* Deprecated `get_cell_style()` and `set_cell_style()` in favor of newly introduced wrapper functions `wb_get_cell_style()` and `wb_set_cell_style()`. [306](https://github.com/JanMarvin/openxlsx2/issues/306)

* Improvements to `wb_clone_worksheet()`. Cloning of chartsheets as well as worksheets containing charts, pivot tables, drawings and tables is now possible or tweaked. Slicers are removed from the cloned worksheet. [305](https://github.com/JanMarvin/openxlsx2/issues/305)

* Allow writing class `data.table`. [313](https://github.com/JanMarvin/openxlsx2/issues/313)

* Provide `na.numbers` for reading functions, that convert numbers to `NA` in R output. Handle `na.strings` in `write_xlsx()`. [301](https://github.com/JanMarvin/openxlsx2/issues/301)

* Add new option to add sparklines with various style options to worksheets: `wb_add_sparklines()`. Sparklines can be created with `create_sparklines()`. The manual page contains an example. [280](https://github.com/JanMarvin/openxlsx2/issues/280)

* Add new options to data validation. allow type custom, add arguments `errorStyle`, `errorTitle`, `error`, `promptTitle`, `prompt`. [271](https://github.com/JanMarvin/openxlsx2/issues/271)

* Provide `wb_clone_sheet_style()`. This improves upon the now deprecated`cloneSheetStyle()` that existed as an early draft. [233](https://github.com/JanMarvin/openxlsx2/issues/233)

* `wb$add_data()` now checks earlier for missing `x` argument.  [246](https://github.com/JanMarvin/openxlsx2/issues/165)

## Internal changes

* Worksheets added to a `wbWorkbook` no longer contain default references to the `drawings` and `vmlDrawings` directories. Previously, these references were added as `rId1` and `rId2` even if the worksheet did not contain any drawing (e.g., an image or a chart) or vml drawing (e.g., a comment or a button). In such cases certain third party software, strictly following the references in worksheet or `Content_Types` complained about missing files and the import of such files failed completely. [311](https://github.com/JanMarvin/openxlsx2/pull/311)

* Implement loading of user defined chartShapes. Previously this was not implemented instead the previous logic assumed that every sheet has a matching drawing. With chartShapes this no longer is true. The number of drawings and the number of worksheets/chartsheets must not match. [323](https://github.com/JanMarvin/openxlsx2/pull/323)

* When loading files with charts, they are now imported into the `wbWorkbook` object. Previously they were simply copied. This will allow easier interaction with charts in the future. [304](https://github.com/JanMarvin/openxlsx2/pull/304)

* Moving the data validation code from the workbook to the worksheet. Also, `data_validation_list()` is no longer stored in `dataValidationLst`. It has been moved to `extLst`, fixing a bug when saving and adding another data validation list. The code for retrieving the date origin from a workbook has been improved and `get_date_origin(wb, origin = TRUE)` now returns the origin as an integer from a `wbWorkbook`. [299](https://github.com/JanMarvin/openxlsx2/pull/299)

* Removed level4 from XML functions. There was only a single use case for a level4 function that has been solved differently. If level4 is needed, this can be solved using a level3 and additional level2 functions. In addition xml_nodes now return nodes for all reachable nestings, therefore `xml_node("<a/><a/>", "a")` will now return a character vector of length two. For `xml_node("<a/><b/>", "a")` only a single character vector is returned. [280](https://github.com/JanMarvin/openxlsx2/issues/280)

* Changes to various internal pugixml functions, to improve handling of XML strings. [279](https://github.com/JanMarvin/openxlsx2/issues/279)

* Provide internal helper `xml_rm_child()` to remove children of XML strings. [273](https://github.com/JanMarvin/openxlsx2/issues/273)

* Fixes a bug in `update_cell()` that slowed down writing on worksheets with data. In addition, this function has been cleaned up and further improved. It is no longer exported, as users only need to use `wb_add_data()` or `write_data()`, each of which calls `update_cell()` under the hood. [275](https://github.com/JanMarvin/openxlsx2/pull/275) [276](https://github.com/JanMarvin/openxlsx2/pull/276)

* Various (mostly internal) changes to `conditional_formatting`. Created `style_mgr` integration for `dxf` (cf-styles) and cleaned up internal code. The syntax has changed slightly, see [conditional formatting vignette](https://janmarvin.github.io/openxlsx2/articles/conditional-formatting.html) for reference. Add `whitespace` argument to `read_xml()`. [268](https://github.com/JanMarvin/openxlsx2/issues/268)

## Breaking changes

* Order of arguments in reading functions `wb_to_df()`, `wb_read()` and `read_xls()` has changed.


***************************************************************************


# openxlsx2 0.2.1

## New features

* Data adding functions now ship a `dims` argument that can be used to determine the `startCol` and `startRow` for any `x` object that is added to the worksheet. Works with `add_data()`, `add_data_table()`, `add_formula()` and their underlying `write_` functions as well as with the wrappers.

* Provide optional `na.strings` argument when writing data to sheets. It can be used to add a custom character string when writing numeric data.

* Improve writing `NA`, `NaN`, and `-Inf`/`Inf`. `NA` will be converted to `#N/A`; `NaN` will be converted to `#VALUE!`; `Inf` will be converted to `#NUM!`. The same conversion is not applied when reading from a workbook. [256](https://github.com/JanMarvin/openxlsx2/pull/256)

* Many `wbWorkbook` methods now contain default sheet values of `current_sheet()` or `next_sheet()` (e.g., `$add_worksheet(sheet = next_sheet())`, `$write_data(sheet = curret_sheet()`).  These internal waiver functions allow the `wbWorkbook` object to use default expectations for what sheet to interact with.  This allows the easier workflow of `wb$add_worksheet()$add_data(x = data.frame())` where `$add_worksheet()` knows to add a new worksheet (with a default name), sets that new worksheet to the current worksheet, and then `$add_data()` picks up the new sheet and places the data there. [165](https://github.com/JanMarvin/openxlsx2/issues/165), [179](https://github.com/JanMarvin/openxlsx2/pull/179)

* New functions `wb_add_cell_style()` and `wb$add_cell_style()` to simplify the creation of cell styles for cells on the sheet. This provides a fast way to create cell styles for regions on the worksheet. The cells for which the cell format is to be created must already exist on the worksheet. If the cells already contain a cell format, it will be preserved, except for the updated cell format entries, which will always be created. The function is applied to a continuous cell of the worksheet.
[230](https://github.com/JanMarvin/openxlsx2/pull/230)

* New functions `wb_add_numfmt()` and `wb$add_numfmt()` to simplify the creation of number formats for cells on the sheet. This provides a fast way to create number formats for regions on the worksheet. The cells for which the number format is to be created must already exist on the worksheet. If the cells already contain a cell style, it will be preserved, except for the number format, which will always be created. The function is applied to a continuous cell of the worksheet.
[229](https://github.com/JanMarvin/openxlsx2/pull/229)

* New functions `wb_add_font()` and `wb$add_font()` to simplify the creation of fonts for cells on the sheet. This provides a fast way to create fonts for regions on the worksheet. The cells for which the font is to be created must already exist on the worksheet. If the cells already contain a cell style, it will be preserved, except for the font, which will always be created. The function is applied to a continuous cell of the worksheet.
[228](https://github.com/JanMarvin/openxlsx2/pull/228)

* New functions `wb_add_fill()` and `wb$add_fill()` to simplify the creation of fills for cells on the sheet. This provides a fast way to create color filled regions on the worksheet. The cells for which the fill is to be created must already exist on the worksheet. If the cells already contain a cell style, it will be preserved, except for the filled color, which will always be created. The function is applied to a continuous cell of the worksheet and allows to change the color of every n-th column or row.
[222](https://github.com/JanMarvin/openxlsx2/pull/222)

* New functions `wb_add_border()` and `wb$add_border()` to simplify the creation of borders for cells on the sheet. This is especially useful when creating surrounding borders with different border styles for various cells. The cells for which the border is to be created must already exist on the worksheet. If the cells already contain a cell style, it will be preserved, except for the border, which will always be created. The function is applied to a continuous cell of the worksheet and allows to change the horizontal and vertical internal border grid independently. [220](https://github.com/JanMarvin/openxlsx2/pull/220)

* Enable reading tables with `wb_to_df()`. Tables are handled similar to defined names. [193](https://github.com/JanMarvin/openxlsx2/pull/193)

* Several enhancements have been added for checking and validation worksheet names
[165](https://github.com/JanMarvin/openxlsx2/issues/165)
  * When adding a new worksheet via `wbWorkbook$add_worksheet()` the provided name is checked for illegal characters (see note in **Breaking changes**)
  * `wbWorkbook$get_sheet_names()` (`wb_get_sheet_names()` wrapper) added. These return both the _formatted_ and original sheet names
  * `wbWorkbook$set_sheet_names()` (`wb_set_sheet_names()`) added
    * these make `names.wbWorkbook()` and `names<-.wbWorkbook()` deprecated
    * `wbWorkbook$setSheetName()` deprecated
  * `clean_worksheet_names()` added to support removing characters that are not allowed in worksheet names

## Fixes

* Various fixes to enable handling of non unicode R environments [243](https://github.com/JanMarvin/openxlsx2/issues/243)

* Fix an issue with broken pageSetup reference causing corrupt excel files [216](https://github.com/JanMarvin/openxlsx2/issues/216)

* Fix reading and writing comments from workbooks that already provide comments [209](https://github.com/JanMarvin/openxlsx2/pull/209)

* Fix an issue with broken xml in Excels vml files and enable opening xlsm files with `wb$open()` [202](https://github.com/JanMarvin/openxlsx2/pull/202)

* Fix reading and writing on non UTF-8 systems [198](https://github.com/JanMarvin/openxlsx2/pull/198) [199](https://github.com/JanMarvin/openxlsx2/pull/199) [207](https://github.com/JanMarvin/openxlsx2/pull/207)

* Instruct parser to import nodes with whitespaces. This fixes a complaint in spreadsheet software. [189](https://github.com/JanMarvin/openxlsx2/pull/189)

* Fix reading file without row attribute. [187](https://github.com/JanMarvin/openxlsx2/pull/187) [190](https://github.com/JanMarvin/openxlsx2/pull/190)

* Remove reference to `printerSettings.bin` when loading. This binary blob is not included and the reference caused file corruption warnings. [185](https://github.com/JanMarvin/openxlsx2/pull/185)

* Fix loading and writing xlsx files with with `workbook$extLst`. Previously if the loaded sheet contains a slicer, a second `extLst` was added which confused spreadsheet software. Now both are combined into a single node.

* Fix writing xlsx file with multiple entries of conditional formatting type databar on any sheet. [174](https://github.com/JanMarvin/openxlsx2/pull/174)

* Cell fields cm, ph and vm are now implemented for reading and writing. This is the first step to handle functions that use metadata. [173](https://github.com/JanMarvin/openxlsx2/pull/173)

* `wbWorkbook`: `$open()` no longer overwrites the `$path` field to the temporary file [171](https://github.com/JanMarvin/openxlsx2/pull/171)
* `xl_open()` works (better) on Windows [170](https://github.com/JanMarvin/openxlsx2/issues/170)

## Breaking changes

* When writing to existing workbooks, the default value for `removeCellStyle` is now `FALSE`. Therefore if a cell contains a style, it is attempted to replace the value, but not the style of the cell itself.

* `wb_conditional_formatting()` is deprecated in favor of `wb_add_conditional_formatting()` and `wbWorkbook$add_conditional_formatting()`.
  * `type` must now match exactly one of: `"expression"`, `"colorScale"`, `"dataBar"`, `"duplicatedValues"`, `"containsText"`, `"notContainsText"`, `"beginsWith"`, `"endsWith"`, `"between"`, `"topN"`, `"bottomN"`

* Assigning a new worksheet with an illegal character now prompts an error [165](https://github.com/JanMarvin/openxlsx2/issues/165).  See `?clean_worksheet_name` for an easy method of replacing bad characters.

* `openxlsx2Coerce()` (which was called on `x` objects when adding data to a workbook) has been removed.  Users can no longer pass some arbitrary objects and will need to format these objects appropriately or rely on `as.data.frame` methods  [167](https://github.com/JanMarvin/openxlsx2/issues/167)
* `xl_open(file = )` is no longer valid and will throw a warning; first argument has been changes to `x` to highlight that `xl_open()` can be called on a file path or a `wbWorkbook` object [171](https://github.com/JanMarvin/openxlsx2/pull/171)

## Internal changes

* Remove `wb$createFontNode()` which was never used.

* Switch to modern xlsx template, when creating workbooks. Imported workbooks will use the imported template

* Rewrite `wb$tables` to use a data frame approach. This simplifies the code a bit and makes it easier to implement more upcoming changes [191](https://github.com/JanMarvin/openxlsx2/pull/191)

* Update of internal pugixml library

* The two functions `write_data()` and `write_datatable()` now use the same internal function `write_data_table()` to add data to the sheet. This simplifies the code and ensures that both functions are tested. In the same pull request, the documentation has been updated and the `stack=` option, which was not present before, has been removed [175](https://github.com/JanMarvin/openxlsx2/pull/175)

* `wbWorkbook$validate_sheet()` added as an object methods

* private `wbWorkbook` field `original_sheet_names` added to track the original names passed to sheets
* private `$get_sheet()` removed in favor of more explicit
* private `wbWorkbook` methods additions:
  * `$get_sheet_id_max()`, `$get_sheet_index()` for getting ids
  * `$get_sheet_name()` for getting a sheet name
  * `$set_single_sheet_name()` for setting sheet names
  * `$pappend()` general private appending
  * `$validate_new_sheet()` for checking new sheet names
  * `$append_workbook_field()` for `self$workbook[[field]]`
  * `$append_sheet_rels()` for `self$worksheet_rels[[sheet]]`
  * `$get_worksheet()` to replace `$ws()`


***************************************************************************

# openxlsx2 0.2.0

* Added a `NEWS.md` file to track changes to the package.
* First public release
