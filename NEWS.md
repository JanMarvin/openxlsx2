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

* New function `wb_colour()` to ease working with colour vectors used in `openxlsx2` styles. [292](https://github.com/JanMarvin/openxlsx2/issues/292)

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
