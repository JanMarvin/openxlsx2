# openxlsx2 (development version)

## New features

* Several enhancements have been added for checking and validation worksheet names
[165](https://github.com/JanMarvin/openxlsx2/issues/165)
  * When adding a new worksheet via `wbWorkbook$add_worksheet()` the provided name is checked for illegal characters (see note in **Breaking changes**)
  * `wbWorkbook$get_sheet_names()` (`wb_get_sheet_names()` wrapper) added. These return both the _formatted_ and original sheet names
  * `wbWorkbook$set_sheet_names()` (`wb_set_sheet_names()`) added
    * these make `names.wbWorkbook()` and `names<-.wbWorkbook()` deprecated
    * `wbWorkbook$setSheetName()` deprecated
  * `clean_worksheet_names()` added to support removing characters that are not allowed in worksheet names

## Fixes

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

* Assigning a new worksheet with an illegal character now prompts an error [165](https://github.com/JanMarvin/openxlsx2/issues/165).  See `?clean_worksheet_name` for an easy method of replacing bad characters.

* `openxlsx2Coerce()` (which was called on `x` objects when adding data to a workbook) has been removed.  Users can no longer pass some arbitrary objects and will need to format these objects appropriately or rely on `as.data.frame` methods  [167](https://github.com/JanMarvin/openxlsx2/issues/167)
* `xl_open(file = )` is no longer valid and will throw a warning; first argument has been changes to `x` to highlight that `xl_open()` can be called on a file path or a `wbWorkbook` object [171](https://github.com/JanMarvin/openxlsx2/pull/171)

## Internal changes

* Rewrite `wb$tables` to use a data frame approach. This simplifies the code a bit and makes it easier to implement more upcoming changes [191](https://github.com/JanMarvin/openxlsx2/pull/191)

* Update of internal pugixml library

* The two functions `write_data()` and `write_datatable()` now use the same internal function `write_data_table()` to add data to the sheet. This simplifies the code and ensures that both functions are tested. In the same pull request, the documentation has been updated and the `stack=` option, which was not present before, has been removed [175](https://github.com/JanMarvin/openxlsx2/pull/175)

* `wbWorkbook$validate_sheet()` added as an object methods

* private `wbWorkbook` field `original_sheet_names` added to track the original names passed to sheets

* private `wbWorkbook` methods added:
  * `$get_sheet_id_max()`, `$get_sheet_index()` for getting ids
  * `$get_sheet_name()` for getting a sheet name
  * `$set_single_sheet_name()` for setting sheet names
  * `$pappend()` general private appending
  * `$validate_new_sheet()` for checking new sheet names
  * `$append_workbook_field()` for `self$workbook[[field]]`
  * `$append_sheet_rels()` for `self$worksheet_rels[[sheet]]`
  * `$get_worksheet()` to replace `$ws()`

# openxlsx2 0.2.0

* Added a `NEWS.md` file to track changes to the package.
* First public release
