# openxlsx2 (development version)

## Fixes

* fix writing xlsx file with multiple entries of conditional formatting type databar on any sheet. [174](https://github.com/JanMarvin/openxlsx2/pull/174) 

* Cell fields cm, ph and vm are now implemented for reading and writing. This is the first step to handle functions that use metadata. [173](https://github.com/JanMarvin/openxlsx2/pull/173)

* `wbWorkbook`: `$open()` no longer overwrites the `$path` field to the temporary file [171](https://github.com/JanMarvin/openxlsx2/pull/171)
* `xl_open()` works (better) on Windows [170](https://github.com/JanMarvin/openxlsx2/issues/170)

## Breaking changes

* `openxlsx2Coerce()` (which was called on `x` objects when adding data to a workbook) has been removed.  Users can no longer pass some arbitrary objects and will need to format these objects appropriately or rely on `as.data.frame` methods  [167](https://github.com/JanMarvin/openxlsx2/issues/167)
* `xl_open(file = )` is no longer valid and will throw a warning; first argument has been changes to `x` to highlight that `xl_open()` can be called on a file path or a `wbWorkbook` object [171](https://github.com/JanMarvin/openxlsx2/pull/171)

# openxlsx2 0.2.0

* Added a `NEWS.md` file to track changes to the package.
* First public release
