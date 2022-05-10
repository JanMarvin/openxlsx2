# openxlsx2 (development version)

* `openxlsx2Coerce()` (which was called on `x` objects when adding data to a workbook) has been removed.  Users can no longer pass some arbitrary objects and will need to format these objects appropriately or rely on `as.data.frame` methods  [167](https://github.com/JanMarvin/openxlsx2/issues/167)

# openxlsx2 0.2.0

* Added a `NEWS.md` file to track changes to the package.
* First public release
