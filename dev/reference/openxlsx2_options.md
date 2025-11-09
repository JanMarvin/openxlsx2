# Options consulted by openxlsx2

The openxlsx2 package allows the user to set global options to simplify
formatting:

If the built-in defaults don't suit you, set one or more of these
options. Typically, this is done in the `.Rprofile` startup file

- `options("openxlsx2.borderColor" = "black")`

- `options("openxlsx2.borderStyle" = "thin")`

- `options("openxlsx2.dateFormat" = "mm/dd/yyyy")`

- `options("openxlsx2.datetimeFormat" = "yyyy-mm-dd hh:mm:ss")`

- `options("openxlsx2.maxWidth" = NULL)` (Maximum width allowed in OOXML
  is 250)

- `options("openxlsx2.minWidth" = NULL)`

- `options("openxlsx2.numFmt" = NULL)`

- `options("openxlsx2.paperSize" = 9)` corresponds to a A4 paper size

- `options("openxlsx2.orientation" = "portrait")` page orientation

- `options("openxlsx2.sheet.default_name" = "Sheet")`

- `options("openxlsx2.rightToLeft" = NULL)`

- `options("openxlsx2.soon_deprecated" = FALSE)` Set to `TRUE` if you
  want a warning if using some functions deprecated recently in
  openxlsx2

- `options("openxlsx2.creator")` A default name for the creator of new
  `wbWorkbook` object with
  [`wb_workbook()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_workbook.md)
  or new comments with
  [`wb_add_comment()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_comment.md)

- `options("openxlsx2.thread_id")` the default person id when adding a
  threaded comment to a cell with
  [`wb_add_thread()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_thread.md)

- `options("openxlsx2.accountingFormat" = 4)`

- `options("openxlsx2.currencyFormat" = 4)`

- `options("openxlsx2.commaFormat" = 3)`

- `options("openxlsx2.percentageFormat" = 10)`

- `options("openxlsx2.scientificFormat" = 48)`

- `options("openxlsx2.string_nums" = TRUE)` numerics in character
  columns will be converted. `"1"` will be written as `1`

- `options("openxlsx2.na.strings" = "#N/A")` consulted by
  [`write_xlsx()`](https://janmarvin.github.io/openxlsx2/dev/reference/write_xlsx.md),
  [`wb_add_data()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data.md)
  and
  [`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data_table.md).

- `options("openxlsx2.compression_level" = 6)` compression level for the
  output file. Increasing compression and time consumed from 1-9.
