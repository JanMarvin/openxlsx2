test_that("xf", {

  xlsxFile <- system.file("extdata", "loadExample.xlsx", package = "openxlsx2")
  wb <- wb_load(xlsxFile)

  xlsxFile <- system.file("extdata", "readTest.xlsx", package = "openxlsx2")
  wb1 <- wb_load(xlsxFile)

  # likely not ordered entirely correct
  exp <- c(
    "numFmtId", "fontId", "fillId", "borderId", "xfId",
    "applyFill", "applyBorder", "applyFont", "applyAlignment",
    "applyNumberFormat", "applyProtection",
    "pivotButton", "quotePrefix",
    "horizontal", "indent", "justifyLastLine", "readingOrder",
    "relativeIndent", "shrinkToFit", "textRotation", "vertical",
    "wrapText", "extLst", "hidden", "locked"
  )

  input <- wb$styles_mgr$styles$cellXfs
  got <- read_xf(as_xml(input))
  expect_equal(sort(exp), sort(names(got)))

  exp_dim <- c(74, 25)
  expect_equal(exp_dim, dim(got))

  expect_equal(input,
               write_xf(got[exp]))

  expect_warning(
    got <- read_xf(read_xml('<xf numFmtId="0" foo="0"/>')),
    "foo: not found in xf name table"
  )

})

test_that("border", {

  xlsxFile <- system.file("extdata", "loadExample.xlsx", package = "openxlsx2")
  wb <- wb_load(xlsxFile)

  # likely not ordered entirely correct
  exp <- c("left", "right", "top", "bottom",
           "diagonal", "diagonalDown", "diagonalUp", "end",
           "horizontal",  "outline",  "start", "vertical"
  )

  input <- wb$styles_mgr$styles$borders
  got <- read_border(as_xml(input))
  expect_equal(sort(exp), sort(names(got)))

  exp_dim <- c(58, 12)
  expect_equal(exp_dim, dim(got))

  expect_equal(input,
               write_border(got[exp]))

  expect_warning(
    got <- read_border(read_xml('<border foo="0"/>')),
    "foo: not found in border name table"
  )

})

test_that("cellStyle", {

  xlsxFile <- system.file("extdata", "loadExample.xlsx", package = "openxlsx2")
  wb <- wb_load(xlsxFile)

  # likely not ordered entirely correct
  exp <- c("name", "xfId", "builtinId", "customBuiltin", "extLst", "hidden", "iLevel")

  input <- wb$styles_mgr$styles$cellStyles
  got <- read_cellStyle(as_xml(input))
  expect_equal(sort(exp), sort(names(got)))

  exp_dim <- c(2, 7)
  expect_equal(exp_dim, dim(got))

  expect_equal(input,
               write_cellStyle(got[exp]))

  expect_warning(
    got <- read_cellStyle(read_xml('<cellStyle foo="0"/>')),
    "foo: not found in cellstyle name table"
  )

})

test_that("tableStyle", {

  # tableStyle elements exist only with custom tableStyles
  xlsxFile <- system.file("extdata", "tableStyles.xlsx", package = "openxlsx2")
  wb <- wb_load(xlsxFile)

  # likely not ordered entirely correct
  exp <- c("name", "pivot", "count", "xr9:uid", "table", "tableStyleElement")

  input <- wb$styles_mgr$styles$tableStyles
  got <- read_tableStyle(as_xml(input))
  expect_equal(sort(exp), sort(names(got)))

  exp_dim <- c(1, 6)
  expect_equal(exp_dim, dim(got))

  expect_equal(input,
               write_tableStyle(got[exp]))

  expect_warning(
    got <- read_tableStyle(read_xml('<tableStyle foo="0"/>')),
    "foo: not found in tablestyle name table"
  )

})

test_that("dxf", {

  xlsxFile <- system.file("extdata", "loadExample.xlsx", package = "openxlsx2")
  wb <- wb_load(xlsxFile)

  # likely not ordered entirely correct
  exp <- c("font", "fill", "alignment", "border", "extLst",   "numFmt",
           "protection")

  input <- wb$styles_mgr$styles$dxf
  got <- read_dxf(as_xml(input))
  expect_equal(sort(exp), sort(names(got)))

  exp_dim <- c(2, 7)
  expect_equal(exp_dim, dim(got))

  expect_equal(input,
               write_dxf(got[exp]))

  expect_warning(
    got <- read_dxf(read_xml('<dxf><foo bar="0"/></dxf>')),
    "foo: not found in dxf name table"
  )

})

test_that("colors", {

  xlsxFile <- system.file("extdata", "oxlsx2_sheet.xlsx", package = "openxlsx2")
  wb <- wb_load(xlsxFile)

  # likely not ordered entirely correct
  exp <- c("mruColors", "indexedColors")

  input <- wb$styles_mgr$styles$colors
  got <- read_colors(as_xml(input))
  expect_equal(sort(exp), sort(names(got)))

  exp_dim <- c(1, 2)
  expect_equal(exp_dim, dim(got))

  expect_equal(input,
               write_colors(got[exp]))

  expect_warning(
    got <- read_colors(read_xml('<colors><foo bar="0"/></colors>')),
    "foo: not found in color name table"
  )

})
