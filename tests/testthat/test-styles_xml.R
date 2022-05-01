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
