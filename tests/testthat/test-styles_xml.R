test_that("xf", {

  xlsxFile <- system.file("extdata", "loadExample.xlsx", package = "openxlsx2")
  wb <- loadWorkbook(xlsxFile)

  xlsxFile <- system.file("extdata", "readTest.xlsx", package = "openxlsx2")
  wb1 <- loadWorkbook(xlsxFile)

  # not ordered
  exp <- c("applyAlignment", "applyBorder", "applyFill", "applyFont",
           "applyNumberFormat", "applyProtection", "borderId", "fillId",
           "fontId", "numFmtId", "pivotButton", "quotePrefix", "xfId",
           "horizontal", "indent", "justifyLastLine", "readingOrder",
           "relativeIndent", "shrinkToFit", "textRotation", "vertical",
           "wrapText", "extLst", "hidden", "locked")

  input <- wb$styles_mgr$styles$cellXfs
  got <- openxlsx2:::read_xf(as_xml(input))
  expect_equal(sort(exp), sort(names(got)))

  exp <- c(74, 25)
  expect_equal(exp, dim(got))

  expect_equal(input,
               openxlsx2:::write_xf(got))

  expect_warning(
    got <- openxlsx2:::read_xf(read_xml('<xf numFmtId="0" foo="0"/>')),
    "\"foo\": not found in xf name table"
  )

})

