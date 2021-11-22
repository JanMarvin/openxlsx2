
test_that("genBaseContent_Type() works", {
  x <- genBaseContent_Type()
  expect_length(x, 8L)
  expect_type(x, "character")
})

test_that("genBaseShapeVML() works", {
  x <- genBaseShapeVML("a", "b")
  expect_length(x, 1L)
  expect_type(x, "character")

  x <- genBaseShapeVML("visible", "b")
  expect_length(x, 1L)
  expect_type(x, "character")
})

test_that("genClientData() works", {
  x <- genClientData(1, 1, TRUE, 1, 1)
  expect_length(x, 1L)
  expect_type(x, "character)")

  x <- genClientData(1, 1, FALE, 1, 1)
  expect_length(x, 1L)
  expect_type(x, "character)")
})

test_that("genBaseCore() works", {
  x <- genBaseCore()
  expect_length(x, 1L)
  expect_type(x, "character")
})

test_that("genBaseWorkbook.xml.rels() works", {
  x <- genBaseWorkbook.xml.rels()
  expect_length(x, 3L)
  expect_type(x, "character")
})

test_that("genBaseWorkbook() works", {
  x <- genBaseWorkbook()
  expect_length(x, 10L)
  expect_type(x, "list")
  expect_null(x$workbookProtection)
  expect_null(x$sheets)
  expect_null(x$externalReferences)
  expect_null(x$definedNames)
  expect_null(x$calcPr)
  expect_null(x$alternateContent)
  expect_null(x$pivotCaches)
  expect_null(x$extLst)
  expect_type(x$workbookPr, "character")
  expect_length(x$workbookPr, 1L)
  expect_type(x$bookViews, "character")
  expect_length(x$bookViews, 1L)
})

test_that("genBaseSheetRels() works", {
  x <- genBaseSheetRels(1L)
  expect_length(x, 2L)
  expect_type(x, "character")
})

test_that("genBaseStyleSheet() works", {
  x <- genBaseStyleSheet()
  expect_length(x, 11L)
  expect_type(x, "class")
})

test_that("genBasePic() works", {
  x <- genBasePic(1L)
  expect_length(x, 1L)
  expect_type(x, "character")
})

test_that("genBaseTheme() works", {
  x <- genBaseTheme()
  expect_length(x, 1L)
  expect_type(x, "character")
})

test_that("genPrinterSettings() works", {
  x <- genPrinterSettings()
  expect_length(x, 1L)
  expect_type(x, "character")
})

test_that("gen_databar_extlst() works", {
  x <- gen_databar_extlst(
    guid = 1,
    sqref = 1,
    posColour = 1,
    negColour = 1,
    values = list(1, 2),
    border = 1,
    gradient = 1
  )
  expect_length(x, 1L)
  expect_type(x, "character")
})

test_that("contentTypePivotXML() works", {
  x <- contentTypePivotXML(1)
  expect_length(x, 3L)
  expect_type(x, "character")
})

test_that("contentTypeSlicerCacheXML() works", {
  x <- genPrinterSettings()
  expect_length(x, 1L)
  expect_type(x, "character")
})

test_that("genBaseSlicerXML() works", {
  x <- genBaseSlicerXML()
  expect_length(x, 1L)
  expect_type(x, "character")
})

test_that("genSlicerCachesExtLst() works", {
  x <- genSlicerCachesExtLst(1)
  expect_length(x, 1L)
  expect_type(x, "character")
})
