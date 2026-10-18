# Files of the Apache POI project (Apache License 2.0), fetched from GitHub on a full run only.
# They are downloaded once into the testfiles folder and not shipped with the package.
poi_file <- function(x) {
  testthat::skip_on_cran()
  if (Sys.getenv("openxlsx2_testthat_fullrun") == "") skip("Only run in a full run")

  test_dir <- testthat::test_path("testfiles")
  if (!dir.exists(test_dir)) dir.create(test_dir, recursive = TRUE)
  fl <- file.path(test_dir, paste0("poi_", x))

  if (!file.exists(fl)) {
    url <- paste0(
      "https://raw.githubusercontent.com/apache/poi/dc577a8cbcdec22ae5547615502d63b2051f35a6/test-data/spreadsheet/", x
    )
    tryCatch(
      suppressWarnings(download.file(url, destfile = fl, quiet = TRUE, mode = "wb")),
      error = function(e) NULL
    )
  }

  if (!file.exists(fl)) skip(paste("Testfile", x, "could not be acquired."))
  fl
}

test_that("xlsb with phonetic runs of several entries (poi 51519) is imported", {
  wb <- wb_load(poi_file("51519.xlsb"))
  expect_equal(wb$get_sheet_names(), c(Sheet1 = "Sheet1", Sheet2 = "Sheet2", Sheet3 = "Sheet3"))

  df <- wb_to_df(wb, sheet = 1, col_names = FALSE)
  expect_equal(dim(df), c(18L, 4L))
  expect_equal(df[1, "A"], "RecordSpec")
  expect_equal(df[2, "B"], "\u8c4a\u7530")
  expect_equal(df[3, "B"], "\u65e5\u672c\u30aa\u30e9\u30af\u30eb")
  expect_equal(df[18, "B"], "\u798f\u89b3\u5efa")
  expect_equal(as.numeric(df[18, "A"]), 17)
})

test_that("xlsb with a cell record in a FRT block (poi 62815) is imported", {
  wb <- wb_load(poi_file("62815.xlsb"))
  expect_equal(wb$get_sheet_names(), c(RkNumber = "RkNumber"))

  df <- wb_to_df(wb, col_names = FALSE)
  expect_equal(nrow(df), 282L)
  expect_equal(df$A[1:4], c(10000000, 1200455, 0.01, 12004.55))
})

test_that("xls with an empty shared string table (poi 57456) is imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  wb <- wb_load(poi_file("57456.xls"))
  expect_equal(wb$get_sheet_names(), c(Sheet2 = "Sheet2"))
  expect_equal(wb$worksheets[[1]]$dimension, "<dimension ref=\"A1:U1047\"/>")
})

test_that("xlsb values, dates and formulas of poi files are imported", {
  wb <- wb_load(poi_file("sample.xlsb"))
  expect_equal(wb$get_sheet_names(), c(Sheet1 = "Sheet1", `rich test` = "rich test", Sheet3 = "Sheet3"))
  df <- wb_to_df(wb, col_names = FALSE)
  expect_equal(df[1, "A"], "Lorem")
  expect_equal(as.numeric(df[10, "B"]), 4995)

  wb <- wb_load(poi_file("date.xlsb"))
  df <- wb_to_df(wb, col_names = FALSE)
  expect_s3_class(df$A, "Date")
  expect_equal(df$A, as.Date("2013-01-12"))

  wb <- wb_load(poi_file("bug66682.xlsb"))
  df <- wb_to_df(wb, col_names = FALSE)
  expect_equal(df[2, "B"], "TRUE")
  expect_equal(df[3, "B"], "FALSE")
  expect_equal(wb_to_df(wb, col_names = FALSE, show_formula = TRUE)[2, "C"], "B2")

  wb <- wb_load(poi_file("testVarious.xlsb"))
  df <- wb_to_df(wb, col_names = FALSE)
  expect_equal(df[1, "B"], "This is a string")
  expect_equal(as.numeric(df[2, "B"]), 13)
  expect_equal(as.numeric(df[4, "B"]), 3.03)
  expect_match(wb$worksheets[[1]]$hyperlinks[[1]], "ref=\"B12\"", fixed = TRUE)
})

test_that("xlsb comments and hyperlinks of poi files are imported", {
  wb <- wb_load(poi_file("comments.xlsb"))
  cmt <- wb_get_comment(wb, sheet = 1)
  expect_equal(cmt$ref, c("A1", "A2", "A3", "A4"))
  expect_equal(cmt$comment[3], "comment top row3 (index2)")

  wb <- wb_load(poi_file("hyperlink.xlsb"))
  expect_equal(wb_to_df(wb, col_names = FALSE)$A, "http://tika.apache.org")
  expect_length(wb$worksheets[[1]]$hyperlinks, 1L)
})

poi_xls <- function(x, sheet = 1) {
  wb <- suppressWarnings(wb_load(poi_file(x)))
  df <- suppressMessages(wb_to_df(wb, sheet = sheet, col_names = FALSE))
  list(wb = wb, df = df)
}

test_that("xls files of poi with merged cells and many strings are imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")

  x <- poi_xls("31749.xls")
  expect_equal(unname(x$wb$get_sheet_names()), "_Proj")
  expect_equal(x$df["1", "A"], "(Template Projection)")
  expect_equal(as.numeric(x$df["3", "C"]), 2004)
  expect_gt(length(x$wb$worksheets[[1]]$mergeCells), 100L)

  x <- poi_xls("51670.xls")
  expect_equal(unname(x$wb$get_sheet_names()), "SBNReport")
  expect_equal(x$df["2", "A"], "Date Printed : 01/08/2011")
  expect_equal(x$df["2", "G"], "Share Repurchase Report")

  x <- poi_xls("bug55505.xls")
  expect_equal(x$df["1", "A"], "Facility Management Software System")
  expect_equal(x$df["2", "A"], "Transaction Totals by Location")

  x <- poi_xls("56450.xls")
  expect_equal(unname(x$wb$get_sheet_names()), "Sheet")
  expect_equal(x$df["1", "A"], "X52 20120")
})

test_that("xls files of poi with double byte characters are imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")

  x <- poi_xls("12843-1.xls")
  nms <- unname(x$wb$get_sheet_names())
  expect_equal(nms[1:3], c("910617", "910701", "910715"))
  expect_equal(nms[7], "POTS-\u8868")
  expect_equal(
    x$df["1", "A"],
    "\u4e2d\u83ef\u96fb\u4fe1\u7db2\u8def\u5bb9\u91cf\u53ca\u4f7f\u7528\u73fe\u6cc1\u96d9\u9031\u5831\u8868"
  )

  x <- poi_xls("44861.xls", sheet = 2)
  expect_equal(unname(x$wb$get_sheet_names()), c("reporting", "data", "Feuil3"))
  expect_equal(x$df["2", "B"], "Fonds JUMP LS")
  expect_equal(x$df["5", "A"], "France")
  expect_equal(as.numeric(x$df["5", "B"]), 0.4108838948063051, tolerance = 1e-12)
  expect_equal(poi_xls("44861.xls")$df["3", "A"], "R\u00e9partition par pays")
})

test_that("xls files of poi with charts, arrays and formulas are imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")

  x <- poi_xls("49581.xls")
  expect_equal(unname(x$wb$get_sheet_names()), c("TestSheet001", "TestSheet002", "TestSheet003"))
  expect_equal(as.numeric(unlist(x$df[1, ])), c(1, 10, 100, -100))
  expect_gt(length(x$wb$charts), 0L)

  x <- poi_xls("testArraysAndTables.xls")
  expect_equal(x$df["1", "A"], "Arrays")
  expect_equal(as.numeric(x$df["2", "C"]), 1.5)
  expect_equal(as.numeric(x$df["2", "D"]), 2.7)

  x <- poi_xls("FormulaEvalTestData.xls")
  expect_equal(length(x$wb$get_sheet_names()), 4L)
  expect_equal(x$df["6", "A"], "DATA SECTION")
  expect_equal(x$df["6", "B"], "INTEGRAL")
  expect_equal(suppressMessages(wb_to_df(x$wb, sheet = 2, col_names = FALSE))["1", "O"], "NPV")

  x <- poi_xls("44840.xls")
  expect_equal(unname(x$wb$get_sheet_names()), c("Sheet1", "Sheet2", "Sheet3"))
  expect_equal(x$df["2", "A"], "Testing")

  x <- poi_xls("45565.xls")
  expect_equal(unname(x$wb$get_sheet_names()), c("Sheet1", "Sheet2", "Sheet3"))
})
