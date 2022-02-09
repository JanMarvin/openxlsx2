
test_that("Maintaining Named Regions on Load", {
  ## create named regions
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  addWorksheet(wb, "Sheet 2")

  ## specify region
  writeData(wb, sheet = 1, x = iris, startCol = 1, startRow = 1)
  createNamedRegion(
    wb = wb,
    sheet = 1,
    name = "iris",
    rows = seq_len(nrow(iris) + 1),
    cols = seq_len(ncol(iris))
  )

  ## using writeData 'name' argument
  writeData(wb, sheet = 1, x = iris, name = "iris2", startCol = 10)

  ## Named region size 1
  writeData(wb, sheet = 2, x = 99, name = "region1", startCol = 3, startRow = 3)

  ## save file for testing
  out_file <- temp_xlsx()
  saveWorkbook(wb, out_file, overwrite = TRUE)

  expect_equal(object = getNamedRegions(wb), expected = getNamedRegions(out_file))

  df1 <- read.xlsx(wb, namedRegion = "iris")
  df2 <- read.xlsx(out_file, namedRegion = "iris")
  expect_equal(df1, df2)

  df1 <- read.xlsx(wb, namedRegion = "region1")
  expect_s3_class(df1, "data.frame")
  expect_equal(nrow(df1), 0)
  expect_equal(ncol(df1), 1)

  df1 <- read.xlsx(wb, namedRegion = "region1", colNames = FALSE)
  expect_s3_class(df1, "data.frame")
  expect_equal(nrow(df1), 1)
  expect_equal(ncol(df1), 1)

  # nonsense
  # df1 is a single value and this single value is now used as rowName
  expect_warning(df1 <- read.xlsx(wb, namedRegion = "region1", rowNames = TRUE))
  expect_s3_class(df1, "data.frame")
  expect_equal(nrow(df1), 0)
  expect_equal(ncol(df1), 0)
})

test_that("Correctly Loading Named Regions Created in Excel", {

  # Load an excel workbook (in the repo, it's located in the /inst folder;
  # when installed on the user's system, it is located in the installation folder
  # of the package)
  filename <- system.file("extdata", "namedRegions.xlsx", package = "openxlsx2")

  # Load this workbook. We will test read.xlsx by passing both the object wb and
  # the filename. Both should produce the same results.
  wb <- loadWorkbook(filename)

  # NamedTable refers to Sheet1!$C$5:$D$8
  table_f <- read.xlsx(filename,
    namedRegion = "NamedTable"
  )
  table_w <- read.xlsx(wb,
    namedRegion = "NamedTable"
  )

  expect_equal(object = table_f, expected = table_w)
  expect_equal(object = class(table_f), expected = "data.frame")
  expect_equal(object = ncol(table_f), expected = 2)
  expect_equal(object = nrow(table_f), expected = 3)

  # NamedCell refers to Sheet1!$C$2
  # This proeduced an error in an earlier version of the pacage when the object
  # wb was passed, but worked correctly when the filename was passed to read.xlsx
  cell_f <- read.xlsx(filename,
    namedRegion = "NamedCell",
    colNames = FALSE,
    rowNames = FALSE
  )

  cell_w <- read.xlsx(wb,
    namedRegion = "NamedCell",
    colNames = FALSE,
    rowNames = FALSE
  )

  expect_equal(object = cell_f, expected = cell_w)
  # expect_equal(object = class(cell_f), expected = "data.frame")
  expect_equal(object = NCOL(cell_f), expected = 1)
  expect_equal(object = NROW(cell_f), expected = 1)

  # NamedCell2 refers to Sheet1!$C$2:$C$2
  cell2_f <- read.xlsx(filename,
    namedRegion = "NamedCell2",
    colNames = FALSE,
    rowNames = FALSE
  )

  cell2_w <- read.xlsx(wb,
    namedRegion = "NamedCell2",
    colNames = FALSE,
    rowNames = FALSE
  )

  expect_equal(object = cell2_f, expected = cell2_w)
  # expect_equal(object = class(cell2_f), expected = "data.frame")
  expect_equal(object = NCOL(cell2_f), expected = 1)
  expect_equal(object = NROW(cell2_f), expected = 1)
})


test_that("Load names from an Excel file with funky non-region names", {
  filename <- system.file("extdata", "namedRegions2.xlsx", package = "openxlsx2")
  wb <- loadWorkbook(filename)
  names <- getNamedRegions(wb)
  sheets <- attr(names, "sheet")
  positions <- attr(names, "position")

  expect_true(length(names) == length(sheets))
  expect_true(length(names) == length(positions))
  expect_equal(
    head(names, 5),
    c("barref", "barref", "fooref", "fooref", "IQ_CH")
  )
  expect_equal(
    sheets,
    c(
      "Sheet with space", "Sheet1", "Sheet with space", "Sheet1",
      rep("", 26)
    )
  )
  expect_equal(positions, c("B4", "B4", "B3", "B3", rep("", 26)))

  names2 <- getNamedRegions(filename)
  expect_equal(names, names2)
})
