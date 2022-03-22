
test_that("Maintaining Named Regions on Load", {
  ## create named regions
  wb <- wb_workbook()
  wb$addWorksheet("Sheet 1")
  wb$addWorksheet("Sheet 2")

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
  wb_save(wb, out_file, overwrite = TRUE)

  expect_equal(object = getNamedRegions(wb), expected = getNamedRegions(out_file))

  df1 <- read.xlsx(wb, namedRegion = "iris")
  df2 <- read.xlsx(out_file, namedRegion = "iris")
  expect_equal(df1, df2)

  df1 <- read.xlsx(wb, namedRegion = "region1", skipEmptyCols = FALSE)
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
  expect_equal(object = class(cell_f), expected = "data.frame")
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
  expect_equal(object = class(cell2_f), expected = "data.frame")
  expect_equal(object = NCOL(cell2_f), expected = 1)
  expect_equal(object = NROW(cell2_f), expected = 1)
})


test_that("Load names from an Excel file with funky non-region names", {
  filename <- system.file("extdata", "namedRegions2.xlsx", package = "openxlsx2")
  wb <- loadWorkbook(filename)
  dn <- getNamedRegions(wb)

  expect_equal(
    head(dn$name, 5),
    c("barref", "barref", "fooref", "fooref", "IQ_CH")
  )
  expect_equal(
    dn$sheets,
    c(
      "Sheet with space", "Sheet1", "Sheet with space", "Sheet1",
      rep("", 26)
    )
  )
  expect_equal(dn$coords, c("B4", "B4", "B3", "B3", rep("", 26)))

  dn2 <- getNamedRegions(filename)
  expect_equal(dn, dn2)
})


test_that("Missing rows in named regions", {
  temp_file <- temp_xlsx()

  wb <- wb_workbook()
  wb$addWorksheet("Sheet 1")

  ## create region
  writeData(wb, sheet = 1, x = iris[1:11, ], startCol = 1, startRow = 1)
  deleteData(wb, sheet = 1, cols = 1:2, rows = c(6, 6))

  createNamedRegion(
    wb = wb,
    sheet = 1,
    name = "iris",
    rows = 1:(5 + 1),
    cols = 1:2
  )
  createNamedRegion(
    wb = wb,
    sheet = 1,
    name = "iris2",
    rows = 1:(5 + 2),
    cols = 1:2
  )

  ## iris region is rows 1:6 & cols 1:2
  ## iris2 region is rows 1:7 & cols 1:2

  ## row 6 columns 1 & 2 are blank
  expect_equal(getNamedRegions(wb)$name, c("iris", "iris2"))
  expect_equal(getNamedRegions(wb)$sheet, c("Sheet 1", "Sheet 1"))
  expect_equal(getNamedRegions(wb)$coords, c("A1:B6", "A1:B7"))

  ######################################################################## from Workbook

  ## Skip empty rows
  x <- read.xlsx(xlsxFile = wb, namedRegion = "iris", colNames = TRUE, skipEmptyRows = TRUE)
  expect_equal(dim(x), c(4, 2))

  x <- read.xlsx(xlsxFile = wb, namedRegion = "iris2", colNames = TRUE, skipEmptyRows = TRUE)
  expect_equal(dim(x), c(5, 2))


  ## Keep empty rows
  x <- read.xlsx(xlsxFile = wb, namedRegion = "iris", colNames = TRUE, skipEmptyRows = FALSE)
  expect_equal(dim(x), c(5, 2))

  x <- read.xlsx(xlsxFile = wb, namedRegion = "iris2", colNames = TRUE, skipEmptyRows = FALSE)
  expect_equal(dim(x), c(6, 2))



  ######################################################################## from file
  wb_save(wb, temp_file)

  ## Skip empty rows
  x <- read.xlsx(xlsxFile = temp_file, namedRegion = "iris", colNames = TRUE, skipEmptyRows = TRUE)
  expect_equal(dim(x), c(4, 2))

  x <- read.xlsx(xlsxFile = temp_file, namedRegion = "iris2", colNames = TRUE, skipEmptyRows = TRUE)
  expect_equal(dim(x), c(5, 2))


  ## Keep empty rows
  x <- read.xlsx(xlsxFile = temp_file, namedRegion = "iris", colNames = TRUE, skipEmptyRows = FALSE)
  expect_equal(dim(x), c(5, 2))

  x <- read.xlsx(xlsxFile = temp_file, namedRegion = "iris2", colNames = TRUE, skipEmptyRows = FALSE)
  expect_equal(dim(x), c(6, 2))

  unlink(temp_file)
})


test_that("Missing columns in named regions", {
  temp_file <- temp_xlsx()

  wb <- wb_workbook()
  wb$addWorksheet("Sheet 1")

  ## create region
  writeData(wb, sheet = 1, x = iris[1:11, ], startCol = 1, startRow = 1)
  deleteData(wb, sheet = 1, cols = 2, rows = 1:12, gridExpand = TRUE)

  createNamedRegion(
    wb = wb,
    sheet = 1,
    name = "iris",
    rows = 1:5,
    cols = 1:2
  )

  createNamedRegion(
    wb = wb,
    sheet = 1,
    name = "iris2",
    rows = 1:5,
    cols = 1:3
  )

  ## iris region is rows 1:5 & cols 1:2
  ## iris2 region is rows 1:5 & cols 1:3

  ## row 6 columns 1 & 2 are blank
  expect_equal(getNamedRegions(wb)$name, c("iris", "iris2"), ignore_attr = TRUE)
  expect_equal(getNamedRegions(wb)$sheet, c("Sheet 1", "Sheet 1"))
  expect_equal(getNamedRegions(wb)$coords, c("A1:B5", "A1:C5"))

  ######################################################################## from Workbook

  ## Skip empty cols
  x <- read.xlsx(xlsxFile = wb, namedRegion = "iris", colNames = TRUE, skipEmptyCols = TRUE)
  expect_equal(dim(x), c(4, 1))

  x <- read.xlsx(xlsxFile = wb, namedRegion = "iris2", colNames = TRUE, skipEmptyCols = TRUE)
  expect_equal(dim(x), c(4, 2))


  ## Keep empty cols
  x <- read.xlsx(xlsxFile = wb, namedRegion = "iris", colNames = TRUE, skipEmptyCols = FALSE)
  expect_equal(dim(x), c(4, 2))

  x <- read.xlsx(xlsxFile = wb, namedRegion = "iris2", colNames = TRUE, skipEmptyCols = FALSE)
  expect_equal(dim(x), c(4, 3))



  ######################################################################## from file
  wb_save(wb, temp_file)

  ## Skip empty cols
  x <- read.xlsx(xlsxFile = temp_file, namedRegion = "iris", colNames = TRUE, skipEmptyCols = TRUE)
  expect_equal(dim(x), c(4, 1))

  x <- read.xlsx(xlsxFile = temp_file, namedRegion = "iris2", colNames = TRUE, skipEmptyCols = TRUE)
  expect_equal(dim(x), c(4, 2))


  ## Keep empty cols
  x <- read.xlsx(xlsxFile = temp_file, namedRegion = "iris", colNames = TRUE, skipEmptyCols = FALSE)
  expect_equal(dim(x), c(4, 2))

  x <- read.xlsx(xlsxFile = temp_file, namedRegion = "iris2", colNames = TRUE, skipEmptyCols = FALSE)
  expect_equal(dim(x), c(4, 3))

  unlink(temp_file)
})


test_that("Matching Substrings breaks reading named regions", {
  temp_file <- temp_xlsx()

  wb <- wb_workbook()
  wb$addWorksheet("table")
  wb$addWorksheet("table2")

  t1 <- head(iris)
  t1$Species <- as.character(t1$Species)
  t2 <- head(mtcars)

  writeData(wb, sheet = "table", x = t1, name = "t", startCol = 3, startRow = 12)
  writeData(wb, sheet = "table2", x = t2, name = "t2", startCol = 5, startRow = 24, rowNames = TRUE)

  writeData(wb, sheet = "table", x = head(t1, 3), name = "t1", startCol = 9, startRow = 3)
  writeData(wb, sheet = "table2", x = head(t2, 3), name = "t22", startCol = 15, startRow = 12, rowNames = TRUE)

  wb_save(wb, temp_file)

  r1 <- getNamedRegions(wb)
  expect_equal(r1$sheet, c("table", "table2", "table", "table2"))
  expect_equal(r1$coords, c("C12:G18", "E24:P30", "I3:M6", "O12:Z15"))
  expect_equal(r1$name, c("t", "t2", "t1", "t22"))

  r2 <- getNamedRegions(temp_file)
  expect_equal(r2$sheet, c("table", "table2", "table", "table2"))
  expect_equal(r1$coords, c("C12:G18", "E24:P30", "I3:M6", "O12:Z15"))
  expect_equal(r2$name, c("t", "t2", "t1", "t22"))


  ## read file named region
  expect_equal(t1, read.xlsx(xlsxFile = temp_file, namedRegion = "t"), ignore_attr = TRUE)
  expect_equal(t2, read.xlsx(xlsxFile = temp_file, namedRegion = "t2", rowNames = TRUE), ignore_attr = TRUE)
  expect_equal(head(t1, 3), read.xlsx(xlsxFile = temp_file, namedRegion = "t1"), ignore_attr = TRUE)
  expect_equal(head(t2, 3), read.xlsx(xlsxFile = temp_file, namedRegion = "t22", rowNames = TRUE), ignore_attr = TRUE)

  ## read Workbook named region
  expect_equal(t1, read.xlsx(xlsxFile = wb, namedRegion = "t"), ignore_attr = TRUE)
  expect_equal(t2, read.xlsx(xlsxFile = wb, namedRegion = "t2", rowNames = TRUE), ignore_attr = TRUE)
  expect_equal(head(t1, 3), read.xlsx(xlsxFile = wb, namedRegion = "t1"), ignore_attr = TRUE)
  expect_equal(head(t2, 3), read.xlsx(xlsxFile = wb, namedRegion = "t22", rowNames = TRUE), ignore_attr = TRUE)

  unlink(temp_file)
})


test_that("Read namedRegion from specific sheet", {

  filename <- system.file("extdata", "namedRegions3.xlsx", package = "openxlsx2")

  namedR <- "MyRange"
  sheets <- getSheetNames(filename)

  # read the correct sheets
  expect_equal(data.frame(X1 = "S1A1", X2 = "S1B1", stringsAsFactors = FALSE), read.xlsx(filename, sheet = "Sheet1", namedRegion = namedR, rowNames = FALSE, colNames = FALSE), ignore_attr = TRUE)
  expect_equal(data.frame(X1 = "S2A1", X2 = "S2B1", stringsAsFactors = FALSE), read.xlsx(filename, sheet = which(sheets %in% "Sheet2"), namedRegion = namedR, rowNames = FALSE, colNames = FALSE), ignore_attr = TRUE)
  expect_equal(data.frame(X1 = "S3A1", X2 = "S3B1", stringsAsFactors = FALSE), read.xlsx(filename, sheet = "Sheet3", namedRegion = namedR, rowNames = FALSE, colNames = FALSE), ignore_attr = TRUE)

  # Warning: Workbook has no such named region. (Wrong namedRegion selected.)
  expect_error(read.xlsx(filename, sheet = "Sheet2", namedRegion = "MyRage", rowNames = FALSE, colNames = FALSE))

  # Warning: Workbook has no such named region on this sheet. (Correct namedRegion, but wrong sheet selected.)
  expect_error(read.xlsx(filename, sheet = "Sheet4", namedRegion = namedR, rowNames = FALSE, colNames = FALSE))
})


test_that("Overwrite and delete named regions", {
  temp_file <- temp_xlsx()

  wb <- wb_workbook()
  wb$addWorksheet("Sheet 1")

  ## create region
  writeData(wb, sheet = 1, x = iris[1:11, ], startCol = 1,
            startRow = 1, name = "iris")

  init_nr <- getNamedRegions(wb)
  expect_equal(init_nr$coords, "A1:E12")

  # no overwrite
  expect_error({
    writeData(wb, sheet = 1, x = iris[1:11, ], startCol = 1,
              startRow = 1, name = "iris")
  })

  expect_error({
    createNamedRegion(
      wb = wb,
      sheet = 1,
      name = "iris",
      rows = 1:5,
      cols = 1:2
    )
  })

  # overwrite
  createNamedRegion(
    wb = wb,
    sheet = 1,
    name = "iris",
    rows = 1:5,
    cols = 1:2,
    overwrite = TRUE
  )

  # check midification
  modify_nr <- getNamedRegions(wb)
  expect_equal(modify_nr$coords, "A1:B5")
  expect_true("iris" %in% modify_nr)

  # delete name region
  deleteNamedRegion(wb, "iris")
  expect_false("iris" %in% getNamedRegions(wb)$name)

  createNamedRegion(
    wb = wb,
    sheet = 1,
    name = "iris",
    rows = 1:5,
    cols = 1:2
  )

  # removing a worksheet removes the named region as well
  wb <- wb_remove_worksheet(wb, 1)
  expect_true(is.null(getNamedRegions(wb)))

})
