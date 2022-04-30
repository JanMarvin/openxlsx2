test_that("Deleting worksheets", {
  tempFile <- temp_xlsx()
  genWS <- function(wb, sheetName) {
    wb$add_worksheet(sheetName)
    write_datatable(wb, sheetName, data.frame("X" = sprintf("This is sheet: %s", sheetName)), colNames = FALSE)
  }

  wb <- wb_workbook()
  genWS(wb, "Sheet 1")
  genWS(wb, "Sheet 2")
  genWS(wb, "Sheet 3")
  expect_equal(names(wb), c("Sheet 1", "Sheet 2", "Sheet 3"))

  wb$remove_worksheet(sheet = 1)
  expect_equal(names(wb), c("Sheet 2", "Sheet 3"))

  wb$remove_worksheet(sheet = 1)
  expect_equal(names(wb), c("Sheet 3"))

  ## add to end
  genWS(wb, "Sheet 1")
  genWS(wb, "Sheet 2")
  expect_equal(names(wb), c("Sheet 3", "Sheet 1", "Sheet 2"))

  wb_save(wb, tempFile, overwrite = TRUE)

  ## re-load & re-order worksheets

  wb <- wb_load(tempFile)
  expect_equal(names(wb), c("Sheet 3", "Sheet 1", "Sheet 2"))

  write_data(wb, sheet = "Sheet 2", x = iris[1:10, 1:4], startRow = 5)
  test <- read_xlsx(wb, "Sheet 2", startRow = 5)
  rownames(test) <- seq_len(nrow(test))
  attr(test, "tt") <- NULL
  attr(test, "types") <- NULL
  expect_equal(iris[1:10, 1:4], test)

  write_data(wb, sheet = 1, x = iris[1:20, 1:4], startRow = 5)
  test <- read_xlsx(wb, "Sheet 3", startRow = 5)
  rownames(test) <- seq_len(nrow(test))
  attr(test, "tt") <- NULL
  attr(test, "types") <- NULL
  expect_equal(iris[1:20, 1:4], test)

  wb$remove_worksheet(sheet = 1)
  expect_equal("This is sheet: Sheet 1", read_xlsx(wb, 1, startRow = 1)[[1]])

  wb$remove_worksheet(sheet = 2)
  expect_equal("This is sheet: Sheet 1", read_xlsx(wb, 1, startRow = 1)[[1]])

  wb$remove_worksheet(sheet = 1)
  expect_equal(names(wb), character(0))

  unlink(tempFile, recursive = TRUE, force = TRUE)
  rm(wb)
})
