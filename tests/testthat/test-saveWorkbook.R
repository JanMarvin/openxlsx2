
test_that("test return values for saveWorkbook", {

  tempFile <- temp_xlsx()
  wb <- createWorkbook()
  addWorksheet(wb, "name")
  expect_equal(tempFile, saveWorkbook(wb, tempFile))

  expect_error( saveWorkbook(wb, tempFile))


  expect_equal(tempFile, saveWorkbook(wb, tempFile, overwrite = TRUE))
  unlink(tempFile, recursive = TRUE, force = TRUE)

})

# regression test for a typo
test_that("regression test for #248", {

  # Basic data frame
  df <- data.frame(number = 1:3, percent = 4:6/100)
  tempFile <- temp_xlsx()

  # no formatting
  expect_silent(write.xlsx(df, tempFile, borders = "columns", overwrite = TRUE))

  # Change column class to percentage
  class(df$percent) <- "percentage"
  expect_silent(write.xlsx(df, tempFile, borders = "columns", overwrite = TRUE))
})


# test for hyperrefs
test_that("creating hyperlinks", {

  # prepare a file
  tempFile <- temp_xlsx()
  wb <- createWorkbook()
  sheet <- "test"
  addWorksheet(wb, sheet)
  img <- "D:/somepath/somepicture.png"

  # warning: col and row provided, but not required
  expect_warning(
    linkString <- makeHyperlinkString(col = 1, row = 4,
                                      text = "test.png", file = img))

  linkString2 <- makeHyperlinkString(text = "test.png", file = img)

  # col and row not needed
  expect_equal(linkString, linkString2)

  # write file without errors
  writeFormula(wb, sheet, x = linkString, startCol = 1, startRow = 1)
  expect_silent(saveWorkbook(wb, tempFile, overwrite = TRUE))

  # TODO: add a check that the written xlsx file contains linkString

})

test_that("writeData2", {
  # create a workbook and add some sheets
  wb <- createWorkbook()

  addWorksheet(wb, "sheet1")
  writeData2(wb, "sheet1", mtcars, colNames = TRUE, rowNames = TRUE)

  addWorksheet(wb, "sheet2")
  writeData2(wb, "sheet2", cars, colNames = FALSE)

  addWorksheet(wb, "sheet3")
  writeData2(wb, "sheet3", letters)

  addWorksheet(wb, "sheet4")
  writeData2(wb, "sheet4", as.data.frame(Titanic), startRow = 2, startCol = 2)

  file <- tempfile(fileext = ".xlsx")
  saveWorkbook(wb, file = file, overwrite = TRUE)

  wb1 <- loadWorkbook(file)

  expect_true(
    all.equal(mtcars, wb_to_df(wb1, "sheet1", rowNames = TRUE),
            check.attributes = FALSE)
  )

  expect_equivalent(cars, wb_to_df(wb1, "sheet2", colNames = FALSE))

  expect_equal(letters,
               as.character(wb_to_df(wb1, "sheet3", colNames = FALSE))
  )

  expect_equivalent(wb_to_df(wb1, "sheet4"),
                    as.data.frame(Titanic, stringsAsFactors = FALSE))

  file.remove(file)
})
