
test_that("test return values for wb_save", {

  tempFile <- temp_xlsx()
  wb <- wb_add_worksheet(wb_workbook(), "name")
  expect_equal(tempFile, wb_save(wb, tempFile))

  expect_error( wb_save(wb, tempFile))


  expect_equal(tempFile, wb_save(wb, tempFile, overwrite = TRUE))
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
  sheet <- "test"
  wb <- wb_add_worksheet(wb_workbook(), sheet)
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
  expect_silent(wb_save(wb, tempFile, overwrite = TRUE))

  # TODO: add a check that the written xlsx file contains linkString

})
