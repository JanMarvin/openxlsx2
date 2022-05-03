test_that("Protect Workbook", {
  wb <- wb_workbook()
  wb$add_worksheet("s1")

  wb$protect(password = "abcdefghij")

  expect_true(wb$workbook$workbookProtection == "<workbookProtection hashPassword=\"FEF1\" lockStructure=\"0\" lockWindows=\"0\"/>")

  wb$protect(protect = FALSE, password = "abcdefghij", lockStructure = TRUE, lockWindows = TRUE)
  expect_true(is.null(wb$workbook$workbookProtection))
})

test_that("Reading protected Workbook", {
  temp_file <- temp_xlsx()

  wb <- wb_workbook()
  wb$add_worksheet("s1")
  wb_protect(wb, password = "abcdefghij")
  wb_save(wb, temp_file)

  wb2 <- wb_load(file = temp_file)
  # Check that the order of the sub-elements is preserved
  n1 <- names(wb2$workbook)
  n2 <- names(wb$workbook)[names(wb$workbook) != "apps"]
  expect_equal(n1, n2)

  file.remove(temp_file)
})
