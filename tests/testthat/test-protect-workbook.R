test_that("Protect Workbook", {
  wb <- wb_workbook()
  names(wb$workbook)

  wb$add_worksheet("s1")

  wb$protect(password = "abcdefghij")

  expect_equal(wb$workbook$workbookProtection, "<workbookProtection hashPassword=\"FEF1\" lockStructure=\"0\" lockWindows=\"0\"/>")


  # this creates a corrupted workbook:
  # 1) this nulls the reference
  wb$protect(protect = FALSE, password = "abcdefghij", lock_structure = TRUE, lock_windows = TRUE)
  expect_null(wb$workbook$workbookProtection)

  # 2) this creates the reference, but at the wrong position (at the end, not at 6)
  wb$protect(password = "abcdefghij", lock_structure = TRUE)

  names(wb$workbook)
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

})
