test_that("names", {
  tmp <- temp_xlsx()

  wb <- wb_workbook()
  wb$add_worksheet("S1")$add_worksheet("S2 & S3")$add_worksheet("S3 <> S4")

  exp <- c("S1", "S2 & S3", "S3 <> S4")
  got <- names(wb$get_sheet_names())
  expect_identical(exp, got)

  # no issues with saving
  expect_error(wb_save(wb, tmp), NA)

  wb <- wb_load(tmp)
  got <- names(wb$get_sheet_names())
  expect_identical(exp, got)

  expect_error(wb$set_sheet_names(new = c("S1", "S2", "S2")), "duplicates")
  expect_error(wb$set_sheet_names(new = c("A", "B")), "same length")

  wb <- wb_workbook()
  expect_error(wb$set_sheet_names(new = "S1"), "does not contain any sheets")

  wb$add_worksheet("S1")
  expect_error(wb$set_sheet_names(new = paste(rep("a", 32), collapse = "")), "31 chars")
  file.remove(tmp)
})
