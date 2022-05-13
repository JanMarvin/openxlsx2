test_that("clean_sheet", {

  xlsx <- system.file("extdata", "oxlsx2_sheet.xlsx", package = "openxlsx2")
  wb <- wb_load(xlsx)

  # remove merged cells
  wb <- wb_clean_sheet(wb = wb, sheet = 1, numbers = FALSE, characters = FALSE, styles = FALSE, merged_cells = TRUE)
  expect_identical(character(), wb$worksheets[[1]]$mergeCells)

  # remove styles
  wb <- wb_clean_sheet(wb = wb, sheet = 1, numbers = FALSE, characters = FALSE, styles = TRUE)
  expect_true(all(wb$worksheets[[1]]$sheet_data$cc$c_s == ""))

  # remove numbers
  wb$worksheets[[1]]$clean_sheet(numbers = TRUE)
  got <- wb_to_df(wb, dims = "B5:I16")
  expect_true(all(is.na(got)))

  # remove characters
  wb$worksheets[[1]]$clean_sheet(characters = TRUE)
  expect_true(is.na(wb_to_df(wb, dims = "B2", colNames = FALSE)))

})
