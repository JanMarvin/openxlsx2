test_that("names", {
  tmp <- temp_xlsx()

  wb <- wb_workbook()
  wb$add_worksheet("S1")$add_worksheet("S2 & S3")$add_worksheet("S3 <> S4")

  exp <- c("S1", "S2 & S3", "S3 <> S4")
  got <- names(wb)
  expect_equal(exp, got)

  expect_error(wb_save(wb, tmp), NA)

  expect_error(wb <- wb_load(tmp), NA)
  got <- names(wb)
  expect_equal(exp, got)

  expect_error(names(wb) <- c("S1", "S2", "S2"), "Worksheet names must be unique.")

  expect_error(
    expect_warning(
      names(wb) <- c("A", "B"),
      # First hint: something is not right.
      "longer object length is not a multiple of shorter object length"
    ),
    # Second hint: something is not right!
    "names vector must have length equal to number of worksheets in Workbook"
  )

  wb <- wb_workbook()
  # TODO this should throw a warning or an error
  expect_silent(names(wb) <- "S1")

  wb$add_worksheet("S1")
  expect_warning(names(wb) <- paste0(letters, letters, collapse = ""), "Worksheet names must less than 32 characters. Truncating names...")
  file.remove(tmp)
})
