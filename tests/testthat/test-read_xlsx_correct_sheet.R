test_that("read_xlsx correct sheet", {

  fl <- testfile_path("readTest.xlsx")
  sheet_names <- wb_load(fl)$get_sheet_names() %>% unname()

  expected_sheet_names <- c(
    "Sheet1", "Sheet2", "Sheet 3",
    "Sheet 4", "Sheet 5", "Sheet 6",
    "1", "11", "111", "1111", "11111", "111111"
  )

  expect_equal(object = sheet_names, expected = expected_sheet_names)

  expect_equal(read_xlsx(file = fl, sheet = 7),  data.frame(x = 1),      ignore_attr = TRUE)
  expect_equal(read_xlsx(file = fl, sheet = 8),  data.frame(x = 11),     ignore_attr = TRUE)
  expect_equal(read_xlsx(file = fl, sheet = 9),  data.frame(x = 111),    ignore_attr = TRUE)
  expect_equal(read_xlsx(file = fl, sheet = 10), data.frame(x = 1111),   ignore_attr = TRUE)
  expect_equal(read_xlsx(file = fl, sheet = 11), data.frame(x = 11111),  ignore_attr = TRUE)
  expect_equal(read_xlsx(file = fl, sheet = 12), data.frame(x = 111111), ignore_attr = TRUE)

  expect_equal(read_xlsx(file = fl, sheet = "1"),      data.frame(x = 1),      ignore_attr = TRUE)
  expect_equal(read_xlsx(file = fl, sheet = "11"),     data.frame(x = 11),     ignore_attr = TRUE)
  expect_equal(read_xlsx(file = fl, sheet = "111"),    data.frame(x = 111),    ignore_attr = TRUE)
  expect_equal(read_xlsx(file = fl, sheet = "1111"),   data.frame(x = 1111),   ignore_attr = TRUE)
  expect_equal(read_xlsx(file = fl, sheet = "11111"),  data.frame(x = 11111),  ignore_attr = TRUE)
  expect_equal(read_xlsx(file = fl, sheet = "111111"), data.frame(x = 111111), ignore_attr = TRUE)
})
