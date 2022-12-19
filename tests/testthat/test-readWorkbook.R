
test_that("check_xlsx_path() gives correct errors [#477]", {
  x <- tempfile()
  expect_error(check_xlsx_path(x))

  x <- tempfile(fileext = ".XLSX")
  file.create(x)
  expect_error(check_xlsx_path(x), NA)
  file.remove(x)

  x <- tempfile(fileext = ".zip")
  file.create(x)
  expect_warning(check_xlsx_path(x))
  expect_warning(check_xlsx_path(x, warn = FALSE), NA)
  file.remove(x)

  x <- tempfile(fileext = ".xls")
  file.create(x)
  expect_error(check_xlsx_path(x))
  file.remove(x)

  x <- tempfile(fileext = ".xlsm")
  file.create(x)
  expect_error(check_xlsx_path(x))
  file.remove(x)

  x <- tempfile(fileext = ".xlsb")
  file.create(x)
  expect_error(check_xlsx_path(x))
  file.remove(x)
})
