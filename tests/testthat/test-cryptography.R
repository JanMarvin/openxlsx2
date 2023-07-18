
test_that("test that encryption works", {

  input  <- temp_xlsx()
  output <- temp_xlsx()

  # write input
  wb_workbook()$add_worksheet()$add_data(x = mtcars)$save(input)

  # encrypt it
  expect_equal(
    msoc("enc", input, output, "openxlsx2"),
    0
  )

  expect_error(
    msoc("enc", output, output, "openxlsx2"),
    "ERR already encrypted"
  )

  # does not work
  expect_error(
    expect_warning(
      wb_to_df(output)
    )
  )

  # decrypt it
  expect_equal(
    msoc("dec", output, output, "openxlsx2"),
    0
  )

  tmp <- temp_xlsx()
  wbin <- wb_workbook()$add_worksheet()$
    add_data(x = head(iris))$save(tmp, password = "openxlsx2")

  wbout <- wb_load(tmp, password = "openxlsx2")
  exp <- wb_to_df(wbin)
  got <- wb_to_df(wbout)

  expect_equal(exp, got)

})
