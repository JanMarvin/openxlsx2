test_that("read_xlsx() does not change random seed", {
  rs <- .Random.seed
  expect_identical(rs, .Random.seed)
  tf <- temp_xlsx()
  expect_identical(rs, .Random.seed)
  write_xlsx(data.frame(a = 1), tf)
  expect_identical(rs, .Random.seed)
  read_xlsx(tf)
  expect_identical(rs, .Random.seed)
  unlink(tf)
})

test_that("wb_add_mschart() does not alter the seed", {

  skip_if_not_installed("mschart")

  require(mschart)

  rs <- .Random.seed

  ### Scatter
  scatter <- ms_scatterchart(data = iris, x = "Sepal.Length",
                             y = "Sepal.Width", group = "Species")

  wb <- wb_workbook()$
    add_worksheet()$add_mschart(graph = scatter)$
    add_worksheet()$add_mschart(graph = scatter)

  expect_identical(rs, .Random.seed)
  expect_false(wb$charts$chart[1] == wb$charts$chart[2])

})
