test_that("Deleting worksheets", {
  tempFile <- temp_xlsx()
  genWS <- function(wb, sheetName) {
    wb$add_worksheet(sheetName)
    wb$add_data_table(sheetName, data.frame("X" = sprintf("This is sheet: %s", sheetName)), col_names = TRUE)
  }

  wb <- wb_workbook()
  genWS(wb, "Sheet 1")
  genWS(wb, "Sheet 2")
  genWS(wb, "Sheet 3")
  expect_equal(names(wb$get_sheet_names()), c("Sheet 1", "Sheet 2", "Sheet 3"))

  wb$remove_worksheet(sheet = 1)
  expect_equal(names(wb$get_sheet_names()), c("Sheet 2", "Sheet 3"))

  wb$remove_worksheet(sheet = 1)
  expect_equal(names(wb$get_sheet_names()), c("Sheet 3"))

  ## add to end
  genWS(wb, "Sheet 1")
  genWS(wb, "Sheet 2")
  expect_equal(names(wb$get_sheet_names()), c("Sheet 3", "Sheet 1", "Sheet 2"))

  wb_save(wb, tempFile, overwrite = TRUE)

  ## re-load & re-order worksheets

  wb <- wb_load(tempFile)
  expect_equal(names(wb$get_sheet_names()), c("Sheet 3", "Sheet 1", "Sheet 2"))

  wb$add_data(sheet = "Sheet 2", x = iris[1:10, 1:4], start_row = 5)
  test <- read_xlsx(wb, "Sheet 2", start_row = 5)
  rownames(test) <- seq_len(nrow(test))
  attr(test, "tt") <- NULL
  attr(test, "types") <- NULL
  expect_equal(iris[1:10, 1:4], test)

  wb$add_data(sheet = 1, x = iris[1:20, 1:4], start_row = 5)
  test <- read_xlsx(wb, "Sheet 3", start_row = 5)
  rownames(test) <- seq_len(nrow(test))
  attr(test, "tt") <- NULL
  attr(test, "types") <- NULL
  expect_equal(iris[1:20, 1:4], test)

  wb$remove_worksheet(sheet = 1)
  expect_equal(read_xlsx(wb, 1, start_row = 1)[[1]], "This is sheet: Sheet 1")

  wb$remove_worksheet(sheet = 2)
  expect_equal(read_xlsx(wb, 1, start_row = 1)[[1]], "This is sheet: Sheet 1")

  wb$remove_worksheet(sheet = 1)
  expect_equal(names(wb$get_sheet_names()), character(0))

  unlink(tempFile, recursive = TRUE, force = TRUE)

  wb <- wb_load(file = system.file("extdata", "oxlsx2_sheet.xlsx", package = "openxlsx2"))
  expect_silent(wb$remove_worksheet())

})

test_that("removing leading chartsheets works", {

  # We remove sheet 1 which is a chartsheet and the only chartsheet in the file.
  # We remember the first worksheet id. And even though the first worksheet
  # could be sheet 1, we keep it at 2. Otherwise our reference counter would get
  # in trouble. Similar things could happen if all worksheets are removed and
  # only chartsheets remain. Though that is currently not implemented.
  fl <- testfile_path("mtcars_chart.xlsx")
  tmp <- temp_xlsx()
  wb <- wb_load(fl)$
    remove_worksheet(1)

  wb$save(tmp)
  tmp_dir <- paste0(tempdir(), "/openxlsx2_unzip")
  dir.create(tmp_dir)
  unzip(tmp, exdir = tmp_dir)

  exp <- c("sheet1.xml", "sheet2.xml", "sheet3.xml")
  got <- dir(paste0(tmp_dir, "/xl/worksheets"), pattern = "*.xml")
  expect_equal(exp, got)

  unlink(tmp_dir, recursive = TRUE)

  ### avoid duplicated names
  wb <- wb_workbook()$
    add_worksheet()$
    add_worksheet()$
    remove_worksheet(1)$
    add_worksheet()

  expect_equal(wb$sheet_names, c("Sheet 2", "Sheet 3"))

  ###
  skip_if_not_installed("mschart")

  require(mschart)
  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = mtcars)

  # create wb_data object
  dat <- wb_data(wb, 1)

  # call ms_scatterplot
  data_plot <- ms_scatterchart(
    data = dat,
    x = "mpg",
    y = c("disp", "hp"),
    labels = c("disp", "hp")
  )

  wb$
    add_chartsheet()$
    add_mschart(graph = data_plot)$
    add_worksheet()$
    add_data(x = "something")

  wb$remove_worksheet(2)

  wb$save(tmp)
  tmp_dir <- paste0(tempdir(), "/openxlsx2_unzip")
  dir.create(tmp_dir)
  unzip(tmp, exdir = tmp_dir)

  got <- dir(paste0(tmp_dir, "/xl/worksheets"), pattern = "*.xml")
  exp <- c("sheet1.xml", "sheet2.xml")
  expect_equal(got, exp)

  unlink(tmp_dir, recursive = TRUE)

})

test_that("bookViews activeTab attribute is updated", {

  wb <- wb_workbook()$
    add_worksheet()$
    add_worksheet()

  wb$set_active_sheet(2)
  got <- wb$workbook$bookViews
  exp <- "<bookViews><workbookView activeTab=\"1\"/></bookViews>"
  expect_equal(got, exp)

  wb$remove_worksheet(1)
  got <- wb$workbook$bookViews
  exp <- "<bookViews><workbookView activeTab=\"0\"/></bookViews>"
  expect_equal(got, exp)

})
