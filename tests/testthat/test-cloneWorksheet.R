test_that("clone Worksheet with data", {

  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")
  wb$add_data("Sheet 1", 1)
  wb$clone_worksheet("Sheet 1", "Sheet 2")

  file_name <- testfile_path("cloneWorksheetExample.xlsx")
  refwb <- wb_load(file = file_name)

  expect_equal(wb$get_sheet_names(), refwb$get_sheet_names())
  expect_equal(wb_get_order(wb), wb_get_order(refwb))
})


test_that("clone empty Worksheet", {

  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")
  wb$clone_worksheet("Sheet 1", "Sheet 2")

  file_name <- testfile_path("cloneEmptyWorksheetExample.xlsx")
  refwb <- wb_load(file = file_name)

  expect_equal(wb$get_sheet_names(), refwb$get_sheet_names())
  expect_equal(wb_get_order(wb), wb_get_order(refwb))
})


test_that("clone Worksheet with table", {

  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")
  wb$add_data_table(sheet = "Sheet 1", x = iris, tableName = "iris")
  wb$add_data_table(sheet = 1, x = mtcars, tableName = "mtcars", startCol = 10)

  # FIXME a wild drawing2.xml appears in wb$Content_Types
  wb$clone_worksheet("Sheet 1", "Clone 1")

  old <- wb_validate_sheet(wb, "Sheet 1")
  new <- wb_validate_sheet(wb, "Clone 1")

  expect_equal(nrow(wb$tables), 4)
  expect_equal(nrow(wb$tables$tab_sheet == old), nrow(wb$tables$tab_sheet == new))
  relships <- rbindlist(xml_attr(unlist(wb$worksheets_rels), "Relationship"))
  relships$typ <- basename(relships$Type)
  relships$target <- basename(relships$Target)

  got <- relships[relships$typ == "table", c("Id", "typ", "target")]

  exp <- structure(list(
    Id = c("rId1", "rId2", "rId1", "rId2"),
    typ = c("table", "table", "table", "table"),
    target = c("table1.xml", "table2.xml", "table3.xml", "table4.xml")
  ),
  row.names = c(1L, 2L, 3L, 4L),
  class = "data.frame")

  expect_equal(got, exp)

})

test_that("copy cells", {

  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = mtcars)$
    add_fill(dims = "A1:F1", color = wb_color("yellow"))

  dat <- wb_data(wb, dims = "A1:D4", colNames = FALSE)

  # FIXME there is a bug with next_sheet() in clone_worksheets()
  wb$
    # 1:1 copy to M2
    clone_worksheet(old = 1, new = "Clone1")$
    copy_cells(data = dat, dims = "M2", as_value = FALSE, as_ref = FALSE, transpose = FALSE)$
    # 1:1 transposed copy to A20
    clone_worksheet(old = 1, new = "Clone2")$
    copy_cells(data = dat, dims = "A20", as_value = FALSE, as_ref = FALSE, transpose = TRUE)$
    # reference transposed copy to A20
    clone_worksheet(old = 1, new = "Clone3")$
    copy_cells(data = dat, dims = "A20", as_value = FALSE, as_ref = TRUE, transpose = TRUE)$
    # value copy to A20
    clone_worksheet(old = 1, new = "Clone4")$
    copy_cells(data = dat, dims = "A20", as_value = TRUE, as_ref = FALSE, transpose = FALSE)$
    # transposed value copy to A20
    clone_worksheet(old = 1, new = "Clone5")$
    copy_cells(data = dat, dims = "A20", as_value = TRUE, as_ref = FALSE, transpose = TRUE)

  got <- wb_data(wb, sheet = 2, dims = "M2:P5", colNames = FALSE)
  expect_equal(dat, got, ignore_attr = TRUE)

  got <- wb_data(wb, sheet = 3, dims = "A20:D23", colNames = FALSE)
  expect_equal(unlist(t(dat)), unlist(got), ignore_attr = TRUE)

  exp <- c("'Sheet 1'!A1", "'Sheet 1'!B1", "'Sheet 1'!C1", "'Sheet 1'!D1")
  got <- wb_data(wb, sheet = 4, dims = "A20:D23", colNames = FALSE, showFormula = TRUE)[[1]]
  expect_equal(exp, got)

  got <- wb_data(wb, sheet = 5, dims = "A20:D23", colNames = FALSE)
  expect_equal(dat, got, ignore_attr = TRUE)

  got <- wb_data(wb, sheet = 6, dims = "A20:D23", colNames = FALSE)
  expect_equal(unlist(t(dat)), unlist(got), ignore_attr = TRUE)

})

test_that("copy_cells works with hyperlinks and empty cells in transpose", {

  fl <- testfile_path("Single_hyperlink.xlsx")
  wb_in <- wb_load(fl)

  dat <- wb_data(wb_in, 1, dims = "A1:B2", col_names = FALSE)
  wb_in$copy_cells(data = dat, dims = "A3")

  dat <- wb_data(wb_in, 1, dims = "A1:B3", col_names = FALSE)
  wb_in$copy_cells(data = dat, dims = "D1", transpose = TRUE)

  exp <- c("A1", "A3", "D1", "F1")
  got <- vapply(wb_in$worksheets[[1]]$hyperlinks, function(x) x$ref, "")
  expect_equal(exp, got)

  cc <- wb_in$worksheets[[1]]$sheet_data$cc
  exp <- rep("0", 4)
  got <- cc$v[cc$r %in% c("A1", "A3", "D1", "F1")]
  expect_equal(exp, got)

})

test_that("cloning comments works", {

  tmp <- temp_xlsx()

  c1 <- create_comment(text = "this is a comment",  author = "")

  # cloning comments from loaded worksheet did not work
  wb <- wb_workbook()$add_worksheet()$add_comment(dims = "A1", comment = c1)$save(tmp)
  wb <- wb_load(tmp)$clone_worksheet()

  expect_equal(wb$comments[[1]][[1]]$ref, wb$comments[[2]][[1]]$ref)

})

test_that("wb_set_header_footer() works", {
  wb <- wb_workbook()

  # Add example data
  wb$add_worksheet("S1")$add_data(x = 1:400)
  wb$set_header_footer(
    sheet = 1,
    header = c("&[Date]", "ALL HEAD CENTER 2", "&[Page] / &[Pages]"),
    footer = c("&[Path]&[File]", NA, "&[Tab]"),
    first_header = c(NA, "Center Header of First Page", NA),
    first_footer = c(NA, "Center Footer of First Page", NA)
  )

  exp <- list(
    oddHeader = list("&amp;D", "ALL HEAD CENTER 2", "&amp;P / &amp;N"),
    oddFooter = list("&amp;Z&amp;F", NULL, "&amp;A"),
    evenHeader = list(),
    evenFooter = list(),
    firstHeader = list(NULL, "Center Header of First Page", NULL),
    firstFooter = list(NULL, "Center Footer of First Page", NULL)
  )
  got <- wb$worksheets[[1]]$headerFooter
  expect_equal(exp, got)

})

test_that("cloning from workbooks works", {

  ## FIXME these tests should be improved, right now they only check the
  ## existance of a worksheet

  # create a second workbook
  wb <- wb_workbook()$
    add_worksheet("NOT_SUM")$
    add_data(x = head(iris))$
    add_fill(dims = "A1:B2", color = wb_color("yellow"))$
    add_border(dims = "B2:C3")

  ## styled cells
  fl <- system.file("extdata", "oxlsx2_sheet.xlsx", package = "openxlsx2")
  wb_in <- wb_load(fl)

  wb$clone_worksheet(old = "SUM", new = "SUM", from = wb_in)
  exp <- c("NOT_SUM", "SUM")
  got <- wb$get_sheet_names() %>% unname()
  expect_equal(exp, got)

  wb$clone_worksheet(old = "SUM", new = "SUM_clone")
  exp <- c("NOT_SUM", "SUM", "SUM_clone")
  got <- wb$get_sheet_names() %>% unname()
  expect_equal(exp, got)

  wb$clone_worksheet(old = "SUM", new = "SUM2", from = wb_in)
  exp <- c("NOT_SUM", "SUM", "SUM_clone", "SUM2")
  got <- wb$get_sheet_names() %>% unname()
  expect_equal(exp, got)

  ## clone table

  wb_in <- wb_workbook()$add_worksheet("tab")$add_data_table(x = mtcars)

  wb$clone_worksheet(old = "tab", new = "tab", from = wb_in)
  exp <- c("NOT_SUM", "SUM", "SUM_clone", "SUM2", "tab")
  got <- wb$get_sheet_names() %>% unname()
  expect_equal(exp, got)

  # clone it twice
  expect_warning(wb$clone_worksheet(old = "tab", new = "tab", from = wb_in))
  exp <- c("NOT_SUM", "SUM", "SUM_clone", "SUM2", "tab", "tab (1)")
  got <- wb$get_sheet_names() %>% unname()
  expect_equal(exp, got)

  ## clone a chart

  skip_if_not_installed("mschart")
  library(mschart)
  ## Add mschart to worksheet (adds data and chart)
  scatter <- ms_scatterchart(data = iris, x = "Sepal.Length", y = "Sepal.Width", group = "Species")
  scatter <- chart_settings(scatter, scatterstyle = "marker")

  wb_ms <- wb_workbook() %>%
    wb_add_worksheet("chart") %>%
    wb_add_mschart(dims = "F4:L20", graph = scatter)

  wb$clone_worksheet(old = "chart", new = "chart_1", from = wb_ms)
  exp <- c("NOT_SUM", "SUM", "SUM_clone", "SUM2", "tab", "tab (1)", "chart_1")
  got <- wb$get_sheet_names() %>% unname()
  expect_equal(exp, got)

  wb$clone_worksheet(old = "chart", new = "chart_2", from = wb_ms)
  exp <- c("NOT_SUM", "SUM", "SUM_clone", "SUM2", "tab", "tab (1)", "chart_1", "chart_2")
  got <- wb$get_sheet_names() %>% unname()
  expect_equal(exp, got)

  ## clone images

  img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")

  wb_img <- wb_workbook()$
    add_worksheet()$
    add_image("Sheet 1", dims = "C5", file = img, width = 6, height = 5)

  wb$clone_worksheet(old = "Sheet 1", new = "img", from = wb_img)
  exp <- c("NOT_SUM", "SUM", "SUM_clone", "SUM2", "tab", "tab (1)", "chart_1", "chart_2", "img")
  got <- wb$get_sheet_names() %>% unname()
  expect_equal(exp, got)

  wb$clone_worksheet(old = "Sheet 1", new = "img2", from = wb_img)
  exp <- c("NOT_SUM", "SUM", "SUM_clone", "SUM2", "tab", "tab (1)", "chart_1", "chart_2", "img", "img2")
  got <- wb$get_sheet_names() %>% unname()
  expect_equal(exp, got)

  ## clone drawing, borders function and shared strings
  wb_ex <- wb_load(testfile_path("loadExample.xlsx"))

  wb$clone_worksheet(old = "testing", new = "test", from = wb_ex)
  exp <- c("NOT_SUM", "SUM", "SUM_clone", "SUM2", "tab", "tab (1)", "chart_1", "chart_2", "img", "img2", "test")
  got <- wb$get_sheet_names() %>% unname()
  expect_equal(exp, got)

})

test_that("cloning column and row styles works", {

  tmp <- temp_xlsx()

  ### prepare a worksheet to clone from
  wb <- wb_workbook()$
    add_worksheet()$add_data(x = head(mtcars))$
    add_worksheet()$add_data(x = head(iris))

  new_fill <- create_fill(patternType = "solid", fgColor = wb_color(hex = "yellow"))

  wb$styles_mgr$add(new_fill, "new_fill")
  new_cellxfs <- create_cell_style(
    num_fmt_id = 0,
    fill_id = wb$styles_mgr$get_fill_id("new_fill")
  )
  wb$styles_mgr$add(new_cellxfs, "new_styles")

  wb$worksheets[[1]]$sheet_data$row_attr[2, "customFormat"] <- "1"
  wb$worksheets[[1]]$sheet_data$row_attr[2, "s"] <- wb$styles_mgr$get_xf_id("new_styles")

  cols <- openxlsx2:::wb_create_columns(wb, sheet = 1, cols = seq_along(mtcars))
  cols[cols$min == 11, "style"] <- "1"
  wb$worksheets[[1]]$fold_cols(cols)

  # otherwise the style color somehow is omitted in the data area
  wb$set_cell_style(sheet = 1, dims = "A2:K2", style = wb$styles_mgr$get_xf_id("new_styles"))
  wb$set_cell_style(sheet = 1, dims = "K1:K7", style = wb$styles_mgr$get_xf_id("new_styles"))

  wb$save(tmp)
  rm(wb, cols, new_cellxfs, new_fill)

  ### clone the actual worksheet
  wb_mtcars <- wb_load(tmp)

  wb <- wb_workbook()$add_worksheet()
  wb$clone_worksheet(old = "Sheet 1", "mtcars", from = wb_mtcars)
  wb$clone_worksheet(old = "Sheet 2", "iris", from = wb_mtcars)

  exp <- "<col min=\"11\" max=\"11\" style=\"2\" width=\"8.43\"/>"
  got <- wb$worksheets[[2]]$cols_attr[2]
  expect_equal(exp, got)

  exp <- "3"
  got <- wb$worksheets[[2]]$sheet_data$row_attr[2, "s"]
  expect_equal(exp, got)

})
