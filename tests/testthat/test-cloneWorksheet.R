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
  wb$add_data_table(sheet = "Sheet 1", x = iris, table_name = "iris")
  wb$add_data_table(sheet = 1, x = mtcars, table_name = "mtcars", start_col = 10)

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

  dat <- wb_data(wb, dims = "A1:D4", col_names = FALSE)

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

  got <- wb_data(wb, sheet = 2, dims = "M2:P5", col_names = FALSE)
  expect_equal(dat, got, ignore_attr = TRUE)

  got <- wb_data(wb, sheet = 3, dims = "A20:D23", col_names = FALSE)
  expect_equal(unlist(t(dat)), unlist(got), ignore_attr = TRUE)

  exp <- structure(
    c("'Sheet 1'!A1", "'Sheet 1'!B1", "'Sheet 1'!C1", "'Sheet 1'!D1"),
    class = c("character", "formula")
  )
  got <- wb_data(wb, sheet = 4, dims = "A20:D23", col_names = FALSE, show_formula = TRUE)[[1]]
  expect_equal(exp, got)

  got <- wb_data(wb, sheet = 5, dims = "A20:D23", col_names = FALSE)
  expect_equal(unlist(dat), unlist(got), ignore_attr = TRUE)

  got <- wb_data(wb, sheet = 6, dims = "A20:D23", col_names = FALSE)
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
  got <- rbindlist(xml_attr(wb_in$worksheets[[1]]$hyperlinks, "hyperlink"))$ref
  expect_equal(exp, got)

  cc <- wb_in$worksheets[[1]]$sheet_data$cc
  exp <- rep("0", 4)
  got <- cc$v[cc$r %in% c("A1", "A3", "D1", "F1")]
  expect_equal(exp, got)

})

test_that("cloning comments works", {

  tmp <- temp_xlsx()
  on.exit(unlink(tmp), add = TRUE)

  c1 <- wb_comment(text = "this is a comment", author = "")

  # cloning comments from loaded worksheet did not work
  wb <- wb_workbook()$add_worksheet()$add_comment(dims = "A1", comment = c1)$save(tmp)
  wb <- wb_load(tmp)$clone_worksheet()

  expect_equal(wb$comments[[1]][[1]]$ref, wb$comments[[2]][[1]]$ref)

})

test_that("cloning from workbooks works", {

  ## FIXME these tests should be improved, right now they only check the
  ## existence of a worksheet

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
  got <- unname(wb$get_sheet_names())
  expect_equal(got, c("NOT_SUM", "SUM"))

  wb$clone_worksheet(old = "SUM", new = "SUM_clone")
  exp <- c("NOT_SUM", "SUM", "SUM_clone")
  got <- unname(wb$get_sheet_names())
  expect_equal(exp, got)

  wb$clone_worksheet(old = "SUM", new = "SUM2", from = wb_in)
  exp <- c("NOT_SUM", "SUM", "SUM_clone", "SUM2")
  got <- unname(wb$get_sheet_names())
  expect_equal(exp, got)

  ## clone table

  wb_in <- wb_workbook()$add_worksheet("tab")$add_data_table(x = mtcars)

  wb$clone_worksheet(old = "tab", new = "tab", from = wb_in)
  exp <- c("NOT_SUM", "SUM", "SUM_clone", "SUM2", "tab")
  got <- unname(wb$get_sheet_names())
  expect_equal(exp, got)

  # clone it twice
  expect_warning(wb$clone_worksheet(old = "tab", new = "tab", from = wb_in), "already exists")
  exp <- c("NOT_SUM", "SUM", "SUM_clone", "SUM2", "tab", "tab (1)")
  got <- unname(wb$get_sheet_names())
  expect_equal(exp, got)

  ## clone a chart

  skip_if_not_installed("mschart")
  library(mschart)
  ## Add mschart to worksheet (adds data and chart)
  scatter <- ms_scatterchart(data = iris, x = "Sepal.Length", y = "Sepal.Width", group = "Species")
  scatter <- chart_settings(scatter, scatterstyle = "marker")

  wb_ms <- wb_workbook()$
    add_worksheet("chart")$
    add_mschart(dims = "F4:L20", graph = scatter)

  wb$clone_worksheet(old = "chart", new = "chart_1", from = wb_ms)
  exp <- c("NOT_SUM", "SUM", "SUM_clone", "SUM2", "tab", "tab (1)", "chart_1")
  got <- unname(wb$get_sheet_names())
  expect_equal(exp, got)

  wb$clone_worksheet(old = "chart", new = "chart_2", from = wb_ms)
  exp <- c("NOT_SUM", "SUM", "SUM_clone", "SUM2", "tab", "tab (1)", "chart_1", "chart_2")
  got <- unname(wb$get_sheet_names())
  expect_equal(exp, got)

  ## clone images

  img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")

  wb_img <- wb_workbook()$
    add_worksheet()$
    add_image("Sheet 1", dims = "C5", file = img, width = 6, height = 5)

  wb$clone_worksheet(old = "Sheet 1", new = "img", from = wb_img)
  exp <- c("NOT_SUM", "SUM", "SUM_clone", "SUM2", "tab", "tab (1)", "chart_1", "chart_2", "img")
  got <- unname(wb$get_sheet_names())
  expect_equal(exp, got)

  wb$clone_worksheet(old = "Sheet 1", new = "img2", from = wb_img)
  exp <- c("NOT_SUM", "SUM", "SUM_clone", "SUM2", "tab", "tab (1)", "chart_1", "chart_2", "img", "img2")
  got <- unname(wb$get_sheet_names())
  expect_equal(exp, got)

  ## clone drawing, borders function and shared strings
  wb_ex <- wb_load(testfile_path("loadExample.xlsx"))

  wb$clone_worksheet(old = "testing", new = "test", from = wb_ex)
  exp <- c("NOT_SUM", "SUM", "SUM_clone", "SUM2", "tab", "tab (1)", "chart_1", "chart_2", "img", "img2", "test")
  got <- unname(wb$get_sheet_names())
  expect_equal(exp, got)

})

test_that("cloning column and row styles works", {

  tmp <- temp_xlsx()
  on.exit(unlink(tmp), add = TRUE)

  ### prepare a worksheet to clone from
  wb <- wb_workbook()$
    add_worksheet()$add_data(x = head(mtcars))$
    add_worksheet()$add_data(x = head(iris))

  new_fill <- create_fill(pattern_type = "solid", fg_color = wb_color(hex = "yellow"))

  wb$styles_mgr$add(new_fill, "new_fill")
  new_cellxfs <- create_cell_style(
    num_fmt_id = 0,
    fill_id = wb$styles_mgr$get_fill_id("new_fill")
  )
  wb$styles_mgr$add(new_cellxfs, "new_styles")

  wb$worksheets[[1]]$sheet_data$row_attr[2, "customFormat"] <- "1"
  wb$worksheets[[1]]$sheet_data$row_attr[2, "s"] <- wb$styles_mgr$get_xf_id("new_styles")

  # wb_create_columns is an internal function in openxlsx2.
  cols <- wb_create_columns(wb, sheet = 1, cols = seq_along(mtcars))
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

  exp <- "<col min=\"11\" max=\"11\" style=\"1\" width=\"8.43\"/>"
  got <- wb$worksheets[[2]]$cols_attr[2]
  expect_equal(exp, got)

  exp <- "1"
  got <- wb$worksheets[[2]]$sheet_data$row_attr[2, "s"]
  expect_equal(exp, got)

})

test_that("cloning slicers throws warning", {

  wb <- wb_workbook()$add_worksheet()$add_data(x = mtcars)

  df <- wb_data(wb, sheet = 1)

  wb$add_pivot_table(
    df, dims = "A3", slicer = "vs", rows = "cyl", cols = "gear", data = "disp",
    pivot_table = "mtcars"
  )$
    add_slicer(x = df, slicer = "vs", pivot_table = "mtcars")

  expect_warning(
    expect_warning(
      wb$clone_worksheet(old = "Sheet 2", new = "Sheet 3"),
      "Cloning slicers is not yet supported. It will not appear on the sheet."
    ),
    "experimental"
  )

})

test_that("cloning sheets with multiple images works", {

  png1 <- tempfile(fileext = "plot1.png")
  png2 <- tempfile(fileext = "plot2.png")
  jpg1 <- tempfile(fileext = "plot1.jpg")
  jpg2 <- tempfile(fileext = "plot2.jpg")

  png(png1)
  plot(1)
  dev.off()

  png(png2)
  plot(1:2)
  dev.off()

  jpeg(jpg1)
  plot(1:3)
  dev.off()

  jpeg(jpg2)
  plot(1:4)
  dev.off()

  wb_old <- wb_workbook()$add_worksheet()$
    add_image(dims = "A1:A1", file = png1)$
    add_image(dims = "A2:A2", file = png2)$
    add_image(dims = "A3:A3", file = jpg1)$
    add_image(dims = "A4:A4", file = jpg2)

  wb <- wb_workbook()
  wb$clone_worksheet(old = 1, new = "Clone1", from = wb_old)

  exp <- list(
    c(
      "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image1.png\"/>",
      "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image2.png\"/>",
      "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image3.jpg\"/>",
      "<Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image4.jpg\"/>"
    )
  )
  got <- wb$drawings_rels
  expect_equal(exp, got)

  exp <- c("<Default Extension=\"png\" ContentType=\"image/png\"/>", "<Default Extension=\"jpg\" ContentType=\"image/jpg\"/>")
  got <- wb$Content_Types[10:11]
  expect_equal(exp, got)

})

test_that("cloning charts and chartEx works", {

  skip_if_not_installed("encharter")

  wb <- wb_workbook()$add_worksheet("data")$add_data(x = mtcars)

  wf <- encharter::encharter("waterfall")$add_series(data = "data!B2:B5")
  bp <- encharter::encharter("barplot")$add_series(data = "data!B2:B5")

  wb$add_worksheet("wf")$add_encharter(graph = wf)
  wb$add_worksheet("bp")$add_encharter(graph = bp)

  wb2 <- wb_workbook()
  wb2$clone_worksheet("data", "data", from = wb)
  wb2$clone_worksheet("bp", "bp", from = wb)
  expect_warning(wb2$clone_worksheet("wf", "wf", from = wb), "experimental")
  expect_warning(wb2$clone_worksheet("wf", "wf2"), "experimental")

  exp <- c(TRUE, FALSE)
  got <- nzchar(wb2$charts$chart)
  expect_equal(exp, got)

  exp <- c(TRUE, TRUE)
  got <- nzchar(wb2$charts$chartEx)
  expect_equal(exp, got)

  exp <- c(wb$charts$colors, wb$charts$colors)
  got <- wb2$charts$colors
  expect_equal(exp, got)

  expect_true(grepl("Target=\"colors2.xml\"", wb2$charts$relsEx[2], fixed = TRUE))
  expect_true(grepl("Target=\"style2.xml\"", wb2$charts$relsEx[2], fixed = TRUE))

  exp <- c("chart1.xml", "chartEx1.xml", "chartEx2.xml")
  got <- vapply(wb2$drawings_rels, function(x) basename(rbindlist(xml_attr(x, "Relationship"))$Target), "")
  expect_equal(exp, got)

  exp <- c(1L, 2L, 3L)
  got <- vapply(wb2$worksheets[2:4], function(x) x$relships$drawing, 0L)
  expect_equal(exp, got)

  xl_names <- function(x) {
    unique(stringi::stri_extract_all_regex(x, "_xlchart\\.v1\\.[0-9]+", omit_no_match = TRUE)[[1]])
  }
  expect_length(intersect(xl_names(wb2$charts$chartEx[1]), xl_names(wb2$charts$chartEx[2])), 0)

  nr <- wb2$get_named_regions()
  expect_true(all(nr$name %in% c(xl_names(wb2$charts$chartEx[1]), xl_names(wb2$charts$chartEx[2]))))
  expect_equal(nrow(nr), 4L)
  expect_true(all(nr$value == "'data'!$B$2:$B$5"))

})

test_that("cloning defined names works", {

  wb <- wb_workbook()$
    add_worksheet("my data")$
    add_data(x = mtcars, with_filter = TRUE)$
    add_named_region(dims = "A2:A5", name = "glob")$
    add_named_region(dims = "B2:B5", name = "loc", local_sheet = TRUE)$
    add_worksheet("other")$
    add_named_region(dims = "A1", name = "elsewhere")
  wb$page_setup(sheet = 1, print_title_rows = 1)

  expect_warning(wb$clone_worksheet("my data", "copy $1"), "experimental")

  nr <- wb$get_named_regions()
  nr <- nr[nr$localSheetId == "2", c("name", "value")]
  nr <- nr[order(nr$name, method = "radix"), ]
  rownames(nr) <- NULL

  exp <- data.frame(
    name = c("_xlnm.Print_Titles", "_xlnm._FilterDatabase", "glob", "loc"),
    value = c("'copy $1'!$1:$1", "'copy $1'!$A$1:$K$33", "'copy $1'!$A$2:$A$5", "'copy $1'!$B$2:$B$5"),
    stringsAsFactors = FALSE
  )
  expect_equal(exp, nr)

  expect_equal(length(wb$workbook$definedNames), 9L)

  wb2 <- wb_workbook()$add_worksheet("S1")$add_worksheet("S12")
  wb2$workbook$definedNames <- c(
    "<definedName name=\"a\">S1!$A$1,S12!$A$1</definedName>",
    "<definedName name=\"b\">SUM(S1!A1:B2)</definedName>",
    "<definedName name=\"c\">S12!$A$1</definedName>",
    "<definedName name=\"d\">'S1'!$A$1</definedName>"
  )
  expect_warning(wb2$clone_worksheet("S1", "N"), "experimental")

  exp <- c(
    "<definedName name=\"a\" localSheetId=\"2\">'N'!$A$1,S12!$A$1</definedName>",
    "<definedName name=\"b\" localSheetId=\"2\">SUM('N'!A1:B2)</definedName>",
    "<definedName name=\"d\" localSheetId=\"2\">'N'!$A$1</definedName>"
  )
  got <- wb2$workbook$definedNames[5:7]
  expect_equal(exp, got)

})

test_that("cloning copies comments, threaded comments and hyperlinks", {

  wb <- wb_workbook()$add_worksheet("S")
  wb$add_person(name = "me")
  pid <- wb$get_person(name = "me")$id
  wb$add_comment(dims = "A1", comment = "note")
  wb$add_thread(dims = "B2", person_id = pid, comment = "thread")
  wb$add_data(dims = "A5", x = "link")
  wb$add_hyperlink(dims = "A5", target = "https://example.com")

  expect_warning(wb$clone_worksheet("S", "C"), "experimental")

  rels <- wb$worksheets_rels[[2]]
  expect_true(any(grepl("Target=\"../drawings/vmlDrawing2.vml\"", rels, fixed = TRUE)))
  expect_true(any(grepl("Target=\"../comments2.xml\"", rels, fixed = TRUE)))
  expect_true(any(grepl("Target=\"../threadedComments/threadedComment2.xml\"", rels, fixed = TRUE)))
  expect_true(any(grepl("Target=\"https://example.com\" TargetMode=\"External\"", rels, fixed = TRUE)))
  expect_true(any(grepl("/xl/threadedComments/threadedComment2.xml", wb$Content_Types, fixed = TRUE)))
  expect_equal(length(wb$comments), 2L)
  expect_equal(length(wb$threadComments), 2L)

  # threaded comments and vml shapes get new ids
  ids1 <- rbindlist(xml_attr(wb$threadComments[[1]], "threadedComment"))$id
  ids2 <- rbindlist(xml_attr(wb$threadComments[[2]], "threadedComment"))$id
  expect_length(intersect(ids1, ids2), 0)
  authors <- vapply(wb$comments[[2]], function(x) x$author, "")
  expect_true(paste0("tc=", ids2) %in% authors)
  expect_false(paste0("tc=", ids1) %in% authors)

  spid <- function(x) unique(regmatches(x, gregexpr("_x0000_s[0-9]+", x))[[1]])
  expect_length(intersect(spid(wb$vml[[1]]), spid(wb$vml[[2]])), 0)
  expect_match(wb$vml[[2]], '<o:idmap v:ext="edit" data="2"/>', fixed = TRUE)

  wb2 <- wb_workbook()
  expect_warning(wb2$clone_worksheet("S", "S", from = wb), "experimental")
  expect_equal(wb2$persons, wb$persons)
  expect_true(any(grepl("persons/person.xml", wb2$workbook.xml.rels, fixed = TRUE)))
  expect_true(any(grepl("vml", wb2$Content_Types, fixed = TRUE)))

  wb3 <- wb_workbook()$add_worksheet("B")
  wb3$add_person(name = "you")
  wb3$add_thread(dims = "A1", person_id = wb3$get_person(name = "you")$id, comment = "other")
  expect_warning(wb3$clone_worksheet("S", "S", from = wb), "experimental")
  expect_equal(wb3$get_person()$displayName, c("you", "me"))
})

test_that("cloning copies form controls, named sheet views and diagrams", {

  wb <- wb_load(testfile_path("form_control.xlsx"))
  n <- length(wb$ctrlProps)
  expect_warning(wb$clone_worksheet(1, "clone"), "experimental")
  expect_equal(length(wb$ctrlProps), 2L * n)
  targets <- rbindlist(xml_attr(wb$worksheets_rels[[2]], "Relationship"))$Target
  exp <- sprintf("../ctrlProps/ctrlProp%s.xml", seq_len(n) + n)
  expect_equal(sort(grep("ctrlProp", targets, value = TRUE)), sort(exp))

  wb <- wb_load(testfile_path("named_sv.xlsx"))
  wb$clone_worksheet("Sheet2", "clone")
  expect_equal(length(wb$namedSheetViews), 2L)
  expect_true(any(grepl("../namedSheetViews/namedSheetView2.xml", wb$worksheets_rels[[3]], fixed = TRUE)))

  wb <- wb_load(testfile_path("diagram.xlsx"))
  n <- length(wb$diagrams)
  wb$clone_worksheet("Sheet1", "clone")
  drels <- wb$drawings_rels[[wb$worksheets[[3]]$relships$drawing]]
  targets <- basename(rbindlist(xml_attr(drels, "Relationship"))$Target)
  expect_length(intersect(targets, basename(wb$diagrams[seq_len(n)])), 0)
  expect_true(all(targets %in% basename(wb$diagrams)))
  expect_true(all(sprintf("/xl/diagrams/%s", targets) %in% rbindlist(xml_attr(wb$Content_Types[grepl("Override", wb$Content_Types)], "Override"))$PartName))
})

test_that("cloning copies shapes on top of charts", {

  wb <- wb_load(testfile_path("unemployment-nrw202208.xlsx"))
  n <- length(wb$drawings)
  wb$clone_worksheet(17, "clone")

  sheet_drawing <- wb$worksheets[[length(wb$worksheets)]]$relships$drawing
  expect_equal(sheet_drawing, n + 1L)
  expect_equal(length(wb$drawings), n + 2L)
  expect_equal(xml_node_name(wb$drawings[[n + 2L]]), "c:userShapes")

  targets <- rbindlist(xml_attr(wb$drawings_rels[[sheet_drawing]], "Relationship"))$Target
  crt <- cdigit(basename(grep("charts/chart", targets, value = TRUE)), as_integer = TRUE)
  expect_length(crt, 1)
  expect_match(wb$charts$rels[crt], sprintf("../drawings/drawing%s.xml", n + 2L), fixed = TRUE)
})

test_that("cloning pivot tables to another workbook drops the pivot table", {

  wb <- wb_workbook()$add_worksheet("data")$add_data(x = mtcars)
  df <- wb_data(wb, sheet = "data", dims = "A1:K33")
  wb$add_pivot_table(df, rows = "cyl", data = "mpg")

  wb2 <- wb_workbook()
  expect_warning(
    expect_warning(
      wb2$clone_worksheet(2, "pivot", from = wb),
      "Cloning pivot tables over workbooks is not yet supported."
    ),
    "experimental"
  )
  expect_false(any(grepl("pivotTable", wb2$worksheets_rels[[1]], fixed = TRUE)))
})

test_that("cloning images into a workbook with images keeps them apart", {

  png1 <- tempfile(fileext = ".png")
  png2 <- tempfile(fileext = ".png")
  png3 <- tempfile(fileext = ".png")
  on.exit(unlink(c(png1, png2, png3)), add = TRUE)
  for (fl in c(png1, png2, png3)) {
    png(fl)
    plot(1:3)
    dev.off()
  }

  wb_old <- wb_workbook()$add_worksheet()$
    add_image(dims = "A1", file = png1)$
    add_image(dims = "A5", file = png2)

  wb <- wb_workbook()$add_worksheet()$add_image(dims = "A1", file = png3)
  wb$clone_worksheet(old = 1, new = "Clone1", from = wb_old)

  targets <- rbindlist(xml_attr(wb$drawings_rels[[2]], "Relationship"))$Target
  expect_equal(targets, c("../media/image2.png", "../media/image3.png"))
  expect_equal(names(wb$media), c("image1.png", "image2.png", "image3.png"))
})

test_that("cloning copies embedded objects and chart sheets", {

  fl <- testfile_path("mtcars_chart.xlsx")

  wb <- wb_load(fl)
  expect_warning(wb$clone_worksheet("Emb", "clone"), "experimental")
  expect_equal(basename(wb$embeddings), c("Microsoft_Word_Document.docx", "Microsoft_Word_Document1.docx"))
  expect_true(any(grepl("../embeddings/Microsoft_Word_Document1.docx", wb$worksheets_rels[[5]], fixed = TRUE)))

  # the ole object, its vml shape and the drawing use new shape ids
  expect_match(wb$worksheets[[4]]$oleObjects, 'shapeId="4097"', fixed = TRUE)
  expect_match(wb$worksheets[[5]]$oleObjects, 'shapeId="5121"', fixed = TRUE)
  expect_match(wb$vml[[2]], '<o:idmap v:ext="edit" data="5"/>', fixed = TRUE)
  expect_match(wb$vml[[2]], "_x0000_s5121", fixed = TRUE)
  expect_match(wb$drawings[[wb$worksheets[[5]]$relships$drawing]], "_x0000_s5121", fixed = TRUE)

  wb_from <- wb_load(fl)
  wb <- wb_workbook()
  for (sheet in wb_from$get_sheet_names()) {
    if (sheet == "Emb") {
      expect_warning(wb$clone_worksheet(sheet, sheet, from = wb_from), "embedded objects")
    } else {
      wb$clone_worksheet(sheet, sheet, from = wb_from)
    }
  }
  expect_equal(basename(wb$embeddings), "Microsoft_Word_Document.docx")
  expect_equal(names(wb$media), "image1.png")
  expect_true(wb$is_chartsheet[[1]])
  expect_true(any(grepl("Extension=\"docx\"", wb$Content_Types, fixed = TRUE)))

  # formulas referring to other workbooks keep their external references
  expect_equal(length(wb$externalLinks), 3L)
  expect_equal(wb$externalLinks, wb_from$externalLinks)
  expect_equal(sum(grepl("externalLinks/externalLink", wb$workbook.xml.rels, fixed = TRUE)), 3L)
  cc <- wb$worksheets[[3]]$sheet_data$cc
  exp <- c("VLOOKUP(B5,[1]SUM!$B$6:$C$6,2,FALSE)", "[2]IrisSample!$A$2", "'[3]1'!$A$2")
  expect_equal(cc$f[cc$f != ""], exp)

  # a second clone reuses the external references
  wb$clone_worksheet("Ext", "Ext2", from = wb_from)
  expect_equal(length(wb$externalLinks), 3L)

  # relationship ids are not always rIdN
  wb <- wb_workbook()
  wb$workbook.xml.rels <- sub("rId", "Rx", wb$workbook.xml.rels, fixed = TRUE)
  expect_silent(wb$clone_worksheet("Ext", "Ext", from = wb_from))
  expect_equal(length(wb$externalLinks), 3L)

  # different external references are added and the formulas updated
  wb <- wb_load(fl)
  wb$externalLinks[[1]] <- sub("SUM", "SUMX", wb$externalLinks[[1]])
  wb$clone_worksheet("Ext", "Ext2", from = wb_from)
  expect_equal(length(wb$externalLinks), 4L)
  cc <- wb$worksheets[[5]]$sheet_data$cc
  exp <- c("VLOOKUP(B5,[4]SUM!$B$6:$C$6,2,FALSE)", "[2]IrisSample!$A$2", "'[3]1'!$A$2")
  expect_equal(cc$f[cc$f != ""], exp)
})

test_that("cloned charts refer to the cloned sheet", {

  skip_if_not_installed("encharter")

  wb <- wb_workbook()$
    add_worksheet("S")$add_data(x = mtcars)$
    add_worksheet("other data")$add_data(x = mtcars)

  dat <- wb_data(wb, sheet = "S", dims = "A1:K33")
  wb$add_encharter(sheet = "S", dims = "M2", graph = encharter::ec("barplot")$add_series(data = dat, name = mpg, label = cyl))
  wb$add_encharter(sheet = "S", dims = "M20", graph = encharter::ec("barplot")$add_series(data = "other data!B2:B5"))

  wb$clone_worksheet("S", "my copy")

  refs <- function(x) unique(regmatches(x, gregexpr("<c:f>[^<]*</c:f>", x))[[1]])

  exp <- c("<c:f>'my copy'!$A$1</c:f>", "<c:f>'my copy'!$B$2:$B$33</c:f>", "<c:f>'my copy'!$A$2:$A$33</c:f>")
  expect_equal(refs(wb$charts$chart[3]), exp)

  exp <- "<c:f>'other data'!$B$2:$B$5</c:f>"
  expect_equal(refs(wb$charts$chart[4]), exp)

  exp <- c("<c:f>'S'!$A$1</c:f>", "<c:f>'S'!$B$2:$B$33</c:f>", "<c:f>'S'!$A$2:$A$33</c:f>")
  expect_equal(refs(wb$charts$chart[1]), exp)
})

test_that("cloning a chart and an image declares each extension once", {

  skip_if_not_installed("encharter")

  img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")
  bp  <- encharter::encharter("barplot")$add_series(data = "data!B2:B5")

  wb <- wb_workbook()$add_worksheet("data")$add_data(x = mtcars)$
    add_encharter(graph = bp, dims = "M2:S15")$
    add_image(dims = "M17", file = img)

  wb2 <- wb_workbook()$add_worksheet()
  wb2$clone_worksheet("data", "data", from = wb)
  wb2$clone_worksheet("data", "data2", from = wb)

  ext <- stringi::stri_match_first_regex(wb2$Content_Types, "<Default Extension=\"([^\"]+)\"")[, 2]
  ext <- ext[!is.na(ext)]
  expect_false(anyDuplicated(ext) > 0)
  expect_true("jpg" %in% ext)
})

test_that("cloning from another workbook carries dxf styles, media and cell styles", {

  img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")

  wb <- wb_workbook()$add_worksheet("cf")$add_data(x = data.frame(a = 1:3))$
    add_conditional_formatting(dims = "A2:A4", rule = ">1")$
    add_worksheet("img")$add_image(dims = "B2", file = img)$
    add_worksheet("other")$add_image(dims = "B2", file = system.file("extdata", "einstein.jpg", package = "openxlsx2"))
  # the named style adds a cell style xf that the target workbook does not have
  wb$add_named_style(name = "Note")
  wb$set_cell_style(sheet = "cf", dims = "A1", style = 1)
  expect_match(wb$styles_mgr$styles$cellXfs[2], "xfId=\"1\"")
  wb$add_formula(sheet = "cf", x = "SEQUENCE(2)", dims = "C1", cm = TRUE)

  wb2 <- wb_workbook()$add_worksheet()$add_fill(dims = "A1", color = wb_color("red"))
  wb2$add_conditional_formatting(dims = "A1", rule = ">5", style = "bad_style")

  expect_warning(wb2$clone_worksheet("cf", "cf", from = wb), "with cell metadata is experimental")
  cf  <- wb2$worksheets[[2]]$conditionalFormatting
  ids <- as.integer(regmatches(cf$cf, regexpr("(?<=dxfId=\")[0-9]+", cf$cf, perl = TRUE)))
  expect_true(all(ids < length(wb2$styles_mgr$styles$dxfs)))
  expect_equal(wb2$styles_mgr$styles$dxfs[ids + 1L], wb$styles_mgr$styles$dxfs[1])
  expect_true(all(wb2$worksheets[[2]]$sheet_data$cc$c_cm == ""))

  xf_ids <- as.integer(rbindlist(xml_attr(wb2$styles_mgr$styles$cellXfs, "xf"))$xfId)
  expect_true(all(is.na(xf_ids) | xf_ids < length(wb2$styles_mgr$styles$cellStyleXfs)))

  # only dxf styles used by the sheet are added, identical ones are reused
  n_dxf <- length(wb2$styles_mgr$styles$dxfs)
  wb2$clone_worksheet("img", "img", from = wb)
  expect_length(wb2$media, 1)
  expect_equal(length(wb2$styles_mgr$styles$dxfs), n_dxf)
  expect_warning(wb2$clone_worksheet("cf", "cf2", from = wb), "with cell metadata is experimental")
  expect_equal(length(wb2$styles_mgr$styles$dxfs), n_dxf)
})

test_that("cloning warns only about experimental elements", {
  wb <- wb_workbook()$add_worksheet("plain")$add_data(x = mtcars)$
    add_fill(dims = "A1:B2", color = wb_color("yellow"))$
    add_data_table(x = iris, dims = "M1")
  expect_silent(wb$clone_worksheet("plain", "plain copy"))

  wb$add_named_region(sheet = "plain", dims = "A1:A3", name = "first")
  expect_warning(wb$clone_worksheet("plain", "named copy"), "with named regions is experimental")
})

test_that("cloning warns about query tables, custom table styles and metadata", {
  wb <- wb_load(testfile_path("connection.xlsx"))
  expect_warning(wb$clone_worksheet(1, "query copy"), "query tables")

  src <- wb_load(testfile_path("tableStyles.xlsx"))
  wb  <- wb_workbook()$add_worksheet()
  expect_warning(wb$clone_worksheet(1, "styled", from = src), "with custom table styles is experimental")
  expect_silent(src$clone_worksheet(1, "styled copy"))

  src <- wb_load(testfile_path("formula.xlsx"))
  wb  <- wb_workbook()$add_worksheet()
  expect_warning(wb$clone_worksheet(1, "dynamic", from = src), "with cell metadata is experimental")
})

test_that("cloning warns about rich values only for sheets with rich values", {
  wb <- wb_load(testfile_path("richData_datatypes.xlsx"))
  wb$add_worksheet("plain")$add_data(x = 1)
  expect_silent(wb$clone_worksheet("plain", "plain copy"))
  expect_warning(wb$clone_worksheet(1, "rich copy"), "with rich values is experimental")
  cc <- wb$worksheets[[4]]$sheet_data$cc
  expect_equal(sum(nzchar(cc$c_vm)), 12L)
})
