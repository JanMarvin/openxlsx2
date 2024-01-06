
test_that("Loading readTest.xlsx Sheet 1", {

  fl <- testfile_path("readTest.xlsx")
  wb <- wb_load(fl)

  sst <- wb$sharedStrings
  attr(sst, "text") <- NULL

  # in r/testthat-helpers.R
  expect_equal(expected_shared_strings(), sst)
})

test_that("Loading multiple pivot tables: loadPivotTables.xlsx works", {

  ## loadPivotTables.xlsx is a file with 3 pivot tables and 2 of them have the same reference data (pivotCacheDefinition)
  fl <- testfile_path("loadPivotTables.xlsx")
  wb <- wb_load(fl)

  # Check that wb is correctly loaded
  sheet_names <- c("iris",
                   "iris_pivot",
                   "penguins",
                   "penguins_pivot1",
                   "penguins_pivot2")

  expect_equal(wb$sheet_names, sheet_names)

  # Check number of 'pivotTables'
  expect_equal(length(wb$pivotTables),
               3)
  # Check number of 'pivotCacheDefinition'
  expect_equal(length(wb$pivotDefinitions),
               2)
})

test_that("Load and saving a file with Threaded Comments works", {

  tmp <- temp_xlsx()

  ## loadThreadComment.xlsx is a simple xlsx file that uses Threaded Comment.
  fl <- testfile_path("loadThreadComment.xlsx")
  expect_silent(wb <- wb_load(fl))

  exp <- "<t>[Threaded comment]\n\nYo"
  got <- substr(wb$comments[[1]][[1]]$comment, 1, 25)
  expect_equal(exp, got)

  # Check that wb can be saved without error
  expect_silent(wb_save(wb, file = tmp))
  wb <- wb_load(tmp)

  got <- substr(wb$comments[[1]][[1]]$comment, 1, 25)
  expect_equal(exp, got)

})

test_that("Read and save file with inlineStr", {

  ## loadThreadComment.xlsx is a simple xlsx file that uses Threaded Comment.
  fl <- testfile_path("inlineStr.xlsx")
  wb <- wb_load(fl)
  wb_df <- wb_read(wb)
  attr(wb_df, "tt") <- NULL
  attr(wb_df, "types") <- NULL

  df <- data.frame(
    this = c("is an xlsx file", "written with writexl::write_xlsx"),
    it = c("cannot be read", "with open.xlsx::read.xlsx"),
    stringsAsFactors = FALSE)
  rownames(df) <- c(2L, 3L)

  # compare file imported with inlineStr
  expect_equal(df, wb_df)

  df_read_xlsx <- read_xlsx(fl)
  attr(df_read_xlsx, "tt") <- NULL
  attr(df_read_xlsx, "types") <- NULL

  df_wb_read <- wb_read(fl)
  attr(df_wb_read, "tt") <- NULL
  attr(df_wb_read, "types") <- NULL

  expect_equal(df, df_read_xlsx)
  expect_equal(df, df_wb_read)

  tmp_xlsx <- temp_xlsx()
  # Check that wb can be saved without error and reimported
  expect_identical(tmp_xlsx, wb_save(wb, file = tmp_xlsx)$path)
  wb_df_re <- wb_read(wb_load(tmp_xlsx))
  attr(wb_df_re, "tt") <- NULL
  attr(wb_df_re, "types") <- NULL
  expect_equal(wb_df, wb_df_re)

})

# tests for getChildlessNode returns the content of every node, single node or not. the name has only historical meaning
test_that("read nodes", {

  # read single node
  test <- "<xf numFmtId=\"0\" fontId=\"4\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\"/></xf>"
  that <- xml_node(test, "xf")
  expect_equal(test, that)

  # real life example <foo/> and <foo>...</foo> mixed
  cellXfs <- "<cellXfs count=\"8\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/><xf numFmtId=\"0\" fontId=\"3\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/><xf numFmtId=\"0\" fontId=\"5\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/><xf numFmtId=\"0\" fontId=\"6\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/><xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/><xf numFmtId=\"0\" fontId=\"7\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/><xf numFmtId=\"0\" fontId=\"4\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\"/></xf><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>"
  that <- xml_node(cellXfs, "cellXfs", "xf")
  test <- c("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>",
            "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>",
            "<xf numFmtId=\"0\" fontId=\"3\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>",
            "<xf numFmtId=\"0\" fontId=\"5\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>",
            "<xf numFmtId=\"0\" fontId=\"6\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>",
            "<xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>",
            "<xf numFmtId=\"0\" fontId=\"7\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>",
            "<xf numFmtId=\"0\" fontId=\"4\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\"/></xf>",
            "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
  )
  expect_equal(test, that)

  # test <foos/>
  test <- "<xfs bla=\"0\"/>"
  that <- xml_node(test, "xf")
  expect_equal(character(0), that)

  # test <foo/>
  test <- "<b/><b/>"
  that <- xml_node(test, "b")
  test <- c(
    "<b/>",
    "<b/>"
  )
  expect_equal(test, that)

  # test <foo>...</foo>
  test <- "<b>a</b><b/>"
  that <- xml_node(test, "b")
  test <- c("<b>a</b>", "<b/>")
  expect_equal(test, that)

  # test <foos><foo/></foos>
  test <- "<xfs><xf/></xfs>"
  that <- xml_node(test, "xfs", "xf")
  test <- "<xf/>"
  expect_equal(test, that)

})

test_that("sheet visibility", {

  # example is rather slow (lots of hidden cols)
  fl <- testfile_path("ColorTabs3.xlsx")
  tmp_dir <- temp_xlsx()

  exp_sheets <- c("Nums", "Chars", "hidden")
  exp_vis <- c("visible", "visible", "hidden")

  # after load
  wb <- wb_load(fl)
  wb_sheets <- wb$get_sheet_names()
  wb_vis <- wb_get_sheet_visibility(wb)

  # save
  wb_save(wb, tmp_dir)

  # re-import
  wb2 <- wb_load(tmp_dir)
  wb2_sheets <- wb$get_sheet_names()
  wb2_vis <- wb_get_sheet_visibility(wb)

  expect_equal(exp_sheets, names(wb_sheets))
  expect_equal(exp_vis, wb_vis)

  expect_equal(exp_sheets, names(wb2_sheets))
  expect_equal(exp_vis, wb2_vis)
})


test_that("additional wb tests", {

  # no data on sheet
  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")

  expect_message(expect_null(wb_to_df(wb, sheet = "Sheet 1")), "no data")

  # wb_to_df
  xlsxFile <- testfile_path("readTest.xlsx")
  wb1 <- wb_load(xlsxFile)

  # showFormula
  exp <- data.frame(Var7 = "1/0", row.names = "2")
  got <- wb_to_df(wb1, showFormula = TRUE, rows = 1:2, cols = 8)
  expect_equal(exp, got, ignore_attr = TRUE)
  expect_equal(names(exp), names(got))

  # detectDates
  exp <- data.frame(Var5 = as.Date("2015-02-07"), row.names = "2")
  got <- wb_to_df(wb1, showFormula = TRUE, rows = 1:2, cols = 6)
  expect_equal(exp, got, ignore_attr = TRUE)
  expect_equal(names(exp), names(got))

  # types
  # Var1 is requested as character
  exp <- data.frame(Var1 = c("TRUE", "TRUE", "TRUE", "FALSE"),
                    Var3 = c(1.00, NaN, 1.34, NA))
  got <- wb_to_df(wb1, cols = c(1, 4),
                  types = c("Var1" = 0, "Var3" = 1))[seq_len(4), ]
  expect_equal(exp, got, ignore_attr = TRUE)
  expect_equal(names(exp), names(got))
})

test_that("test headerFooter", {

  # Plain text headers and footers
  header <- c('ODD HEAD LEFT', 'ODD HEAD CENTER', 'ODD HEAD RIGHT')
  footer <- c('ODD FOOT RIGHT', 'ODD FOOT CENTER', 'ODD FOOT RIGHT')
  evenHeader <- c('EVEN HEAD LEFT', 'EVEN HEAD CENTER', 'EVEN HEAD RIGHT')
  evenFooter <- c('EVEN FOOT RIGHT', 'EVEN FOOT CENTER', 'EVEN FOOT RIGHT')
  firstHeader <- c('TOP', 'OF FIRST', 'PAGE')
  firstFooter <- c('BOTTOM', 'OF FIRST', 'PAGE')

  # Add Sheet 1
  wb <- wb_workbook()
  wb$add_worksheet(
    'Sheet 1',
    header = header,
    footer = footer,
    evenHeader = evenHeader,
    evenFooter = evenFooter,
    firstHeader = firstHeader,
    firstFooter = firstFooter
  )

  # Modified headers and footers to make them Arial 8
  header <- paste0('&"Arial"&8', header)
  footer <- paste0('&"Arial"&8', footer)
  evenHeader <- paste0('&"Arial"&8', evenHeader)
  evenFooter <- paste0('&"Arial"&8', evenFooter)
  firstHeader <- paste0('&"Arial"&8', firstHeader)
  firstFooter <- paste0('&"Arial"&8', firstFooter)

  # Add Sheet 2
  wb$add_worksheet(
    'Sheet 2',
    header = header,
    footer = footer,
    evenHeader = evenHeader,
    evenFooter = evenFooter,
    firstHeader = firstHeader,
    firstFooter = firstFooter
  )
  wb$add_data(sheet = 1, 1:400)
  wb$add_data(sheet = 2, 1:400)

  tmp1 <- temp_xlsx()
  # Save workbook
  wb_save(wb, tmp1, overwrite = TRUE)
  # Load workbook and save again
  wb2 <- wb_load(tmp1)

  expect_equal(lapply(wb$worksheets[[1]]$headerFooter, unlist),
               wb2$worksheets[[1]]$headerFooter)

  expect_equal(lapply(wb$worksheets[[2]]$headerFooter, unlist),
               wb2$worksheets[[2]]$headerFooter)

})


test_that("load workbook with chartsheet", {

  fl <- testfile_path("mtcars_chart.xlsx")

  expect_silent(z <- wb_load(fl))
  expect_silent(z <- wb_load(fl, sheet = "Chart1"))
  expect_silent(z <- wb_load(fl, sheet = "test"))
  # explicitly request the chartsheet
  expect_silent(z <- wb_load(fl, sheet = 1))
  expect_silent(z <- wb_load(fl, sheet = 2))

  expect_equal(read_xlsx(fl, sheet = "test"), mtcars, ignore_attr = TRUE)
  expect_equal(read_xlsx(fl, sheet = 2), mtcars, ignore_attr = TRUE)

  # sheet found, but contains no data
  expect_error(read_xlsx(fl, sheet = "Chart1"), "Requested sheet is a chartsheet. No data to return")
  expect_error(read_xlsx(fl, sheet = 1), "Requested sheet is a chartsheet. No data to return")
})


test_that("Content Types is not modified", {

  # Content Types should remain identical during saving. All modifications should remain
  # temporary because otherwise they are applied over and over and over again during saving
  wb <- wb_load(file = testfile_path("loadExample.xlsx"))
  pre <- wb$Content_Types
  wb$save(temp_xlsx())
  post <- wb$Content_Types
  expect_equal(pre, post)

})

test_that("Sheet not found", {
  temp <- temp_xlsx()

  wb <- wb_workbook()$
    add_worksheet("Test")$
    add_worksheet("Test ")$
    add_worksheet("Test & Test")
  wb$save(temp)

  expect_error(
    read_xlsx(temp, "Tes"),
    paste0("No such sheet in the workbook. Workbook contains:\nTest\nTest \nTest & Test")
  )

})

test_that("loading slicers works", {

  fl <- testfile_path("loadExample.xlsx")
  wb <- wb_load(file = fl, calc_chain = TRUE)

  exp <- "<calcPr calcId=\"152511\" fullCalcOnLoad=\"1\"/>"
  got <- wb$workbook$calcPr
  expect_equal(exp, got)

  exp <- c(
    "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>",
    "<Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>",
    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>",
    "<Relationship Id=\"rId0\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>",
    "<Relationship Id=\"rId0\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>",
    "<Relationship Id=\"rId0\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet3.xml\"/>",
    "<Relationship Id=\"rId0\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet4.xml\"/>",
    "<Relationship Id=\"rId8\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain\" Target=\"calcChain.xml\"/>",
    "<Relationship Id=\"rId20001\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition\" Target=\"pivotCache/pivotCacheDefinition1.xml\"/>",
    "<Relationship Id=\"rId20002\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition\" Target=\"pivotCache/pivotCacheDefinition2.xml\"/>",
    "<Relationship Id=\"rId100001\" Type=\"http://schemas.microsoft.com/office/2007/relationships/slicerCache\" Target=\"slicerCaches/slicerCache1.xml\"/>"
  )
  got <- wb$workbook.xml.rels
  expect_equal(exp, got)

  wb <- wb_load(file = fl, calc_chain = FALSE)
  got <- wb$workbook.xml.rels
  expect_equal(exp[-8], got)


  op <- options("openxlsx2.disableFullCalcOnLoad" = TRUE)
  on.exit(options(op), add = TRUE)

  wb <- wb_load(file = fl, calc_chain = TRUE)

  exp <- "<calcPr calcId=\"152511\"/>"
  got <- wb$workbook$calcPr
  expect_equal(exp, got)

  wb <- wb_load(file = fl, calc_chain = FALSE)

  got <- wb$workbook$calcPr
  expect_null(got)

  exp <- character()
  got <- wb$calcChain
  expect_equal(exp, got)

  # check the default once again
  wb <- wb_load(file = fl)

  got <- wb$workbook$calcPr
  expect_null(got)

  exp <- character()
  got <- wb$calcChain
  expect_equal(exp, got)

})

test_that("vml target is updated on load", {

  fl <- testfile_path("mtcars_chart.xlsx")
  wb <- wb_load(fl)

  exp <- "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" Target=\"../drawings/vmlDrawing1.vml\"/>"
  got <- wb$worksheets_rels[[4]][2]
  expect_equal(exp, got)

})

test_that("sheetView is not switched", {

  wb <- wb_load(file = testfile_path("loadExample.xlsx"))

  exp <- "<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>"
  got <- wb$worksheets[[1]]$sheetViews
  expect_equal(exp, got)

  exp <- "<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews>"
  got <- wb$worksheets[[2]]$sheetViews
  expect_equal(exp, got)

})

test_that("Loading a workbook with property preserves it.", {
  wb <- wb_workbook(title = "x", creator = "y", subject = "z", category = "aa", keywords = "ab", comments = "ac", manager = "ad", company = "ae")$add_worksheet()
  tmp <- temp_xlsx()
  wb$save(file = tmp)

  wb2 <- wb_load(tmp)
  exp <- c(
    `title` = "x", `subject` = "z", `creator` = "y", `keywords` = "ab",
    `comments` = "ac",
    `modifier` = "y", `category` = "aa",
    manager = "ad", company = "ae"
  )
  sel <- names(exp) # ignore creation date
  got <- wb2$get_properties()
  expect_equal(exp, got[sel])

  wb2$set_properties(title = "xyz")
  expect_equal(wb2$get_properties()[["title"]], "xyz")

  wb2$set_properties(subject = "aaa")
  expect_equal(wb2$get_properties()[["subject"]], "aaa")

  wb2$set_properties(creator = "bbb")
  expect_equal(wb2$get_properties()[["creator"]], "bbb")

  wb2$set_properties(keywords = "ccc")
  expect_equal(wb2$get_properties()[["keywords"]], "ccc")

  wb2$set_properties(comments = "ddd")
  expect_equal(wb2$get_properties()[["comments"]], "ddd")

  wb2$set_properties(category = "eee")
  expect_equal(wb2$get_properties()[["category"]], "eee")

  wb2$set_properties(manager = "fff")
  expect_equal(wb2$get_properties()[["manager"]], "fff")

  wb2$set_properties(company = "ggg")
  expect_equal(wb2$get_properties()[["company"]], "ggg")
})

test_that("failing to unzip works as expected", {
  tmp <- temp_xlsx()
  writeLines("", tmp)
  expect_error(wb <- wb_load(tmp), "Unable to open and load file")

  wb_workbook()$add_worksheet()$save(tmp)
  expect_silent(wb <- wb_load(tmp))
})
