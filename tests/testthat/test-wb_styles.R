test_that("test add_border()", {

  wb <- wb_workbook()
  wb$add_worksheet("S1")$add_data("S1", mtcars)
  expect_silent(wb$add_border(1, dims = "A1:K1", left_border = NULL, right_border = NULL, top_border = NULL, bottom_border = "double"))

  # check xf
  exp <- c("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>",
           "<xf applyBorder=\"1\" borderId=\"1\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>",
           "<xf applyBorder=\"1\" borderId=\"2\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>",
           "<xf applyBorder=\"1\" borderId=\"3\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>"
  )
  got <- wb$styles_mgr$styles$cellXfs

  expect_equal(exp, got)

  # check borders
  exp <- c("<border><left/><right/><top/><bottom/><diagonal/></border>",
           "<border><left><color rgb=\"FF000000\"/></left><right/><top><color rgb=\"FF000000\"/></top><bottom style=\"double\"><color rgb=\"FF000000\"/></bottom><diagonal/></border>",
           "<border><left/><right><color rgb=\"FF000000\"/></right><top><color rgb=\"FF000000\"/></top><bottom style=\"double\"><color rgb=\"FF000000\"/></bottom><diagonal/></border>",
           "<border><left/><right/><top><color rgb=\"FF000000\"/></top><bottom style=\"double\"><color rgb=\"FF000000\"/></bottom><diagonal/></border>"
  )
  got <- wb$styles_mgr$styles$borders

  expect_equal(exp, got)

})

test_that("test add_fill()", {

  wb <- wb_workbook()
  wb$add_worksheet("S1")$add_data("S1", mtcars)
  expect_silent(wb$add_fill("S1", dims = "D5:G6", color = c(rgb = "FFFFFF00")))

  # check xf
  exp <- c("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>",
           "<xf applyFill=\"1\" borderId=\"0\" fillId=\"2\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>"
  )
  got <- wb$styles_mgr$styles$cellXfs

  expect_equal(exp, got)


  # check fill
  exp <- c("<fill><patternFill patternType=\"none\"/></fill>",
           "<fill><patternFill patternType=\"gray125\"/></fill>",
           "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFFFF00\"/></patternFill></fill>"
  )
  got <- wb$styles_mgr$styles$fills

  expect_equal(exp, got)

  # every_nth_col/row
  wb <- wb_workbook()
  wb$add_worksheet("S1", gridLines = FALSE)$add_data("S1", matrix(1:20, 4, 5))
  wb$add_fill("S1", dims = "A1:E6", color = c(rgb = "FFFFFF00"), every_nth_col = 2)
  wb$add_fill("S1", dims = "A1:E6", color = c(rgb = "FF00FF00"), every_nth_row = 2)

  exp <- c("<fill><patternFill patternType=\"none\"/></fill>",
           "<fill><patternFill patternType=\"gray125\"/></fill>",
           "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFFFF00\"/></patternFill></fill>",
           "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF00FF00\"/></patternFill></fill>"
  )
  got <- wb$styles_mgr$styles$fills

  expect_equal(exp, got)

  exp <- c("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>",
           "<xf applyFill=\"1\" borderId=\"0\" fillId=\"2\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>",
           "<xf applyFill=\"1\" fillId=\"2\"/>",
           "<xf applyFill=\"1\" borderId=\"0\" fillId=\"3\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>",
           "<xf applyFill=\"1\" fillId=\"3\"/>")
  got <- wb$styles_mgr$styles$cellXfs

  expect_equal(exp, got)

  # check the actual styles
  exp <- c("", "1", "", "1", "", "3", "3", "3", "3", "3", "", "1", "",
           "1", "", "3", "3", "3", "3", "3", "", "1", "", "1", "")
  got <- wb$worksheets[[1]]$sheet_data$cc$c_s
  expect_equal(exp, got)


})

test_that("test add_font()", {

  wb <- wb_workbook()
  wb$add_worksheet("S1")$add_data("S1", mtcars)
  expect_silent(wb$add_font("S1", dims = "A1:K1", color = c(rgb = "FFFFFF00")))

  # check xf
  exp <- c("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>",
           "<xf applyFont=\"1\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"0\" xfId=\"0\"/>"
  )
  got <- wb$styles_mgr$styles$cellXfs

  expect_equal(exp, got)


  # check font
  exp <- c("<font><sz val=\"11\"/><color rgb=\"FF000000\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>",
           "<font><color rgb=\"FFFFFF00\"/><name val=\"Calibri\"/><sz val=\"11\"/></font>"
  )
  got <- wb$styles_mgr$styles$fonts

  expect_equal(exp, got)

  # check the actual styles
  exp <- c("1", "1", "1", "1", "1", "1", "1", "1", "1", "1", "1", "",
           "", "", "", "", "", "", "", "")
  got <- head(wb$worksheets[[1]]$sheet_data$cc$c_s, 20)
  expect_equal(exp, got)

})

test_that("test add_numfmt()", {

  wb <- wb_workbook()
  wb$add_worksheet("S1")$add_data("S1", mtcars)
  expect_silent(wb$add_numfmt("S1", dims = "A1:A33", numfmt = 2))
  expect_silent(wb$add_numfmt("S1", dims = "F1:F33", numfmt = "#.0"))

  # check xf
  exp <- c("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>",
           "<xf applyNumberFormat=\"1\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"2\" xfId=\"0\"/>",
           "<xf applyNumberFormat=\"1\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"165\" xfId=\"0\"/>"
  )
  got <- wb$styles_mgr$styles$cellXfs

  expect_equal(exp, got)


  # check numfmt
  exp <- "<numFmt numFmtId=\"165\" formatCode=\"#.0\"/>"
  got <- wb$styles_mgr$styles$numFmts

  expect_equal(exp, got)

  # check the actual styles
  exp <- c("1", "", "", "", "", "2", "", "", "", "", "", "1", "", "",
           "", "", "2", "", "", "")
  got <- head(wb$worksheets[[1]]$sheet_data$cc$c_s, 20)
  expect_equal(exp, got)

})

