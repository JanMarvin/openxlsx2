test_that("test add_border()", {

  wb <- wb_workbook()
  wb$add_worksheet("S1")$add_data("S1", mtcars)
  expect_silent(wb$add_border(1, dims = "A1:K1", left_border = NULL, right_border = NULL, top_border = NULL, bottom_border = "double"))

  # check xf
  exp <- c("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>",
           "<xf applyBorder=\"1\" borderId=\"1\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>",
           "<xf applyBorder=\"1\" borderId=\"2\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>",
           "<xf applyBorder=\"1\" borderId=\"3\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>",
           "<xf applyBorder=\"1\" borderId=\"3\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>",
           "<xf applyBorder=\"1\" borderId=\"3\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>",
           "<xf applyBorder=\"1\" borderId=\"3\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>",
           "<xf applyBorder=\"1\" borderId=\"3\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>",
           "<xf applyBorder=\"1\" borderId=\"3\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>",
           "<xf applyBorder=\"1\" borderId=\"3\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>",
           "<xf applyBorder=\"1\" borderId=\"3\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>",
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

test_that("test add_border()", {

  wb <- wb_workbook()
  wb$add_worksheet("S1")$add_data("S1", mtcars)
  expect_silent(wb$add_fill("S1", dims = "D5:G6", color = c(rgb = "FFFFFF00")))

  # check xf
  exp <- c("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>",
           "<xf applyFill=\"1\" borderId=\"0\" fillId=\"2\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>",
           "<xf applyFill=\"1\" borderId=\"0\" fillId=\"2\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>",
           "<xf applyFill=\"1\" borderId=\"0\" fillId=\"2\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>",
           "<xf applyFill=\"1\" borderId=\"0\" fillId=\"2\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>",
           "<xf applyFill=\"1\" borderId=\"0\" fillId=\"2\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>",
           "<xf applyFill=\"1\" borderId=\"0\" fillId=\"2\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>",
           "<xf applyFill=\"1\" borderId=\"0\" fillId=\"2\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>",
           "<xf applyFill=\"1\" borderId=\"0\" fillId=\"2\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>"
  )
  got <- wb$styles_mgr$styles$cellXfs

  expect_equal(exp, got)


  # check fill
  exp <- c("<fill><patternFill patternType=\"none\"/></fill>", "<fill><patternFill patternType=\"gray125\"/></fill>",
           "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFFFF00\"/></patternFill></fill>"
  )
  got <- wb$styles_mgr$styles$fills

  expect_equal(exp, got)

})

