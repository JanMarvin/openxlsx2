test_that("wb_clone_sheet_style", {

  # clone style to empty sheet (creates cells and style)
  fl <- system.file("extdata", "oxlsx2_sheet.xlsx", package = "openxlsx2")
  wb <- wb_load(fl)$add_worksheet("copy")
  wb <- wb_clone_sheet_style(wb, "SUM", "copy")
  expect_equal(dim(wb$worksheets[[1]]$sheet_data$cc),
               dim(wb$worksheets[[1]]$sheet_data$cc))
  expect_equal(dim(wb$worksheets[[1]]$sheet_data$row_attr),
               dim(wb$worksheets[[2]]$sheet_data$row_attr))

  # clone style to sheet with data
  fl <- system.file("extdata", "oxlsx2_sheet.xlsx", package = "openxlsx2")
  wb <- wb_load(fl)$add_worksheet("copy")$add_data(x = mtcars, startRow = 5, startCol = 2)
  wb <- wb_clone_sheet_style(wb, "SUM", "copy")
  expect_equal(c(36, 13), dim(wb$worksheets[[2]]$sheet_data$row_attr))

  # clone style on cloned and cleaned worksheet
  fl <- system.file("extdata", "oxlsx2_sheet.xlsx", package = "openxlsx2")
  wb <- wb_load(fl)$clone_worksheet("SUM", "clone")
  wb <- wb$clean_sheet(sheet = "clone", numbers = TRUE, characters = TRUE, styles = TRUE, merged_cells = FALSE)
  wb <- wb_clone_sheet_style(wb, "SUM", "clone")

  # sort for this test, does not matter later, because we will sort prior to saving
  ord <- match(
    wb$worksheets[[1]]$sheet_data$cc$r,
    wb$worksheets[[2]]$sheet_data$cc$r
  )

  expect_equal(
    wb$worksheets[[1]]$sheet_data$cc$c_s,
    wb$worksheets[[2]]$sheet_data$cc$c_s[ord]
  )

  # output if copying from empty sheet
  fl <- system.file("extdata", "oxlsx2_sheet.xlsx", package = "openxlsx2")
  wb <- wb_load(fl)$add_worksheet("copy")
  expect_message(
    expect_message(
      wb <- wb_clone_sheet_style(wb, "copy", "SUM"),
      "'from' has no sheet data styles to clone"
    ),
    "'from' has no row styles to clone"
  )

})


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
           "<border><left><color rgb=\"FF000000\"/></left><top><color rgb=\"FF000000\"/></top><bottom style=\"double\"><color rgb=\"FF000000\"/></bottom></border>",
           "<border><right><color rgb=\"FF000000\"/></right><top><color rgb=\"FF000000\"/></top><bottom style=\"double\"><color rgb=\"FF000000\"/></bottom></border>",
           "<border><top><color rgb=\"FF000000\"/></top><bottom style=\"double\"><color rgb=\"FF000000\"/></bottom></border>"
  )
  got <- wb$styles_mgr$styles$borders

  expect_equal(exp, got)


  wb <- wb_workbook()
  wb$add_worksheet("S1")$add_border(1, dims = "A1:K1", left_border = NULL, right_border = NULL, top_border = NULL, bottom_border = "double")

  exp <- c("1", "3", "3", "3", "3", "3", "3", "3", "3", "3", "2")
  got <- wb$worksheets[[1]]$sheet_data$cc$c_s
  expect_equal(exp, got)

})

test_that("test add_fill()", {

  wb <- wb_workbook()
  wb$add_worksheet("S1")$add_data("S1", mtcars)
  expect_silent(wb$add_fill("S1", dims = "D5:G6", color = wb_colour(hex = "FFFFFF00")))

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
  wb$add_fill("S1", dims = "A1:E6", color = wb_colour(hex = "FFFFFF00"), every_nth_col = 2)
  wb$add_fill("S1", dims = "A1:E6", color = wb_colour(hex = "FF00FF00"), every_nth_row = 2)

  exp <- c("<fill><patternFill patternType=\"none\"/></fill>",
           "<fill><patternFill patternType=\"gray125\"/></fill>",
           "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFFFF00\"/></patternFill></fill>",
           "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF00FF00\"/></patternFill></fill>"
  )
  got <- wb$styles_mgr$styles$fills

  expect_equal(exp, got)

  exp <- c("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>",
           "<xf applyFill=\"1\" borderId=\"0\" fillId=\"2\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>",
           "<xf applyFill=\"1\" borderId=\"0\" fillId=\"3\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>")
  got <- wb$styles_mgr$styles$cellXfs

  expect_equal(exp, got)

  # check the actual styles
  exp <- c("", "1", "", "1", "", "2", "2", "2", "2", "2", "", "1", "",
           "1", "", "2", "2", "2", "2", "2", "", "1", "", "1", "", "2",
           "2", "2", "2", "2")
  got <- wb$worksheets[[1]]$sheet_data$cc$c_s
  expect_equal(exp, got)


  wb <- wb_workbook()
  wb$add_worksheet("S1")$add_fill("S1", dims = "D5:G6", color = wb_colour(hex = "FFFFFF00"))

  exp <- rep("1", 8)
  got <- wb$worksheets[[1]]$sheet_data$cc$c_s
  expect_equal(exp, got)

})

test_that("test add_font()", {

  wb <- wb_workbook()
  wb$add_worksheet("S1")$add_data("S1", mtcars)
  expect_silent(wb$add_font("S1", dims = "A1:K1", color = wb_colour(hex = "FFFFFF00")))

  # check xf
  exp <- c(
    "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>",
    "<xf applyFont=\"1\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"0\" xfId=\"0\"/>"
  )
  got <- wb$styles_mgr$styles$cellXfs

  expect_equal(exp, got)


  # check font
  exp <- c(
    "<font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>",
    "<font><color rgb=\"FFFFFF00\"/><name val=\"Calibri\"/><sz val=\"11\"/></font>"
  )
  got <- wb$styles_mgr$styles$fonts

  expect_equal(exp, got)

  # check the actual styles
  exp <- c("1", "1", "1", "1", "1", "1", "1", "1", "1", "1", "1", "",
           "", "", "", "", "", "", "", "")
  got <- head(wb$worksheets[[1]]$sheet_data$cc$c_s, 20)
  expect_equal(exp, got)

  wb <- wb_workbook()
  wb$add_worksheet("S1")$add_font("S1", dims = "A1:K1", color = wb_colour(hex = "FFFFFF00"))

  exp <- rep("1", 11)
  got <- wb$worksheets[[1]]$sheet_data$cc$c_s
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


  wb <- wb_workbook()
  wb$add_worksheet("S1")$add_numfmt("S1", dims = "A1:A33", numfmt = 2)

  exp <- rep("1", 33)
  got <- wb$worksheets[[1]]$sheet_data$cc$c_s
  expect_equal(exp, got)

})

test_that("test add_cell_style()", {

  wb <- wb_workbook()
  wb$add_worksheet("S1")$add_data("S1", mtcars)
  expect_silent(wb$add_cell_style("S1", dims = "A1:A33", textRotation = "45"))
  expect_silent(wb$add_cell_style("S1", dims = "F1:F33", horizontal = "center"))

  # check xf
  exp <- c("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>",
           "<xf borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"><alignment textRotation=\"45\"/></xf>",
           "<xf borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"><alignment horizontal=\"center\"/></xf>"
  )
  got <- wb$styles_mgr$styles$cellXfs

  expect_equal(exp, got)

  # check the actual styles
  exp <- c("1", "", "", "", "", "2", "", "", "", "", "", "1", "", "",
           "", "", "2", "", "", "")
  got <- head(wb$worksheets[[1]]$sheet_data$cc$c_s, 20)
  expect_equal(exp, got)


  wb <- wb_workbook()
  wb$add_worksheet("S1")$add_cell_style("S1", dims = "A1:A33", textRotation = "45")

  exp <- rep("1", 33)
  got <- wb$worksheets[[1]]$sheet_data$cc$c_s
  expect_equal(exp, got)

  ###
  exp <- "<xf applyAlignment=\"1\" applyBorder=\"1\" applyFill=\"1\" applyFont=\"1\" applyNumberFormat=\"1\" applyProtection=\"1\" borderId=\"1\" fillId=\"1\" fontId=\"1\" numFmtId=\"1\" pivotButton=\"0\" quotePrefix=\"0\" xfId=\"1\"><alignment horizontal=\"1\" indent=\"1\" justifyLastLine=\"1\" readingOrder=\"1\" relativeIndent=\"1\" shrinkToFit=\"1\" textRotation=\"1\" vertical=\"1\" wrapText=\"1\"/><extLst extLst=\"1\"/><protection hidden=\"1\" locked=\"1\"/></xf>"
  got <- create_cell_style(
    borderId = "1",
    fillId = "1",
    fontId = "1",
    numFmtId = "1",
    pivotButton = "0",
    quotePrefix = "0",
    xfId = "1",
    horizontal = "1",
    indent = "1",
    justifyLastLine = "1",
    readingOrder = "1",
    relativeIndent = "1",
    shrinkToFit = "1",
    textRotation = "1",
    vertical = "1",
    wrapText = "1",
    extLst = "1",
    hidden = "1",
    locked = "1"
  )
  expect_equal(exp, got)

})

test_that("add_style", {

  # without name
  num <- create_numfmt(numFmtId = "165", formatCode = "#.#")
  wb <- wb_workbook() %>% wb_add_style(num)

  exp <- num
  got <- wb$styles_mgr$styles$numFmts
  expect_equal(exp, got)

  exp <- structure(list(typ = "numFmt", id = "165", name = "num"),
                   row.names = c(NA, -1L),
                   class = "data.frame")
  got <- wb$styles_mgr$numfmt
  expect_equal(exp, got)

  # with name
  wb <- wb_workbook() %>% wb_add_style(num, "num")

  exp <- num
  got <- wb$styles_mgr$styles$numFmts
  expect_equal(exp, got)

  exp <- structure(list(typ = "numFmt", id = "165", name = "num"),
                   row.names = c(NA, -1L),
                   class = "data.frame")
  got <- wb$styles_mgr$numfmt
  expect_equal(exp, got)

})

test_that("assigning styles to loaded workbook works", {

  wb <- wb_load(file = system.file("extdata", "oxlsx2_sheet.xlsx", package = "openxlsx2"))

  # previously it would break on xml_import, because NA was returned
  expect_silent(wb$add_font()$add_font())

})

test_that("get & set cell style(s)", {

  # set a style in b1
  wb <- wb_workbook()$add_worksheet()$
    add_numfmt(dims = "B1", numfmt = "#,0")

  # get style from b1 to assign it to a1
  numfmt <- wb$get_cell_style(dims = "B1")
  expect_equal("1", numfmt)

  # assign style to a1
  pre <- wb$get_cell_style(dims = "A1")
  expect_equal("", pre)

  expect_silent(wb$set_cell_style(dims = "A1", style = numfmt))

  post <- wb$get_cell_style(dims = "A1")
  expect_equal("1", post)

  s_a1_b1 <- wb$get_cell_style(dims = "A1:B1")
  expect_silent(wb$set_cell_style(dims = "A2:B2", style = s_a1_b1))
  s_a2_b2 <- wb$get_cell_style(dims = "A2:B2")
  expect_equal(s_a1_b1, s_a2_b2)

})

test_that("get_cell_styles()", {

  wb <- wb_workbook()$
    add_worksheet(gridLines = FALSE)$
    # add title
    add_data(dims = "B2", x = "MTCARS Title")$
    add_font(dims = "B2", bold = "1", size = "16")$           # style 1
    # add data
    add_data(x = head(mtcars), startCol = 2, startRow = 3)$
    add_fill(dims = "B3:L3", color = wb_colour("turquoise"))$ # style 2 unused
    add_font(dims = "B3:L3", color = wb_colour("white"))$     # style 3
    add_border(dims = "B9:L9",
               bottom_color = wb_colour(hex = "FF000000"),
               bottom_border = "thin",
               left_border = "",
               right_border = "",
               top_border = "")

  exp <- "1"
  got <- wb$get_cell_style(dims = "B2")
  expect_equal(exp, got)

  exp <- "<xf applyFont=\"1\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"0\" xfId=\"0\"/>"
  got <- get_cell_styles(wb, 1, "B2")
  expect_equal(exp, got)

  exp <- "3"
  got <- wb$get_cell_style(dims = "B3")
  expect_equal(exp, got)

  exp <- "<xf applyFill=\"1\" applyFont=\"1\" borderId=\"0\" fillId=\"2\" fontId=\"2\" numFmtId=\"0\" xfId=\"0\"/>"
  got <- get_cell_styles(wb, 1, "B3")
  expect_equal(exp, got)

  wb$add_cell_style(dims = "B3:L3",
                    textRotation = "45",
                    horizontal = "center",
                    vertical = "center",
                    wrapText = "1")

  exp <- "<xf applyFill=\"1\" applyFont=\"1\" borderId=\"0\" fillId=\"2\" fontId=\"2\" numFmtId=\"0\" xfId=\"0\"><alignment horizontal=\"center\" textRotation=\"45\" vertical=\"center\" wrapText=\"1\"/></xf>"
  got <- get_cell_styles(wb, 1, "B3")
  expect_equal(exp, got)

})

test_that("applyCellStyle works", {

  wb <- wb_workbook()$
    add_worksheet()$
    add_fill(dims = "B2:G8", color = wb_colour("yellow"))$
    add_data(dims = "C3", x = Sys.Date())$
    add_data(dims = "E3", x = Sys.Date(), applyCellStyle = FALSE)$
    add_data(dims = "E5", x = Sys.Date(), removeCellStyle = TRUE)$
    add_data(dims = "A1", x = Sys.Date())

  cc <- wb$worksheets[[1]]$sheet_data$cc
  exp <- c("3", "2", "1", "3")
  got <- cc[cc$r %in% c("A1", "C3", "E3", "E5"), "c_s"]
  expect_equal(exp, got)

})

test_that("style names are xml", {
  sheet <- mtcars[1:6, 1:6]

  wb <- wb_workbook() %>%
    wb_add_worksheet("test") %>%
    wb_add_data(x = "Title", startCol = 1, startRow = 1) %>%
    wb_add_font(dims = "A1", bold = "1", size = "14") %>%
    wb_add_data(x = sheet, colNames = TRUE, startCol = 1, startRow = 2, removeCellStyle = TRUE) %>%
    wb_add_cell_style(dims = "B2:F2", horizontal = "right") %>%
    wb_add_font(dims = "A2:F2", bold = "1", size = "11") %>%
    wb_add_numfmt(dims = "B3:D8", numfmt = 2) %>%
    wb_add_font(dims = "B3:D8", italic = "1", size = "11") %>%
    wb_add_fill(dims = "B3:D8", color = wb_colour("orange")) %>%
    wb_add_fill(dims = "C5", color = wb_colour("black")) %>%
    wb_add_font(dims = "C5", color = wb_colour("white"))

  exp <- list(
    numFmts = NULL,
    fonts = c(
      "<font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>",
      "<font><b val=\"1\"/><color rgb=\"FF000000\"/><name val=\"Calibri\"/><sz val=\"14\"/></font>",
      "<font><b val=\"1\"/><color rgb=\"FF000000\"/><name val=\"Calibri\"/><sz val=\"11\"/></font>",
      "<font><color rgb=\"FF000000\"/><i val=\"1\"/><name val=\"Calibri\"/><sz val=\"11\"/></font>",
      "<font><color rgb=\"FFFFFFFF\"/><name val=\"Calibri\"/><sz val=\"11\"/></font>"
    ),
    fills = c(
      "<fill><patternFill patternType=\"none\"/></fill>",
      "<fill><patternFill patternType=\"gray125\"/></fill>", "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFFA500\"/></patternFill></fill>",
      "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF000000\"/></patternFill></fill>"
    ),
    borders = "<border><left/><right/><top/><bottom/><diagonal/></border>",
    cellStyleXfs = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>",
    cellXfs = c(
      "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>",
      "<xf applyFont=\"1\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"0\" xfId=\"0\"/>",
      "<xf borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"><alignment horizontal=\"right\"/></xf>",
      "<xf applyFont=\"1\" borderId=\"0\" fillId=\"0\" fontId=\"2\" numFmtId=\"0\" xfId=\"0\"/>",
      "<xf applyFont=\"1\" borderId=\"0\" fillId=\"0\" fontId=\"2\" numFmtId=\"0\" xfId=\"0\"><alignment horizontal=\"right\"/></xf>",
      "<xf applyNumberFormat=\"1\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"2\" xfId=\"0\"/>",
      "<xf applyFont=\"1\" applyNumberFormat=\"1\" borderId=\"0\" fillId=\"0\" fontId=\"3\" numFmtId=\"2\" xfId=\"0\"/>",
      "<xf applyFill=\"1\" applyFont=\"1\" applyNumberFormat=\"1\" borderId=\"0\" fillId=\"2\" fontId=\"3\" numFmtId=\"2\" xfId=\"0\"/>",
      "<xf applyFill=\"1\" applyFont=\"1\" applyNumberFormat=\"1\" borderId=\"0\" fillId=\"3\" fontId=\"3\" numFmtId=\"2\" xfId=\"0\"/>",
      "<xf applyFill=\"1\" applyFont=\"1\" applyNumberFormat=\"1\" borderId=\"0\" fillId=\"3\" fontId=\"4\" numFmtId=\"2\" xfId=\"0\"/>"
    ),
    cellStyles = "<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>",
    dxfs = NULL,
    tableStyles = NULL,
    indexedColors = NULL,
    extLst = NULL
  )
  got <- wb$styles_mgr$styles
  expect_equal(exp, got)

})

test_that("add numfmt is no longer slow", {

  beg <- "1900-1-1"
  end <- "2022-11-18"

  dat <- seq(
    from = as.POSIXct(beg, tz = "UTC"),
    to   = as.POSIXct(end, tz = "UTC"),
    by   = "day"
  )

  # when writing this creates the 29Feb1900 (#421)
  out <- data.frame(
    date = dat,
    chr  = as.character(dat),
    num  = seq_along(dat) - 1
  )

  wb <- wb_workbook() %>%
    wb_add_worksheet()

  # just a tiny test to check that this does not run forever
  expect_silent(
    wb <- wb %>%
      wb_add_data(x = out) %>%
      wb_add_numfmt(dims = "C1:C44882", numfmt = "#.0")
  )

  expect_silent(
    wb_workbook()$add_worksheet()$
      add_fill(dims = "A1:Z10000", color = wb_color("yellow"))
  )

  mm <- matrix(NA, ncol = 26, nrow = 10000)
  expect_silent(
    wb_workbook()$add_worksheet()$
      add_data(x = mm, colNames = FALSE, na.strings = NULL)$
      add_fill(dims = "A1:Z10000", color = wb_color("yellow"))
  )

  expect_silent(
    wb_workbook()$add_worksheet()$
      add_data(dims = "A1", x = 1, colNames = FALSE)$
      add_data(dims = "A10000", x = 1, colNames = FALSE)$
      add_data(dims = "Z1", x = 5, colNames = FALSE)$
      add_data(dims = "Z10000", x = 5, colNames = FALSE)$
      add_fill(dims = "A1:Z10000", color = wb_color("yellow"))
  )

})

test_that("logical and numeric work too", {

  wb <- wb_workbook() %>% wb_add_worksheet("S1") %>% wb_add_data("S1", mtcars)

  wb2 <- wb %>% wb_add_font("S1", "A1:K1", name = "Arial", bold = "1", size = "14")

  wb3 <- wb %>% wb_add_font("S1", "A2:K2", name = "Arial", bold = TRUE, size = 14)

  expect_equal(
    wb2$styles_mgr$styles$fonts,
    wb3$styles_mgr$styles$fonts
  )

})

test_that("create_tablestyle() works", {

  exp <- "<tableStyle name=\"red_table\" pivot=\"0\" count=\"9\" xr9:uid=\"{CE23B8CA-E823-724F-9713-ASEVX1JWJGYG}\"><tableStyleElement type=\"wholeTable\" dxfId=\"8\"/><tableStyleElement type=\"headerRow\" dxfId=\"7\"/><tableStyleElement type=\"totalRow\" dxfId=\"6\"/><tableStyleElement type=\"firstColumn\" dxfId=\"5\"/><tableStyleElement type=\"lastColumn\" dxfId=\"4\"/><tableStyleElement type=\"firstRowStripe\" dxfId=\"3\"/><tableStyleElement type=\"secondRowStripe\" dxfId=\"2\"/><tableStyleElement type=\"firstColumnStripe\" dxfId=\"1\"/><tableStyleElement type=\"secondColumnStripe\" dxfId=\"0\"/></tableStyle>"
  set.seed(123)
  options("openxlsx2_seed" = NULL)
  got <- create_tablestyle(
    name               = "red_table",
    wholeTable         = 8,
    headerRow          = 7,
    totalRow           = 6,
    firstColumn        = 5,
    lastColumn         = 4,
    firstRowStripe     = 3,
    secondRowStripe    = 2,
    firstColumnStripe  = 1,
    secondColumnStripe = 0
  )
  expect_equal(exp, got)

})
