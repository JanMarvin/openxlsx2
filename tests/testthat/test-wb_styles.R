testsetup()

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
  wb <- wb_load(fl)$add_worksheet("copy")$add_data(x = mtcars, start_row = 5, start_col = 2)
  wb <- wb_clone_sheet_style(wb, "SUM", "copy")
  expect_equal(c(36, 13), dim(wb$worksheets[[2]]$sheet_data$row_attr))

  # clone style on cloned and cleaned worksheet
  fl <- system.file("extdata", "oxlsx2_sheet.xlsx", package = "openxlsx2")
  wb <- wb_load(fl)$clone_worksheet("SUM", "clone")
  wb <- wb_clean_sheet(wb, sheet = "clone", numbers = TRUE, characters = TRUE, styles = TRUE, merged_cells = FALSE)
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
           "<border><bottom style=\"double\"><color rgb=\"FF000000\"/></bottom></border>",
           "<border><bottom style=\"double\"><color rgb=\"FF000000\"/></bottom></border>",
           "<border><bottom style=\"double\"><color rgb=\"FF000000\"/></bottom></border>"
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
  wb$add_worksheet("S1", grid_lines = FALSE)$add_data("S1", matrix(1:20, 4, 5))
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
    "<font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Aptos Narrow\"/><family val=\"2\"/><scheme val=\"minor\"/></font>",
    "<font><color rgb=\"FFFFFF00\"/><name val=\"Aptos Narrow\"/><sz val=\"11\"/></font>"
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
  expect_silent(wb$add_cell_style("S1", dims = "A1:A33", text_rotation = "45"))
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
  wb$add_worksheet("S1")$add_cell_style("S1", dims = "A1:A33", text_rotation = "45")

  exp <- rep("1", 33)
  got <- wb$worksheets[[1]]$sheet_data$cc$c_s
  expect_equal(exp, got)

  ###
  exp <- "<xf applyAlignment=\"1\" applyBorder=\"1\" applyFill=\"1\" applyFont=\"1\" applyNumberFormat=\"1\" applyProtection=\"1\" borderId=\"1\" fillId=\"1\" fontId=\"1\" numFmtId=\"1\" pivotButton=\"0\" quotePrefix=\"0\" xfId=\"1\"><alignment horizontal=\"left\" indent=\"1\" justifyLastLine=\"1\" readingOrder=\"1\" relativeIndent=\"1\" shrinkToFit=\"1\" textRotation=\"1\" vertical=\"top\" wrapText=\"1\"/><protection hidden=\"1\" locked=\"1\"/><extLst><ext><foo/></ext></extLst></xf>"
  got <- create_cell_style(
    border_id = "1",
    fill_id = "1",
    font_id = "1",
    num_fmt_id = "1",
    pivot_button = "0",
    quote_prefix = "0",
    xf_id = "1",
    horizontal = "left",
    indent = "1",
    justify_last_line = "1",
    reading_order = "1",
    relative_indent = "1",
    shrink_to_fit = "1",
    text_rotation = "1",
    vertical = "top",
    wrap_text = "1",
    ext_lst = "<extLst><ext><foo/></ext></extLst>",
    hidden = "1",
    locked = "1"
  )
  expect_equal(exp, got)

})

test_that("add_style", {

  # without name
  num <- create_numfmt(numFmtId = "165", formatCode = "#.#")
  wb <- wb_workbook()$add_style(num)

  exp <- num
  got <- wb$styles_mgr$styles$numFmts
  expect_equal(exp, got)

  exp <- structure(list(typ = "numFmt", id = "165", name = "num"),
                   row.names = c(NA, -1L),
                   class = "data.frame")
  got <- wb$styles_mgr$numfmt
  expect_equal(exp, got)

  # with name
  wb <- wb_workbook()$add_style(num, "num")

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
  expect_equal(numfmt, c(B1 = "1"))

  # assign style to a1
  pre <- wb$get_cell_style(dims = "A1")
  expect_equal(pre, c(A1 = ""))

  expect_silent(wb$set_cell_style(dims = "A1", style = numfmt))

  post <- wb$get_cell_style(dims = "A1")
  expect_equal(post, c(A1 = "1"))

  s_a1_b1 <- wb$get_cell_style(dims = "A1:B1")
  expect_silent(wb$set_cell_style(dims = "A2:B2", style = s_a1_b1))
  s_a2_b2 <- wb$get_cell_style(dims = "A2:B2")
  expect_equal(unname(s_a1_b1), unname(s_a2_b2))

})

test_that("get_cell_styles()", {

  wb <- wb_workbook()$
    add_worksheet(grid_lines = FALSE)$
    # add title
    add_data(dims = "B2", x = "MTCARS Title")$
    add_font(dims = "B2", bold = "1", size = "16")$           # style 1
    # add data
    add_data(x = head(mtcars), start_col = 2, start_row = 3)$
    add_fill(dims = "B3:L3", color = wb_colour("turquoise"))$ # style 2 unused
    add_font(dims = "B3:L3", color = wb_colour("white"))$     # style 3
    add_border(dims = "B9:L9",
               bottom_color = wb_colour(hex = "FF000000"),
               bottom_border = "thin",
               left_border = "",
               right_border = "",
               top_border = "")

  got <- wb$get_cell_style(dims = "B2")
  expect_equal(got, c(B2 = "1"))

  exp <- "<xf applyFont=\"1\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"0\" xfId=\"0\"/>"
  got <- get_cell_styles(wb, 1, "B2")
  expect_equal(got, exp)

  got <- wb$get_cell_style(dims = "B3")
  expect_equal(got, c(B3 = "3"))

  exp <- "<xf applyFill=\"1\" applyFont=\"1\" borderId=\"0\" fillId=\"2\" fontId=\"2\" numFmtId=\"0\" xfId=\"0\"/>"
  got <- get_cell_styles(wb, 1, "B3")
  expect_equal(got, exp)

  wb$add_cell_style(dims = "B3:L3",
                    text_rotation = "45",
                    horizontal = "center",
                    vertical = "center",
                    wrap_text = "1")

  exp <- "<xf applyFill=\"1\" applyFont=\"1\" borderId=\"0\" fillId=\"2\" fontId=\"2\" numFmtId=\"0\" xfId=\"0\"><alignment horizontal=\"center\" textRotation=\"45\" vertical=\"center\" wrapText=\"1\"/></xf>"
  got <- get_cell_styles(wb, 1, "B3")
  expect_equal(got, exp)

})

test_that("applyCellStyle works", {

  wb <- wb_workbook()$
    add_worksheet()$
    add_fill(dims = "B2:G8", color = wb_colour("yellow"))$
    add_data(dims = "C3", x = Sys.Date())$
    add_data(dims = "E3", x = Sys.Date(), apply_cell_style = FALSE)$
    add_data(dims = "E5", x = Sys.Date(), remove_cell_style = TRUE)$
    add_data(dims = "A1", x = Sys.Date())

  cc <- wb$worksheets[[1]]$sheet_data$cc
  exp <- c("3", "2", "1", "3")
  got <- cc[cc$r %in% c("A1", "C3", "E3", "E5"), "c_s"]
  expect_equal(exp, got)

})

test_that("style names are xml", {
  sheet <- mtcars[1:6, 1:6]

  wb <- wb_workbook()$
    add_worksheet("test")$
    add_data(x = "Title", start_col = 1, start_row = 1)$
    add_font(dims = "A1", bold = "1", size = "14")$
    add_data(x = sheet, col_names = TRUE, start_col = 1, start_row = 2, remove_cell_style = TRUE)$
    add_cell_style(dims = "B2:F2", horizontal = "right")$
    add_font(dims = "A2:F2", bold = "1", size = "11")$
    add_numfmt(dims = "B3:D8", numfmt = 2)$
    add_font(dims = "B3:D8", italic = "1", size = "11")$
    add_fill(dims = "B3:D8", color = wb_colour("orange"))$
    add_fill(dims = "C5", color = wb_colour("black"))$
    add_font(dims = "C5", color = wb_colour("white"))

  exp <- list(
    numFmts = NULL,
    fonts = c(
      "<font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Aptos Narrow\"/><family val=\"2\"/><scheme val=\"minor\"/></font>",
      "<font><b val=\"1\"/><color rgb=\"FF000000\"/><name val=\"Aptos Narrow\"/><sz val=\"14\"/></font>",
      "<font><b val=\"1\"/><color rgb=\"FF000000\"/><name val=\"Aptos Narrow\"/><sz val=\"11\"/></font>",
      "<font><color rgb=\"FF000000\"/><i val=\"1\"/><name val=\"Aptos Narrow\"/><sz val=\"11\"/></font>",
      "<font><color rgb=\"FFFFFFFF\"/><name val=\"Aptos Narrow\"/><sz val=\"11\"/></font>"
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
    num  = seq_along(dat) - 1,
    stringsAsFactors = FALSE
  )

  wb <- wb_workbook()$
    add_worksheet()

  # just a tiny test to check that this does not run forever
  expect_silent(
    wb$
      add_data(x = out)$
      add_numfmt(dims = "C1:C44882", numfmt = "#.0")
  )

  expect_silent(
    wb_workbook()$add_worksheet()$
      add_fill(dims = "A1:Z10000", color = wb_color("yellow"))
  )

  mm <- matrix(NA, ncol = 26, nrow = 10000)
  expect_silent(
    wb_workbook()$add_worksheet()$
      add_data(x = mm, col_names = FALSE, na.strings = NULL)$
      add_fill(dims = "A1:Z10000", color = wb_color("yellow"))
  )

  expect_silent(
    wb_workbook()$add_worksheet()$
      add_data(dims = "A1", x = 1, col_names = FALSE)$
      add_data(dims = "A10000", x = 1, col_names = FALSE)$
      add_data(dims = "Z1", x = 5, col_names = FALSE)$
      add_data(dims = "Z10000", x = 5, col_names = FALSE)$
      add_fill(dims = "A1:Z10000", color = wb_color("yellow"))
  )

})

test_that("logical and numeric work too", {

  wb <- wb_workbook()$add_worksheet("S1")$add_data("S1", mtcars)

  wb2 <- wb_add_font(wb, "S1", "A1:K1", name = "Arial", bold = "1", size = "14")

  wb3 <- wb_add_font(wb, "S1", "A2:K2", name = "Arial", bold = TRUE, size = 14)

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
    whole_table         = 8,
    header_row          = 7,
    total_row           = 6,
    first_column        = 5,
    last_column         = 4,
    first_row_stripe     = 3,
    second_row_stripe    = 2,
    first_column_stripe  = 1,
    second_column_stripe = 0
  )
  expect_equal(exp, got)

})

test_that("wb_add_cell_style works with logical and numeric", {

  wb <- wb_workbook()$add_worksheet()

  wb$add_cell_style(wrap_text = TRUE, text_rotation = 45)

  exp <- "<xf borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"><alignment textRotation=\"45\" wrapText=\"1\"/></xf>"
  got <- wb$styles_mgr$styles$cellXfs[2]
  expect_equal(exp, got)

})

test_that("wb_add_named_style() works", {

  wb <- wb_workbook()$
    add_worksheet()$
    add_named_style(dims = "A1", name = "Title")$
    add_data(dims = "A1", x = "Title")

  exp <- "<cellStyle name=\"Title\" xfId=\"1\" builtinId=\"15\"/>"
  got <- wb$styles_mgr$styles$cellStyles[2]
  expect_equal(exp, got)

  wb <- wb_workbook()$
    add_worksheet()$
    add_named_style(dims = "A1", name = "Good")$
    add_data(dims = "A1", x = "Title")$
    add_named_style(dims = "A1", name = "Good")

  exp <- 2
  got <- length(wb$styles_mgr$styles$cellStyles)
  expect_equal(exp, got)


  ### run all named styles for coverage
  wb <- wb_workbook()$add_worksheet()

  name <- "Normal"
  dims <- "A1"
  wb$add_data(dims = dims, x = name)

  name <- "Bad"
  dims <- "B1"
  wb <- wb_add_named_style(wb, dims = dims, name = name)
  wb$add_data(dims = dims, x = name)

  name <- "Good"
  dims <- "C1"
  wb <- wb_add_named_style(wb, dims = dims, name = name)
  wb$add_data(dims = dims, x = name)

  name <- "Neutral"
  dims <- "D1"
  wb <- wb_add_named_style(wb, dims = dims, name = name)
  wb$add_data(dims = dims, x = name)

  name <- "Calculation"
  dims <- "A2"
  wb <- wb_add_named_style(wb, dims = dims, name = name)
  wb$add_data(dims = dims, x = name)

  name <- "Check Cell"
  dims <- "B2"
  wb <- wb_add_named_style(wb, dims = dims, name = name)
  wb$add_data(dims = dims, x = name)

  name <- "Explanatory Text"
  dims <- "C2"
  wb <- wb_add_named_style(wb, dims = dims, name = name)
  wb$add_data(dims = dims, x = name)

  name <- "Input"
  dims <- "D2"
  wb <- wb_add_named_style(wb, dims = dims, name = name)
  wb$add_data(dims = dims, x = name)

  name <- "Linked Cell"
  dims <- "E2"
  wb <- wb_add_named_style(wb, dims = dims, name = name)
  wb$add_data(dims = dims, x = name)

  name <- "Note"
  dims <- "F2"
  wb <- wb_add_named_style(wb, dims = dims, name = name)
  wb$add_data(dims = dims, x = name)

  name <- "Output"
  dims <- "G2"
  wb <- wb_add_named_style(wb, dims = dims, name = name)
  wb$add_data(dims = dims, x = name)

  name <- "Warning Text"
  dims <- "H2"
  wb <- wb_add_named_style(wb, dims = dims, name = name)
  wb$add_data(dims = dims, x = name)

  name <- "Heading 1"
  dims <- "A3"
  wb <- wb_add_named_style(wb, dims = dims, name = name)
  wb$add_data(dims = dims, x = name)

  name <- "Heading 2"
  dims <- "B3"
  wb <- wb_add_named_style(wb, dims = dims, name = name)
  wb$add_data(dims = dims, x = name)

  name <- "Heading 3"
  dims <- "C3"
  wb <- wb_add_named_style(wb, dims = dims, name = name)
  wb$add_data(dims = dims, x = name)

  name <- "Heading 4"
  dims <- "D3"
  wb <- wb_add_named_style(wb, dims = dims, name = name)
  wb$add_data(dims = dims, x = name)

  name <- "Title"
  dims <- "E3"
  wb <- wb_add_named_style(wb, dims = dims, name = name)
  wb$add_data(dims = dims, x = name)

  name <- "Total"
  dims <- "F3"
  wb <- wb_add_named_style(wb, dims = dims, name = name)
  wb$add_data(dims = dims, x = name)

  for (i in seq_len(6)) {

    name <- paste0("20% - Accent", i)
    dims <- paste0(int2col(i), "4")
    wb <- wb_add_named_style(wb, dims = dims, name = name)
    wb$add_data(dims = dims, x = name)

    name <- paste0("40% - Accent", i)
    dims <- paste0(int2col(i), "5")
    wb <- wb_add_named_style(wb, dims = dims, name = name)
    wb$add_data(dims = dims, x = name)

    name <- paste0("60% - Accent", i)
    dims <- paste0(int2col(i), "6")
    wb <- wb_add_named_style(wb, dims = dims, name = name)
    wb$add_data(dims = dims, x = name)

    name <- paste0("Accent", i)
    dims <- paste0(int2col(i), "7")
    wb <- wb_add_named_style(wb, dims = dims, name = name)
    wb$add_data(dims = dims, x = name)

  }

  name <- "Comma"
  dims <- "A8"
  wb <- wb_add_named_style(wb, dims = dims, name = name)
  wb$add_data(dims = dims, x = name)

  name <- "Comma [0]"
  dims <- "B8"
  wb <- wb_add_named_style(wb, dims = dims, name = name)
  wb$add_data(dims = dims, x = name)

  name <- "Currency"
  dims <- "C8"
  wb <- wb_add_named_style(wb, dims = dims, name = name)
  wb$add_data(dims = dims, x = name)

  name <- "Currency [0]"
  dims <- "D8"
  wb <- wb_add_named_style(wb, dims = dims, name = name)
  wb$add_data(dims = dims, x = name)

  name <- "Per cent"
  dims <- "E8"
  wb <- wb_add_named_style(wb, dims = dims, name = name)
  wb$add_data(dims = dims, x = name)

  exp <- c(
    "Normal", "Bad", "Good", "Neutral", "Calculation", "Check Cell", "Explanatory Text",
    "Input", "Linked Cell", "Note", "Output", "Warning Text", "Heading 1",
    "Heading 2", "Heading 3", "Heading 4", "Title", "Total", "20% - Accent1",
    "40% - Accent1", "60% - Accent1", "Accent1", "20% - Accent2",
    "40% - Accent2", "60% - Accent2", "Accent2", "20% - Accent3",
    "40% - Accent3", "60% - Accent3", "Accent3", "20% - Accent4",
    "40% - Accent4", "60% - Accent4", "Accent4", "20% - Accent5",
    "40% - Accent5", "60% - Accent5", "Accent5", "20% - Accent6",
    "40% - Accent6", "60% - Accent6", "Accent6", "Comma", "Comma [0]",
    "Currency", "Currency [0]", "Per cent"
  )
  got <- wb$styles_mgr$cellStyle$name
  expect_equal(exp, got)

})

test_that("wb_add_dxfs_style() works", {
  wb <- wb_workbook()$
    add_worksheet()$
    add_dxfs_style(
      name = "nay",
      font_color = wb_color(hex = "FF9C0006"),
      bg_fill = wb_color(hex = "FFFFC7CE")
    )$
    add_dxfs_style(
      name = "yay",
      font_color = wb_color(hex = "FF006100"),
      bg_fill = wb_color(hex = "FFC6EFCE")
    )$
    add_data(x = -5:5)$
    add_data(x = LETTERS[1:11], start_col = 2)$
    add_conditional_formatting(
      dims = wb_dims(cols = 1, rows = 1:11),
      rule = "!=0",
      style = "nay"
    )$
    add_conditional_formatting(
      dims = wb_dims(cols = 1, rows = 1:11),
      rule = "==0",
      style = "yay"
    )

  exp <- data.frame(
    sqref = c("A1:A11", "A1:A11"),
    cf = c(
      "<cfRule type=\"expression\" dxfId=\"0\" priority=\"1\"><formula>A1&lt;&gt;0</formula></cfRule>",
      "<cfRule type=\"expression\" dxfId=\"1\" priority=\"2\"><formula>A1=0</formula></cfRule>"
    ),
    stringsAsFactors = FALSE
  )
  got <- wb$worksheets[[1]]$conditionalFormatting
  expect_equal(exp, got)

  exp <- c("nay", "yay")
  got <- wb$styles_mgr$dxf$name
  expect_equal(exp, got)

  expect_warning(
    wb_workbook()$
    add_worksheet()$
    add_dxfs_style(
      name = "nay",
      font_color = wb_color(hex = "FF9C0006"),
      bg_fill = wb_color(hex = "FFFFC7CE")
    )$
    add_dxfs_style(
      name = "nay",
      font_color = wb_color(hex = "FF006100"),
      bg_fill = wb_color(hex = "FFC6EFCE")
    ),
    "dxfs style names should be unique"
  )

})

test_that("initialized styles remain available", {
  foo_fill <- create_fill(pattern_type = "solid",
                          fg_color = wb_color("blue"))
  foo_font <- create_font(sz = 36, b = TRUE, color = wb_color("yellow"))

  wb <- wb_workbook()
  wb$styles_mgr$add(foo_fill, "foo")
  wb$styles_mgr$add(foo_font, "foo")

  foo_style <- create_cell_style(
    fill_id = wb$styles_mgr$get_fill_id("foo"),
    font_id = wb$styles_mgr$get_font_id("foo")
  )

  wb$styles_mgr$add(foo_style, "foo")

  wb$add_worksheet("test")
  wb$add_data(x = "Foo")
  wb$set_cell_style(dims = "A1", style = wb$styles_mgr$get_xf_id("foo"))

  exp <- c("xf-0", "foo")
  got <- wb$styles_mgr$xf$name
  expect_equal(exp, got)

  tmp <- temp_xlsx()
  wb$save(tmp)
  wb <- wb_load(tmp)

  exp <- c("xf-0", "xf-1")
  got <- wb$styles_mgr$xf$name
  expect_equal(exp, got)

})

test_that("apply styles across columns and rows", {

  wb <- wb_workbook()$
    add_worksheet()$
    add_fill(dims = "C3", color = wb_color("yellow"))$
    set_cell_style_across(style = "C3", cols = "C:D", rows = 3:4)

  exp <- c(
    "<col min=\"1\" max=\"2\" width=\"8.43\"/>",
    "<col min=\"3\" max=\"4\" style=\"1\" width=\"8.43\"/>"
  )
  got <- wb$worksheets[[1]]$cols_attr
  expect_equal(exp, got)

  exp <- data.frame(
    customFormat = c("1", "1"),
    r = c("3", "4"),
    s = c("1", "1"),
    stringsAsFactors = FALSE
  )
  got <- wb$worksheets[[1]]$sheet_data$row_attr[c("customFormat", "r", "s")]
  expect_equal(exp, got)

  # same as above but use cell style id
  wb <- wb_workbook()$add_worksheet()$add_fill(dims = "C3", color = wb_color("yellow"))
  wb$set_cell_style_across(style = wb_get_cell_style(wb, sheet = 1, dims = "C3"), cols = "C:D", rows = 3:4)

  exp <- c(
    "<col min=\"1\" max=\"2\" width=\"8.43\"/>",
    "<col min=\"3\" max=\"4\" style=\"1\" width=\"8.43\"/>"
  )
  got <- wb$worksheets[[1]]$cols_attr
  expect_equal(exp, got)

})

test_that("create_colors_xml() works", {

  skip_if(grDevices::palette()[2] == "red") # R 3.6 has a different palette

  exp <- "<a:clrScheme name=\"Base R\"><a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1><a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"00008B\"/></a:dk2><a:lt2><a:srgbClr val=\"D3D3D3\"/></a:lt2><a:accent1><a:srgbClr val=\"DF536B\"/></a:accent1><a:accent2><a:srgbClr val=\"61D04F\"/></a:accent2><a:accent3><a:srgbClr val=\"2297E6\"/></a:accent3><a:accent4><a:srgbClr val=\"28E2E5\"/></a:accent4><a:accent5><a:srgbClr val=\"CD0BBC\"/></a:accent5><a:accent6><a:srgbClr val=\"F5C710\"/></a:accent6><a:hlink><a:srgbClr val=\"0000FF\"/></a:hlink><a:folHlink><a:srgbClr val=\"A020F0\"/></a:folHlink></a:clrScheme>"
  got <- create_colors_xml()
  expect_equal(exp, got)

  got <- create_colours_xml()
  expect_equal(exp, got)

  wb <- wb_workbook()$set_base_colors(xml = got)
  got <- wb$get_base_colors(xml = TRUE, plot = FALSE)
  expect_equal(exp, got)

  wb <- wb_workbook()$set_base_colours(xml = got)
  got <- wb$get_base_colours(xml = TRUE, plot = FALSE)
  expect_equal(exp, got)

})

test_that("dims work", {

  wb <- wb_workbook()$add_worksheet()

  new_fill <- create_fill(pattern_type = "solid", fg_color = wb_color(hex = "FF334E6F"))
  wb$styles_mgr$add(new_fill, "new_fill")

  xf <- create_cell_style(fill_id = wb$styles_mgr$get_fill_id("new_fill"))
  wb$styles_mgr$add(xf, "xf")

  wb$set_cell_style(
    dims = "A1:A2",
    style = wb$styles_mgr$get_xf_id("xf")
  )

  wb$set_cell_style(
    dims = c("B1", "B2"),
    style = wb$styles_mgr$get_xf_id("xf")
  )

  expect_setequal(wb$worksheets[[1]]$sheet_data$cc$c_s, "1")

})

test_that("update font works", {
  wb <- wb_workbook() |>
    wb_add_worksheet() |>
    wb_add_data(x = letters) |>
    wb_add_font(dims = wb_dims(x = letters), name = "Calibri", size = 20, update = c("name", "size", "scheme"))

  exp <- "<font><color theme=\"1\"/><family val=\"2\"/><name val=\"Calibri\"/><sz val=\"20\"/></font>"
  got <- wb$styles_mgr$styles$fonts[2]
  expect_equal(exp, got)

  # updates only the font color
  wb$add_font(dims = wb_dims(x = letters), color = wb_color("orange"), update = c("color"))

  exp <- "<font><color rgb=\"FFFFA500\"/><family val=\"2\"/><name val=\"Calibri\"/><sz val=\"20\"/></font>"
  got <- wb$styles_mgr$styles$fonts[3]
  expect_equal(exp, got)
})

test_that("adding bg_color and diagonal borders work", {

  wb <- wb_workbook()$
    add_worksheet()$
    add_fill(dims = "B2:D4",
             color = wb_color("white"),
             bg_color = wb_color("black"),
             pattern = "lightUp")$
    add_border(
      dims = "B6:D8",
      diagonal_up = "dashed",
      diagonal_down = "dashed",
      diagonal_color = wb_color("red")
    )

  expect_error(
    wb$add_worksheet()$
      add_border(
      dims = "B6:D8",
      diagonal_up = "dashed",
      diagonal_down = "thin",
      diagonal_color = wb_color("red")
    ),
    "there can be only a single diagonal style per cell"
  )

  exp <- "<border diagonalDown=\"1\" diagonalUp=\"1\"><start><color rgb=\"FFFF0000\"/></start><end><color rgb=\"FFFF0000\"/></end><left style=\"thin\"><color rgb=\"FF000000\"/></left><top style=\"thin\"><color rgb=\"FF000000\"/></top><diagonal style=\"dashed\"><color rgb=\"FFFF0000\"/></diagonal></border>"
  got <- wb$styles_mgr$styles$borders[2]
  expect_equal(exp, got)

  exp <- "<fill><patternFill patternType=\"lightUp\"><fgColor rgb=\"FFFFFFFF\"/><bgColor rgb=\"FF000000\"/></patternFill></fill>"
  got <- wb$styles_mgr$styles$fills[3]
  expect_equal(exp, got)

})
