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
           "<xf applyBorder=\"1\" borderId=\"1\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>"
  )
  got <- wb$styles_mgr$styles$cellXfs

  expect_equal(exp, got)

  # check borders
  exp <- c("<border><left/><right/><top/><bottom/><diagonal/></border>",
           "<border><bottom style=\"double\"><color rgb=\"FF000000\"/></bottom></border>"
  )
  got <- wb$styles_mgr$styles$borders

  expect_equal(exp, got)


  wb <- wb_workbook()
  wb$add_worksheet("S1")$add_border(1, dims = "A1:K1", left_border = NULL, right_border = NULL, top_border = NULL, bottom_border = "double")

  exp <- "1"
  got <- unique(wb$worksheets[[1]]$sheet_data$cc$c_s)
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

  exp <- "<numFmt numFmtId=\"200\" formatCode=\"hh:mm:ss AM/PM\"/>"
  got <- create_numfmt(numFmtId = 200, formatCode = "hh:mm:ss AM/PM")
  expect_equal(got, exp)

  exp <- "<numFmt numFmtId=\"200\" formatCode=\"hh:mm:ss A/P\"/>"
  got <- create_numfmt(numFmtId = 200, formatCode = "hh:mm:ss A/P")
  expect_equal(got, exp)

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
  on.exit(unlink(tmp), add = TRUE)

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
    "Only a single diagonal style per cell allowed"
  )

  exp <- "<border diagonalDown=\"1\" diagonalUp=\"1\"><start><color rgb=\"FFFF0000\"/></start><end><color rgb=\"FFFF0000\"/></end><left style=\"thin\"><color rgb=\"FF000000\"/></left><top style=\"thin\"><color rgb=\"FF000000\"/></top><diagonal style=\"dashed\"><color rgb=\"FFFF0000\"/></diagonal></border>"
  got <- wb$styles_mgr$styles$borders[2]
  expect_equal(exp, got)

  exp <- "<fill><patternFill patternType=\"lightUp\"><fgColor rgb=\"FFFFFFFF\"/><bgColor rgb=\"FF000000\"/></patternFill></fill>"
  got <- wb$styles_mgr$styles$fills[3]
  expect_equal(exp, got)

})

test_that("adding borders works", {
  wb <- wb_workbook()$add_worksheet()
  wb$add_data(x = matrix(8 * 10, 10, 8), col_names = FALSE)
  wb$add_fill(dims = "A1:D5", color = wb_color("blue"))
  wb$add_fill(dims = "A6:D10", color = wb_color("red"))
  wb$add_fill(dims = "E1:H5", color = wb_color("green"))
  wb$add_fill(dims = "E6:H10", color = wb_color("white"))

  exp <- c(A1 = "1", B1 = "1", C1 = "1", D1 = "1", E1 = "3", F1 = "3",
           G1 = "3", H1 = "3", A2 = "1", B2 = "1", C2 = "1", D2 = "1", E2 = "3",
           F2 = "3", G2 = "3", H2 = "3", A3 = "1", B3 = "1", C3 = "1", D3 = "1",
           E3 = "3", F3 = "3", G3 = "3", H3 = "3", A4 = "1", B4 = "1", C4 = "1",
           D4 = "1", E4 = "3", F4 = "3", G4 = "3", H4 = "3", A5 = "1", B5 = "1",
           C5 = "1", D5 = "1", E5 = "3", F5 = "3", G5 = "3", H5 = "3", A6 = "2",
           B6 = "2", C6 = "2", D6 = "2", E6 = "4", F6 = "4", G6 = "4", H6 = "4",
           A7 = "2", B7 = "2", C7 = "2", D7 = "2", E7 = "4", F7 = "4", G7 = "4",
           H7 = "4", A8 = "2", B8 = "2", C8 = "2", D8 = "2", E8 = "4", F8 = "4",
           G8 = "4", H8 = "4", A9 = "2", B9 = "2", C9 = "2", D9 = "2", E9 = "4",
           F9 = "4", G9 = "4", H9 = "4", A10 = "2", B10 = "2", C10 = "2",
           D10 = "2", E10 = "4", F10 = "4", G10 = "4", H10 = "4")
  got <- wb$get_cell_style(dims = "A1:H10")
  expect_equal(exp, got)

  wb$add_border(dims = "A1:H10")
  exp <- c(A1 = "5", B1 = "6", C1 = "6", D1 = "6", E1 = "7", F1 = "7",
           G1 = "7", H1 = "8", A2 = "9", B2 = "11", C2 = "11", D2 = "11",
           E2 = "12", F2 = "12", G2 = "12", H2 = "15", A3 = "9", B3 = "11",
           C3 = "11", D3 = "11", E3 = "12", F3 = "12", G3 = "12", H3 = "15",
           A4 = "9", B4 = "11", C4 = "11", D4 = "11", E4 = "12", F4 = "12",
           G4 = "12", H4 = "15", A5 = "9", B5 = "11", C5 = "11", D5 = "11",
           E5 = "12", F5 = "12", G5 = "12", H5 = "15", A6 = "10", B6 = "13",
           C6 = "13", D6 = "13", E6 = "14", F6 = "14", G6 = "14", H6 = "16",
           A7 = "10", B7 = "13", C7 = "13", D7 = "13", E7 = "14", F7 = "14",
           G7 = "14", H7 = "16", A8 = "10", B8 = "13", C8 = "13", D8 = "13",
           E8 = "14", F8 = "14", G8 = "14", H8 = "16", A9 = "10", B9 = "13",
           C9 = "13", D9 = "13", E9 = "14", F9 = "14", G9 = "14", H9 = "16",
           A10 = "17", B10 = "18", C10 = "18", D10 = "18", E10 = "19", F10 = "19",
           G10 = "19", H10 = "20")
  got <- wb$get_cell_style(dims = "A1:H10")
  expect_equal(exp, got)

})

# Formatting a number
test_that("number formatting works", {

  got <- apply_numfmt(1234.5678, "#,##0.00")
  expect_identical(got, "1,234.57")

  got <- apply_numfmt(pi, "#,##0")
  expect_identical(got, "3")

  got <- apply_numfmt(123456, "$#,##0")
  expect_identical(got, "$123,456")

  got <- apply_numfmt(pi, "0,000")
  expect_identical(got, "0,003")

  got <- apply_numfmt(123456789, "#,###,,")
  expect_identical(got, "123")

})

test_that("date and time formatting works", {

  got <- apply_numfmt("2025-01-05", "yyyy-mm-dd")
  expect_identical(got, "2025-01-05")

  got <- apply_numfmt("2025-01-05", "m/d/yyyy")
  expect_identical(got, "1/5/2025")

  got <- apply_numfmt("2025-01-05", "mm/dd/yy")
  expect_identical(got, "01/05/25")

  # Formatting a time
  got <- apply_numfmt("2025-01-05 13:45:30", "hh:mm:ss AM/PM")
  expect_identical(got, "01:45:30 PM")

  got <- apply_numfmt("2025-01-05 13:45:30", "hh:mm:ss")
  expect_identical(got, "13:45:30")

  got <- apply_numfmt("13:45:30", "hh:mm:ss AM/PM")
  expect_identical(got, "01:45:30 PM")

  got <- apply_numfmt("13:45:30", "hh:mm:ss")
  expect_identical(got, "13:45:30")

  x <- structure(42.5, class = c("hms", "difftime"))
  got <- apply_numfmt(x, "hh:mm:ss")
  expect_identical(got, "12:00:00")

  got <- apply_numfmt(x, "YYYY-MM-DD hh:mm:ss")
  expect_identical(got, "1900-02-11 12:00:00")

  # Formatting combined date and time
  got <- apply_numfmt("2025-01-05 13:45:30", "yy-mmm-dd HH:mm:ss")
  expect_identical(got, "25-Jan-05 13:45:30")

  got <- apply_numfmt(as.POSIXct("1900-01-12 08:17:47", "UTC"), "[h]:mm:ss")
  expect_equal(got, "296:17:47")

  got <- apply_numfmt("1900-01-12 08:17:47", "[h]:mm:ss")
  expect_equal(got, "296:17:47")

})

testthat::test_that("comprehensive duration works", {
  # The original 296 hours test
  val <- as.POSIXct("1900-01-12 08:17:47", tz = "UTC")

  # Total Hours
  expect_equal(apply_numfmt(val, "[h]:mm:ss"), "296:17:47")

  # Total Minutes
  # (12 days * 1440) + (8 * 60) + 17 = 17777
  expect_equal(apply_numfmt(val, "[m]:ss"), "17777:47")

  # Total Seconds
  # (17777 * 60) + 47 = 1066667
  expect_equal(apply_numfmt(val, "[s]"), "1066667")

  # Mixed format (Excel supports this)
  expect_equal(apply_numfmt(val, "[h] \"hours and\" m \"minutes\""), "296 hours and 17 minutes")

  expect_equal(apply_numfmt("1900-01-12 08:17:47", "[h] \"hours and\" m \"minutes\""), "296 hours and 17 minutes")

  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = as.POSIXct("1900-01-12 08:17:47"))$
    add_numfmt(numfmt = "[h] \"hours and\" m \"minutes\"")

  exp <- "296 hours and 17 minutes"
  got <- wb$to_df(apply_numfmts = TRUE, col_names = FALSE)$A
  expect_equal(got, exp)

})

test_that("special formatting works", {

  # Formatting a fraction
  got <- apply_numfmt(0.75, "# ?/?")
  expect_identical(got, "3/4")

  got <- apply_numfmt(1.75, "# ?/?")
  expect_identical(got, "1 3/4")

  got <- apply_numfmt(-1234.56, "(#,###.00)")
  expect_identical(got, "(1,234.56)")

  fmt <- '#,###.00_);[Red](#,###.00);0.00;"gross receipts for "@'
  exp <- "<numFmt numFmtId=\"164\" formatCode=\"#,###.00_);[Red](#,###.00);0.00;&quot;gross receipts for &quot;@\"/>"
  got <- create_numfmt(formatCode = fmt)
  expect_identical(got, exp)

  got <- apply_numfmt(1234.5678, fmt)
  expect_identical(got, "1,234.57")

  got <- apply_numfmt(-1234.5678, fmt)
  expect_identical(got, "(1,234.57)")

  got <- apply_numfmt(0, fmt)
  expect_identical(got, "0.00")

  got <- apply_numfmt("a", fmt)
  expect_identical(got, "gross receipts for a")

  # Formatting a text
  got <- apply_numfmt("Hello", "Hello @")
  expect_identical(got, "Hello Hello")

  fmt <- "_-[$£-809]* #,##0.00_-;\\-[$£-809]* #,##0.00_-;_-[$£-809]* &quot;-&quot;??_-;_-@_-"
  got <- apply_numfmt(1234.5678, fmt) # currency symbol is tabbed into the front, so at least one whitespace after currency symbol
  expect_identical(got, "£ 1,234.57")

  got <- apply_numfmt(-1234.5678, fmt)
  expect_identical(got, "-£ 1,234.57")

  fmt <- "_- [$£-809]* #,##0.00_-;\\- [$£-809]* #,##0.00_-; [Red]-[$£-809]* &quot;-&quot;??_-;_-@_-"

  got <- apply_numfmt(12345.13456, fmt) # same tabbed currency
  expect_identical(got, "£ 12,345.13")

  got <- apply_numfmt(-12345.13456, fmt)
  expect_identical(got, "- £ 12,345.13")

  got <- apply_numfmt(0, fmt) # no leading minus here
  expect_identical(got, "-£ -")

  fmt <- "[$€-2] #,##0.00_);[Red]([$€-2] #,##0.00)"
  got <- apply_numfmt(12345.13456, fmt)
  expect_identical(got, "€ 12,345.13")

  got <- apply_numfmt(-12345.13456, fmt)
  expect_identical(got, "(€ 12,345.13)")

  got <- apply_numfmt(0, fmt)
  expect_identical(got, "€ 0.00")

  got <- apply_numfmt(0.5, "#.##%")
  expect_identical(got, "50%")

  got <- apply_numfmt(1.75, "#%")
  expect_identical(got, "175%")

  got <- apply_numfmt(-1234.57, "$ #,##0")
  expect_identical(got, "-$ 1,235")

  got <- apply_numfmt(.POSIXct(72901, "UTC"), "hh:mm:ss A/P")
  expect_identical(got, "08:15:01 P")

  got <- apply_numfmt(0, "#,##0; -#,##0; \"Nil\"")
  expect_equal(got, "Nil")

  got <- apply_numfmt(5, "# ?/?")
  expect_equal(got, "5")

  got <- apply_numfmt(1.25, "?/?")
  expect_equal(got, "5/4")

  got <- apply_numfmt(0.66, "?/?")
  expect_equal(got, "2/3")

  got <- apply_numfmt(123456789, "0.00E+00")
  expect_equal(got, "1.23E+08")

})

test_that("", {
  exp <- c(
    "4.00", "4.00", "7.00", "7.00", "8.00", "9.00", "10.00", "10.00",
    "10.00", "11.00", "11.00", "12.00", "12.00", "12.00", "12.00",
    "13.00", "13.00", "13.00", "13.00", "14.00", "14.00", "14.00",
    "14.00", "15.00", "15.00", "15.00", "16.00", "16.00", "17.00",
    "17.00", "17.00", "18.00", "18.00", "18.00", "18.00", "19.00",
    "19.00", "19.00", "20.00", "20.00", "20.00", "20.00", "20.00",
    "22.00", "23.00", "24.00", "24.00", "24.00", "24.00", "25.00"
  )
  got <- apply_numfmt(cars$speed, "#,##0.00")
  expect_identical(got, exp)

  got <- apply_numfmt(cars$speed, rep_len("#,##0.00", nrow(cars)))
  expect_identical(got, exp)
})

test_that("apply_numfmts works", {
  df <- data.frame(
    is_active  = c(TRUE, FALSE, TRUE, NA, FALSE),
    count      = c(10L, 25L, NA, 7L, 15L),
    measure    = c(1.23, NA, 4.56, 7.89, 0.01),
    event_date = as.Date(c("2023-01-01", "2023-02-15", NA, "2023-05-20", "2023-12-31")),
    timestamp  = as.POSIXct(c("2023-01-01 10:00:00", NA, "2023-03-10 14:30:00",
                              "2023-06-01 08:15:00", "2023-12-25 12:00:00"), tz = "UTC"),
    # Manual hms without loading the library:
    daily_time = structure(c(30600, 35100, NA, 40500, 43200), units = "secs", class = c("hms", "difftime")),
    category   = c("A", "B", "C", NA, "E"),
    stringsAsFactors = FALSE
  )

  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = df)$
    set_col_widths(cols = seq_along(df), widths = "auto")

  wb$add_numfmt(dims = wb_dims(x = df, cols = "count"), numfmt = "$0")
  wb$add_numfmt(dims = wb_dims(x = df, cols = "measure"), numfmt = "0.0")
  wb$add_numfmt(dims = wb_dims(x = df, cols = "event_date"), numfmt = "m/d/yyyy")
  wb$add_numfmt(dims = wb_dims(x = df, cols = "timestamp"), numfmt = "yy-mmm-dd HH:mm:ss")
  wb$add_numfmt(dims = wb_dims(x = df, cols = "daily_time"), numfmt = "hh:mm:ss AM/PM")
  wb$add_numfmt(dims = wb_dims(x = df, cols = "category"), numfmt = "cat: @")

  wb$
    set_col_widths(cols = seq_along(df), widths = "auto")

  exp <- structure(
    list(
      is_active = c("TRUE", "FALSE", "TRUE", NA, "FALSE"),
      count = c("$10", "$25", NA, "$7", "$15"),
      measure = c("1.2", NA, "4.6", "7.9", "0.0"),
      event_date = c("1/1/2023", "2/15/2023", NA, "5/20/2023", "12/31/2023"),
      timestamp = c("23-Jan-01 10:00:00", NA, "23-Mar-10 14:30:00", "23-Jun-01 08:15:00", "23-Dec-25 12:00:00"),
      daily_time = c("08:30:00 AM", "09:45:00 AM", NA, "11:15:00 AM", "12:00:00 PM"),
      category = c("cat: A", "cat: B", "cat: C", NA, "cat: E")
    ),
    row.names = 2:6,
    class = "data.frame"
  )
  got <- wb$to_df(apply_numfmts = TRUE)
  expect_equal(got, exp)
})

test_that("escaped numfmt works", {

  fmt  <- "_- [$£-809]* #,##0.00_-;\\- [$£-809]* #,##0.00_-; [Red]-[$£-809]* &quot;-&quot;??_-;_-@_-"
  fmt2 <- "_-[$£-809]* #,##0.00_-;\\-[$£-809]* #,##0.00_-;_-[$£-809]* &quot;-&quot;??_-;_-@_-"

  wb <- wb_workbook()$
    add_worksheet()$
    add_data(dims = "A1", x = -1234.5678)$
    add_data(dims = "A2", x = -1234.5678)$
    add_data(dims = "A3", x = 0)$
    add_numfmt(dims = "A1", numfmt = fmt)$
    add_numfmt(dims = "A2", numfmt = fmt2)$
    add_numfmt(dims = "A3", numfmt = fmt)

  exp <- structure(
    list(
      A = c("- £ 1,234.57", "-£ 1,234.57", "-£ -")
    ),
    row.names = c(NA, 3L),
    class = "data.frame"
  )
  got <- wb_to_df(wb, apply_numfmts = TRUE, col_names = FALSE)
  expect_equal(got, exp)

})

test_that("day names work", {
  val <- "2025-01-05" # This is a Sunday
  expect_identical(apply_numfmt(val, "ddd"), "Sun")
  expect_identical(apply_numfmt(val, "dddd"), "Sunday")
})

test_that("zero vs placeholder difference", {
  expect_identical(apply_numfmt(5, "0.00"), "5.00")
  expect_identical(apply_numfmt(5, "#.##"), "5")
  expect_identical(apply_numfmt(5.4, "#.00"), "5.40")
  expect_identical(apply_numfmt(0.4, "#.00"), ".40")
  expect_identical(apply_numfmt(0.4, "0.00"), "0.40")
})

test_that("escaped literals", {
  # The 'm' should be a literal 'm', not a month
  expect_identical(apply_numfmt(123, "\\m#"), "m123")
})

test_that("padded durations", {
  expect_equal(apply_numfmt(1 / 24, "[hh]:mm"), "01:00")
  expect_equal(apply_numfmt("1899-12-31 01:00:00", "[hh]:mm"), "01:00")
  expect_equal(apply_numfmt(pi, "[hh]:mm"), "75:23")
})

test_that("duration minute ambiguity", {
  # In a standard date, 'm' is a month
  expect_identical(apply_numfmt("2025-01-05", "yyyy m d"), "2025 1 5")

  # Inside a duration context, 'm' must be minutes
  # 1900-01-12 08:17:47 is 296 hours and 17 minutes
  val <- "1900-01-12 08:17:47"
  expect_equal(apply_numfmt(val, "[h] m"), "296 17")
  expect_equal(apply_numfmt(val, "[h] mm"), "296 17")
})

test_that("conditional bracket formatting", {
  # Format: if < 1000, show as is; if >= 1000, show in thousands with a 'k'
  fmt <- "[<1000]#,##0;[>=1000]#,##0,\"k\""

  expect_identical(apply_numfmt(500, fmt), "500")
  expect_identical(apply_numfmt(1500, fmt), "2k")

  # Color-based conditionals (the colors are usually stripped or ignored in text output)
  fmt_color <- "[Red][<100]0;[Blue][>=100]0"
  expect_identical(apply_numfmt(50, fmt_color), "50")
  expect_identical(apply_numfmt(150, fmt_color), "150")


  fmt <- "[<1000]#,##0;[>=1000]#,##0,\"k\""
  fmt_color <- "[Red][<100]0;[Blue][>=100]0"
  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = c(500, 1500, 50, 150))$
    add_numfmt(dims = "A1:A2", numfmt = fmt)$
    add_numfmt(dims = "A3:A4", numfmt = fmt_color)

  exp <- c("500", "2k", "50", "150")
  got <- wb$to_df(apply_numfmts = TRUE, col_names = FALSE)$A

  expect_equal(got, exp)

})

test_that("mandatory and optional fractional zeros", {
  # Trigger: num_parts > 1 and nchar(frac_str) > mandatory_frac_len
  # Format '0.0#' has 1 mandatory zero.
  # Value 5.40 -> should trim the trailing 0 because it's in a '#' position.
  expect_identical(apply_numfmt(5.4, "0.0#"), "5.4")

  # Value 5.45 -> should keep the 5 because it's not a zero.
  expect_identical(apply_numfmt(5.45, "0.0#"), "5.45")

  # Value 5.00 -> should keep one zero because of the '0' in '0.0#'
  expect_identical(apply_numfmt(5, "0.0#"), "5.0")
})

test_that("conditional else branch", {
  # Format: If <100, show 'Small'; If <200, show 'Medium'; otherwise show the number.
  # The last section "0.00" has no brackets and triggers the 'else' logic.
  fmt <- "[<100]\"Small\";[<200]\"Medium\";0.00"

  # Hits the 'else' section (target_fmt <- sec)
  expect_identical(apply_numfmt(250, fmt), "250.00")
})

test_that("rounding works similar to R", {
  exp <- sprintf("%.1f", round(seq(from = 0, to = 1, by = .05), digits = 1))
  got <- apply_numfmt(seq(from = 0, to = 1, by = .05), "0.0")
  expect_equal(got, exp)

  got <- apply_numfmt(-0.05, "0.0")
  expect_equal(got, "-0.0")
})

test_that("extensive hms works", {
  df <- data.frame(
    daily_time = structure(c(3061200, 3511200, NA, 4051200, 4321200), units = "secs", class = c("hms", "difftime")),
    stringsAsFactors = FALSE
  )

  wb <- wb_workbook()$
    add_worksheet(zoom = 130)$
    add_data(x = df)

  wb$add_numfmt(dims = wb_dims(x = df, cols = "daily_time"), numfmt = "hh:mm:ss")

  exp <- structure(
    list(
      daily_time = structure(c(-2206014000, -2205564000, NA, -2205024000, -2204754000), class = c("POSIXct", "POSIXt"), tzone = "UTC")
    ),
    row.names = 2:6,
    class = "data.frame"
  )
  got <- wb$to_df()
  expect_equal(got, exp)

  exp <- structure(
    list(
      daily_time = c("10:20:00", "15:20:00", NA, "21:20:00", "00:20:00")
    ),
    row.names = 2:6,
    class = "data.frame"
  )
  got <- wb$to_df(apply_numfmts = TRUE)
  expect_equal(got, exp)

})

test_that("removing fill works", {
  wb <- wb_workbook()$add_worksheet()$add_data(x = head(mtcars))
  wb$add_border(dims = wb_dims(x = head(mtcars), select = "data"))

  wb$add_fill(dims = "A1:B2", color = wb_color("blue"))
  wb$add_fill(dims = "A2:B2", color = NULL)

  styles <- wb$get_cell_style(dims = "A2:B2")
  xml <- wb$styles_mgr$styles$cellXfs[as.numeric(styles) + 1]

  xml_df <- rbindlist(xml_attr(xml, "xf"))
  expect_true(is.null(xml_df$applyFill))
  expect_equal(unique(xml_df$fillId), "0")
})

test_that("removing numfmt works", {
  ddims <- wb_dims(x = head(mtcars), select = "data")
  wb <- wb_workbook()$add_worksheet()$add_data(x = head(mtcars))
  wb$add_border(dims = ddims)

  wb$add_fill(dims = ddims, color = wb_color("yellow"))
  wb$add_numfmt(dims = ddims, numfmt = "0.00%")
  wb$add_numfmt(dims = "A2:B2", numfmt = NULL)

  styles <- wb$get_cell_style(dims = "A2:B2")
  xml <- wb$styles_mgr$styles$cellXfs[as.numeric(styles) + 1]

  xml_df <- rbindlist(xml_attr(xml, "xf"))
  expect_true(is.null(xml_df$applyNumberFormat))
  expect_equal(unique(xml_df$numFmtId), "0")
})

test_that("removing font works", {
  ddims <- wb_dims(x = head(mtcars), select = "data")
  wb <- wb_workbook()$add_worksheet()$add_data(x = head(mtcars))

  wb$add_border(dims = ddims)
  wb$add_fill(dims = ddims, color = wb_color("yellow"))
  wb$add_font(dims = ddims, name = "Arial", size = 16)
  wb$add_font(dims = "A2:B2", update = NULL)

  styles <- wb$get_cell_style(dims = "A2:B2")
  xml <- wb$styles_mgr$styles$cellXfs[as.numeric(styles) + 1]

  xml_df <- rbindlist(xml_attr(xml, "xf"))
  expect_true(is.null(xml_df$applyFont))
  expect_equal(unique(xml_df$fontId), "0")
})

test_that("removing border works", {
  ddims <- wb_dims(x = head(mtcars), select = "data")
  wb <- wb_workbook()$add_worksheet()$add_data(x = head(mtcars))

  wb$add_border(dims = ddims)
  wb$add_fill(dims = ddims, color = wb_color("yellow"))
  wb$add_font(dims = ddims, name = "Arial", size = 16)
  wb$add_border(dims = "A2:B2", update = NULL)

  styles <- wb$get_cell_style(dims = "A2:B2")
  xml <- wb$styles_mgr$styles$cellXfs[as.numeric(styles) + 1]

  xml_df <- rbindlist(xml_attr(xml, "xf"))
  expect_true(is.null(xml_df$applyBorder))
  expect_equal(unique(xml_df$borderId), "0")
})

test_that("wb_set_col_widths() works", {
  wb <- wb_workbook()$add_worksheet()$
    set_col_widths(cols = 2:7, width = "auto")
  exp <- character()
  got <- wb$worksheets[[1]]$cols_attr
  expect_equal(got, exp)


  wb$add_data(x = cars, dims = "B2")$
    set_col_widths(cols = 1:4, width = "auto")

  exp <- c(
    "<col min=\"2\" max=\"2\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"5.711\"/>",
    "<col min=\"3\" max=\"3\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"4.711\"/>"
  )
  got <- wb$worksheets[[1]]$cols_attr
  expect_equal(got, exp)


  wb <- wb_workbook()$add_worksheet()$
    set_col_widths(cols = 2:3, width = 6)

  exp <- "<col min=\"2\" max=\"3\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"6.711\"/>"
  got <- wb$worksheets[[1]]$cols_attr
  expect_equal(got, exp)


  wb <- wb_workbook()$add_worksheet()
  wb$set_col_widths(cols = 1:4, width = 4)
  wb$set_col_widths(cols = 7:8, width = 4)

  exp <- c(
    "<col min=\"1\" max=\"4\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"4.711\"/>",
    "<col min=\"7\" max=\"8\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"4.711\"/>"
  )
  got <- wb$worksheets[[1]]$cols_attr
  expect_equal(got, exp)

  wb <- wb_workbook()$add_worksheet()
  wb$set_col_widths(cols = 3:7, width = 5)
  wb$set_col_widths(cols = 2:8, width = 4)

  exp <- "<col min=\"2\" max=\"8\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"4.711\"/>"
  got <- wb$worksheets[[1]]$cols_attr
  expect_equal(got, exp)

})

test_that("adding the same numfmts twice works", {
  wb1 <- wb_workbook()$add_worksheet()
  wb1$styles_mgr$add("<numFmt numFmtId=\"900\" formatCode=\"0.0\"/>", "foo")
  wb1$styles_mgr$add("<numFmt numFmtId=\"900\" formatCode=\"0.0\"/>", "foo")

  exp <- "<numFmt numFmtId=\"900\" formatCode=\"0.0\"/>"
  got <- wb1$styles_mgr$styles$numFmts
  expect_equal(got, exp)

  got <- nrow(wb1$styles_mgr$numfmt)
  expect_equal(got, 1L)
})

test_that("checking  the same numfmts twice works", {
  wb <- wb_workbook()$add_worksheet()
  expect_warning(
    wb$styles_mgr$get_xf_id(name = "my_awesome_style"),
    "Could not find style\\(s\\): my_awesome_style"
  )
  wb$styles_mgr$add("<numFmt numFmtId=\"900\" formatCode=\"0.0\"/>", "foo")
  wb$styles_mgr$get_numfmt_id(name = "foo")
  expect_warning(
    wb$styles_mgr$get_numfmt_id(name = c("foo", "my_awesome_style")),
    "Could not find style\\(s\\): my_awesome_style"
  )

})

test_that("applying styles works", {
  xl <- system.file("extdata", "oxlsx2_sheet.xlsx", package = "openxlsx2")
  exp <- "96.55\u00a0%"
  got <- wb_to_df(xl, apply_numfmts = TRUE, dims = "I7", col_names = FALSE)[["I"]]
  expect_equal(got, exp)
})

test_that("apply_numfmt handles AM/PM regardless of system locale", {
  orig_locale <- Sys.getlocale("LC_TIME")
  tryCatch({
    Sys.setlocale("LC_TIME", "en_GB.UTF-8")
  }, error = function(e) {
    skip("Target locale en_GB.UTF-8 not available on this system")
  })
  on.exit(Sys.setlocale("LC_TIME", orig_locale), add = TRUE)

  got <- apply_numfmt("13:45:30", "hh:mm:ss AM/PM")
  expect_identical(got, "01:45:30 PM")

  got_short <- apply_numfmt("13:45:30", "hh:mm:ss A/P")
  expect_identical(got_short, "01:45:30 P")
})
