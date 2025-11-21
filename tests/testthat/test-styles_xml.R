test_that("xf", {

  xlsxFile <- testfile_path("loadExample.xlsx")
  wb <- wb_load(xlsxFile)

  xlsxFile <- testfile_path("readTest.xlsx")
  wb1 <- wb_load(xlsxFile)

  # likely not ordered entirely correct
  exp <- c(
    "numFmtId", "fontId", "fillId", "borderId", "xfId",
    "applyFill", "applyBorder", "applyFont", "applyAlignment",
    "applyNumberFormat", "applyProtection",
    "pivotButton", "quotePrefix",
    "horizontal", "indent", "justifyLastLine", "readingOrder",
    "relativeIndent", "shrinkToFit", "textRotation", "vertical",
    "wrapText", "extLst", "hidden", "locked"
  )

  input <- wb$styles_mgr$styles$cellXfs
  got <- read_xf(as_xml(input))
  expect_equal(sort(exp), sort(names(got)))

  exp_dim <- c(74, 25)
  expect_equal(exp_dim, dim(got))

  expect_equal(input,
               write_xf(got[exp]))

  expect_warning(
    got <- read_xf(read_xml('<xf numFmtId="0" foo="0"/>')),
    "foo: not found in xf name table"
  )

})

test_that("border", {

  xlsxFile <- testfile_path("loadExample.xlsx")
  wb <- wb_load(xlsxFile)

  # likely not ordered entirely correct
  exp <- c("left", "right", "top", "bottom",
           "diagonal", "diagonalDown", "diagonalUp", "end",
           "horizontal",  "outline",  "start", "vertical"
  )

  input <- wb$styles_mgr$styles$borders
  got <- read_border(as_xml(input))
  expect_equal(sort(exp), sort(names(got)))

  exp_dim <- c(58, 12)
  expect_equal(exp_dim, dim(got))

  expect_equal(input,
               write_border(got[exp]))

  expect_warning(
    got <- read_border(read_xml('<border foo="0"/>')),
    "foo: not found in border name table"
  )

})

test_that("cellStyle", {

  xlsxFile <- testfile_path("loadExample.xlsx")
  wb <- wb_load(xlsxFile)

  # likely not ordered entirely correct
  exp <- c("name", "xfId", "xr:uid", "builtinId", "customBuiltin", "extLst", "hidden", "iLevel")

  input <- wb$styles_mgr$styles$cellStyles
  got <- read_cellStyle(as_xml(input))
  expect_equal(sort(exp), sort(names(got)))

  exp_dim <- c(2, 8)
  expect_equal(exp_dim, dim(got))

  expect_equal(input,
               write_cellStyle(got[exp]))

  expect_warning(
    got <- read_cellStyle(read_xml('<cellStyle foo="0"/>')),
    "foo: not found in cellstyle name table"
  )

})

test_that("tableStyle", {


  # tableStyle elements exist only with custom tableStyles
  xlsxFile <- testfile_path("tableStyles.xlsx")
  wb <- wb_load(xlsxFile)

  # likely not ordered entirely correct
  exp <- c("name", "pivot", "count", "xr9:uid", "table", "tableStyleElement")

  input <- wb$styles_mgr$styles$tableStyles
  got <- read_tableStyle(as_xml(input))
  expect_equal(sort(exp), sort(names(got)))

  exp_dim <- c(1, 6)
  expect_equal(exp_dim, dim(got))

  expect_equal(input,
               write_tableStyle(got[exp]))

  expect_warning(
    got <- read_tableStyle(read_xml('<tableStyle foo="0"/>')),
    "foo: not found in tablestyle name table"
  )

})

test_that("dxf", {

  xlsxFile <- testfile_path("loadExample.xlsx")
  wb <- wb_load(xlsxFile)

  # likely not ordered entirely correct
  exp <- c("font", "fill", "alignment", "border", "extLst",   "numFmt",
           "protection")

  input <- wb$styles_mgr$styles$dxfs
  got <- read_dxf(as_xml(input))
  expect_equal(sort(exp), sort(names(got)))

  exp_dim <- c(2, 7)
  expect_equal(exp_dim, dim(got))

  expect_equal(input,
               write_dxf(got[exp]))

  expect_warning(
    got <- read_dxf(read_xml('<dxf><foo bar="0"/></dxf>')),
    "foo: not found in dxf name table"
  )

})

test_that("colors", {

  xlsxFile <- testfile_path("oxlsx2_sheet.xlsx")
  wb <- wb_load(xlsxFile)

  # likely not ordered entirely correct
  exp <- c("mruColors", "indexedColors")

  input <- wb$styles_mgr$styles$colors
  got <- read_colors(as_xml(input))
  expect_equal(sort(exp), sort(names(got)))

  exp_dim <- c(1, 2)
  expect_equal(exp_dim, dim(got))

  expect_equal(input,
               write_colors(got[exp]))

  expect_warning(
    got <- read_colors(read_xml('<colors><foo bar="0"/></colors>')),
    "foo: not found in color name table"
  )

})

test_that("reading xf node extLst works", {
  xml <- "<xf borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"><extLst><ext><foo/></ext></extLst></xf>"
  xf  <- read_xml(xml)

  df_xf <- read_xf(xml_doc_xf = xf)
  got   <- write_xf(df_xf)
  expect_equal(xml, got)

  df_xf$extLst <- "<extLst></foo/></extLst>"
  expect_error(write_xf(df_xf), "failed to load xf child")
})

## FIXME silent tests should be replaced with something better
test_that("updating borders works", {

  # A1:B10, C1:D10
  wb <- wb_workbook()$add_worksheet()$
    add_border(dims = "A1:B10")$
    add_border(dims = "C1:D10")

  # A1:D1, and another at A4:D4
  wb$add_border(dims = "A1:D1", update = TRUE)
  exp <- c(A1 = "13", B1 = "15", C1 = "16", D1 = "14")
  got <- wb$get_cell_style(dims = "A1:D1")
  expect_equal(exp, got)

  wb$add_border(dims = "A4:D4", top_border = NULL, update = TRUE)
  exp <- c(A4 = "17", B4 = "19", C4 = "20", D4 = "18")
  got <- wb$get_cell_style(dims = "A4:D4")
  expect_equal(exp, got)

  wb$add_worksheet()$
    add_border(dims = "B2:D4", bottom_border = "thick", left_border = "thick", right_border = "thick", top_border = "thick")

  expect_silent(wb$add_border(dims = "C3:E5", update = TRUE))

  ## check single cell
  wb$add_worksheet()$
    add_border(dims = "B2:B4", bottom_border = "thick", left_border = "thick", right_border = "thick", top_border = "thick")

  # to update the inner cell, both the style and the color must be NULL
  expect_silent(
    wb$add_border(dims = "B3",
                  top_border = NULL, left_border = NULL, right_border = NULL, bottom_border = "double", bottom_color = wb_color("blue"),
                  update = TRUE)
  )

  # update it a second time
  expect_silent(
    wb$add_border(dims = "B3:B3",
                  bottom_border = NULL, left_border = NULL, right_border = NULL, top_border = "dashed", top_color = wb_color("red"),
                  update = TRUE)
  )

  # horizontal overlap
  wb$add_worksheet()$
    add_border(dims = "B2:E2", bottom_border = "thick", left_border = "thick", right_border = "thick", top_border = "thick")

  expect_silent(
    wb$add_border(dims = "C2:F2", bottom_border = "dashed", left_border = "dashed", right_border = "dashed", top_border = "dashed", update = TRUE)
  )

  # vertical overlap
  wb$add_worksheet()$
    add_border(dims = "B2:B5", bottom_border = "thick", left_border = "thick", right_border = "thick", top_border = "thick")

  expect_silent(
    wb$add_border(dims = "B3:B6", bottom_border = "dashed", left_border = "dashed", right_border = "dashed", top_border = "dashed", update = TRUE)
  )

})

test_that("create_font() works", {
  got <- create_font(
    b = TRUE, # Boolean
    charset = "1", # Numeric-as-string
    color = wb_color(hex = "FFDDAA00"), # wbColour object (ARGB yellow)
    condense = TRUE, # Boolean
    extend = TRUE, # Boolean
    family = "10", # Numeric-as-string (must be 0-14)
    i = TRUE, # Boolean
    name = "Impact", # String
    outline = TRUE, # Boolean
    scheme = "major", # From 'minor', 'major', 'none'
    shadow = TRUE, # Boolean
    strike = TRUE, # Boolean
    sz = 36, # Font size
    u = "double", # From valid underline list
    vert_align = "superscript" # From valid alignment list
  )

  exp <- "<font><b val=\"1\"/><charset val=\"1\"/><color rgb=\"FFDDAA00\"/><condense val=\"1\"/><extend val=\"1\"/><family val=\"10\"/><i val=\"1\"/><name val=\"Impact\"/><outline val=\"1\"/><scheme val=\"major\"/><shadow val=\"1\"/><strike val=\"1\"/><sz val=\"36\"/><u val=\"double\"/><vertAlign val=\"superscript\"/></font>"
  expect_equal(exp, got)
})
