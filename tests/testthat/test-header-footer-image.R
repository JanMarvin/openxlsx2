img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")

test_that("images land in the section they were asked for", {
  wb <- wb_workbook()$add_worksheet()
  for (loc in c("header", "footer")) {
    for (pos in c("left", "center", "right")) {
      wb$add_header_footer_image(file = img, position = pos, location = loc)
    }
  }

  vml <- wb$vml[[wb$worksheets[[1]]$relships$vmlDrawingHF]]
  ids <- regmatches(vml, gregexpr('(?<=<v:shape id=")[^"]+', vml, perl = TRUE))[[1]]
  expect_equal(ids, c("LH", "CH", "RH", "LF", "CF", "RF"))

  hf <- wb$worksheets[[1]]$headerFooter
  expect_equal(hf$oddHeader, rep("&amp;G", 3))
  expect_equal(hf$oddFooter, rep("&amp;G", 3))

  expect_length(wb$media, 6L)
  expect_length(wb$vml_rels[[1]], 6L)
})

test_that("the shape carries the size and the locks Excel writes", {
  wb <- wb_workbook()$add_worksheet()
  wb$add_header_footer_image(file = img, position = "right", width = 2, height = 1)

  vml <- wb$vml[[1]]
  expect_match(vml, "width:144pt;height:72pt;z-index:1", fixed = TRUE)
  expect_match(vml, '<o:lock v:ext="edit" rotation="t" aspectratio="f"/>', fixed = TRUE)
  expect_match(vml, '<v:imagedata o:relid="rId1" o:title="einstein"/>', fixed = TRUE)

  wb2 <- wb_workbook()$add_worksheet()
  wb2$add_header_footer_image(file = img, position = "right", width = 72, height = 36,
                              units = "pt")
  expect_match(wb2$vml[[1]], "width:72pt;height:36pt", fixed = TRUE)
})

test_that("a section holds a single image", {
  wb <- wb_workbook()$add_worksheet()
  wb$add_header_footer_image(file = img, position = "right")
  expect_error(wb$add_header_footer_image(file = img, position = "right"),
               "already holds an image")
  expect_silent(wb$add_header_footer_image(file = img, position = "right",
                                           location = "footer"))
})

test_that("parts, content types and relationships are complete", {
  wb <- wb_workbook()$add_worksheet()
  wb$add_header_footer_image(file = img, position = "center")
  wb$add_header_footer_image(file = img, position = "center", location = "footer")

  expect_length(grep('Extension="vml"', wb$Content_Types), 1L)
  expect_length(grep('Extension="jpg" ContentType="image/jpeg"', wb$Content_Types), 1L)

  rel <- grep("vmlDrawing", wb$worksheets_rels[[1]], value = TRUE)
  expect_length(rel, 1L)
  expect_equal(wb$worksheets[[1]]$legacyDrawingHF,
               sprintf('<legacyDrawingHF r:id="%s"/>',
                       sub('.*Id="(rId\\d+)".*', "\\1", rel)))
  expect_match(wb$vml_rels[[1]][1], 'Target="../media/image1.jpg"', fixed = TRUE)

  fl <- temp_xlsx()
  wb$save(fl)
  parts <- utils::unzip(fl, list = TRUE)$Name
  expect_true(all(c("xl/drawings/vmlDrawing1.vml",
                    "xl/drawings/_rels/vmlDrawing1.vml.rels",
                    "xl/media/image1.jpg") %in% parts))

  exdir <- tempfile()
  dir.create(exdir)
  utils::unzip(fl, exdir = exdir)
  sheet <- paste(readLines(file.path(exdir, "xl/worksheets/sheet1.xml"), warn = FALSE),
                 collapse = "")
  expect_match(sheet, "<oddHeader>&amp;C&amp;G</oddHeader>", fixed = TRUE)
  expect_match(sheet, "<oddFooter>&amp;C&amp;G</oddFooter>", fixed = TRUE)
  expect_match(sheet, '<legacyDrawingHF r:id="rId1"/>', fixed = TRUE)
})

test_that("header images and comments use separate vml parts", {
  wb <- wb_workbook()$add_worksheet()
  wb$add_comment(dims = "A1", comment = "note")
  wb$add_header_footer_image(file = img, position = "right")

  ws <- wb$worksheets[[1]]
  expect_false(identical(ws$relships$vmlDrawing, ws$relships$vmlDrawingHF))
  expect_length(wb$vml, 2L)
  expect_match(wb$vml[[ws$relships$vmlDrawingHF]], "_x0000_t75", fixed = TRUE)
  expect_true(nzchar(ws$legacyDrawingHF))
  expect_length(grep("vmlDrawing", wb$worksheets_rels[[1]]), 2L)
})

test_that("bad input is rejected", {
  wb <- wb_workbook()$add_worksheet()

  expect_error(wb$add_header_footer_image(file = "no-such-file.png"), "does not exist")

  txt <- tempfile(fileext = ".txt")
  writeLines("x", txt)
  expect_error(wb$add_header_footer_image(file = txt), "not an image type")

  expect_error(wb$add_header_footer_image(file = img, position = "middle"))
  expect_error(wb$add_header_footer_image(file = img, units = "furlong"))
})

test_that("a loaded workbook keeps header images apart from comments", {
  wb <- wb_workbook()$add_worksheet()
  wb$add_comment(dims = "A1", comment = "note")
  wb$add_header_footer_image(file = img, position = "right")
  fl <- temp_xlsx()
  wb$save(fl)

  wb2 <- wb_load(fl)
  expect_equal(wb2$worksheets[[1]]$relships$vmlDrawing, 1L)
  expect_equal(wb2$worksheets[[1]]$relships$vmlDrawingHF, 2L)

  # a second image has to extend the part that is already there
  wb2$add_header_footer_image(file = img, position = "left")
  expect_length(wb2$vml, 2L)
  expect_equal(wb2$worksheets[[1]]$legacyDrawingHF, '<legacyDrawingHF r:id="rId3"/>')
  expect_length(grep("vmlDrawing", wb2$worksheets_rels[[1]]), 2L)

  fl2 <- temp_xlsx()
  wb2$save(fl2)
  exdir <- tempfile()
  dir.create(exdir)
  utils::unzip(fl2, exdir = exdir)
  sheet <- paste(readLines(file.path(exdir, "xl/worksheets/sheet1.xml"), warn = FALSE),
                 collapse = "")
  expect_length(gregexpr("legacyDrawingHF", sheet, fixed = TRUE)[[1]], 1L)
  expect_match(sheet, "<oddHeader>&amp;L&amp;G&amp;R&amp;G</oddHeader>", fixed = TRUE)
})

test_that("a workbook without comments loads the header part alone", {
  wb <- wb_workbook()$add_worksheet()
  wb$add_header_footer_image(file = img, position = "center")
  fl <- temp_xlsx()
  wb$save(fl)

  wb2 <- wb_load(fl)
  expect_length(wb2$worksheets[[1]]$relships$vmlDrawing, 0L)
  expect_equal(wb2$worksheets[[1]]$relships$vmlDrawingHF, 1L)
})

test_that("a cloned sheet gets a header part of its own", {
  wb <- wb_workbook()$add_worksheet("a")
  wb$add_header_footer_image(file = img, position = "right")
  wb$clone_worksheet("a", "b")

  expect_length(wb$vml, 2L)
  expect_equal(wb$worksheets[[1]]$relships$vmlDrawingHF, 1L)
  expect_equal(wb$worksheets[[2]]$relships$vmlDrawingHF, 2L)

  # a picture added to the copy must not turn up in the header of the original
  wb$add_header_footer_image(sheet = "b", file = img, position = "left")
  expect_length(gregexpr("<v:shape ", wb$vml[[1]], fixed = TRUE)[[1]], 1L)
  expect_length(gregexpr("<v:shape ", wb$vml[[2]], fixed = TRUE)[[1]], 2L)
  expect_equal(wb$worksheets[[1]]$headerFooter$oddHeader, c("", "", "&amp;G"))
})

test_that("removing a sheet drops its header part", {
  wb <- wb_workbook()$add_worksheet("a")$add_worksheet("b")
  wb$add_header_footer_image(sheet = "a", file = img, position = "right")
  wb$remove_worksheet("a")

  fl <- temp_xlsx()
  wb$save(fl)
  expect_false(any(grepl("\\.vml$", utils::unzip(fl, list = TRUE)$Name)))
})
