test_that("read_xlsx from different sources", {

  skip_online_checks()

  ## URL
  xlsxFile <- "https://github.com/JanMarvin/openxlsx2/raw/main/inst/extdata/oxlsx2_sheet.xlsx"
  df_url <- read_xlsx(xlsxFile)

  ## File
  xlsxFile <- system.file("extdata", "oxlsx2_sheet.xlsx", package = "openxlsx2")
  df_file <- read_xlsx(xlsxFile)

  expect_equal(df_url, df_file, label = "Read from URL")


  ## Non-existing URL
  xlsxFile <- "https://github.com/JanMarvin/openxlsx2/raw/main/inst/extdata/readTest2.xlsx"
  expect_error(suppressWarnings(read_xlsx(xlsxFile)))


  ## Non-existing File
  xlsxFile <- file.path(dirname(system.file("extdata", "readTest.xlsx", package = "openxlsx2")), "readTest00.xlsx")
  expect_error(read_xlsx(xlsxFile), regexp = "File does not exist.")
})


test_that("wb_load from different sources", {

  skip_online_checks()

  ## URL
  xlsxFile <- "https://github.com/JanMarvin/openxlsx2/raw/main/inst/extdata/oxlsx2_sheet.xlsx"
  wb_url <- wb_load(xlsxFile)

  ## File
  xlsxFile <- system.file("extdata", "oxlsx2_sheet.xlsx", package = "openxlsx2")
  wb_file <- wb_load(xlsxFile)

  # Loading from URL vs local not equal
  expect_equal_workbooks(wb_url, wb_file, ignore_fields = c("path"))
})


test_that("get_date_origin from different sources", {

  xlsxFile <- testfile_path("readTest.xlsx")
  origin_url <- get_date_origin(xlsxFile)

  expect_equal(origin_url, "1900-01-01")
})


test_that("read html source without r attribute on cell", {

  # sheet without row attribute
  # original from https://www.atih.sante.fr/sites/default/files/public/content/3968/fichier_complementaire_ccam_descriptive_a_usage_pmsi_2021_v2.xlsx
  wb <- wb_load(testfile_path("fichier_complementaire_ccam_descriptive_a_usage_pmsi_2021_v2.xlsx"))

  expect_equal(dim(wb_to_df(wb, sheet = 1)), c(46L, 1L))
  expect_equal(dim(wb_to_df(wb, sheet = 2)), c(31564L, 52L))
  expect_equal(names(wb_to_df(wb, sheet = 1)), "PRÉSENTATION DU DOCUMENT")

  # This file has a few cells with row names, the majority has none. check that
  # we did not create duplicates while loading
  expect_false(any(duplicated(wb$worksheets[[1]]$sheet_data$cc)))
  expect_false(any(duplicated(wb$worksheets[[2]]$sheet_data$cc)))

})


test_that("read <br> node in vml", {

  # test
  expect_silent(wb <- wb_load(testfile_path("macro2.xlsm")))

})


test_that("read vml from sheet two works as expected", {

  # test
  expect_silent(wb <- wb_load(testfile_path("vml_numbering.xlsx")))

  expect_equal(length(wb$vml), 1L)
  expect_equal(length(wb$vml_rels), 1L)

})


test_that("encoding", {

  fl <- testfile_path("umlauts.xlsx")
  wb <- wb_load(fl)
  expect_equal("äöüß", names(wb$get_sheet_names()))

  exp <- structure(list("hähä" = "ÄÖÜ", "höhö" = "äöüß"),
                   row.names = 2L, class = "data.frame",
                   tt = structure(list("hähä" = "s", "höhö" = "s"),
                                  row.names = 2L, class = "data.frame"),
                   types = c(A = 0, B = 0))

  expect_equal(wb_to_df(wb, keep_attributes = TRUE), exp)

  fl <- testfile_path("eurosymbol.xlsx")
  wb <- wb_load(fl)
  got <- wb$sharedStrings

  if (isTRUE(l10n_info()[["UTF-8"]])) {
    exp <- structure(
      c("<si><t xml:space=\"preserve\">€</t></si>",
        "<si><t xml:space=\"preserve\">ä</t></si>"),
      uniqueCount = "2", text = c("€", "ä")
    )
  } else {
    exp <- structure(
      c("<si><t xml:space=\"preserve\">\200</t></si>",
        "<si><t xml:space=\"preserve\">ä</t></si>"),
      uniqueCount = "2", text = c("\200", "ä")
    )
  }

  expect_equal(exp, got)


  exp <- "<xml>\n <a0>äöüß</a0>\n <A0>ÄÖÜ</A0>\n <a1>€</a1>\n</xml>"
  got <- paste(capture.output(
    read_xml(system.file("extdata", "unicode.xml", package = "openxlsx2"))
  ), collapse = "\n")
  expect_equal(exp, got)

  exp <- "<xml><a0>äöüß</a0><A0>ÄÖÜ</A0><a1>€</a1></xml>"
  got <- read_xml(system.file("extdata", "unicode.xml", package = "openxlsx2"),
                  pointer = FALSE)
  expect_equal(exp, got)

})

test_that("reading charts", {

  temp <- temp_xlsx()

  wb <- wb_load(testfile_path("unemployment-nrw202208.xlsx"))

  exp <- c("", "", "", "", "", "", "", "", "", "", "", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes\" Target=\"../drawings/drawing18.xml\"/></Relationships>", "", "", "")
  got <- wb$charts$rels
  expect_equal(exp, got)

  img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")

  which(wb$get_sheet_names() == "Uebersicht_Quoten")
  expect_warning(
    wb$add_image(19, file = img, startRow = 5, startCol = 3, width = 6, height = 5),
    "'start_col/start_row' is deprecated."
  )

  wb$save(temp)

  # check that we wrote a chartshape
  xlsx_unzip <- paste0(tempdir(), "/unzip")
  dir.create(xlsx_unzip)


  unzip(temp, exdir = xlsx_unzip)
  overrides <- xml_node(read_xml(paste0(xlsx_unzip, "/[Content_Types].xml"), pointer = FALSE), "Types", "Override")
  expect_match(overrides, "chartshapes", all = FALSE)

  unlink(xlsx_unzip, recursive = TRUE)


  # check that the image is valid and was placed on the correct sheet and drawing
  exp <- c(
    "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"#Inhaltsverzeichnis!A1\"/>",
    "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" Target=\"../charts/chart14.xml\"/>",
    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image6.jpeg\"/>",
    "<Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image7.jpg\"/>"
  )
  got <- wb$drawings_rels[[20]]
  expect_equal(exp, got)


  wb$add_worksheet()
  wb$save(temp)

  # remove worksheet
  wb <- wb_load(testfile_path("unemployment-nrw202208.xlsx"))
  rmsheet <- length(wb$worksheets) - 2
  wb$remove_worksheet(rmsheet)

  expect_no_match(unlist(wb$worksheets_rels), "drawing21.xml")
  expect_equal(wb$drawings[[21]], "")
  expect_equal(wb$drawings_rels[[21]], "")

})

test_that("load file with xml namespace", {

  skip_online_checks()

  fl <- "https://github.com/ycphs/openxlsx/files/8480120/2022-04-12-11-42-36-DP_Melanges1.xlsx"

  expect_warning(
    wb <- wb_load(fl),
    "has been removed from the xml files, for example"
  )
  expect_null(getOption("openxlsx2.namespace_xml"))

})

test_that("reading file with macro and custom xml", {

  temp <- temp_xlsx()

  wb <- wb_load(testfile_path("gh_issue_416.xlsm"))
  wb$save(temp)
  wb <- wb_load(temp)

  exp <- "<sheetPr codeName=\"Sheet1\"/>"
  got <- wb$worksheets[[1]]$sheetPr
  expect_equal(exp, got)

  exp <- "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"2\" name=\"Source\"><vt:lpwstr>openxlsx2</vt:lpwstr></property></Properties>"
  got <- wb$custom
  expect_equal(exp, got)
})

test_that("load file with connection", {

  temp <- temp_xlsx()

  wb <- wb_load(testfile_path("connection.xlsx"))

  expect_false(is.null(wb$customXml))
  expect_length(wb$customXml, 3)

  wb$save(temp)

  wb <- wb_load(temp)
  expect_length(wb$customXml, 3)
  expect_match(wb$customXml[1], "customXml/_rels/item1.xml.rels")
  expect_match(wb$customXml[2], "customXml/item1.xml")
  expect_match(wb$customXml[3], "customXml/itemProps1.xml")

})

test_that("calcChain is updated", {

  temp <- temp_xlsx()

  fl <- testfile_path("overwrite_formula.xlsx")

  wb <- wb_load(fl, calc_chain = TRUE)

  exp <- "<calcChain xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><c r=\"A1\" i=\"1\" l=\"1\"/></calcChain>"
  got <- wb$calcChain
  expect_equal(exp, got)

  expect_silent(wb$save(temp))

  wb$add_data(dims = "A1", x = "Formula overwritten")

  exp <- character()
  got <- wb$calcChain
  expect_equal(exp, got)

  expect_silent(wb$save(temp))

})

test_that("read workbook with chart extension", {

  fl <- testfile_path("charts.xlsx")

  wb <- wb_load(fl)

  expect_warning(
    wb$clone_worksheet(),
    "The file you have loaded contains chart extensions. At the moment, cloning worksheets can damage the output."
  )

})

test_that("reading of formControl works", {

  temp <- temp_xlsx()

  wb <- wb_load(testfile_path("form_control.xlsx"))

  exp <- c(
    "<formControlPr xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" objectType=\"CheckBox\" checked=\"Checked\" lockText=\"1\" noThreeD=\"1\"/>",
    "<formControlPr xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" objectType=\"Radio\" checked=\"Checked\" firstButton=\"1\" lockText=\"1\" noThreeD=\"1\"/>",
    "<formControlPr xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" objectType=\"Radio\" lockText=\"1\" noThreeD=\"1\"/>",
    "<formControlPr xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" objectType=\"CheckBox\" lockText=\"1\" noThreeD=\"1\"/>",
    "<formControlPr xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" objectType=\"Drop\" dropStyle=\"combo\" dx=\"15\" noThreeD=\"1\" sel=\"0\" val=\"0\"/>"
  )
  got <- wb$ctrlProps
  expect_equal(exp, got)

  wb$save(temp)
  wb <- wb_load(temp)

  got <- wb$ctrlProps
  expect_equal(exp, got)

})

test_that("reading xml escapes works", {

  skip_online_checks()

  fl <- "https://github.com/ycphs/openxlsx/files/10032200/sample_data.xlsx"
  wb <- wb_load(fl)

  df <- wb_to_df(wb)

  exp <- "US & Canada"
  got <- unique(df$colB)
  expect_equal(exp, got)


  df <- wb_to_df(wb, showFormula = TRUE)

  exp <- c("US & Canada", "B2")
  got <- df$colB
  expect_equal(exp, got)

})

test_that("reading multiple slicers on a pivot table works", {

  wb <- wb_load(testfile_path("gh_issue_504.xlsx"))

  expect_equal(length(wb$slicers), 1L)
  expect_equal(length(wb$slicerCaches), 2L)

  exp <- c(
    "<Override PartName=\"/xl/slicers/slicer1.xml\" ContentType=\"application/vnd.ms-excel.slicer+xml\"/>",
    "<Override PartName=\"/xl/slicerCaches/slicerCache1.xml\" ContentType=\"application/vnd.ms-excel.slicerCache+xml\"/>",
    "<Override PartName=\"/xl/slicerCaches/slicerCache2.xml\" ContentType=\"application/vnd.ms-excel.slicerCache+xml\"/>"
  )
  got <- wb$Content_Types[14:16]
  expect_equal(exp, got)

  exp <- c(
    "<Override PartName=\"/xl/slicers/slicer1.xml\" ContentType=\"application/vnd.ms-excel.slicer+xml\"/>",
    "<Override PartName=\"/xl/slicerCaches/slicerCache1.xml\" ContentType=\"application/vnd.ms-excel.slicerCache+xml\"/>",
    "<Override PartName=\"/xl/slicerCaches/slicerCache2.xml\" ContentType=\"application/vnd.ms-excel.slicerCache+xml\"/>"
  )
  got <- wb$Content_Types[14:16]
  expect_equal(exp, got)

  exp <- c(
    "<Relationship Id=\"rId20001\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition\" Target=\"pivotCache/pivotCacheDefinition1.xml\"/>",
    "<Relationship Id=\"rId100001\" Type=\"http://schemas.microsoft.com/office/2007/relationships/slicerCache\" Target=\"slicerCaches/slicerCache1.xml\"/>",
    "<Relationship Id=\"rId100002\" Type=\"http://schemas.microsoft.com/office/2007/relationships/slicerCache\" Target=\"slicerCaches/slicerCache2.xml\"/>"
  )
  got <- wb$workbook.xml.rels[6:8]
  expect_equal(exp, got)
})

test_that("reading slicer for tables works", {
  fl <- testfile_path("table_slicer.xlsx")
  wb <- wb_load(fl)
  expect_match(wb$workbook$extLst, "<x14:slicerCache r:id=\"rId100001\"/>")
})

test_that("hyperlinks work", {

  tmp <- temp_xlsx()
  wb_load(testfile_path("Single_hyperlink.xlsx"))$save(tmp)

  temp_uzip <- paste0(tempdir(), "/unzip_openxlsx2")
  dir.create(temp_uzip)
  unzip(tmp, exdir = temp_uzip)

  exp <- "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1h\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://www.github.com/JanMarvin/openxlsx2\" TargetMode=\"External\"/></Relationships>"
  got <- read_xml(paste0(temp_uzip, "/xl/worksheets/_rels/sheet1.xml.rels"), pointer = FALSE)
  expect_equal(exp, got)

  unlink(temp_uzip, recursive = TRUE)
})

test_that("reading richData content works", {

  tmp <- temp_xlsx()
  expect_silent(wb_load(testfile_path("pic_in_cell.xlsx"))$save(tmp))

})

test_that("reading timeline works", {

  tmp <- temp_xlsx()
  fl  <- testfile_path("timeline.xlsx")

  wb <- wb_load(fl)
  expect_true(wb$worksheets[[2]]$relships$timeline == 1)

  wb$save(tmp)

  # save again, so that all sheets in workbook.xml.rels have a correct id. Not
  # sure why we keep the update rel ids once we have saved a file saving.
  wb2 <- wb_load(tmp)
  wb2$save(tmp)

  expect_true(wb$timelines == wb2$timelines)
  expect_true(all(wb$Content_Types %in% wb2$Content_Types))
  expect_true(all(wb$workbook.xml.rels == wb2$workbook.xml.rels))

  expect_warning(
    wb$clone_worksheet(old = "Sheet 2", new = "Sheet3"),
    "Cloning timelines is not yet supported. It will not appear on the sheet."
  )

})
