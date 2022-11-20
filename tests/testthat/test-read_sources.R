test_that("read_xlsx from different sources", {

  skip_if_offline()

  ## URL
  xlsxFile <- "https://github.com/JanMarvin/openxlsx2/raw/main/inst/extdata/readTest.xlsx"
  df_url <- read_xlsx(xlsxFile)

  ## File
  xlsxFile <- system.file("extdata", "readTest.xlsx", package = "openxlsx2")
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

  skip_if_offline()

  ## URL
  xlsxFile <- "https://github.com/JanMarvin/openxlsx2/raw/main/inst/extdata/readTest.xlsx"
  wb_url <- wb_load(xlsxFile)

  ## File
  xlsxFile <- system.file("extdata", "readTest.xlsx", package = "openxlsx2")
  wb_file <- wb_load(xlsxFile)

  # Loading from URL vs local not equal
  expect_equal_workbooks(wb_url, wb_file, ignore_fields = "datetimeCreated")
})


test_that("get_date_origin from different sources", {

  skip_if_offline()

  ## URL
  xlsxFile <- "https://github.com/JanMarvin/openxlsx2/raw/main/inst/extdata/readTest.xlsx"
  origin_url <- get_date_origin(xlsxFile)

  ## File
  xlsxFile <- system.file("extdata", "readTest.xlsx", package = "openxlsx2")
  origin_file <- get_date_origin(xlsxFile)

  ## check
  expect_equal(origin_url, origin_file)
  expect_equal(origin_url, "1900-01-01")
})


test_that("read html source without r attribute on cell", {

  skip_if_offline()

  # sheet without row attribute
  # original from https://www.atih.sante.fr/sites/default/files/public/content/3968/fichier_complementaire_ccam_descriptive_a_usage_pmsi_2021_v2.xlsx
  wb <- wb_load("https://github.com/JanMarvin/openxlsx-data/raw/main/fichier_complementaire_ccam_descriptive_a_usage_pmsi_2021_v2.xlsx")

  expect_equal(c(46, 1), dim(wb_to_df(wb, sheet = 1)))
  expect_equal(c(31564, 52), dim(wb_to_df(wb, sheet = 2)))
  expect_equal("PRÉSENTATION DU DOCUMENT", names(wb_to_df(wb, sheet = 1)))

  # This file has a few cells with row names, the majority has none. check that
  # we did not create duplicates while loading
  expect_true(!any(duplicated(wb$worksheets[[1]]$sheet_data$cc)))
  expect_true(!any(duplicated(wb$worksheets[[2]]$sheet_data$cc)))

})


test_that("read <br> node in vml", {

  skip_if_offline()

  # test
  expect_silent(wb <- wb_load("https://github.com/JanMarvin/openxlsx-data/raw/main/macro2.xlsm"))
})


test_that("encoding", {

  skip_if_offline()

  fl <- "https://github.com/JanMarvin/openxlsx-data/raw/main/umlauts.xlsx"
  wb <- wb_load(fl)
  expect_equal("äöüß", names(wb$get_sheet_names()))

  exp <- structure(list("hähä" = "ÄÖÜ", "höhö" = "äöüß"),
                   row.names = 2L, class = "data.frame",
                   tt = structure(list("hähä" = "s", "höhö" = "s"),
                                  row.names = 2L, class = "data.frame"),
                   types = c(A = 0, B = 0))

  expect_equal(exp, wb_to_df(wb))

  fl <- system.file("extdata", "eurosymbol.xlsx", package = "openxlsx2")
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

  skip_if_offline()

  temp <- temp_xlsx()

  wb <- wb_load("https://github.com/JanMarvin/openxlsx-data/raw/main/unemployment-nrw202208.xlsx")

  exp <- c("", "", "", "", "", "", "", "", "", "", "", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes\" Target=\"../drawings/drawing18.xml\"/></Relationships>", "", "", "")
  got <- wb$charts$rels
  expect_equal(exp, got)

  img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")

  which(wb$get_sheet_names() == "Uebersicht_Quoten")
  wb$add_image(19, img, startRow = 5, startCol = 3, width = 6, height = 5)
  wb$save(temp)

  # check that we wrote a chartshape
  xlsx_unzip <- paste0(tempdir(), "/unzip")
  dir.create(xlsx_unzip)


  unzip(temp, exdir = xlsx_unzip)
  overrides <- xml_node(read_xml(paste0(xlsx_unzip, "/[Content_Types].xml"), pointer = FALSE), "Types", "Override")
  unlink(xlsx_unzip, recursive = TRUE)

  expect_true(any(grepl("chartshapes", overrides)))

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
  wb <- wb_load("https://github.com/JanMarvin/openxlsx-data/raw/main/unemployment-nrw202208.xlsx")
  rmsheet <- length(wb$worksheets) - 2
  wb$remove_worksheet(rmsheet)

  expect_false(any(grepl("drawing21.xml", unlist(wb$worksheets_rels))))
  expect_equal("", wb$drawings[[21]])
  expect_equal("", wb$drawings_rels[[21]])

})

test_that("load file with xml namespace", {

  skip_if_offline()

  fl <- "https://github.com/ycphs/openxlsx/files/8480120/2022-04-12-11-42-36-DP_Melanges1.xlsx"

  expect_warning(
    wb <- wb_load(fl),
    "has been removed from the xml files, for example"
  )
  expect_null(getOption("openxlsx2.namespace_xml"))

})

test_that("reading file with macro and custom xml", {

  skip_if_offline()

  temp <- temp_xlsx()

  wb <- wb_load("https://github.com/JanMarvin/openxlsx-data/raw/main/gh_issue_416.xlsm")
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

  skip_if_offline()

  temp <- temp_xlsx()

  wb <- wb_load("https://github.com/JanMarvin/openxlsx-data/raw/main/connection.xlsx")

  expect_true(!is.null(wb$customXml))
  expect_equal(3, length(wb$customXml))

  wb$save(temp)

  wb <- wb_load(temp)
  expect_equal(3, length(wb$customXml))

  expect_true(grepl("customXml/_rels/item1.xml.rels", wb$customXml[1]))
  expect_true(grepl("customXml/item1.xml", wb$customXml[2]))
  expect_true(grepl("customXml/itemProps1.xml", wb$customXml[3]))

})

test_that("calcChain is updated", {

  skip_if_offline()

  fl <- "https://github.com/JanMarvin/openxlsx-data/raw/main/overwrite_formula.xlsx"

  wb <- wb_load(fl)$
    add_data(dims = "A1", x = "Formula overwritten")

  exp <- character()
  got <- wb$calcChain
  expect_equal(exp, got)

})

test_that("read workbook with chart extension", {

  skip_if_offline()

  fl <- "https://github.com/JanMarvin/openxlsx-data/raw/main/charts.xlsx"

  wb <- wb_load(fl)

  expect_warning(
    wb$clone_worksheet(),
    "The file you have loaded contains chart extensions. At the moment, cloning worksheets can damage the output."
  )

})

test_that("reading of formControl works", {

  skip_if_offline()

  temp <- temp_xlsx()

  wb <- wb_load("https://github.com/JanMarvin/openxlsx-data/raw/main/form_control.xlsx")

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
