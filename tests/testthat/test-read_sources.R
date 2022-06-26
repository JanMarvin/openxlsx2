test_that("read_xlsx from different sources", {

  ## URL
  xlsxFile <- "https://github.com/ycphs/openxlsx/raw/master/inst/extdata/readTest.xlsx"
  df_url <- read_xlsx(xlsxFile)

  ## File
  xlsxFile <- system.file("extdata", "readTest.xlsx", package = "openxlsx2")
  df_file <- read_xlsx(xlsxFile)

  expect_true(all.equal(df_url, df_file), label = "Read from URL")


  ## Non-existing URL
  xlsxFile <- "https://github.com/ycphs/openxlsx/raw/master/inst/extdata/readTest2.xlsx"
  expect_error(suppressWarnings(read_xlsx(xlsxFile)))


  ## Non-existing File
  xlsxFile <- file.path(dirname(system.file("extdata", "readTest.xlsx", package = "openxlsx2")), "readTest00.xlsx")
  expect_error(read_xlsx(xlsxFile), regexp = "File does not exist.")
})


test_that("wb_load from different sources", {

  ## URL
  xlsxFile <- "https://github.com/ycphs/openxlsx/raw/master/inst/extdata/readTest.xlsx"
  wb_url <- wb_load(xlsxFile)

  ## File
  xlsxFile <- system.file("extdata", "readTest.xlsx", package = "openxlsx2")
  wb_file <- wb_load(xlsxFile)

  # Loading from URL vs local not equal
  expect_equal_workbooks(wb_url, wb_file, ignore_fields = "datetimeCreated")
})


test_that("get_date_origin from different sources", {

  ## URL
  xlsxFile <- "https://github.com/ycphs/openxlsx/raw/master/inst/extdata/readTest.xlsx"
  origin_url <- get_date_origin(xlsxFile)

  ## File
  xlsxFile <- system.file("extdata", "readTest.xlsx", package = "openxlsx2")
  origin_file <- get_date_origin(xlsxFile)

  ## check
  expect_equal(origin_url, origin_file)
  expect_equal(origin_url, "1900-01-01")
})


test_that("read html source without r attribute on cell", {

  # sheet without row attribute
  # original from https://www.atih.sante.fr/sites/default/files/public/content/3968/fichier_complementaire_ccam_descriptive_a_usage_pmsi_2021_v2.xlsx
  wb <- wb_load("https://github.com/JanMarvin/openxlsx2/files/8702731/fichier_complementaire_ccam_descriptive_a_usage_pmsi_2021_v2.xlsx")

  expect_equal(c(46, 1), dim(wb_to_df(wb, sheet = 1)))
  expect_equal(c(31564, 52), dim(wb_to_df(wb, sheet = 2)))
  expect_equal("PRÉSENTATION DU DOCUMENT", names(wb_to_df(wb, sheet = 1)))

  # This file has a few cells with row names, the majority has none. check that
  # we did not create duplicates while loading
  expect_true(!any(duplicated(wb$worksheets[[1]]$sheet_data$cc)))
  expect_true(!any(duplicated(wb$worksheets[[2]]$sheet_data$cc)))

})


test_that("read <br> node in vml", {
  # prepare test
  temp_dir <- tempfile()
  temp_file <- file.path(temp_dir, "macro2.xlsm")
  temp_zip  <- paste0(temp_file, ".zip")
  dir.create(temp_dir, recursive = TRUE, showWarnings = FALSE)
  download.file("https://github.com/JanMarvin/openxlsx2/files/8773595/macro2.xlsm.zip", temp_zip)
  zip::unzip(temp_zip, exdir = temp_dir)

  # test
  expect_silent(wb <- wb_load(temp_file))

  # clean up
  unlink(temp_dir, recursive = TRUE, force = TRUE)
})


test_that("encoding", {

  fl <- "https://github.com/JanMarvin/openxlsx2/files/8779041/umlauts.xlsx"
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


  exp <- "<xml>\n<a0>äöüß</a0>\n<A0>ÄÖÜ</A0>\n<a1>€</a1>\n</xml>"
  got <- paste(capture.output(
    read_xml(system.file("extdata", "unicode.xml", package = "openxlsx2")
    )), collapse = "\n")
  expect_equal(exp, got)

  got <- read_xml(system.file("extdata", "unicode.xml", package = "openxlsx2"),
                  pointer = FALSE)
  expect_equal(exp, got)

})
