test_that("reading xlsb works", {

  skip_if(.Platform$endian == "big", "XLSB: skipping on big-endian system")
  skip_online_checks()

  xlsxFile <- testfile_path("openxlsx2_example.xlsb")

  wb <- wb_load(xlsxFile)
  df_xlsb <- wb_to_df(wb)

  xlsx <- system.file("extdata", "openxlsx2_example.xlsx",
                      package = "openxlsx2")
  df_xlsx <- wb_to_df(xlsx)

  expect_equal(df_xlsb, df_xlsx)

  df_xlsb <- wb_to_df(wb, show_formula = TRUE)
  df_xlsx <- wb_to_df(xlsx, show_formula = TRUE)
  expect_equal(df_xlsb, df_xlsx)

})

test_that("reading complex xlsb works", {

  skip_if(.Platform$endian == "big", "XLSB: skipping on big-endian system")
  skip_online_checks()

  xlsxFile <- testfile_path("hyperlink.xlsb")

  wb <- wb_load(xlsxFile)

  # chartsheets
  exp <- c(TRUE, FALSE, FALSE, FALSE)
  got <- wb$is_chartsheet
  expect_equal(exp, got)

  # named regions
  exp <- c("Numbers", "characters", "_xlnm._FilterDatabase")
  got <- unique(wb$get_named_regions()$name)
  expect_equal(exp, got)

  # tables
  exp <- "Table1"
  got <- wb$get_tables(sheet = 2)$tab_name
  expect_equal(exp, got)

  # comments
  exp <- c(
    "<t>Jan Marvin Garbuszus:</t>",
    "<t xml:space=\"preserve\">\n</t>",
    "<t>A new note!</t>"
  )
  got <- wb$comments[[1]][[1]]$comment
  expect_equal(exp, got)

  # hyperlinks
  exp <- "Sheet1!E3"
  got <- xml_attr(wb$worksheets[[2]]$hyperlinks[[1]], "hyperlink")[[1]][["location"]]
  expect_equal(exp, got)

})

test_that("worksheets with real world formulas", {

  skip_if(.Platform$endian == "big", "XLSB: skipping on big-endian system")
  skip_online_checks()

  xlsxFile <- testfile_path("nhs-core-standards-for-eprr-v6.1.xlsb")

  wb <- wb_load(xlsxFile)

  exp <- c("Control", "EPRR Core Standards", "Deep dive",
           "Interoperable capabilities ", "Lookups", "Calculations")
  got <- names(wb$get_sheet_names())
  expect_equal(exp, got)

  xlsxFile <- testfile_path("array_fml.xlsb")

  wb <- wb_load(xlsxFile)

  exp <- structure(
    list(
      A = c("1", "2", NA, "6", "#VALUE!"),
      B = c("2", "3", NA, "a", NA),
      C = c(3, 4, NA, NA, NA),
      D = c(4, 5, NA, NA, NA),
      E = c(5, 6, NA, NA, NA)
    ),
    row.names = c(NA, 5L),
    class = "data.frame"
  )
  got <- wb_to_df(wb, col_names = FALSE)
  expect_equal(exp, got)

  exp <- structure(
    list(
      A = c("1", "A1+1", NA, "SUM({1,2,3})", "SUMIFS(A1:A2,'[1]foo'!B2:B3,{\"A\"})"),
      B = structure(c("A1+1", "B1+1", NA, "{\"a\"}", NA), class = c("character", "formula")),
      C = structure(c("B1+1", "C1+1", NA, NA, NA), class = c("character", "formula")),
      D = structure(c("C1+1", "D1+1", NA, NA, NA), class = c("character", "formula")),
      E = structure(c("D1+1", "E1+1", NA, NA, NA), class = c("character", "formula"))
    ),
    row.names = c(NA, 5L),
    class = "data.frame"
  )
  got <- wb_to_df(wb, col_names = FALSE, show_formula = TRUE)
  expect_equal(exp, got)


  xlsxFile <- testfile_path("match_escape.xlsb")

  wb <- wb_load(xlsxFile)

  exp <- structure(
    "IF(ISNA(MATCH(G681,{\"Annual\",\"\"\"5\"\"\",\"\"\"1+4\"\"\"},0)),1,MATCH(G681,{\"Annual\",\"\"\"5\"\"\",\"\"\"1+4\"\"\"},0))",
    class = c("character", "formula")
  )
  got <- wb_to_df(wb, col_names = FALSE, show_formula = TRUE)$A
  expect_equal(exp, got)

})

test_that("xlsb formulas", {

  skip_if(.Platform$endian == "big", "XLSB: skipping on big-endian system")
  fl <- testfile_path("formula_checks.xlsb")
  wb <- wb_load(fl)

  exp <- c("", "D2:E2", "A1,B1", "A1 A2", "1+1", "1-1", "1*1", "1/1",
          "1%", "1^1", "1=1", "1&gt;1", "1&gt;=1", "1&lt;1", "1&lt;=1",
          "1&lt;&gt;1", "+A3", "-R2", "(1)", "SUM(1, )", "1", "2.5",
          "\"a\"", "\"A\"&amp;\"B\"", "Sheet2!B2", "'[1]Sheet3'!A2")
  got <- unique(wb$worksheets[[1]]$sheet_data$cc$f)
  expect_equal(exp, got)

  fl <- testfile_path("formula_checks.xlsx")
  xl <- wb_load(fl)

  expect_equal(
    xl$worksheets[[2]]$sheet_data$cc$f,
    wb$worksheets[[2]]$sheet_data$cc$f
  )

})

test_that("shared formulas are detected correctly", {

  skip_if(.Platform$endian == "big", "XLSB: skipping on big-endian system")
  xlsb <- testfile_path("formula_checks.xlsb")
  xlsx <- testfile_path("formula_checks.xlsx")

  wbb <- wb_load(xlsb)
  wbx <- wb_load(xlsx)

  expect_equal(
    wbb$worksheets[[4]]$sheet_data$cc,
    wbx$worksheets[[4]]$sheet_data$cc
  )

})

test_that("loading custom sheet view in xlsb files works", {

  skip_if(.Platform$endian == "big", "XLSB: skipping on big-endian system")
  skip_online_checks()

  fl <- testfile_path("custom_sheet_view.xlsb")

  wb <- wb_load(fl)

  exp <- "<customSheetViews><customSheetView guid=\"{101E93B2-5AEC-EA4B-A037-FE6837BBF571}\" view=\"pageLayout\"><selection pane=\"topLeft\" activeCell=\"E23\" sqref=\"E23\"/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><printOptions gridLines=\"1\"/><pageSetup copies=\"1\" firstPageNumber=\"1\" fitToHeight=\"1\" fitToWidth=\"1\" horizontalDpi=\"300\" paperSize=\"9\" scale=\"100\" verticalDpi=\"300\"/></customSheetView><customSheetView guid=\"{29911425-4959-8744-A301-8A3516CFBBC5}\" view=\"pageLayout\"><selection pane=\"topLeft\" activeCell=\"A1\" sqref=\"A1\"/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><printOptions gridLines=\"1\"/><pageSetup copies=\"1\" firstPageNumber=\"1\" fitToHeight=\"1\" fitToWidth=\"1\" horizontalDpi=\"300\" paperSize=\"9\" scale=\"100\" verticalDpi=\"300\"/></customSheetView><customSheetView guid=\"{E70FF13A-F852-294B-B121-766AE66F850F}\"><selection pane=\"topLeft\" activeCell=\"A1\" sqref=\"A1\"/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><printOptions gridLines=\"1\"/><pageSetup copies=\"1\" firstPageNumber=\"1\" fitToHeight=\"1\" fitToWidth=\"1\" horizontalDpi=\"300\" paperSize=\"9\" scale=\"100\" verticalDpi=\"300\"/></customSheetView><customSheetView guid=\"{F233A099-1241-ED48-AE78-14CEAAA87C95}\" hiddenColumns=\"1\"><selection pane=\"topLeft\" activeCell=\"F1\" sqref=\"F1:K1048576\"/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><printOptions gridLines=\"1\"/><pageSetup copies=\"1\" firstPageNumber=\"1\" fitToHeight=\"1\" fitToWidth=\"1\" horizontalDpi=\"300\" paperSize=\"9\" scale=\"100\" verticalDpi=\"300\"/></customSheetView><customSheetView guid=\"{16591214-8D55-8A47-82DC-2F7E1C7FCAA3}\" filter=\"1\" showAutoFilter=\"1\" hiddenColumns=\"1\"><selection pane=\"topLeft\" activeCell=\"A1\" sqref=\"A1\"/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><printOptions gridLines=\"1\"/><pageSetup copies=\"1\" firstPageNumber=\"1\" fitToHeight=\"1\" fitToWidth=\"1\" horizontalDpi=\"300\" paperSize=\"9\" scale=\"100\" verticalDpi=\"300\"/><autoFilter ref=\"A1:K33\"><filterColumn colId=\"0\"><customFilters><customFilter operator=\"greaterThan\" val=\"15\"/></customFilters></filterColumn></autoFilter></customSheetView><customSheetView guid=\"{AB8C9049-5EEE-4D41-AFBB-D05935ABA15C}\" filter=\"1\" showAutoFilter=\"1\" hiddenColumns=\"1\"><selection pane=\"topLeft\" activeCell=\"A1\" sqref=\"A1\"/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><printOptions gridLines=\"1\"/><pageSetup copies=\"1\" firstPageNumber=\"1\" fitToHeight=\"1\" fitToWidth=\"1\" horizontalDpi=\"300\" paperSize=\"9\" scale=\"100\" verticalDpi=\"300\"/><autoFilter ref=\"A1:K33\"><filterColumn colId=\"1\"><filters blank=\"0\"><filter val=\"4\"/></filters></filterColumn></autoFilter></customSheetView></customSheetViews>"
  got <- wb$worksheets[[1]]$customSheetViews
  expect_equal(exp, got)
})

test_that("xlsb formula line breaks are handled", {

  skip_if(.Platform$endian == "big", "XLSB: skipping on big-endian system")
  skip_online_checks()

  fl <- testfile_path("line_break.xlsb")

  fml <- "IF(A1 = 1,\"Value \n\"&A1,)"
  wb <- wb_workbook()$add_worksheet()$
    add_data(x = 1)$
    add_formula(x = fml, dims = "B1")$
    add_cell_style(dims = "B1", wrap_text = TRUE)$
    set_row_heights(rows = 1, heights = 30)

  wb2 <- wb_load(fl)
  exp <- "IF( A1= 1,\"Value \n\"&A1, )"
  fml2 <- wb2$to_df(show_formula = TRUE, col_names = FALSE)[1, "B"]
  expect_equal(exp,  fml2)

  fml  <- gsub(" ", "", fml)
  fml2 <- gsub(" ", "", fml2)

  expect_equal(fml, fml2)
})


test_that("xlsb formula line breaks are handled", {

  skip_if(.Platform$endian == "big", "XLSB: skipping on big-endian system")
  skip_online_checks()

  ## increase the testing a bit more, for now this only checks that the file
  ## can be imported
  fl <- testfile_path("test_coverage.xlsb")
  warns <- capture_warnings(wb <- wb_load(fl))

  unhandled_warns <- grep("Worksheet contains unhandled conditional formatting", warns, value = TRUE)
  expect_length(unhandled_warns, 0)
  expect_equal(length(wb$get_sheet_names()), 3L)

  exp <- c("Align: left", "Align: center", "Align: right")
  got <- wb$to_df(dims = "A1:A3", col_names = FALSE)$A
  expect_equal(got, exp)

  fl <- testfile_path("test_coverage2.xlsb")
  suppressWarnings(wb <- wb_load(fl))
  cc <- wb$worksheets[[2]]$sheet_data$cc

  exp <- c("SUM({1,2,3,4})", "SUM({\"1\",\"2\",\"3\",\"4\"})",
           "SUM({1,#NUM!,3,4})", "SUM({TRUE,FALSE,TRUE})")
  got <- cc$f[cc$r %in% paste0("D", 1:4)]
  expect_equal(got, exp)

})

test_that("pivot tables are imported from xlsb", {

  skip_if(.Platform$endian == "big", "XLSB: skipping on big-endian system")
  skip_online_checks()

  wb <- wb_load(testfile_path("loadExample.xlsb"))

  expect_length(wb$pivotTables, 2L)
  expect_length(wb$pivotDefinitions, 2L)
  expect_match(wb$workbook$pivotCaches, "<pivotCache cacheId=\"24\"", fixed = TRUE)

  expect_match(wb$pivotTables[[1]], "<dataField name=\"Sum of Sepal Length\" fld=\"0\"", fixed = TRUE)
  expect_match(wb$pivotTables[[1]], "<location ref=\"G2:K6\"", fixed = TRUE)
  expect_match(wb$pivotTables[[2]], "subtotal=\"average\"", fixed = TRUE)
  expect_match(wb$pivotTables[[1]], "outline=\"1\"", fixed = TRUE)
  expect_match(wb$pivotTables[[1]], "<rowItems count=", fixed = TRUE)
  expect_match(wb$pivotDefinitions[[1]], "<worksheetSource name=\"Table2\"/>", fixed = TRUE)
  expect_match(wb$pivotDefinitions[[1]], "<s v=\"virginica\"/><s v=\"setosa\"/><s v=\"versicolor\"/>", fixed = TRUE)

  expect_match(wb$slicers, "<slicer name=\"Species\" cache=\"Slicer_Species\" caption=\"Species\"", fixed = TRUE)
  expect_match(wb$slicerCaches, "<pivotTable tabId=\"1\" name=\"PivotTable1\"/>", fixed = TRUE)
  expect_match(wb$slicerCaches, "<tabular pivotCacheId=\"2\"><items count=\"3\">", fixed = TRUE)
  expect_match(wb$pivotDefinitions[[1]], "<x14:pivotCacheDefinition pivotCacheId=\"2\"/>", fixed = TRUE)
  expect_match(wb$worksheets[[1]]$extLst, "<x14:slicerList><x14:slicer r:id=\"rId5\"/>", fixed = TRUE, all = FALSE)

  tmp <- temp_xlsx()
  on.exit(unlink(tmp), add = TRUE)
  wb$save(tmp)
  wb2 <- wb_load(tmp)
  expect_length(wb2$pivotTables, 2L)
  expect_length(wb2$slicers, 1L)
})

test_that("conditional formatting is imported from xlsb", {

  skip_if(.Platform$endian == "big", "XLSB: skipping on big-endian system")
  skip_online_checks()

  wb <- wb_load(testfile_path("loadExample.xlsb"))
  expect_equal(wb$worksheets[[2]]$conditionalFormatting$sqref, "I2:N5")
  expect_match(
    wb$worksheets[[2]]$conditionalFormatting$cf,
    "<cfRule type=\"cellIs\" operator=\"greaterThan\" dxfId=\"0\" priority=\"1\"><formula>0.5</formula></cfRule>",
    fixed = TRUE
  )
  expect_equal(
    wb$styles_mgr$styles$dxfs[1],
    "<dxf><font><color rgb=\"FF9C0006\"/></font><fill><patternFill><bgColor rgb=\"FFFFC7CE\"/></patternFill></fill></dxf>"
  )

  wb <- wb_load(testfile_path("test_coverage.xlsb"))
  expect_match(wb$worksheets[[1]]$conditionalFormatting$cf[1], "<iconSet iconSet=\"5Arrows\" reverse=\"1\"><cfvo type=\"num\" val=\"-67\"/>", fixed = TRUE)
})

test_that("pivot cache records are imported from xlsb unless the cache is refreshed", {

  skip_if(.Platform$endian == "big", "XLSB: skipping on big-endian system")
  skip_online_checks()

  wb <- wb_load(testfile_path("loadExample.xlsb"))
  expect_false(grepl("refreshOnLoad", wb$pivotDefinitions[[1]], fixed = TRUE))
  expect_match(wb$pivotRecords[[1]], "<r><n v=\"", fixed = TRUE)
  expect_match(wb$pivotRecords[[1]], "<x v=\"", fixed = TRUE)

  tmp <- temp_xlsx()
  on.exit(unlink(tmp), add = TRUE)
  wb$save(tmp)
  wb2 <- wb_load(tmp)
  expect_length(wb2$pivotTables, 2L)
  expect_match(wb2$pivotRecords[[1]], "<r>", fixed = TRUE)

  op <- options("openxlsx2.pivot_refresh" = TRUE)
  on.exit(options(op), add = TRUE)

  wb <- wb_load(testfile_path("loadExample.xlsb"))
  expect_match(wb$pivotDefinitions[[1]], "refreshOnLoad=\"1\"", fixed = TRUE)
  expect_match(wb$pivotRecords[[1]], "count=\"0\"", fixed = TRUE)
})

test_that("default column width is imported from xlsb", {
  skip_if(.Platform$endian == "big", "XLSB: skipping on big-endian system")
  skip_on_cran()
  wb <- wb_load(testfile_path("loadExample.xlsb"))
  expect_match(wb$worksheets[[1]]$sheetFormatPr, "defaultColWidth=\"8.832", fixed = TRUE)
})

test_that("xlsb styles use documented attribute values", {
  skip_if(.Platform$endian == "big", "XLSB: skipping on big-endian system")
  skip_on_cran()
  wb <- wb_load(testfile_path("test_coverage.xlsb"))
  expect_true(any(grepl("horizontal=\"centerContinuous\"", wb$styles_mgr$styles$cellXfs, fixed = TRUE)))
  expect_false(any(grepl("center-across", wb$styles_mgr$styles$cellXfs, fixed = TRUE)))

  wb <- wb_load(testfile_path("nhs-core-standards-for-eprr-v6.1.xlsb"))
  expect_false(any(grepl("<color hex=", wb$styles_mgr$styles$borders, fixed = TRUE)))
  expect_true(any(grepl("<color rgb=", wb$styles_mgr$styles$borders, fixed = TRUE)))
})

test_that("xlsb files created from the xls_cov workbooks match their xlsx", {
  skip_if(.Platform$endian == "big", "XLSB: skipping on big-endian system")
  skip_on_cran()

  files <- c(
    "xls_cov_views", "xls_cov_features", "xls_cov_names_formulas",
    "xls_cov_drawings", "xls_cov_pivot_external", "xls_cov_values",
    "xls_cov2_values", "xls_cov2_names", "xls_cov2_shapes", "xls_cov2_pivots"
  )

  for (nm in files) {
    xlsb <- testfile_path(paste0(nm, ".xlsb"))
    xlsx <- testfile_path(paste0(nm, ".xlsx"))
    wb_b <- suppressWarnings(wb_load(xlsb))
    wb_x <- wb_load(xlsx)
    expect_equal(wb_b$get_sheet_names(), wb_x$get_sheet_names(), info = nm)

    tmp <- tempfile(fileext = ".xlsx")
    suppressWarnings(wb_b$save(tmp))
    expect_equal(length(wb_load(tmp)$worksheets), length(wb_x$worksheets), info = nm)
    unlink(tmp)
  }

  for (nm in c("xls_cov2_values", "xls_cov2_shapes")) {
    wb_b <- wb_load(testfile_path(paste0(nm, ".xlsb")))
    wb_x <- wb_load(testfile_path(paste0(nm, ".xlsx")))
    for (i in seq_along(wb_x$worksheets)) {
      if (wb_x$is_chartsheet[i]) next
      for (fml in c(FALSE, TRUE)) {
        expect_equal(
          suppressMessages(wb_to_df(wb_b, sheet = i, col_names = FALSE, show_formula = fml)),
          suppressMessages(wb_to_df(wb_x, sheet = i, col_names = FALSE, show_formula = fml)),
          ignore_attr = TRUE, info = paste(nm, i, fml)
        )
      }
    }
  }
})

test_that("xlsb formulas, names and page breaks are imported", {
  skip_if(.Platform$endian == "big", "XLSB: skipping on big-endian system")
  skip_on_cran()

  wb_b <- suppressWarnings(wb_load(testfile_path("xls_cov_values.xlsb")))
  wb_x <- wb_load(testfile_path("xls_cov_values.xlsx"))
  expect_equal(wb_b$styles_mgr$styles$fills[grepl("gradientFill", wb_b$styles_mgr$styles$fills)],
               wb_x$styles_mgr$styles$fills[grepl("gradientFill", wb_x$styles_mgr$styles$fills)])
  expect_equal(unlist(wb_b$worksheets[[1]]$rowBreaks), unlist(wb_x$worksheets[[1]]$rowBreaks))
  expect_equal(unlist(wb_b$worksheets[[1]]$colBreaks), unlist(wb_x$worksheets[[1]]$colBreaks))

  fb <- wb_to_df(wb_b, col_names = FALSE, show_formula = TRUE)
  fx <- wb_to_df(wb_x, col_names = FALSE, show_formula = TRUE)
  expect_equal(fb[[3]], fx[[3]])

  nb <- wb_get_named_regions(wb_load(testfile_path("xls_cov_names_formulas.xlsb")))
  nx <- wb_get_named_regions(wb_load(testfile_path("xls_cov_names_formulas.xlsx")))
  expect_true(all(nx$name %in% nb$name))
  m <- match(nx$name, nb$name)
  expect_equal(gsub("'", "", nb$value[m]), gsub("'", "", nx$value))

  wb_b <- suppressWarnings(wb_load(testfile_path("xls_cov_features.xlsb")))
  wb_x <- wb_load(testfile_path("xls_cov_features.xlsx"))
  types <- function(wb) {
    cf <- unlist(wb$worksheets[[3]]$conditionalFormatting)
    sort(unname(unlist(regmatches(cf, gregexpr('type="[A-Za-z0-9]+"', cf)))))
  }
  expect_equal(types(wb_b), types(wb_x))
})

test_that("xlsb conditional formatting rule types match the xlsx", {
  skip_if(.Platform$endian == "big", "XLSB: skipping on big-endian system")
  skip_on_cran()

  wb_b <- suppressWarnings(wb_load(testfile_path("xls_cov3_cf.xlsb")))
  wb_x <- wb_load(testfile_path("xls_cov3_cf.xlsx"))

  rules <- function(wb, i) {
    cf <- wb$worksheets[[i]]$conditionalFormatting$cf
    r <- unlist(regmatches(cf, gregexpr("<cfRule [^>]*>", cf)))
    r <- gsub(" (dxfId|priority)=\"[0-9]+\"", "", r)
    r <- gsub(" aboveAverage=\"1\"|/?>$", "", r)
    vapply(strsplit(r, " "), function(x) paste(sort(x), collapse = " "), "")
  }

  for (i in 1:5) {
    expect_equal(sort(rules(wb_b, i)), sort(rules(wb_x, i)), info = i)
  }

  icons <- function(wb) {
    cf <- wb$worksheets[[3]]$conditionalFormatting$cf
    unlist(regmatches(cf, gregexpr("<iconSet iconSet=\"[^\"]+\"( reverse=\"1\")?", cf)))
  }
  expect_equal(icons(wb_b), icons(wb_x))
  expect_true(any(grepl("3TrafficLights2", icons(wb_b), fixed = TRUE)))
})

test_that("xlsb top10 autofilter is imported", {
  skip_if(.Platform$endian == "big", "XLSB: skipping on big-endian system")
  skip_on_cran()

  wb <- suppressWarnings(wb_load(testfile_path("xls_cov_features.xlsb")))
  expect_match(wb$worksheets[[1]]$autoFilter, '<top10 top="1" val="5" filterVal="230"/>', fixed = TRUE)
})

test_that("xlsb filters, strings and error formulas of xls_cov4_misc are imported", {
  skip_if(.Platform$endian == "big", "XLSB: skipping on big-endian system")
  skip_on_cran()

  wb_b <- suppressWarnings(wb_load(testfile_path("xls_cov4_misc.xlsb")))
  wb_x <- wb_load(testfile_path("xls_cov4_misc.xlsx"))

  expect_equal(wb_b$get_sheet_names(), wb_x$get_sheet_names())

  tab <- wb_b$tables$tab_xml
  expect_match(tab, '<customFilter operator="equal" val="a*"/>', fixed = TRUE)
  expect_match(tab, '<colorFilter dxfId="0" cellColor="1"/>', fixed = TRUE)
  expect_match(tab, '<iconFilter iconSet="3TrafficLights1" iconId="0"/>', fixed = TRUE)
  expect_match(tab, '<dynamicFilter type="aboveAverage" val="50"/>', fixed = TRUE)

  af <- wb_b$worksheets[[2]]$autoFilter
  expect_match(af, '<dateGroupItem year="2020" month="1" dateTimeGrouping="month"/>', fixed = TRUE)
  expect_match(af, '<dateGroupItem year="2020" month="3" day="5" dateTimeGrouping="day"/>', fixed = TRUE)
  expect_match(af, '<top10 top="0" percent="1" val="20" filterVal="16"/>', fixed = TRUE)
  expect_match(af, '<dynamicFilter type="belowAverage" val="10.5"/>', fixed = TRUE)
  expect_match(af, '<customFilters><customFilter operator="notEqual" val="x"/>', fixed = TRUE)
  expect_match(af, '<customFilters and="1"><customFilter operator="notEqual" val="Fiat"/>', fixed = TRUE)

  for (i in c(3, 5, 6)) {
    for (fml in c(FALSE, TRUE)) {
      expect_equal(
        wb_to_df(wb_b, sheet = i, col_names = FALSE, show_formula = fml),
        wb_to_df(wb_x, sheet = i, col_names = FALSE, show_formula = fml),
        ignore_attr = TRUE, info = paste(i, fml)
      )
    }
  }

  expect_length(wb_b$externalLinks, 1L)

  expect_match(wb_b$worksheets[[7]]$autoFilter, '<filter val="TRUE"/>', fixed = TRUE)
})
