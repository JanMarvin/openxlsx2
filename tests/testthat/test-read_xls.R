test_that("reading xls values works", {
  skip_if_big_endian("XLS")
  fl <- testfile_path("xls_values.xls")
  wb <- wb_load(fl)

  expect_equal(unname(wb$get_sheet_names()), c("Data", 'Sp&c<e>s "q"', "Hidden"))

  df <- wb_to_df(wb, sheet = 1, col_names = FALSE, dims = "A1:A10")
  expect_equal(
    df$A,
    c(1, -5, 0.1, 123456.789, 1e20, 2147483647, -2147483648, 1 / 3, 0, 1234567.5)
  )

  df <- wb_to_df(wb, sheet = 1, col_names = FALSE)
  strs <- df$E[grepl("^s[0-9]{3}_", df$E)]
  expect_length(strs, 400)
  expect_true(all(nchar(strs) == 35))
  expect_equal(sum(nchar(unlist(df[sapply(df, is.character)])) == 9000, na.rm = TRUE), 1)
})

test_that("reading xls styles works", {
  skip_if_big_endian("XLS")
  fl <- testfile_path("xls_styles.xls")
  wb <- wb_load(fl)

  expect_equal(wb$worksheets[[1]]$mergeCells, "<mergeCell ref=\"G1:H2\"/>")

  fills <- paste(wb$styles_mgr$styles$fills, collapse = "")
  expect_match(fills, "FFFFFF00")
  expect_match(fills, "FF99CCFF")
  expect_match(fills, "FFFF99CC")

  nf <- paste(wb$styles_mgr$styles$numFmts, collapse = "")
  expect_match(nf, "0.0%", fixed = TRUE)
  expect_match(nf, "EUR", fixed = TRUE)

  ra <- wb$worksheets[[1]]$sheet_data$row_attr
  expect_equal(as.numeric(ra[ra$r == "3", "ht"]), 40, tolerance = 0.01)
  expect_equal(ra[ra$r == "5", "hidden"], "1")
})

test_that("reading xls formulas works", {
  skip_if_big_endian("XLS")
  fl <- testfile_path("xls_formulas.xls")
  wb <- wb_load(fl)

  f <- wb_to_df(wb, sheet = 1, col_names = FALSE, show_formula = TRUE, dims = "B1:F1")
  expect_equal(as.character(f$B), "A1*2+1")
  expect_equal(as.character(f$C), "SUM(A1:A5)")
  expect_equal(as.character(f$D), "IF(A1>2,\"big\",\"small\")")
  expect_equal(as.character(f$E), "Sheet2!A1+'My Sheet'!B2")
  expect_equal(as.character(f$F), "SUM(Sheet2:Sheet3!A1:A3)")

  v <- wb_to_df(wb, sheet = 1, col_names = FALSE, dims = "C1:F1")
  expect_equal(c(v$C, v$E, v$F), c(22.5, 103, 66))

  arr <- wb_to_df(wb, sheet = 1, col_names = FALSE, dims = "J1:J3")
  expect_equal(arr$J, c(3, 6, 9))

  dn <- paste(wb$workbook$definedNames, collapse = "")
  expect_match(dn, "Total", fixed = TRUE)
  expect_match(dn, "Rng", fixed = TRUE)
  expect_match(dn, "LocalTotal", fixed = TRUE)
  expect_match(dn, "localSheetId=\"0\"", fixed = TRUE)
})

test_that("reading an xls saved by Excel works", {
  skip_if_big_endian("XLS")
  fl <- testfile_path("xls_source.xls")
  wb <- wb_load(fl)

  expect_true(all(c("Values", "Styles", "Formulas", "Links", "Misc", "Hidden") %in% wb$get_sheet_names()))

  df <- wb_to_df(wb, sheet = 1)
  expect_equal(df$num, c(1, -5, 0.1, 123456.789, 1e20, 1 / 3))
  expect_equal(df$int, c(0L, 1L, -1L, 2147483647L, -2147483647L, 65536L))
  expect_equal(df$chr[2:5], c("b & <c>", "d \"q\"", "  lead", "\u00fc\u20ac"))
  expect_equal(df$lgl, c(TRUE, FALSE, TRUE, FALSE, TRUE, FALSE))
  expect_equal(df$date, as.Date("2024-02-28") + 0:5)

  f <- wb_to_df(wb, sheet = 3, col_names = FALSE, show_formula = TRUE, dims = "B1:K1")
  expect_equal(
    unlist(f, use.names = FALSE),
    c(
      "A1*2+1", "SUM(A1:A5)", "IF(A1>2,\"big\",\"small\")", "Values!B2+Styles!A2",
      "SUM(Values!B2:B6)", "IFERROR(1/0,\"err\")", "VLOOKUP(2,A1:B5,2,FALSE)",
      "SUMPRODUCT(A1:A5,A1:A5)", "TEXT(TODAY(),\"yyyy\")", "COUNTIFS(A1:A5,\">1\",A1:A5,\"<5\")"
    )
  )

  filled <- wb_to_df(wb, sheet = 3, col_names = FALSE, show_formula = TRUE, dims = "O1:O20")
  expect_equal(as.character(filled$O), sprintf("A%d*2", 1:20))
})

test_that("reading xls hyperlinks, comments and validation works", {
  skip_if_big_endian("XLS")
  fl <- testfile_path("xls_source.xls")
  wb <- wb_load(fl)

  hl <- wb$worksheets[[4]]$hyperlinks
  expect_length(hl, 3)
  expect_match(hl[1], "ref=\"A1\"", fixed = TRUE)
  expect_match(hl[1], "tooltip=\"tip\"", fixed = TRUE)
  expect_match(hl[2], "location=\"Values!A1\"", fixed = TRUE)

  rels <- paste(wb$worksheets_rels[[4]], collapse = "")
  expect_match(rels, "Target=\"https://example.com/\"", fixed = TRUE)
  expect_match(rels, "Target=\"mailto:a@example.com\"", fixed = TRUE)

  cmt <- wb_get_comment(wb, sheet = 4)
  expect_true(cmt$ref %in% c("B1", "B2"))
  expect_match(paste(unlist(wb$comments), collapse = ""), "a comment", fixed = TRUE)

  dv <- wb$worksheets[[4]]$dataValidations
  expect_length(dv, 1)
  expect_match(dv, "type=\"whole\"", fixed = TRUE)
  expect_match(dv, "sqref=\"F1:F5\"", fixed = TRUE)
  expect_match(dv, "<formula1>1</formula1><formula2>10</formula2>", fixed = TRUE)
})

test_that("reading an xls table works", {
  skip_if_big_endian("XLS")
  fl <- testfile_path("xls_source.xls")
  wb <- wb_load(fl)

  tabs <- wb$tables
  skip_if(NROW(tabs) == 0L, "xls_source.xls has no table yet")

  expect_equal(tabs$tab_name, "tab")
  expect_equal(tabs$tab_ref, "D1:E4")
  expect_match(tabs$tab_xml, "<tableColumn id=\"1\" name=\"g\"/>", fixed = TRUE)
  expect_match(tabs$tab_xml, "<tableColumn id=\"2\" name=\"v\"/>", fixed = TRUE)
})

test_that("reading xls conditional formatting works", {
  skip_if_big_endian("XLS")
  fl <- testfile_path("xls_source.xls")
  wb <- wb_load(fl)

  cf <- wb$worksheets[[4]]$conditionalFormatting
  expect_equal(cf$sqref, "G1:G5")
  expect_match(cf$cf, "type=\"expression\"", fixed = TRUE)
  expect_match(cf$cf, "<formula>G1&gt;2</formula>", fixed = TRUE)

  # the formatting is a dxf with a font color and a fill
  dxfs <- wb$styles_mgr$styles$dxfs
  expect_length(dxfs, 1)
  expect_match(dxfs, "<font><color rgb=", fixed = TRUE)
  expect_match(dxfs, "<fill><patternFill", fixed = TRUE)
})

test_that("reading xls rich text works", {
  skip_if_big_endian("XLS")
  fl <- testfile_path("xls_source.xls")
  wb <- wb_load(fl)

  sst <- unlist(wb$sharedStrings)
  rich <- sst[grepl("<r>", sst, fixed = TRUE)]
  skip_if(length(rich) == 0L, "xls_source.xls has no rich text yet")

  expect_match(rich[1], "<b/>", fixed = TRUE)
  expect_match(rich[1], "FFFF0000", fixed = TRUE)
  expect_equal(wb_to_df(wb, "Misc", dims = "A1", col_names = FALSE)[[1]], "plain bold red")
})

test_that("reading xls page margins works", {
  skip_if_big_endian("XLS")
  fl <- testfile_path("xls_source.xls")
  wb <- wb_load(fl)

  margins <- wb$worksheets[[1]]$pageMargins
  expect_length(margins, 1)
  expect_match(margins, "<pageMargins left=\"[0-9.]+\" right=\"[0-9.]+\" top=\"[0-9.]+\" bottom=\"[0-9.]+\"")
})

test_that("reading xls color scales, data bars and icon sets works", {
  skip_if_big_endian("XLS")
  fl <- testfile_path("xls_source.xls")
  wb <- wb_load(fl)

  cf <- wb$worksheets[[5]]$conditionalFormatting
  skip_if(length(cf) == 0L || !any(grepl("colorScale", cf$cf)), "xls_source.xls has no color scale yet")

  expect_true(any(grepl("<colorScale>", cf$cf, fixed = TRUE)))
  expect_true(any(grepl("<dataBar", cf$cf, fixed = TRUE)))
  expect_true(any(grepl("<iconSet iconSet=\"3Arrows\"", cf$cf, fixed = TRUE)))
})

test_that("pictures in xls are imported", {
  skip_if_big_endian("XLS")
  skip_on_cran()
  wb <- wb_load(testfile_path("xls_source.xls"))
  skip_if(length(wb$media) == 0L, "xls_source.xls has no pictures yet")
  expect_equal(length(wb$media), 2L)
  expect_true(any(grepl("xdr:pic", unlist(wb$drawings), fixed = TRUE)))
  tmp <- temp_xlsx()
  on.exit(unlink(tmp))
  wb$save(tmp)
  expect_equal(length(wb_load(tmp)$media), 2L)
})

test_that("loadExample.xls matches loadExample.xlsx", {
  skip_if_big_endian("XLS")
  skip_on_cran()
  a <- wb_load(testfile_path("loadExample.xlsx"))
  b <- wb_load(testfile_path("loadExample.xls"))
  expect_equal(wb_get_sheet_names(b), wb_get_sheet_names(a))
  expect_equal(length(b$tables), length(a$tables))
  expect_equal(length(b$media), length(a$media))
  for (i in c(1, 3, 4)) {
    x <- wb_to_df(a, sheet = i, col_names = FALSE)
    y <- wb_to_df(b, sheet = i, col_names = FALSE)
    expect_equal(dim(y), dim(x))
    for (j in seq_along(x)) {
      xn <- suppressWarnings(as.numeric(x[[j]]))
      yn <- suppressWarnings(as.numeric(y[[j]]))
      num <- !is.na(xn) & !is.na(yn)
      expect_equal(yn[num], xn[num], tolerance = 1e-9)
      expect_equal(y[[j]][!num], x[[j]][!num])
    }
  }
})

test_that("charts in xls are imported", {
  skip_if_big_endian("XLS")
  skip_on_cran()
  wb <- wb_load(testfile_path("loadExample.xls"))
  charts <- unlist(wb$charts)
  expect_true(any(grepl("<c:scatterChart>", charts, fixed = TRUE)))
  expect_true(any(grepl("IrisSample!$D$2:$D$51", charts, fixed = TRUE)))
  expect_true(any(grepl("<c:barChart>", charts, fixed = TRUE)))
  expect_true(any(grepl("schemeClr val=\"accent1\"", charts, fixed = TRUE)))
  expect_true(any(grepl("<a:t>Petal</a:t>", charts, fixed = TRUE)))
  expect_true(any(grepl("<a:t> Width</a:t>", charts, fixed = TRUE)))

  drawings <- unlist(wb$drawings)
  expect_equal(sum(lengths(regmatches(drawings, gregexpr("<xdr:pic>", drawings, fixed = TRUE)))), 2L)
  expect_true(any(grepl("prst=\"smileyFace\"", drawings, fixed = TRUE)))

  tmp <- temp_xlsx()
  on.exit(unlink(tmp))
  wb$save(tmp)
  expect_equal(unlist(wb_load(tmp)$charts), charts)
})

test_that("chart sheets in xls are imported", {
  skip_if_big_endian("XLS")
  skip_on_cran()
  wb <- wb_load(testfile_path("xls_source.xls"))
  skip_if(!any(wb$is_chartsheet), "xls_source.xls has no chart sheet yet")
  expect_true(any(grepl("<c:chart", unlist(wb$charts), fixed = TRUE)))
})

test_that("document properties and workbook views of an xls are imported", {
  skip_if_big_endian("XLS")
  skip_on_cran()
  wb <- wb_load(testfile_path("loadExample.xls"))
  expect_match(wb$core, "<dc:creator>Alex</dc:creator>", fixed = TRUE)
  expect_match(wb$core, "<dcterms:created xsi:type=\"dcterms:W3CDTF\">2014-03-07T16:08:25Z</dcterms:created>", fixed = TRUE)
  expect_match(wb$workbook$bookViews, "activeTab=\"1\"", fixed = TRUE)
})

test_that("sheet format and outline settings of an xls are imported", {
  skip_if_big_endian("XLS")
  skip_on_cran()
  wb <- wb_load(testfile_path("xls_source.xls"))
  expect_match(wb$worksheets[[2]]$sheetFormatPr, "outlineLevelRow=\"1\"", fixed = TRUE)
})

test_that("named cell styles of an xls are imported", {
  skip_if_big_endian("XLS")
  skip_on_cran()
  wb <- wb_load(testfile_path("xls_source.xls"))
  cs <- wb$styles_mgr$styles$cellStyles
  expect_true(any(grepl("name=\"Normal\" xfId=\"0\" builtinId=\"0\"", cs, fixed = TRUE)))
  expect_true(any(grepl("name=\"Bad\".*builtinId=\"27\"", cs)))
  expect_gt(length(wb$styles_mgr$styles$cellStyleXfs), 1L)
})

test_that("sheet dimensions of an xls are imported", {
  skip_if_big_endian("XLS")
  skip_on_cran()
  wb <- wb_load(testfile_path("xls_source.xls"))
  expect_equal(wb$worksheets[[1]]$dimension, "<dimension ref=\"A1:H7\"/>")
})

test_that("ignored errors of an xls are imported", {
  skip_if_big_endian("XLS")
  skip_on_cran()
  wb <- wb_load(testfile_path("xls_source.xls"))
  ie <- unlist(lapply(wb$worksheets, function(x) x$ignoredErrors))
  expect_match(ie, "sqref=\"G1:G3\" numberStoredAsText=\"1\"", fixed = TRUE)
  expect_match(ie, "sqref=\"H1:H2\" evalError=\"1\" emptyCellReference=\"1\"", fixed = TRUE)
})

test_that("a single-cell sheet of an xls has the dimension A1", {
  skip_if_big_endian("XLS")
  skip_on_cran()
  wb <- wb_load(testfile_path("xls_source.xls"))
  expect_equal(wb$worksheets[[8]]$dimension, "<dimension ref=\"A1\"/>")
})

test_that("sort state of an xls auto filter is imported", {
  skip_if_big_endian("XLS")
  skip_on_cran()
  wb <- wb_load(testfile_path("xls_source.xls"))
  expect_equal(
    wb$worksheets[[5]]$autoFilter,
    "<autoFilter ref=\"J1:K11\"><sortState ref=\"J2:K11\"><sortCondition descending=\"1\" ref=\"J2:J11\"/></sortState></autoFilter>"
  )
})

test_that("charts of an xls get cached values and the default data label block", {
  skip_if_big_endian("XLS")
  skip_on_cran()
  wb <- wb_load(testfile_path("loadExample.xls"))
  chart <- paste(unlist(wb$charts), collapse = "")
  expect_match(chart, "<c:f>IrisSample!$D$2:$D$51</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"50\"/><c:pt idx=\"0\"><c:v>1.9</c:v>", fixed = TRUE)
  expect_match(chart, "<c:dLbls><c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/>", fixed = TRUE)
})

test_that("legends, data bars and filter names of an xls are imported", {
  skip_if_big_endian("XLS")
  skip_on_cran()
  wb <- wb_load(testfile_path("loadExample.xls"))
  expect_match(paste(unlist(wb$charts), collapse = ""), "<c:legendPos val=\"r\"/><c:overlay", fixed = TRUE)

  wb <- wb_load(testfile_path("xls_source.xls"))
  expect_match(paste(unlist(wb$charts), collapse = ""), "<c:legendPos val=\"r\"/><c:layout><c:manualLayout>", fixed = TRUE)
  expect_false(any(grepl("showValue=\"0\"", unlist(lapply(wb$worksheets, function(x) x$conditionalFormatting)), fixed = TRUE)))
  expect_match(wb$workbook$definedNames, "_xlnm._FilterDatabase\" localSheetId=\"4\" hidden=\"1\">'Misc'!$J$1:$K$11", fixed = TRUE, all = FALSE)
})

test_that("data bars of an xls get the extension Excel uses for them", {
  skip_if_big_endian("XLS")
  skip_on_cran()
  wb <- wb_load(testfile_path("xls_source.xls"))
  ws <- wb$worksheets[[5]]
  expect_match(ws$conditionalFormatting$cf[grepl("dataBar", ws$conditionalFormatting$cf)], "<dataBar><cfvo type=\"min\"/>", fixed = TRUE)
  expect_match(ws$extLst, "<x14:dataBar minLength=\"0\" maxLength=\"100\" border=\"1\"", fixed = TRUE)
})

test_that("data_only skips everything but the cells", {
  skip_if_big_endian("XLS")
  fl <- testfile_path("xls_source.xls")
  skip_if(!file.exists(fl))

  wb <- wb_load(fl, sheet = 3, data_only = TRUE)
  full <- wb_load(fl, sheet = 3)
  expect_equal(length(wb$charts), 0L)
  expect_equal(length(wb$tables), 0L)
  expect_message(a <- wb_to_df(wb, col_names = FALSE), "contains no data")
  expect_message(b <- wb_to_df(full, col_names = FALSE), "contains no data")
  expect_equal(a, b)
})

test_that("pivot tables are imported", {
  skip_if_big_endian("XLS")
  fl <- testfile_path("xls_pivot.xls")
  skip_if(!file.exists(fl))

  wb <- wb_load(fl)
  expect_equal(length(wb$pivotTables), 2L)
  expect_equal(length(wb$pivotDefinitions), 1L)

  pt1 <- wb$pivotTables[[1]]
  expect_match(pt1, '<location ref="A3:C14"', fixed = TRUE)
  expect_match(pt1, '<rowFields count="2"><field x="0"/><field x="1"/></rowFields>', fixed = TRUE)
  expect_match(pt1, '<colFields count="1"><field x="-2"/></colFields>', fixed = TRUE)
  expect_match(pt1, '<pageField fld="4"', fixed = TRUE)
  expect_match(pt1, 'subtotal="average"', fixed = TRUE)

  def <- wb$pivotDefinitions[[1]]
  expect_match(def, '<worksheetSource ref="A1:E13" sheet="data"/>', fixed = TRUE)
  expect_match(def, '<cacheField name="Species"', fixed = TRUE)
  expect_match(def, '<s v="a"/><s v="b"/><s v="c"/>', fixed = TRUE)

  tmp <- temp_xlsx()
  wb$save(tmp)
  wb2 <- wb_load(tmp)
  expect_equal(length(wb2$pivotTables), 2L)

  expect_equal(length(wb_load(fl, data_only = TRUE, sheet = 1)$pivotTables), 0L)
})

test_that("shape numbers of comments do not overlap", {
  skip_if_big_endian("XLS")
  fl <- testfile_path("xls_comments.xls")
  skip_if(!file.exists(fl))

  wb <- wb_load(fl)
  ids <- lapply(wb$vml, function(x) as.integer(regmatches(x, gregexpr("(?<=_x0000_s)[0-9]+", x, perl = TRUE))[[1]]))
  expect_equal(sum(duplicated(unlist(ids))), 0L)
  expect_equal(lengths(ids), lengths(wb$comments))
  expect_match(wb$vml[[1]], 'data="1,2"', fixed = TRUE)
})

test_that("xls paths with a tilde are expanded", {
  skip_if_big_endian("XLS")
  skip_on_cran()
  home <- tempfile()
  dir.create(home)
  file.copy(testfile_path("loadExample.xls"), file.path(home, "loadExample.xls"))
  old <- Sys.getenv("HOME")
  Sys.setenv(HOME = home)
  on.exit(Sys.setenv(HOME = old))
  wb <- wb_load("~/loadExample.xls")
  expect_equal(length(wb$get_sheet_names()), 4L)
})

test_that("conditional formatting of xls uses the full colors", {
  skip_if_big_endian("XLS")
  skip_on_cran()
  wb <- wb_load(testfile_path("loadExample.xls"))
  dxfs <- wb$styles_mgr$styles$dxfs
  expect_match(dxfs, "FFFFC7CE", fixed = TRUE, all = FALSE)
  expect_match(dxfs, "FF9C0006", fixed = TRUE, all = FALSE)
})

test_that("pivot table style of xls is imported", {
  skip_if_big_endian("XLS")
  skip_on_cran()
  a <- wb_load(testfile_path("loadExample.xlsx"))
  b <- wb_load(testfile_path("loadExample.xls"))
  style <- function(wb) regmatches(wb$pivotTables, regexpr("<pivotTableStyleInfo[^>]*>", wb$pivotTables))
  expect_equal(style(b), style(a))
})

test_that("pivot field options of xls are imported", {
  skip_if_big_endian("XLS")
  skip_on_cran()
  a <- wb_load(testfile_path("xls_cov2_pivots.xlsx"))
  b <- wb_load(testfile_path("xls_cov2_pivots.xls"))
  attrs <- function(wb, i, what) {
    x <- unlist(regmatches(wb$pivotTables[i], gregexpr("<pivotField\\b[^>]*>", wb$pivotTables[i])))
    sort(unlist(regmatches(x, gregexpr(paste0(what, '="[^"]*"'), x))))
  }
  for (i in 10:11) expect_equal(attrs(b, i, "sortType"), attrs(a, i, "sortType"))
  expect_equal(attrs(b, 13, "subtotalTop"), attrs(a, 13, "subtotalTop"))
  for (w in c("dragToRow", "dragToCol", "dragToPage")) expect_equal(attrs(b, 20, w), attrs(a, 20, w))
  expect_length(attrs(b, 10, "sortType"), 3L)
})

test_that("shared string counts of xls are imported", {
  skip_if_big_endian("XLS")
  skip_on_cran()
  wb <- wb_load(testfile_path("loadExample.xls"))
  expect_equal(attr(wb$sharedStrings, "uniqueCount"), as.character(length(wb$sharedStrings)))
})

test_that("page layout view of xls is imported", {
  skip_if_big_endian("XLS")
  skip_on_cran()
  fl <- testfile_path("xls_source.xls")
  raw_xls <- readBin(fl, "raw", file.size(fl))
  plv <- as.raw(c(0x8b, 0x08, 0x10, 0x00, 0x8b, 0x08))
  at <- which(raw_xls == plv[1])
  at <- at[vapply(at, function(i) all(raw_xls[i + 0:5] == plv), NA)][1]
  raw_xls[at + 16] <- as.raw(80)
  raw_xls[at + 18] <- as.raw(3)
  tmp <- tempfile(fileext = ".xls")
  on.exit(unlink(tmp), add = TRUE)
  writeBin(raw_xls, tmp)
  wb <- wb_load(tmp)
  views <- unlist(lapply(wb$worksheets, function(x) x$sheetViews))
  expect_equal(sum(grepl('view="pageLayout"', views, fixed = TRUE)), 1L)
  expect_true(any(grepl('zoomScalePageLayoutView="80"', views, fixed = TRUE)))
})

test_that("page break preview of xls is imported", {
  skip_if_big_endian("XLS")
  skip_on_cran()
  fl <- testfile_path("xls_source.xls")
  raw_xls <- readBin(fl, "raw", file.size(fl))
  w2 <- as.raw(c(0x3e, 0x02, 0x12, 0x00))
  at <- which(raw_xls == w2[1])
  at <- at[vapply(at, function(i) all(raw_xls[i + 0:3] == w2), NA)][1]
  raw_xls[at + 5] <- as.raw(bitwOr(as.integer(raw_xls[at + 5]), 0x08L))
  raw_xls[at + 14] <- as.raw(60)
  raw_xls[at + 15] <- as.raw(0)
  tmp <- tempfile(fileext = ".xls")
  on.exit(unlink(tmp), add = TRUE)
  writeBin(raw_xls, tmp)
  wb <- wb_load(tmp)
  views <- unlist(lapply(wb$worksheets, function(x) x$sheetViews))
  expect_true(any(grepl('view="pageBreakPreview"', views, fixed = TRUE)))
  expect_true(any(grepl('zoomScaleSheetLayoutView="60"', views, fixed = TRUE)))
  expect_true(any(grepl('zoomScale="60"', views, fixed = TRUE)))
})

test_that("rels of chart sheets are read if no worksheet has rels", {
  skip_if_big_endian("XLS")
  skip_on_cran()
  tmp <- tempfile("chartsheet_rels")
  dir.create(tmp)
  on.exit(unlink(tmp, recursive = TRUE), add = TRUE)
  wb <- wb_workbook()$add_worksheet()$add_data(x = 1)$add_chartsheet()
  fl <- tempfile(fileext = ".xlsx")
  on.exit(unlink(fl), add = TRUE)
  wb$save(fl)
  fls <- utils::unzip(fl, exdir = tmp)
  dir.create(file.path(tmp, "xl", "chartsheets", "_rels"), recursive = TRUE, showWarnings = FALSE)
  dir.create(file.path(tmp, "xl", "drawings"), showWarnings = FALSE)
  sheet <- grep("chartsheets/sheet[0-9]+[.]xml$", fls, value = TRUE)
  xml <- paste(readLines(sheet, warn = FALSE), collapse = "")
  xml <- sub("</chartsheet>", "<drawing r:id=\"rId1\"/></chartsheet>", xml, fixed = TRUE)
  if (!grepl("xmlns:r=", xml, fixed = TRUE))
    xml <- sub("<chartsheet ", "<chartsheet xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ", xml, fixed = TRUE)
  writeLines(xml, sheet)
  writeLines(paste0(
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing1.xml\"/>",
    "</Relationships>"
  ), file.path(tmp, "xl", "chartsheets", "_rels", basename(paste0(sheet, ".rels"))))
  writeLines(
    "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>",
    file.path(tmp, "xl", "drawings", "drawing1.xml")
  )
  out <- tempfile(fileext = ".xlsx")
  on.exit(unlink(out), add = TRUE)
  old <- setwd(tmp)
  zip::zip(out, files = list.files(tmp, recursive = TRUE, all.files = TRUE), mode = "mirror")
  setwd(old)

  wb2 <- wb_load(out)
  expect_true(any(grepl("drawing1.xml", unlist(wb2$worksheets_rels), fixed = TRUE)))
})

test_that("xls files created from the xls_cov workbooks are imported", {
  skip_if_big_endian("XLS")
  skip_on_cran()

  expected <- list(
    xls_cov_views = c("freeze", "freeze_row", "freeze_col", "split", "hidden_outline", "protected", "very_hidden"),
    xls_cov_features = c("filter", "merge_val", "cf", "links", "page", "fit"),
    xls_cov_names_formulas = c("data", "calc"),
    xls_cov_drawings = c("img", "chart", "chartsheet"),
    xls_cov_pivot_external = c("src", "pivot", "pivot2", "ext")
  )

  for (nm in names(expected)) {
    fl <- testfile_path(paste0(nm, ".xls"))

    wb <- wb_load(fl)
    expect_equal(wb$get_sheet_names(escape = TRUE), expected[[nm]], ignore_attr = TRUE, info = nm)

    tmp <- tempfile(fileext = ".xlsx")
    wb$save(tmp)
    wb2 <- wb_load(tmp)
    expect_equal(length(wb2$worksheets), length(expected[[nm]]), info = nm)
    unlink(tmp)
  }

  wb <- wb_load(testfile_path("xls_cov_views.xls"))
  panes <- unlist(lapply(wb$worksheets, function(x) x$freezePane))
  expect_equal(sum(grepl('state="frozen"', panes, fixed = TRUE)), 3L)
  expect_equal(sum(grepl('state="split"', panes, fixed = TRUE)), 1L)

  wb <- wb_load(testfile_path("xls_cov_features.xls"))
  expect_true(length(wb$worksheets[[1]]$autoFilter) > 0)
  expect_true(length(wb$worksheets[[2]]$mergeCells) > 0)
  expect_true(length(wb$worksheets[[2]]$dataValidations) > 0)
  expect_true(length(wb$worksheets[[3]]$conditionalFormatting) > 0)
  cf_txt <- unlist(wb$worksheets[[3]]$conditionalFormatting)
  expect_match(cf_txt, 'type="containsText" operator="containsText" text="Mazda"', fixed = TRUE, all = FALSE)
  expect_match(cf_txt, 'type="endsWith" operator="endsWith" text="4"', fixed = TRUE, all = FALSE)
  expect_true(length(wb$worksheets[[4]]$hyperlinks) > 0)

  wb <- wb_load(testfile_path("xls_cov_names_formulas.xls"))
  expect_true(nrow(wb_get_named_regions(wb)) > 0)
  fml <- wb_to_df(wb, sheet = "calc", col_names = FALSE, show_formula = TRUE)
  expect_true(sum(grepl("^[A-Z]", unlist(fml[[3]]))) > 50)

  wb <- wb_load(testfile_path("xls_cov_drawings.xls"))
  expect_true(length(wb$charts) > 0)
  expect_true(length(wb$media) > 0)
})

test_that("xls files created from the xls_cov2 workbooks are imported", {
  skip_if_big_endian("XLS")
  skip_on_cran()

  wb <- wb_load(testfile_path("xls_cov2_pivots.xls"))
  expect_equal(wb$get_sheet_names(escape = TRUE)[1:3], c("src", "functions", "functions2"), ignore_attr = TRUE)
  expect_true(length(wb$pivotTables) >= 20)
  expect_true(any(grepl("defaultSubtotal=\"0\"", wb$pivotTables, fixed = TRUE)))
  # a table with column fields only has no row fields
  expect_true(any(grepl("<colFields", wb$pivotTables, fixed = TRUE) & !grepl("<rowFields", wb$pivotTables, fixed = TRUE)))
  # the lines of the table match the area of the table
  expect_true(any(grepl("<rowItems count=\"11\">", wb$pivotTables, fixed = TRUE)))
  # data fields on the axis are numbered, dates of the cache are kept
  expect_true(any(grepl("<i i=\"3\">", wb$pivotTables, fixed = TRUE)))
  expect_true(any(grepl("<d v=\"2024-01-01T00:00:00\"/>", wb$pivotDefinitions, fixed = TRUE)))
  tmp <- tempfile(fileext = ".xlsx")
  wb$save(tmp)
  expect_equal(length(wb_load(tmp)$pivotTables), length(wb$pivotTables))
  unlink(tmp)

  wb <- wb_load(testfile_path("xls_cov2_values.xls"))
  expect_equal(wb$get_sheet_names(escape = TRUE), c("errors", "numfmts", "numbers", "colors", "fonts"), ignore_attr = TRUE)
  cc <- wb$worksheets[[1]]$sheet_data$cc
  errs <- cc$v[cc$c_t == "e"]
  expect_true(all(c("#N/A", "#DIV/0!", "#VALUE!", "#NUM!", "#NAME?", "#NULL!", "#REF!") %in% errs))
  expect_equal(nrow(wb_to_df(wb, sheet = "numfmts", col_names = FALSE)), 32L)

  wb <- wb_load(testfile_path("xls_cov2_names.xls"))
  expect_equal(wb$get_sheet_names(escape = TRUE), c("a", "b", "c with space"), ignore_attr = TRUE)
  expect_true(nrow(wb_get_named_regions(wb)) > 10)

  wb <- wb_load(testfile_path("xls_cov2_shapes.xls"))
  expect_equal(wb$get_sheet_names(escape = TRUE), c("shapes", "textboxes", "connectors"), ignore_attr = TRUE)
  expect_true(sum(lengths(regmatches(wb$drawings, gregexpr("<xdr:sp[ >]", wb$drawings)))) > 70)
  expect_equal(sum(lengths(regmatches(wb$drawings, gregexpr("<xdr:cxnSp", wb$drawings)))), 6L)
  expect_equal(sum(lengths(regmatches(wb$drawings, gregexpr("<xdr:grpSp>", wb$drawings)))), 6L)
  expect_true(any(grepl("plain text box", wb$drawings, fixed = TRUE)))
  expect_true(any(grepl("<a:prstDash val=\"sysDot\"/>", wb$drawings, fixed = TRUE)))
})

test_that("a single sheet of an xls file can be read without converting the others", {
  skip_if_big_endian("XLS")
  fl <- testfile_path("xls_cov2_values.xls")
  skip_if(fl == "" || !file.exists(fl))

  full <- wb_load(fl)
  nms <- full$get_sheet_names()

  wb <- wb_load(fl, sheet = 3, data_only = TRUE)
  expect_equal(wb_to_df(wb, sheet = 3), wb_to_df(full, sheet = 3))
  expect_null(suppressMessages(suppressWarnings(wb_to_df(wb, sheet = 1))))

  wb <- wb_load(fl, sheet = names(nms)[2], data_only = TRUE)
  expect_equal(wb_to_df(wb, sheet = 2), wb_to_df(full, sheet = 2))

  expect_error(wb_load(fl, sheet = "nope", data_only = TRUE), "No such sheet")
})

test_that("single sheets of xls files keep their strings, dates and names", {
  skip_if_big_endian("XLS")
  for (file in c("xls_values.xls", "xls_cov2_values.xls", "loadExample.xls")) {
    fl <- testfile_path(file)
    skip_if(fl == "" || !file.exists(fl))

    full <- wb_load(fl)
    nms <- full$get_sheet_names()

    for (i in seq_along(nms)) {
      if (full$is_chartsheet[i]) next
      expected <- wb_to_df(full, sheet = i)

      by_index <- wb_load(fl, sheet = i, data_only = TRUE)
      expect_equal(wb_to_df(by_index, sheet = i), expected)

      by_name <- wb_load(fl, sheet = names(nms)[i], data_only = TRUE)
      expect_equal(wb_to_df(by_name, sheet = i), expected)

      expect_equal(by_index$get_sheet_names(), nms)
    }
  }

  fl <- testfile_path("xls_cov2_values.xls")
  skip_if(fl == "" || !file.exists(fl))
  df <- wb_to_df(wb_load(fl, sheet = 3, data_only = TRUE), sheet = 3)
  expect_true(any(vapply(df, is.character, NA)))
  expect_true(any(vapply(df, inherits, NA, c("Date", "POSIXct"))))
})

test_that("xls_cov_values with comments is imported", {
  skip_if_big_endian("XLS")
  skip_on_cran()

  fl <- testfile_path("xls_cov_values.xls")

  wb <- wb_load(fl)
  expect_equal(wb$get_sheet_names(escape = TRUE), "values", ignore_attr = TRUE)
  df <- wb_to_df(wb, sheet = "values")
  expect_true(nrow(df) >= 9)
  expect_equal(nchar(df[[1]][3]), 32767L)

  wb2 <- wb_load(fl, sheet = 1, data_only = TRUE)
  expect_equal(wb_to_df(wb2, sheet = 1), df)
})

test_that("xls files created from the xls_cov3 and xls_cov4 workbooks are imported", {
  skip_if_big_endian("XLS")
  skip_on_cran()

  for (nm in c("xls_cov3_cf", "xls_cov4_misc")) {
    wb_s <- suppressWarnings(wb_load(testfile_path(paste0(nm, ".xls"))))
    wb_x <- wb_load(testfile_path(paste0(nm, ".xlsx")))
    expect_equal(wb_s$get_sheet_names(), wb_x$get_sheet_names(), info = nm)
    expect_equal(length(wb_s$tables$tab_xml), length(wb_x$tables$tab_xml), info = nm)

    tmp <- tempfile(fileext = ".xlsx")
    suppressWarnings(wb_s$save(tmp))
    expect_equal(length(wb_load(tmp)$worksheets), length(wb_x$worksheets), info = nm)
    unlink(tmp)
  }

  wb_s <- wb_load(testfile_path("xls_cov4_misc.xls"))
  expect_length(wb_s$externalLinks, 1L)

  cf <- function(f, i) paste(unlist(suppressWarnings(wb_load(testfile_path(f)))$worksheets[[i]]$conditionalFormatting[2]), collapse = "")
  cf_s <- vapply(1:5, function(i) cf("xls_cov3_cf.xls", i), "")
  expect_match(cf_s[1], "type=\"timePeriod\" timePeriod=\"lastWeek\"", fixed = TRUE)
  expect_match(cf_s[2], "type=\"aboveAverage\" equalAverage=\"1\"", fixed = TRUE)
  expect_match(cf_s[2], "stdDev=\"3\"", fixed = TRUE)
  expect_match(cf_s[3], "iconSet=\"3TrafficLights2\" reverse=\"1\"", fixed = TRUE)
  expect_match(cf_s[5], "type=\"top10\" rank=\"3\" percent=\"0\" bottom=\"1\"", fixed = TRUE)
  expect_match(cf_s[5], "stopIfTrue=\"1\"", fixed = TRUE)

  tab <- wb_s$tables$tab_xml
  expect_match(tab, "<filter val=\"alpha\"/>", fixed = TRUE)
  expect_match(tab, "<customFilter operator=\"greaterThan\" val=\"10\"/>", fixed = TRUE)
  expect_match(tab, "<dynamicFilter type=\"aboveAverage\" val=\"50\"/>", fixed = TRUE)
  expect_match(tab, "<colorFilter dxfId=\"0\" cellColor=\"1\"/>", fixed = TRUE)
  expect_match(tab, "<iconFilter iconSet=\"3TrafficLights1\" iconId=\"0\"/>", fixed = TRUE)
  expect_match(wb_s$styles_mgr$styles$dxfs[1], "FFFF0000", fixed = TRUE)

  af <- wb_s$worksheets[[2]]$autoFilter
  expect_match(af, "<filters blank=\"1\"><filter val=\"3\"/><filter val=\"7\"/></filters>", fixed = TRUE)
  expect_match(af, "<dateGroupItem year=\"2020\" month=\"3\" day=\"5\" dateTimeGrouping=\"day\"/>", fixed = TRUE)
  expect_match(af, "<dynamicFilter type=\"belowAverage\" val=\"10.5\"/>", fixed = TRUE)
})

test_that("sheet names with an apostrophe are quoted in the filter database name", {
  skip_if_big_endian("XLS")
  xml_dir <- tempfile()
  dir.create(file.path(xml_dir, "xl", "worksheets"), recursive = TRUE)
  on.exit(unlink(xml_dir, recursive = TRUE), add = TRUE)
  writeLines(
    paste0(
      '<workbook><sheets><sheet name="O&apos;Brien &amp; Co" sheetId="1" r:id="rId1"/></sheets>',
      '<calcPr calcId="1"/></workbook>'
    ),
    file.path(xml_dir, "xl", "workbook.xml")
  )
  writeLines('<worksheet><autoFilter ref="A1:C5"/></worksheet>', file.path(xml_dir, "xl", "worksheets", "sheet1.xml"))

  openxlsx2:::xls_filter_names(xml_dir, 1L, 0L)

  got <- paste(readLines(file.path(xml_dir, "xl", "workbook.xml"), warn = FALSE), collapse = "")
  expect_match(got, "'O&apos;&apos;Brien &amp; Co'!$A$1:$C$5", fixed = TRUE)
})

# little endian record builders for streams of older xls files
xls_u8 <- function(x) as.raw(x)
xls_u16 <- function(x) writeBin(as.integer(x), raw(), size = 2L)
xls_u32 <- function(x) {
  x <- x %% 4294967296
  as.raw(c(x %% 256, (x %/% 256) %% 256, (x %/% 65536) %% 256, (x %/% 16777216) %% 256))
}
xls_rec <- function(id, ...) {
  p <- c(raw(), ...)
  c(xls_u16(id), xls_u16(length(p)), p)
}
xls_txt <- function(x) c(xls_u8(nchar(x)), charToRaw(x))


xls_ole <- function(book, stream) {
  u8 <- xls_u8
  u16 <- xls_u16
  u32 <- xls_u32
  book <- c(book, raw(max(0L, 4096L - length(book))))
  book <- c(book, raw(ceiling(length(book) / 512) * 512 - length(book)))
  ns <- length(book) / 512

  fat <- rep(0xFFFFFFFF, 128)
  fat[1:2] <- c(0xFFFFFFFD, 0xFFFFFFFE)
  fat[2 + seq_len(ns)] <- c(2 + seq_len(ns - 1L), 0xFFFFFFFE)
  header <- c(
    as.raw(c(0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1)), raw(16), u16(0x3E), u16(3), u16(0xFFFE), u16(9), u16(6), raw(6),
    u32(0), u32(1), u32(1), u32(0), u32(4096), u32(0xFFFFFFFE), u32(0), u32(0xFFFFFFFE), u32(0), u32(0),
    rep(u32(0xFFFFFFFF), 108)
  )
  entry <- function(name, type, child, start, size) {
    nm <- c(as.raw(rbind(as.integer(charToRaw(name)), 0L)), raw(2))
    c(nm, raw(64 - length(nm)), u16(length(nm)), u8(c(type, 1)), u32(0xFFFFFFFF), u32(0xFFFFFFFF), u32(child),
      raw(16), u32(0), raw(16), u32(start), u32(size), u32(0))
  }
  empty <- c(raw(66), raw(2), u32(0xFFFFFFFF), u32(0xFFFFFFFF), u32(0xFFFFFFFF), raw(48))

  directory <- c(entry("Root Entry", 5, 1, 0xFFFFFFFE, 0), entry(stream, 2, 0xFFFFFFFF, 2, length(book)), empty, empty)
  c(header, unlist(lapply(fat, u32)), directory, book)
}

test_that("reading a hand made Excel 2 file works", {
  u16 <- xls_u16
  u8 <- xls_u8
  rec <- xls_rec
  txt <- xls_txt
  attr3 <- function(xf = 0L, fmt = 0L, font = 0L, al = 0L) u8(c(xf, bitwAnd(fmt, 63L) + bitwShiftL(font, 6L), al))
  cell <- function(id, r, c, a, ...) rec(id, u16(r), u16(c), a, ...)
  ref <- function(r, c) c(u16(r), u8(c))
  fml <- function(...) {
    t <- c(...)
    c(u8(length(t)), t)
  }

  x <- c(
    rec(0x0009, u16(0), u16(0x10)),
    rec(0x0031, u16(200), u16(0), txt("Arial")),
    rec(0x0031, u16(240), u16(1), txt("Courier")),
    rec(0x001E, txt("General")), rec(0x001E, txt("0")), rec(0x001E, txt("0.00")), rec(0x001E, txt("0.00%")),
    rec(0x0043, u8(0), u8(0), u8(0), u8(0)),
    rec(0x0043, u8(1), u8(0), u8(0x40 + 2), u8(0x01 + 0x08 + 0x80)),
    rec(0x0043, u8(0), u8(0), u8(0x40 + 3), u8(0x02)),
    rec(0x0024, u8(0), u8(1), u16(3000)),
    rec(0x0000, u16(0), u16(8), u16(0), u16(6)),
    cell(0x0002, 0, 0, attr3(1), u16(42)),
    cell(0x0003, 0, 1, attr3(2), writeBin(3.25, raw())),
    rec(0x0044, u16(1)),
    cell(0x0004, 0, 2, attr3(0x3F), txt("hello")),
    cell(0x0005, 0, 4, attr3(0), u8(1), u8(0)),
    cell(0x0005, 0, 5, attr3(0), u8(0x07), u8(1)),
    cell(0x0001, 1, 0, attr3(2)),
    cell(0x0006, 1, 1, attr3(0), writeBin(45, raw()), u8(0),
         fml(u8(0x44), ref(0, 0), u8(0x44), ref(0, 1), u8(0x1E), u16(2), u8(0x05), u8(0x03))),
    cell(0x0006, 1, 2, attr3(0), u8(c(0, 0, 0, 0, 0, 0, 0xFF, 0xFF)), u8(0),
         fml(u8(0x17), u8(5), charToRaw("a b c"), u8(0x17), u8(0), u8(0x08))),
    rec(0x0007, txt("text result")),
    cell(0x0006, 1, 3, attr3(0), writeBin(7, raw()), u8(0),
         fml(u8(0x25), u16(0), u16(0), u8(0), u8(1), u8(0x42), u8(1), u8(4))),
    cell(0x0006, 2, 0, attr3(0), writeBin(1, raw()), u8(0), fml(u8(0x1D), u8(1))),
    cell(0x0006, 2, 1, attr3(0), writeBin(5, raw()), u8(0), fml(u8(0x01), u16(2), u8(0))),
    cell(0x0006, 3, 0, attr3(0), writeBin(1, raw()), u8(0), fml(u8(0x01), u16(3), u8(0))),
    rec(0x0021, u16(3), u16(4), u8(0), u8(0), u8(0), fml(u8(0x25), u16(0), u16(1), u8(0), u8(0))),
    cell(0x0006, 4, 0, attr3(0), writeBin(1.5, raw()), u8(0), fml(u8(0x43), u16(1), raw(5))),
    cell(0x0006, 5, 0, attr3(0), writeBin(8, raw()), u8(0), fml(u8(0x60))),
    cell(0x0006, 5, 1, attr3(0), u8(c(0, 0, 0, 0, 0, 0, 0xFF, 0xFF)), u8(0), fml(u8(0x60))),
    rec(0x0007, txt("kept")),
    cell(0x0006, 5, 2, attr3(0), u8(c(1, 0, 1, 0, 0, 0, 0xFF, 0xFF)), u8(0), fml(u8(0x60))),
    cell(0x0006, 5, 3, attr3(0), u8(c(2, 0, 7, 0, 0, 0, 0xFF, 0xFF)), u8(0), fml(u8(0x60))),
    rec(0x001C, u16(0), u16(0), u16(7), charToRaw("a note!")),
    rec(0x0008, u16(0), u16(0), u16(1), u16(0), u16(0), u8(0)),
    rec(0x0014, txt("head")), rec(0x0015, txt("foot")),
    rec(0x003E, u8(0), u8(1), u8(1), u8(0), u8(0), u16(0), u16(0), u8(0), u8(0), u8(0), u8(0)),
    rec(0x000A)
  )

  fl <- tempfile(fileext = ".xls")
  on.exit(unlink(fl), add = TRUE)
  writeBin(x, fl)

  wb <- wb_load(fl)
  got <- wb_to_df(wb, col_names = FALSE, show_formula = TRUE, convert = FALSE)
  expect_equal(unname(unlist(got[1, ])), c("42", "3.25", "hello", NA, "TRUE", "#DIV/0!"))
  expect_equal(got[2, 2], "$A$1+$B$1*2")
  expect_equal(got[2, 3], "\"a b c\"&\"\"")
  expect_equal(got[2, 4], "SUM($A$1:$B$1)")
  expect_equal(got[3, 1], "TRUE")
  expect_equal(got[3, 2], "5")
  expect_equal(got[4, 1], "$A$1:$A$2")
  expect_equal(unname(unlist(got[5:6, 1])), c("1.5", "8"))
  expect_equal(unname(unlist(got[6, 2:4])), c("kept", "TRUE", "#DIV/0!"))

  cmt <- wb_get_comment(wb)
  expect_equal(cmt$ref, "A1")
  expect_equal(cmt$comment, "a note!")
})


test_that("reading a hand made Excel 2 file without cell formats works", {
  u16 <- xls_u16
  u8 <- xls_u8
  rec <- xls_rec
  txt <- xls_txt
  attr3 <- function(flags, fmt, font, al) u8(c(flags, bitwAnd(fmt, 63L) + bitwShiftL(font, 6L), al))
  cell <- function(id, r, c, a, ...) rec(id, u16(r), u16(c), a, ...)

  shaded <- attr3(0x40, 2, 1, 0x80 + 0x08 + 0x02)
  x <- c(
    rec(0x0009, u16(0), u16(0x10)),
    rec(0x0031, u16(200), u16(0), txt("Arial")),
    rec(0x0031, u16(240), u16(1), txt("Courier")),
    rec(0x001E, txt("General")), rec(0x001E, txt("0")), rec(0x001E, txt("0.00")),
    rec(0x0000, u16(0), u16(1), u16(0), u16(3)),
    cell(0x0002, 0, 0, shaded, u16(42)),
    cell(0x0002, 0, 1, shaded, u16(43)),
    cell(0x0004, 0, 2, attr3(0xC0, 0, 0, 0), txt("hello")),
    rec(0x0025, u16(255)),
    rec(0x0036, u16(0)),
    rec(0x0018, u8(0), u8(0), u8(5), u8(7), charToRaw("total"), u8(0x25), u16(0), u16(0), u8(0), u8(1)),
    rec(0x000A)
  )

  fl <- tempfile(fileext = ".xls")
  on.exit(unlink(fl), add = TRUE)
  writeBin(x, fl)

  wb <- wb_load(fl)
  cc <- wb$worksheets[[1]]$sheet_data$cc
  expect_equal(cc$v, c("42", "43", ""))
  expect_equal(cc$c_s[1], cc$c_s[2])
  expect_false(cc$c_s[1] == cc$c_s[3])
  xf <- wb$styles_mgr$styles$cellXfs[as.integer(cc$c_s[1]) + 1L]
  expect_match(xf, "<alignment horizontal=\"center\"/>", fixed = TRUE)
  expect_match(wb$styles_mgr$styles$fonts[2], "<b/><sz val=\"12\"/><name val=\"Courier\"/>", fixed = TRUE)
  expect_match(wb$styles_mgr$styles$borders[2], "<left style=\"thin\">", fixed = TRUE)
  expect_true("<numFmt numFmtId=\"166\" formatCode=\"0.00\"/>" %in% wb$styles_mgr$styles$numFmts)
  expect_match(wb$styles_mgr$styles$cellXfs[as.integer(cc$c_s[3]) + 1L], "<protection locked=\"1\" hidden=\"1\"/>", fixed = TRUE)
  expect_equal(length(wb$workbook$definedNames), 0L)
})

test_that("reading hand made Excel 3 and 4 files works", {
  u16 <- xls_u16
  u8 <- xls_u8
  u32 <- xls_u32
  rec <- xls_rec
  txt <- xls_txt
  dbl <- function(x) writeBin(x, raw())
  fml <- function(...) {
    t <- c(...)
    c(u16(length(t)), t)
  }
  cell <- function(r, c, ...) c(u16(r), u16(c), ...)

  obj <- function(ot, id, ..., visible = 6L) {
    rec(0x005D, u32(1), u16(ot), u16(id), u8(c(0x16, visible)), u16(c(1, 10, 2, 10, 4, 10, 6, 10)), u16(0), u16(0), ...)
  }
  build <- function(b4, codepage = 1252, label = charToRaw("ab")) {
    xf <- function(font, fmt, flags, al, brd) {
      if (b4) {
        rec(0x0443, u8(c(font, fmt)), u16(flags), u8(c(al, 0)), u16(0), u32(brd))
      } else {
        rec(0x0243, u8(c(font, fmt, flags, 0)), u16(al), u16(0), u32(brd))
      }
    }
    format <- function(key, x) if (b4) rec(0x041E, u16(key), txt(x)) else rec(0x001E, txt(x))
    formula <- if (b4) 0x0406 else 0x0206
    c(
      rec(if (b4) 0x0409 else 0x0209, u16(0), u16(0x10)),
      rec(0x0042, u16(codepage)),
      rec(0x0022, u16(1)),
      rec(0x0012, u16(1)),
      rec(0x0013, u16(0x1234)),
      rec(0x003D, u16(0), u16(0), u16(1000), u16(1000), u8(0)),
      rec(0x0231, u16(200), u16(1), u16(8), txt("Arial")),
      rec(0x0231, u16(200), u16(0), u16(0), txt("Arial")),
      format(0, "General"), format(1, "0.00"), format(2, "0.0%"),
      xf(0, 0, 0x0004, 0, 0),
      xf(0, 1, 0x0001, 2 + 8, 0x01010101 + 0x00000102),
      xf(1, 2, 0, if (b4) 0x10 else 0x30, 0),
      rec(0x0293, u16(0), txt("Normal")),
      rec(0x0293, u16(0x8000), u8(0), u8(255)),
      rec(0x0092, u16(1), u8(c(255, 0, 0, 0))),
      rec(0x0218, u16(0), u8(0), u8(5), u16(7), charToRaw("total"), u8(0x25), u16(0), u16(0), u8(0), u8(1)),
      rec(0x0218, u16(0x20), u8(0), u8(1), u16(7), u8(6), u8(0x25), u16(0), u16(1), u8(0), u8(1)),
      rec(0x0218, u16(0), u8(0), u8(3), u16(4), charToRaw("one"), u8(0x24), u16(0), u8(0)),
      rec(0x0200, u16(0), u16(3), u16(0), u16(4), u16(0)),
      rec(0x0041, u16(1), u16(1), u16(1), u16(1), u8(0)),
      rec(0x001A, u16(1), u16(5)),
      rec(0x001B, u16(1), u16(2)),
      rec(0x007D, u16(0), u16(1), u16(2000), u16(0), u16(0), u16(0)),
      rec(0x0208, u16(0), u16(0), u16(4), u16(0xFF), u32(0), u32(0)),
      rec(0x0203, cell(0, 0, u16(1), dbl(1.5))),
      rec(0x0204, cell(0, 1, u16(0), u16(length(label)), label)),
      rec(0x0205, cell(0, 2, u16(2), u8(1), u8(0))),
      rec(0x0201, cell(0, 3, u16(2))),
      rec(formula, cell(1, 0, u16(1), dbl(3), u16(0),
                        fml(u8(0x44), u16(0), u8(0), u8(0x1E), u16(2), u8(0x05),
                            u8(0x42), u8(1), if (b4) u16(4) else u8(4), u8(0x15)))),
      rec(formula, cell(1, 1, u16(0), u8(c(0, 0, 0, 0, 0, 0, 0xFF, 0xFF)), u16(0), fml(u8(0x17), u8(2), charToRaw("hi")))),
      rec(0x0207, u16(3), charToRaw("res")),
      rec(formula, cell(1, 2, u16(0), dbl(1.5), u16(0), fml(u8(0x43), u16(1), raw(8)))),
      rec(0x001C, u16(0), u16(0), u16(7), charToRaw("a note!")),
      rec(0x001C, u16(1), u16(1), u16(6), charToRaw("long a")),
      rec(0x001C, u16(0xFFFF), u16(0), u16(6), charToRaw("nd cut")),
      rec(formula, cell(1, 3, u16(0), dbl(3.14), u16(0), fml(u8(0x21), if (b4) u16(19) else u8(19)))),
      obj(2, 1, u8(c(9, 10, 1, 0, 12, 0, 1, 0)), u16(1)),
      obj(3, 2, u8(c(9, 10, 0, 0, 12, 1, 2, 0)), u16(0)),
      obj(1, 3, u8(c(12, 0, 1, 0)), u16(2), u8(c(1, 0))),
      obj(6, 4, raw(60)),
      obj(8, 5, raw(60)),
      obj(2, 6, u8(c(9, 10, 1, 0, 12, 0, 1, 0)), u16(0), visible = 1L),
      rec(0x005D, raw(10)),
      rec(if (b4) 0x0409 else 0x0209, u16(0), u16(0x20)),
      rec(0x0203, cell(9, 9, u16(0), dbl(1))),
      rec(0x000A),
      rec(0x023E, u16(0x2EC), u16(0), u16(0), u16(0), u16(0), u16(0), u16(0)),
      rec(0x000A)
    )
  }

  fl <- tempfile(fileext = ".xls")
  on.exit(unlink(fl), add = TRUE)

  for (b4 in c(FALSE, TRUE)) {
    writeBin(build(b4), fl)
    wb <- wb_load(fl)
    got <- wb_to_df(wb, col_names = FALSE, show_formula = TRUE, convert = FALSE, skip_empty_cols = FALSE)
    expect_equal(unname(unlist(got[1, 1:3])), c("1.5", "ab", "TRUE"))
    expect_equal(got[2, 1], "(SUM($A$1*2))")
    expect_equal(got[2, 2], "\"hi\"")
    expect_equal(got[2, 3], "total")
    expect_equal(got[2, 4], "PI()")
    expect_equal(nrow(got), 2L)
    expect_equal(wb$workbook$workbookPr, "<workbookPr date1904=\"1\"/>")
    expect_equal(wb$worksheets[[1]]$sheetProtection, "<sheetProtection sheet=\"1\" password=\"1234\"/>")
    expect_equal(wb$worksheets[[1]]$rowBreaks, "<brk id=\"2\" max=\"16383\" man=\"1\"/>")
    expect_equal(wb$worksheets[[1]]$colBreaks, "<brk id=\"5\" max=\"1048575\" man=\"1\"/>")
    expect_match(wb$worksheets[[1]]$freezePane, "state=\"frozen\"", fixed = TRUE)

    dn <- wb$workbook$definedNames
    expect_equal(dn[1], "<definedName name=\"total\">Sheet1!$A$1:$B$1</definedName>")
    expect_equal(dn[2], "<definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\">Sheet1!$A$1:$B$2</definedName>")
    expect_equal(dn[3], "<definedName name=\"one\">Sheet1!$A$1</definedName>")

    cmt <- wb_get_comment(wb)
    expect_equal(cmt$ref, c("A1", "B2"))
    expect_equal(cmt$comment, c("a note!", "long and cut"))

    drawing <- paste(unlist(wb$drawings), collapse = "")
    expect_equal(lengths(regmatches(drawing, gregexpr("<xdr:sp ", drawing, fixed = TRUE))), 2L)
    expect_equal(lengths(regmatches(drawing, gregexpr("<xdr:cxnSp ", drawing, fixed = TRUE))), 1L)
    expect_match(drawing, "prst=\"roundRect\"", fixed = TRUE)
    expect_match(drawing, "prst=\"ellipse\"", fixed = TRUE)
    expect_match(drawing, "<a:ln w=\"25400\">", fixed = TRUE)
  }

  # text is converted with the code page of the file
  writeBin(build(FALSE, 1252, as.raw(c(0xE4, 0x80))), fl)
  expect_equal(wb_to_df(wb_load(fl), col_names = FALSE, skip_empty_cols = FALSE)[1, 2], "\u00e4\u20ac")
  writeBin(build(TRUE, 65001, as.raw(c(0xF0, 0x9F, 0x98, 0x80))), fl)
  expect_equal(wb_to_df(wb_load(fl), col_names = FALSE, skip_empty_cols = FALSE)[1, 2], "\U0001F600")
})

test_that("reading a hand made Excel 5 file works", {
  u16 <- xls_u16
  u8 <- xls_u8
  u32 <- xls_u32
  rec <- xls_rec
  txt <- xls_txt
  dbl <- function(x) writeBin(x, raw())
  fml <- function(...) {
    t <- c(...)
    c(u16(length(t)), t)
  }
  ref3 <- function(op, first, last, ...) c(u8(op), u16(0xFFFF), raw(8), u16(first), u16(last), ...)
  cell <- function(r, c, ...) c(u16(r), u16(c), ...)
  formula <- function(r, c, value, ...) rec(0x0006, cell(r, c, u16(0), dbl(value), u16(0), u32(0), fml(...)))
  xf <- function(font, fmt, tp, al) rec(0x00E0, u16(font), u16(fmt), u16(tp), u8(c(al, 0, 0, 0)), u32(0), u32(0), u16(0))
  bof <- function(type) rec(0x0809, u16(0x0500), u16(type), u16(0), u16(0))

  sheet1 <- c(
    bof(0x0010),
    rec(0x0200, u16(0), u16(8), u16(0), u16(8), u16(0)),
    rec(0x0203, cell(0, 0, u16(1), dbl(1))),
    rec(0x0203, cell(0, 1, u16(1), dbl(2))),
    rec(0x0204, cell(0, 2, u16(1), u16(1), charToRaw("x"))),
    formula(1, 0, 10, ref3(0x7A, 1, 1, u16(0), u8(0))),
    formula(1, 1, 3, ref3(0x7B, 1, 1, u16(0), u16(1), u8(0), u8(0)), u8(0x42), u8(1), u16(4)),
    formula(1, 2, 3, u8(0x43), u16(1), raw(12)),
    formula(2, 0, 2, u8(0x01), u16(2), u16(0)),
    rec(0x04BC, u16(2), u16(3), u8(0), u8(0), u8(0), u8(2), u16(8), u8(0x4C), u16(0xFFFE), u8(0), u8(0x1E), u16(2), u8(0x05)),
    formula(3, 0, 4, u8(0x01), u16(2), u16(0)),
    formula(0, 3, 1, u8(0x01), u16(0), u16(3)),
    rec(0x0221, u16(0), u16(1), u8(3), u8(3), u16(0), u32(0), u16(7), u8(0x25), u16(0), u16(1), u8(0), u8(0)),
    formula(5, 0, 7, u8(0x29), u16(7), u8(0x25), u16(0), u16(1), u8(0), u8(1), u8(0x42), u8(1), u16(4)),
    formula(5, 1, 8, u8(0x44), u16(0), u8(0), u8(0x29), u16(4), u8(0x44), u16(1), u8(0)),
    formula(5, 2, 9, u8(0x2A), u16(0), u8(0), u8(0x2B), u16(0), u16(0), u8(0), u8(0), u8(0x03)),
    formula(5, 3, 11, u8(0x1E), u16(1), u8(0x19), u8(0x04), u16(1), u16(0), u16(0)),
    formula(5, 4, 12, u8(0x7A), u16(1), raw(8), u16(0), u16(0), u16(0), u8(0)),
    formula(5, 5, 13, ref3(0x7C, 0, 0, u16(0), u8(0)), ref3(0x7D, 0, 0, u16(0), u16(0), u16(0))),
    formula(0, 6, 5, u8(0x01), u16(0), u16(6)),
    rec(0x04BC, u16(0), u16(1), u8(6), u8(6), u8(0), u8(2), u16(11), u8(0x4D), u16(0xC000), u16(0xC001), u8(0xFA), u8(0xFB),
        u8(0x42), u8(1), u16(4)),
    formula(1, 6, 6, u8(0x01), u16(0), u16(6)),
    formula(0, 7, 3, u8(0x01), u16(0), u16(7)),
    rec(0x0221, u16(0), u16(1), u8(7), u8(7), u16(0), u32(0), u16(1), u8(0x60)),
    rec(0x001C, u16(0), u16(0), u16(5), charToRaw("BIFF5")),
    rec(0x00A1, u16(9), u16(100), u16(1), u16(1), u16(1), u16(0), u16(600), u16(600), raw(18)),
    rec(0x1234, u16(1)),
    bof(0x0020),
    rec(0x0203, cell(9, 9, u16(1), dbl(1))),
    rec(0x000A),
    rec(0x000A)
  )
  sheet2 <- c(
    bof(0x0010),
    rec(0x0203, cell(0, 0, u16(1), dbl(5))),
    rec(0x0203, cell(1, 0, u16(1), dbl(6))),
    rec(0x000A)
  )
  sheet3 <- c(bof(0x0020), rec(0x000A))
  name <- function() {
    f <- c(u8(0x3B), u16(0xFFFF), raw(8), u16(0), u16(0), u16(0), u16(0), u8(0), u8(1))
    rec(0x0018, u16(0), u8(0), u8(5), u16(length(f)), u16(0), u16(0), raw(4), charToRaw("total"), f)
  }
  bad_name <- rec(0x0018, u16(0), u8(0), u8(3), u16(1), u16(0), u16(0), raw(4), charToRaw("bad"), u8(0x60))
  globals <- function(o1, o2, o3) {
    c(
      bof(0x0005),
      rec(0x0042, u16(1252)),
      rec(0x0012, u16(0)),
      rec(0x003D, u16(0), u16(0), u16(1000), u16(1000), u16(0), u16(0), u16(1), u16(600), u16(0)),
      rec(0x0031, u16(200), u16(0), u16(0x7FFF), u16(400), u16(0), u8(c(0, 0, 0, 0)), txt("Arial")),
      rec(0x041E, u16(164), txt("0.000")),
      xf(0, 0, 0xFFF5, 0), xf(0, 164, 0x0001, 2),
      rec(0x0085, u32(o1), u8(0), u8(0), txt("Data")),
      rec(0x0085, u32(o2), u8(0), u8(0), txt("Other")),
      rec(0x0085, u32(o3), u8(0), u8(2), txt("Chart")),
      name(), bad_name, bof(0x0020), rec(0x000A), rec(0x1234, u16(1)), rec(0x000A)
    )
  }
  o1 <- length(globals(0, 0, 0))
  o2 <- o1 + length(sheet1)
  o3 <- o2 + length(sheet2)
  book <- c(globals(o1, o2, o3), sheet1, sheet2, sheet3)
  fl <- tempfile(fileext = ".xls")
  on.exit(unlink(fl), add = TRUE)

  # BIFF5 workbooks are stored in a stream called Book, but some writers use the name Workbook
  for (stream in c("Book", "Workbook")) {
    writeBin(xls_ole(book, stream), fl)

    wb <- wb_load(fl)
    expect_equal(wb_get_sheet_names(wb), c(Data = "Data", Other = "Other", Chart = "Chart"))
    got <- wb_to_df(wb, sheet = 1, col_names = FALSE, show_formula = TRUE, convert = FALSE, skip_empty_cols = FALSE)
    expect_equal(got[2, 1], "Other!$A$1")
    expect_equal(got[2, 2], "SUM(Other!$A$1:$A$2)")
    expect_equal(got[2, 3], "total")
    expect_equal(got[3, 1], "A1*2")
    expect_equal(got[4, 1], "A2*2")
    expect_equal(got[1, 4], "$A$1:$A$2")
    expect_equal(got[6, 1], "SUM($A$1:$B$2)")
    expect_equal(got[6, 3], "#REF!+#REF!")
    expect_equal(unname(unlist(got[6, c(4, 5)])), c("1", "12"))
    expect_match(got[6, 6], "#REF!", fixed = TRUE)
    expect_equal(unname(unlist(got[1:2, 7])), c("SUM(A1:B2)", "SUM(A2:B3)"))
    expect_null(wb_to_df(wb, sheet = 3, col_names = FALSE))

    dn <- wb$workbook$definedNames
    expect_equal(dn[1], "<definedName name=\"total\">Data!$A$1:$B$1</definedName>")
    expect_equal(dn[2], "<definedName name=\"bad\">#REF!</definedName>")

    cmt <- wb_get_comment(wb, sheet = 1)
    expect_equal(cmt$ref, "A1")
    expect_equal(cmt$comment, "BIFF5")
  }
})

test_that("reading drawings and charts of an Excel 5 file works", {
  u16 <- xls_u16
  u8 <- xls_u8
  u32 <- xls_u32
  rec <- xls_rec
  txt <- xls_txt
  dbl <- function(x) writeBin(x, raw())
  bof <- function(type) rec(0x0809, u16(0x0500), u16(type), u16(0), u16(0))
  obj <- function(ot, id, anchor, ..., visible = 6L) {
    rec(0x005D, u32(1), u16(ot), u16(id), u8(c(0x16, visible)), u16(anchor), u16(0), raw(6), ...)
  }
  styles <- function(back, fore, fls, auto, icv, lns, lnw, lauto) u8(c(back, fore, fls, auto, icv, lns, lnw, lauto))
  textbox <- function(s, id, rot = 0) {
    t <- charToRaw(s)
    runs <- c(u16(0), u16(0), u32(0), u16(length(t)), u16(0), u32(0))
    obj(6, id, c(1, 100, 1, 50, 4, 100, 5, 200), styles(9, 10, 1, 0, 12, 0, 1, 0), u16(0), u16(length(t)), u16(0),
        u16(length(runs)), u16(0), u16(0), u8(c(0x12, 2)), u16(rot), raw(12), u8(0), u8(0), t, raw(length(t) %% 2L), runs)
  }
  picture <- function(id, anchor, cf, header, pixels, palette = raw(), split = 0L) {
    data <- c(header, palette, pixels)
    c(obj(8, id, anchor, styles(9, 10, 0, 0, 12, 0, 1, 0), u16(0), u16(cf), raw(20)),
      rec(0x007F, u16(cf), u16(1), u32(length(data)), data[seq_len(length(data) - split)]),
      if (split > 0L) rec(0x003C, data[length(data) - split + 1L]))
  }
  core <- function(w, h, bits) c(u32(12), u16(w), u16(h), u16(1), u16(bits))
  num <- function(r, c, v) rec(0x0203, u16(r), u16(c), u16(1), dbl(v))

  area3 <- function(r1, r2, c) c(u8(0x7B), u16(0xFFFF), raw(8), u16(0), u16(0), u16(r1), u16(r2), u8(c), u8(c))
  ai <- function(id, rt, f = raw()) rec(0x1051, u8(c(id, rt)), u16(0), u16(0), u16(length(f)), f)
  begin <- rec(0x1033)
  end <- rec(0x1034)
  external <- c(u8(0x7B), u16(1), raw(8), u16(0), u16(0), u16(0), u16(2), u8(1), u8(1))
  series <- function(name, values = area3(0, 2, 1), ...) {
    c(rec(0x1003, u16(1), u16(1), u16(3), u16(3)), begin, ai(0, 1), ai(1, 2, values), ai(2, 2, area3(0, 2, 0)), ai(3, 1),
      rec(0x100D, u16(0xFFFF), u8(nchar(name)), charToRaw(name)), ..., rec(0x1045, u16(0)), end)
  }
  axes <- c(
    rec(0x101D, u16(0), raw(16)), begin, rec(0x1020, u16(1), u16(1), u16(1), u16(1)), rec(0x101E, u8(c(2, 0, 3, 1)), u32(0), raw(16), u16(0)), end,
    rec(0x101D, u16(1), raw(16)), begin, rec(0x101F, dbl(0), dbl(100), dbl(10), dbl(1), dbl(0), u16(0x1F)), rec(0x1021, u16(1)), end
  )
  frame <- function(...) c(rec(0x1002, u32(0), u32(0), u32(400 * 65536), u32(250 * 65536)), begin, ..., end)
  bars <- frame(
    series("Qty", area3(0, 2, 1), rec(0x1006, u16(0xFFFF), u16(0), u16(0), u16(0)), begin, rec(0x100A, u32(0x0000FF), u32(0), u16(1), u16(0)), end),
    rec(0x1041, u16(0), raw(16)), begin, axes, rec(0x1014, raw(16), u16(1), u16(0)), begin, rec(0x1017, u16(0), u16(150), u16(0)), end, end
  )
  lines <- frame(
    series("Failing", external, rec(0x1006, u16(0xFFFF), u16(0), u16(0), u16(0)), begin, rec(0x1009, u32(0), u32(0xFFFFFF), u16(1), u16(0)),
           rec(0x1007, u32(0), u16(0), u16(0), u16(1)), end),
    rec(0x1041, u16(0), raw(16)), begin, axes, rec(0x1014, raw(16), u16(0), u16(0)), begin, rec(0x1018, u16(0)), end, end,
    rec(0x1025, u8(c(2, 2)), u16(1), u32(0), raw(16), u16(0x0100 + 0x0010)), begin, rec(0x100D, u16(0xFFFF), u8(5), charToRaw("Title")), end,
    rec(0x1027, u16(1), u16(0), u16(0)),
    rec(0x1065, u16(2)), rec(0x0204, u16(0), u16(0), u16(1), u16(1), charToRaw("a")), rec(0x0203, u16(1), u16(0), u16(1), dbl(2)),
    rec(0x0201, u16(2), u16(0), u16(1)), rec(0x1234, u16(1))
  )

  pixels24 <- c(0, 0, 255, 0, 255, 0, 0, 0, 255, 0, 0, 255, 255, 255, 0, 0)
  sheet1 <- c(
    bof(0x0010), rec(0x0200, u16(0), u16(3), u16(0), u16(2), u16(0)),
    num(0, 0, 1), num(1, 0, 2), num(2, 0, 3), num(0, 1, 10), num(1, 1, 20), num(2, 1, 30),
    obj(0, 9, c(0, 0, 0, 0, 0, 0, 0, 0), raw(20)),
    obj(2, 1, c(0, 10, 7, 10, 2, 10, 9, 10), styles(9, 8, 1, 0, 12, 0, 2, 0), u16(0), u8(0), u8(0)),
    obj(2, 23, c(0, 10, 7, 10, 2, 10, 9, 10), styles(9, 65, 1, 0, 65, 0, 1, 0), u16(0), u8(0), u8(0)),
    obj(2, 2, c(0, 10, 7, 10, 2, 10, 9, 10), styles(9, 10, 1, 0, 12, 0, 2, 0), u16(1), u8(0), u8(0)),
    obj(3, 3, c(3, 10, 7, 10, 5, 10, 9, 10), styles(9, 10, 0, 0, 12, 1, 1, 0), u16(0), u8(0), u8(0)),
    obj(3, 4, c(3, 10, 7, 10, 5, 10, 9, 10), styles(9, 10, 1, 1, 12, 5, 1, 0), u16(0), u8(0), u8(0)),
    obj(1, 5, c(6, 0, 1, 0, 8, 0, 3, 0), u8(c(12, 0, 3, 0)), u16(2), u8(c(1, 0)), u8(0), u8(0)),
    obj(1, 6, c(6, 0, 1, 0, 8, 0, 3, 0), u8(c(12, 2, 2, 1)), u16(0x0012), u8(c(2, 0)), u8(0), u8(0)),
    obj(1, 7, c(6, 0, 1, 0, 8, 0, 3, 0), u8(c(12, 1, 1, 1)), u16(1), u8(c(3, 0)), u8(0), u8(0)),
    obj(1, 8, c(6, 0, 1, 0, 8, 0, 3, 0), u8(c(12, 3, 0, 0)), u16(0), u8(c(0, 0)), u8(0), u8(0)),
    obj(4, 10, c(0, 0, 0, 0, 1, 0, 1, 0), styles(9, 10, 0, 0, 12, 0, 1, 0), u16(0), u8(0), u8(0)),
    obj(2, 11, c(0, 0, 0, 0, 1, 0, 1, 0), styles(9, 10, 1, 0, 12, 0, 1, 0), u16(0), u8(0), u8(0), visible = 1L),
    textbox("Hello", 12),
    textbox("odd", 13),
    obj(6, 14, c(1, 0, 1, 0, 2, 0, 2, 0), styles(9, 10, 1, 0, 12, 0, 1, 0), u16(0), u16(100), raw(100)),
    picture(15, c(1, 0, 12, 0, 3, 0, 15, 0), 9, core(2, 2, 24), as.raw(pixels24), split = 1L),
    picture(16, c(1, 0, 12, 0, 3, 0, 15, 0), 9, core(2, 2, 1), as.raw(c(0, 0, 0, 0, 0xC0, 0, 0, 0)), as.raw(c(0, 0, 0, 255, 255, 255))),
    picture(17, c(1, 0, 12, 0, 3, 0, 15, 0), 2, core(2, 2, 24), as.raw(pixels24)),
    picture(18, c(1, 0, 12, 0, 3, 0, 15, 0), 9, core(2, 2, 16), as.raw(pixels24)),
    picture(19, c(1, 0, 12, 0, 3, 0, 15, 0), 9, core(0, 2, 24), as.raw(pixels24)),
    picture(20, c(1, 0, 12, 0, 3, 0, 15, 0), 9, core(2, 2, 24), as.raw(pixels24[1:8])),
    picture(21, c(1, 0, 12, 0, 3, 0, 15, 0), 9, u32(40), as.raw(pixels24)),
    obj(8, 22, c(1, 0, 12, 0, 3, 0, 15, 0), styles(9, 10, 0, 0, 12, 0, 1, 0), u16(0), u16(9), raw(20)),
    obj(5, 30, c(3, 0, 1, 0, 9, 0, 14, 0), styles(9, 10, 1, 0, 12, 0, 1, 0), u16(0), raw(40)),
    bof(0x0020), bars, rec(0x000A),
    obj(5, 31, c(3, 0, 16, 0, 9, 0, 30, 0), styles(9, 10, 1, 0, 12, 0, 1, 0), u16(0), raw(40)),
    bof(0x0020), lines, bof(0x0020), rec(0x1234, u16(1)), rec(0x000A), rec(0x000A),
    rec(0x000A)
  )
  sheet2 <- c(bof(0x0020), rec(0x0014, u8(0)), bars, rec(0x1234, u16(1)), rec(0x000A))
  sheet4 <- c(bof(0x0006), rec(0x000A))
  sheet3 <- c(bof(0x0010), rec(0x0200, u16(0), u16(0), u16(0), u16(0), u16(0)), obj(2, 1, c(0, 0, 0, 0, 1, 0, 1, 0), styles(9, 10, 1, 0, 12, 0, 1, 0), u16(0), u8(0), u8(0)), rec(0x000A))
  globals <- function(o1, o2, o3, o4) {
    c(
      bof(0x0005), rec(0x0042, u16(1252)),
      rec(0x0092, u16(2), u8(c(0, 0, 255, 0)), u8(c(0, 255, 0, 0))),
      rec(0x0031, u16(200), u16(0), u16(0x7FFF), u16(400), u16(0), u8(c(0, 0, 0, 0)), txt("Arial")),
      rec(0x00E0, u16(0), u16(0), u16(0xFFF5), u8(c(0, 0, 0, 0)), u32(0), u32(0), u16(0)),
      rec(0x0085, u32(o1), u8(0), u8(0), txt("Data")), rec(0x0085, u32(o2), u8(0), u8(2), txt("Chart1")),
      rec(0x0085, u32(o3), u8(0), u8(0), txt("Late")), rec(0x0085, u32(o4), u8(0), u8(6), txt("Module")), rec(0x000A)
    )
  }
  o1 <- length(globals(0, 0, 0, 0))
  o2 <- o1 + length(sheet1)
  o3 <- o2 + length(sheet2)
  o4 <- o3 + length(sheet3)
  fl <- tempfile(fileext = ".xls")
  on.exit(unlink(fl), add = TRUE)
  writeBin(xls_ole(c(globals(o1, o2, o3, o4), sheet1, sheet2, sheet3, sheet4), "Book"), fl)

  wb <- wb_load(fl)
  expect_equal(unname(wb_get_sheet_names(wb)), c("Data", "Chart1", "Late", "Module"))

  drawings <- unlist(wb$drawings)
  count <- function(pattern, x = drawings) sum(lengths(regmatches(x, gregexpr(pattern, x, fixed = TRUE))))
  expect_equal(count("<xdr:sp "), 8L)
  expect_equal(count("<xdr:cxnSp "), 4L)
  expect_equal(count("<xdr:pic>"), 2L)
  expect_equal(count("<xdr:graphicFrame"), 3L)
  expect_equal(count("prst=\"ellipse\""), 2L)
  expect_equal(count("prst=\"roundRect\""), 1L)
  expect_equal(count("<a:t>Hello</a:t>"), 1L)
  expect_equal(count("<a:t>odd</a:t>"), 1L)
  expect_equal(count("<a:srgbClr val=\"0000FF\"/>"), 9L)
  expect_equal(count("<a:srgbClr val=\"FFFFFF\"/>"), 3L)
  expect_equal(count("<a:srgbClr val=\"FF0000\"/>"), 4L)
  expect_equal(count("<a:ln w=\"25400\">"), 2L)
  expect_equal(count("<a:ln w=\"38100\">"), 1L)
  expect_equal(count("<a:ln w=\"3175\">"), 1L)
  expect_equal(count("<a:noFill/>"), 2L)
  expect_equal(count("<a:tailEnd type=\"triangle\"/>"), 2L)
  expect_equal(count("<a:tailEnd type=\"arrow\"/>"), 1L)
  expect_equal(count("flipH=\"1\""), 2L)
  expect_equal(count("flipV=\"1\""), 2L)
  expect_equal(count("prstDash val=\"sysDash\""), 1L)

  expect_equal(length(wb$media), 2L)
  bmp <- lapply(wb$media, readBin, what = "raw", n = 200L)
  expect_equal(rawToChar(bmp[[1]][1:2]), "BM")
  expect_equal(as.integer(bmp[[1]][55:62]), pixels24[1:8])
  expect_equal(as.integer(bmp[[2]][55:58]), c(0, 0, 0, 0))
  expect_equal(as.integer(bmp[[2]][27:30]), c(1, 0, 1, 0))

  charts <- unlist(wb$charts)
  expect_equal(count("<c:barChart>", charts), 2L)
  expect_equal(count("<c:lineChart>", charts), 1L)
  expect_equal(count("<c:f>Data!$B$1:$B$3</c:f>", charts), 2L)
  expect_equal(count("<c:f>Data!$A$1:$A$3</c:f>", charts), 3L)
  expect_equal(count("<c:v>Qty</c:v>", charts), 2L)
  expect_equal(count("<c:v>Failing</c:v>", charts), 1L)
})

test_that("reading xls files of older versions works", {
  # BIFF3 and BIFF4 are plain record streams, BIFF5 and BIFF7 are compound files
  for (nm in c("testEXCEL_3.xls", "testEXCEL_4.xls", "testEXCEL_5.xls", "testEXCEL_95.xls", "biff4_no_format_no_window2.xls")) {
    fl <- testfile_path(nm)
    wb <- wb_load(fl)
    expect_true(length(wb$get_sheet_names()) >= 1L, info = nm)
    expect_true(nrow(wb_to_df(wb, col_names = FALSE)) > 1L, info = nm)
  }

  wb <- wb_load(testfile_path("testEXCEL_3.xls"))
  got <- wb_to_df(wb, col_names = FALSE)
  expect_equal(dim(got), c(35L, 10L))
  expect_equal(got[1, 1], "Table C-13--Lemons: Utilization of production, by State, 1977/78 to date")

  wb <- wb_load(testfile_path("testEXCEL_5.xls"))
  got <- wb_to_df(wb, col_names = FALSE)
  expect_equal(got[1, 1], "Sample Excel Worksheet - Numbers and their Squares")
  expect_equal(got[4, 2:3], data.frame(B = "Number", C = "Square", row.names = 4L))

  wb <- wb_load(testfile_path("biff4_no_format_no_window2.xls"))
  got <- wb_to_df(wb, col_names = FALSE)
  expect_equal(dim(got), c(27L, 4L))
  expect_equal(got[2, 2], "DOG & CAT PRODUCTS")
})
