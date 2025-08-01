test_that("openxlsx2_types", {

  # test vector types
  expect_equal(openxlsx2_celltype[["short_date"]], openxlsx2_type(Sys.Date()))
  expect_equal(openxlsx2_celltype[["long_date"]], openxlsx2_type(as.POSIXct(Sys.Date())))
  expect_equal(openxlsx2_celltype[["numeric"]], openxlsx2_type(1))
  expect_equal(openxlsx2_celltype[["logical"]], openxlsx2_type(TRUE))
  expect_equal(openxlsx2_celltype[["character"]], openxlsx2_type("a"))
  expect_equal(openxlsx2_celltype[["factor"]], openxlsx2_type(as.factor(1)))

  # even complex numbers
  z <- complex(real = 0.9156691, imaginary = -0.5533739)
  expect_equal(openxlsx2_celltype[["character"]], openxlsx2_type(z))

  # wb_add_data_table() example: data frame with various types
  df <- data.frame(
    "Date" = Sys.Date() - 0:19,
    "T" = TRUE, "F" = FALSE,
    "Time" = Sys.time() - 0:19 * 60 * 60,
    "Cash" = 1:20, "Cash2" = 31:50,
    "hLink" = "https://CRAN.R-project.org/",
    "Percentage" = seq(0, 1, length.out = 20),
    "TinyNumbers" = runif(20) / 1E9, stringsAsFactors = FALSE
  )

  ## openxlsx will apply default Excel styling for these classes
  class(df$Cash) <- c(class(df$Cash), "currency")
  class(df$Cash2) <- c(class(df$Cash2), "accounting")
  class(df$hLink) <- "hyperlink"
  class(df$Percentage) <- c(class(df$Percentage), "percentage")
  class(df$TinyNumbers) <- c(class(df$TinyNumbers), "scientific")

  got <- openxlsx2_type(df)
  exp <- c(
    Date = openxlsx2_celltype[["short_date"]],
    `T` = openxlsx2_celltype[["logical"]],
    `F` = openxlsx2_celltype[["logical"]],
    Time = openxlsx2_celltype[["long_date"]],
    Cash = openxlsx2_celltype[["currency"]],
    Cash2 = openxlsx2_celltype[["accounting"]],
    hLink = openxlsx2_celltype[["hyperlink"]],
    Percentage = openxlsx2_celltype[["percentage"]],
    TinyNumbers = openxlsx2_celltype[["scientific"]]
  )

  expect_equal(exp, got)

  wb <- wb_workbook()$add_worksheet()$add_data(x = df)
  xf <- rbindlist(xml_attr(wb$styles_mgr$styles$cellXfs, "xf"))

  exp <- c("0", "0", "14", "22", "44", "4", "10", "48")
  got <- xf$numFmtId
  expect_equal(exp, got)

})

test_that("wb_page_setup example", {

  wb <- wb_workbook()
  wb$add_worksheet("S1")
  wb$add_worksheet("S2")
  wb$add_data_table(1, x = iris[1:30, ])

  ## landscape page scaled to 50%
  wb$page_setup(sheet = 1, orientation = "landscape", scale = 50)
  exp <- "<pageSetup paperSize=\"9\" orientation=\"landscape\" horizontalDpi=\"300\" verticalDpi=\"300\" fitToHeight=\"0\" fitToWidth=\"0\" scale=\"50\"/>"
  expect_equal(exp, wb$worksheets[[1]]$pageSetup)


  ## portrait page scales to 300% with 0.5in left and right margins
  wb$page_setup(sheet = 2, orientation = "portrait", scale = 300, left = 0.5, right = 0.5)
  exp <- "<pageSetup paperSize=\"9\" orientation=\"portrait\" horizontalDpi=\"300\" verticalDpi=\"300\" fitToHeight=\"0\" fitToWidth=\"0\" scale=\"300\"/>"
  expect_equal(exp, wb$worksheets[[2]]$pageSetup)


  ## print titles
  wb$add_worksheet("print_title_rows")
  wb$add_worksheet("print_title_cols")

  wb$add_data("print_title_rows", rbind(iris, iris, iris, iris))
  wb$add_data("print_title_cols", x = rbind(mtcars, mtcars, mtcars), row_names = TRUE)

  wb$page_setup(sheet = "print_title_rows", print_title_rows = 1) ## first row
  wb$page_setup(sheet = "print_title_cols", print_title_cols = 1, print_title_rows = 1)

  exp <- c(
    "<definedName localSheetId=\"2\" name=\"_xlnm.Print_Titles\">'print_title_rows'!$1:$1</definedName>",
    "<definedName name=\"_xlnm.Print_Titles\" localSheetId=\"3\">'print_title_cols'!$A:$A,'print_title_cols'!$1:$1</definedName>"
  )
  expect_equal(exp, wb$workbook$definedNames)

  tmp <- temp_xlsx()
  expect_silent(wb_save(wb, tmp, overwrite = TRUE))

  # survives write and load
  wb <- wb_load(tmp)
  expect_equal(exp, wb$workbook$definedNames)


})

test_that("amp_split & genHeaderFooterNode", {

  xml <- paste0(
    "<headerFooter differentOddEven=\"0\" differentFirst=\"0\" scaleWithDoc=\"0\" alignWithMargins=\"0\">",
    "<oddHeader>&amp;C&amp;&quot;Times New Roman,Standard&quot;&amp;12&amp;A</oddHeader>",
    "<oddFooter>&amp;C&amp;&quot;Times New Roman,Standard&quot;&amp;12Seite &amp;P</oddFooter>",
    "</headerFooter>"
  )

  exp <- list(
    oddHeader = c("", "&amp;&quot;Times New Roman,Standard&quot;&amp;12&amp;A", ""),
    oddFooter = c("", "&amp;&quot;Times New Roman,Standard&quot;&amp;12Seite &amp;P", "")
  )
  got <- getHeaderFooterNode(xml)
  expect_equal(exp, got)

  exp <- xml
  got <- genHeaderFooterNode(got)
  expect_equal(exp, got)

  xml <- "<headerFooter alignWithMargins=\"0\"><oddFooter>&amp;L^  &amp;D  +&amp;C&amp;R</oddFooter></headerFooter>"

  exp <- list(oddFooter = c("^  &amp;D  +", "", ""))
  got <- getHeaderFooterNode(xml)
  expect_equal(exp, got)

  exp <- "<headerFooter differentOddEven=\"0\" differentFirst=\"0\" scaleWithDoc=\"0\" alignWithMargins=\"0\"><oddFooter>&amp;L^  &amp;D  +</oddFooter></headerFooter>"
  got <- genHeaderFooterNode(got)
  expect_equal(exp, got)

})

test_that("add_sparklines", {

  set.seed(123) # sparklines has a random uri string
  options("openxlsx2_seed" = NULL)

  sparklines <- c(
    create_sparklines("Sheet 1", "A3:L3", "M3", type = "column", first = "1"),
    create_sparklines("Sheet 1", "A2:L2", "M2", markers = "1"),
    create_sparklines("Sheet 1", "A4:L4", "M4", type = "stacked", negative = "1")
  )

  t1 <- AirPassengers
  t2 <- do.call(cbind, split(t1, cycle(t1)))
  dimnames(t2) <- dimnames(.preformat.ts(t1))


  wb <- wb_workbook()$
    add_worksheet("Sheet 1")$
    add_data(x = t2)$
    add_sparklines(sparklines = sparklines)

  exp <- read_xml('<ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
 <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
  <x14:sparklineGroup type="column" displayEmptyCellsAs="gap" first="1" xr2:uid="{6F57B887-24F1-C14A-942C-4C6EF08E87F7}">
   <x14:colorSeries rgb="FF376092" />
   <x14:colorNegative rgb="FFD00000" />
   <x14:colorAxis rgb="FFD00000" />
   <x14:colorMarkers rgb="FFD00000" />
   <x14:colorFirst rgb="FFD00000" />
   <x14:colorLast rgb="FFD00000" />
   <x14:colorHigh rgb="FFD00000" />
   <x14:colorLow rgb="FFD00000" />
   <x14:sparklines>
    <x14:sparkline>
     <xm:f>\'Sheet 1\'!A3:L3</xm:f>
     <xm:sqref>M3</xm:sqref>
    </x14:sparkline>
   </x14:sparklines>
  </x14:sparklineGroup>
  <x14:sparklineGroup displayEmptyCellsAs="gap" markers="1" xr2:uid="{6F57B887-24F1-C14A-942C-459E3EFAA032}">
   <x14:colorSeries rgb="FF376092" />
   <x14:colorNegative rgb="FFD00000" />
   <x14:colorAxis rgb="FFD00000" />
   <x14:colorMarkers rgb="FFD00000" />
   <x14:colorFirst rgb="FFD00000" />
   <x14:colorLast rgb="FFD00000" />
   <x14:colorHigh rgb="FFD00000" />
   <x14:colorLow rgb="FFD00000" />
   <x14:sparklines>
    <x14:sparkline>
     <xm:f>\'Sheet 1\'!A2:L2</xm:f>
     <xm:sqref>M2</xm:sqref>
    </x14:sparkline>
   </x14:sparklines>
  </x14:sparklineGroup>
  <x14:sparklineGroup type="stacked" displayEmptyCellsAs="gap" negative="1" xr2:uid="{6F57B887-24F1-C14A-942C-2B92FF2D7883}">
   <x14:colorSeries rgb="FF376092" />
   <x14:colorNegative rgb="FFD00000" />
   <x14:colorAxis rgb="FFD00000" />
   <x14:colorMarkers rgb="FFD00000" />
   <x14:colorFirst rgb="FFD00000" />
   <x14:colorLast rgb="FFD00000" />
   <x14:colorHigh rgb="FFD00000" />
   <x14:colorLow rgb="FFD00000" />
   <x14:sparklines>
    <x14:sparkline>
     <xm:f>\'Sheet 1\'!A4:L4</xm:f>
     <xm:sqref>M4</xm:sqref>
    </x14:sparkline>
   </x14:sparklines>
  </x14:sparklineGroup>
 </x14:sparklineGroups>
</ext>', pointer = FALSE)
  got <- wb$worksheets[[1]]$extLst
  expect_equal(exp, got)

  expect_error(
    wb$add_sparklines(sparklines = xml_node_create("sparklines", sparklines)),
    "all nodes must match x14:sparklineGroup. Got sparklines"
  )

})

test_that("more sparkline tests", {

  set.seed(123) # sparklines has a random uri string
  options("openxlsx2_seed" = NULL)

  sl1 <- create_sparklines("Sheet 1", "A3:K3", "L3")
  sl2 <- create_sparklines("Sheet 1", "A4:K4", "L4", type = "column", high = TRUE, low = TRUE)
  sl3 <- create_sparklines("Sheet 1", "A5:K5", "L5", type = "stacked", display_empty_cells_as = 0)

  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = mtcars)$
    add_sparklines(sparklines = c(sl1, sl2, sl3))

  exp <- "<ext xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" uri=\"{05C60535-1F16-4fd2-B633-F4F36F0B64E0}\"><x14:sparklineGroups xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"><x14:sparklineGroup displayEmptyCellsAs=\"gap\" xr2:uid=\"{6F57B887-24F1-C14A-942C-4C6EF08E87F7}\"><x14:colorSeries rgb=\"FF376092\"/><x14:colorNegative rgb=\"FFD00000\"/><x14:colorAxis rgb=\"FFD00000\"/><x14:colorMarkers rgb=\"FFD00000\"/><x14:colorFirst rgb=\"FFD00000\"/><x14:colorLast rgb=\"FFD00000\"/><x14:colorHigh rgb=\"FFD00000\"/><x14:colorLow rgb=\"FFD00000\"/><x14:sparklines><x14:sparkline><xm:f>'Sheet 1'!A3:K3</xm:f><xm:sqref>L3</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup><x14:sparklineGroup type=\"column\" displayEmptyCellsAs=\"gap\" high=\"1\" low=\"1\" xr2:uid=\"{6F57B887-24F1-C14A-942C-459E3EFAA032}\"><x14:colorSeries rgb=\"FF376092\"/><x14:colorNegative rgb=\"FFD00000\"/><x14:colorAxis rgb=\"FFD00000\"/><x14:colorMarkers rgb=\"FFD00000\"/><x14:colorFirst rgb=\"FFD00000\"/><x14:colorLast rgb=\"FFD00000\"/><x14:colorHigh rgb=\"FFD00000\"/><x14:colorLow rgb=\"FFD00000\"/><x14:sparklines><x14:sparkline><xm:f>'Sheet 1'!A4:K4</xm:f><xm:sqref>L4</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup><x14:sparklineGroup type=\"stacked\" displayEmptyCellsAs=\"0\" xr2:uid=\"{6F57B887-24F1-C14A-942C-2B92FF2D7883}\"><x14:colorSeries rgb=\"FF376092\"/><x14:colorNegative rgb=\"FFD00000\"/><x14:colorAxis rgb=\"FFD00000\"/><x14:colorMarkers rgb=\"FFD00000\"/><x14:colorFirst rgb=\"FFD00000\"/><x14:colorLast rgb=\"FFD00000\"/><x14:colorHigh rgb=\"FFD00000\"/><x14:colorLow rgb=\"FFD00000\"/><x14:sparklines><x14:sparkline><xm:f>'Sheet 1'!A5:K5</xm:f><xm:sqref>L5</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>"
  got <- wb$worksheets[[1]]$extLst
  expect_equal(exp, got)

})

test_that("add_sparkline_group", {

  set.seed(123) # sparklines has a random uri string
  options("openxlsx2_seed" = NULL)

  sparklines <- c(
    create_sparklines(
      "Sheet 1", "A2:L13", "M2:M13", type = "column"
    ),
    create_sparklines(
      "Sheet 2", "A2:L13", "A14:L14", direction =  "col",
      high = TRUE, low = TRUE,
      color_series = wb_color(hex = "FF323232"),
      color_high = wb_color(hex = "FF00B050"),
      color_low = wb_color(hex = "FFFF0000")
    )
  )

  t1 <- AirPassengers
  t2 <- do.call(cbind, split(t1, cycle(t1)))
  dimnames(t2) <- dimnames(.preformat.ts(t1))


  wb <- wb_workbook()$
    add_worksheet("Sheet 1")$
    add_data(x = t2)$
    add_sparklines(sparklines = sparklines[[1]])$
    add_worksheet("Sheet 2")$
    add_data(x = t2)$
    add_sparklines(sparklines = sparklines[[2]])

  exp <- read_xml('<ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
 <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
  <x14:sparklineGroup type="column" displayEmptyCellsAs="gap" xr2:uid="{6F57B887-24F1-C14A-942C-4C6EF08E87F7}">
    <x14:colorSeries rgb="FF376092"/>
    <x14:colorNegative rgb="FFD00000"/>
    <x14:colorAxis rgb="FFD00000"/>
    <x14:colorMarkers rgb="FFD00000"/>
    <x14:colorFirst rgb="FFD00000"/>
    <x14:colorLast rgb="FFD00000"/>
    <x14:colorHigh rgb="FFD00000"/>
    <x14:colorLow rgb="FFD00000"/>
    <x14:sparklines>
      <x14:sparkline>
        <xm:f>\'Sheet 1\'!A2:L2</xm:f>
        <xm:sqref>M2</xm:sqref>
      </x14:sparkline>
      <x14:sparkline>
        <xm:f>\'Sheet 1\'!A3:L3</xm:f>
        <xm:sqref>M3</xm:sqref>
      </x14:sparkline>
      <x14:sparkline>
        <xm:f>\'Sheet 1\'!A4:L4</xm:f>
        <xm:sqref>M4</xm:sqref>
      </x14:sparkline>
      <x14:sparkline>
        <xm:f>\'Sheet 1\'!A5:L5</xm:f>
        <xm:sqref>M5</xm:sqref>
      </x14:sparkline>
      <x14:sparkline>
        <xm:f>\'Sheet 1\'!A6:L6</xm:f>
        <xm:sqref>M6</xm:sqref>
      </x14:sparkline>
      <x14:sparkline>
        <xm:f>\'Sheet 1\'!A7:L7</xm:f>
        <xm:sqref>M7</xm:sqref>
      </x14:sparkline>
      <x14:sparkline>
        <xm:f>\'Sheet 1\'!A8:L8</xm:f>
        <xm:sqref>M8</xm:sqref>
      </x14:sparkline>
      <x14:sparkline>
        <xm:f>\'Sheet 1\'!A9:L9</xm:f>
        <xm:sqref>M9</xm:sqref>
      </x14:sparkline>
      <x14:sparkline>
        <xm:f>\'Sheet 1\'!A10:L10</xm:f>
        <xm:sqref>M10</xm:sqref>
      </x14:sparkline>
      <x14:sparkline>
        <xm:f>\'Sheet 1\'!A11:L11</xm:f>
        <xm:sqref>M11</xm:sqref>
      </x14:sparkline>
      <x14:sparkline>
        <xm:f>\'Sheet 1\'!A12:L12</xm:f>
        <xm:sqref>M12</xm:sqref>
      </x14:sparkline>
      <x14:sparkline>
        <xm:f>\'Sheet 1\'!A13:L13</xm:f>
        <xm:sqref>M13</xm:sqref>
      </x14:sparkline>
    </x14:sparklines>
  </x14:sparklineGroup>
</x14:sparklineGroups>
</ext>', pointer = FALSE)
  got <- wb$worksheets[[1]]$extLst

  expect_equal(exp, got)

  exp <- read_xml('<ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
 <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
  <x14:sparklineGroup displayEmptyCellsAs="gap" high="1" low="1" xr2:uid="{6F57B887-24F1-C14A-942C-459E3EFAA032}">
    <x14:colorSeries rgb="FF323232"/>
    <x14:colorNegative rgb="FFD00000"/>
    <x14:colorAxis rgb="FFD00000"/>
    <x14:colorMarkers rgb="FFD00000"/>
    <x14:colorFirst rgb="FFD00000"/>
    <x14:colorLast rgb="FFD00000"/>
    <x14:colorHigh rgb="FF00B050"/>
    <x14:colorLow rgb="FFFF0000"/>
    <x14:sparklines>
      <x14:sparkline>
        <xm:f>\'Sheet 2\'!A2:A13</xm:f>
        <xm:sqref>A14</xm:sqref>
      </x14:sparkline>
      <x14:sparkline>
        <xm:f>\'Sheet 2\'!B2:B13</xm:f>
        <xm:sqref>B14</xm:sqref>
      </x14:sparkline>
      <x14:sparkline>
        <xm:f>\'Sheet 2\'!C2:C13</xm:f>
        <xm:sqref>C14</xm:sqref>
      </x14:sparkline>
      <x14:sparkline>
        <xm:f>\'Sheet 2\'!D2:D13</xm:f>
        <xm:sqref>D14</xm:sqref>
      </x14:sparkline>
      <x14:sparkline>
        <xm:f>\'Sheet 2\'!E2:E13</xm:f>
        <xm:sqref>E14</xm:sqref>
      </x14:sparkline>
      <x14:sparkline>
        <xm:f>\'Sheet 2\'!F2:F13</xm:f>
        <xm:sqref>F14</xm:sqref>
      </x14:sparkline>
      <x14:sparkline>
        <xm:f>\'Sheet 2\'!G2:G13</xm:f>
        <xm:sqref>G14</xm:sqref>
      </x14:sparkline>
      <x14:sparkline>
        <xm:f>\'Sheet 2\'!H2:H13</xm:f>
        <xm:sqref>H14</xm:sqref>
      </x14:sparkline>
      <x14:sparkline>
        <xm:f>\'Sheet 2\'!I2:I13</xm:f>
        <xm:sqref>I14</xm:sqref>
      </x14:sparkline>
      <x14:sparkline>
        <xm:f>\'Sheet 2\'!J2:J13</xm:f>
        <xm:sqref>J14</xm:sqref>
      </x14:sparkline>
      <x14:sparkline>
        <xm:f>\'Sheet 2\'!K2:K13</xm:f>
        <xm:sqref>K14</xm:sqref>
      </x14:sparkline>
      <x14:sparkline>
        <xm:f>\'Sheet 2\'!L2:L13</xm:f>
        <xm:sqref>L14</xm:sqref>
      </x14:sparkline>
    </x14:sparklines>
  </x14:sparklineGroup>
</x14:sparklineGroups>
</ext>', pointer = FALSE)
  got <- wb$worksheets[[2]]$extLst

  expect_equal(exp, got)
})

test_that("distinct() works", {

  x <- c("London", "NYC", "NYC", "Berlin", "Madrid", "London", "BERLIN", "berlin")

  exp <- c("London", "NYC", "Berlin", "Madrid")
  got <- distinct(x)
  expect_equal(exp, got)

})

test_that("fix_pt_names() works", {

  x <- c("Foo", "foo", "bar", "FOO", "bar", "x")

  exp <- c("Foo", "foo2", "bar", "FOO3", "bar2", "x")
  got <- fix_pt_names(x)
  expect_equal(exp, got)

})

test_that("validate_colors() works", {

  col <- c("FF0000FF", "#0000FF", "000FF", "#FF000FF", "blue")

  exp <- c("FF0000FF", "FF0000FF", "FFF000FF", "FFF000FF", "FF0000FF")
  got <- validate_color(col)
  expect_equal(exp, got)

  # switch from RGBA to ARGB format
  col <- c("#000000FF", adjustcolor("black"), "#000000", "black")
  exp <- c("FF000000", "FF000000", "FF000000", "FF000000")
  got <- validate_color(col, format = "RGBA")
  expect_equal(got, exp)

  # handle non-standard input (too short, missing hash)
  # short input is padded with F, this might not be the desired color
  col <- c("000000", "#000000", "#00", "00")
  exp <- c("FF000000", "FF000000", "FFFFFF00", "FFFFFF00")
  got <- validate_color(col)
  expect_equal(got, exp)
})

test_that("basename2() works", {

  long_path <- paste0(
    paste0(rep("foldername/", 40), collapse = ""),
    paste0(rep("filename", 40), collapse = ""),
    ".txt"
  )

  # # maybe only broken on old Windows. Errors in 4.1 not in 4.3.1
  # if (to_long(long_path))
  #   expect_error(basename(long_path))

  exp <- "filenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilenamefilename.txt"
  got <- basename2(long_path)
  expect_equal(exp, got)

})

test_that("is_double works", {

  x <- c("0.1", "a")
  exp <- c(TRUE, FALSE)
  got <- is_charnum(x)
  expect_equal(exp, got)

})

test_that("create_hyperlinks() works", {
  exp <- "=HYPERLINK(\"#'Sheet3'!B1\")"
  got <- create_hyperlink(sheet = "Sheet3", row = 1, col = 2)
  expect_equal(exp, got)

  fl <- "testdir/testfile.R"
  exp <- "=HYPERLINK(\"[testdir/testfile.R]Sheet2!J3\", \"Link to File.\")"
  got <- create_hyperlink(sheet = "Sheet2", row = 3, col = 10, file = fl, text = "Link to File.")
  expect_equal(exp, got)

})

test_that("waiver works in hyperlink", {
  wb <- wb_workbook()$
    add_worksheet("Sheet1")$add_worksheet("Sheet2")$add_worksheet("Sheet3")

  ## Internal - No text to display using create_hyperlink() function
  x <- create_hyperlink(sheet = "Sheet3", row = 1, col = 2)
  wb$add_formula(sheet = "Sheet1", x = x, dims = "A2")

  ## Internal - No text to display using create_hyperlink() function
  x <- create_hyperlink(sheet = current_sheet(), row = 1, col = 2)
  wb$add_formula(sheet = "Sheet3", x = x, dims = "A3")

  exp <- '=HYPERLINK(\"#\'Sheet3\'!B1\")'
  got <- unique(unname(unlist(wb$to_df(show_formula = TRUE, col_names = FALSE))))
  expect_equal(exp, got)
})

test_that("create_shape() works", {

  ## heart
  txt <- fmt_txt("openxlsx2 is the \n", bold = TRUE, size = 15) +
    fmt_txt("best", underline = TRUE, bold = TRUE, size = 15) +
    fmt_txt("\n!", bold = TRUE, size = 15)

  heart <- create_shape(
    shape = "heart", text = txt, text_align = "center",
    fill_colour = wb_color("pink"), text_colour = wb_color("theme" = 1))

  ## ribbon
  txt <- fmt_txt("\nthe ") +
    fmt_txt("very", underline = TRUE, font = "Courier", color = wb_color("gold")) +
    fmt_txt(" best")

  ribbon <- create_shape(shape = "ribbon", text = txt, text_align = "center")

  wb <- wb_workbook()$add_worksheet(grid_lines = FALSE)$
    add_drawing(dims = "B2:E11", xml = heart)$
    add_drawing(dims = "B12:E14", xml = ribbon)

  expect_equal(1L, length(wb$drawings))

  expect_silent(create_shape(text = "foo", rotation = 20, fill_transparency = 50))

})

test_that("get_dims works", {
  got <- get_dims(c("A1:A5", "B1:B5"), check = TRUE)
  expect_true(got)

  exp <- list(rows = list(c(1L, 5L)), cols = 1:2)
  got <- get_dims(c("A1:A5", "B1:B5"), check = FALSE)
  expect_equal(exp, got)


  got <- get_dims(c("A1:A5", "B2:B6"), check = TRUE)
  expect_false(got)

  exp <- list(rows = list(c(1L, 5L), c(2L, 6L)), cols = 1:2)
  got <- get_dims(c("A1:A5", "B2:B6"), check = FALSE)
  expect_equal(exp, got)
})
