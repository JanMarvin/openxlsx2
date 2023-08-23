test_that("openxlsx2_types", {

  # test vector types
  expect_equal(openxlsx2_celltype[["short_date"]], openxlsx2_type(Sys.Date()))
  expect_equal(openxlsx2_celltype[["long_date"]], openxlsx2_type(as.POSIXct(Sys.Date())))
  expect_equal(openxlsx2_celltype[["numeric"]], openxlsx2_type(1))
  expect_equal(openxlsx2_celltype[["logical"]], openxlsx2_type(TRUE))
  expect_equal(openxlsx2_celltype[["character"]], openxlsx2_type("a"))
  expect_equal(openxlsx2_celltype[["factor"]], openxlsx2_type(as.factor(1)))

  # even complex numbers
  z <- complex(real = stats::rnorm(1), imaginary = stats::rnorm(1))
  expect_equal(openxlsx2_celltype[["character"]], openxlsx2_type(z))

  # write_datatable example: data frame with various types
  df <- data.frame(
    "Date" = Sys.Date() - 0:19,
    "T" = TRUE, "F" = FALSE,
    "Time" = Sys.time() - 0:19 * 60 * 60,
    "Cash" = paste("$", 1:20), "Cash2" = 31:50,
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
    T = openxlsx2_celltype[["logical"]],
    F = openxlsx2_celltype[["logical"]],
    Time = openxlsx2_celltype[["long_date"]],
    Cash = openxlsx2_celltype[["character"]],
    Cash2 = openxlsx2_celltype[["accounting"]],
    hLink = openxlsx2_celltype[["hyperlink"]],
    Percentage = openxlsx2_celltype[["percentage"]],
    TinyNumbers = openxlsx2_celltype[["scientific"]]
  )


  expect_equal(exp, got)

})


test_that("wb_page_setup example", {

  wb <- wb_workbook()
  wb$add_worksheet("S1")
  wb$add_worksheet("S2")
  wb$add_data_table(1, x = iris[1:30, ])

  ## landscape page scaled to 50%
  wb$page_setup(sheet = 1, orientation = "landscape", scale = 50)
  exp <- "<pageSetup paperSize=\"9\" orientation=\"landscape\" scale = \"50\" fitToWidth=\"0\" fitToHeight=\"0\" horizontalDpi=\"300\" verticalDpi=\"300\"/>"
  expect_equal(exp, wb$worksheets[[1]]$pageSetup)


  ## portrait page scales to 300% with 0.5in left and right margins
  wb$page_setup(sheet = 2, orientation = "portrait", scale = 300, left = 0.5, right = 0.5)
  exp <- "<pageSetup paperSize=\"9\" orientation=\"portrait\" scale = \"300\" fitToWidth=\"0\" fitToHeight=\"0\" horizontalDpi=\"300\" verticalDpi=\"300\"/>"
  expect_equal(exp, wb$worksheets[[2]]$pageSetup)


  ## print titles
  wb$add_worksheet("print_title_rows")
  wb$add_worksheet("print_title_cols")

  wb$add_data("print_title_rows", rbind(iris, iris, iris, iris))
  wb$add_data("print_title_cols", x = rbind(mtcars, mtcars, mtcars), rowNames = TRUE)

  wb$page_setup(sheet = "print_title_rows", printTitleRows = 1) ## first row
  wb$page_setup(sheet = "print_title_cols", printTitleCols = 1, printTitleRows = 1)

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

test_that("distinct() works", {

  x <- c("London", "NYC", "NYC", "Berlin", "Madrid", "London", "BERLIN", "berlin")

  exp <- c("London", "NYC", "Berlin", "Madrid")
  got <- distinct(x)
  expect_equal(exp, got)

})

test_that("validate_colors() works", {

  col <- c("FF0000FF", "#0000FF", "000FF", "#FF000FF", "blue")

  exp <- c("FF0000FF", "FF0000FF", "FFF000FF", "FFF000FF", "FF0000FF")
  got <- validate_color(col)
  expect_equal(exp, got)

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
