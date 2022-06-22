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


test_that("ws_page_setup example", {

  wb <- wb_workbook()
  wb$add_worksheet("S1")
  wb$add_worksheet("S2")
  wb$add_data_table(1, x = iris[1:30, ])
  expect_warning(wb$add_data_table(2, x = iris[1:30, ], xy = c("C", 5)))

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
    "<definedName name=\"_xlnm.Print_Titles\" localSheetId=\"2\">'print_title_rows'!$1:$1</definedName>",
    "<definedName name=\"_xlnm.Print_Titles\" localSheetId=\"3\">'print_title_cols'!$A:$A,'print_title_cols'!$1:$1</definedName>"
  )
  expect_equal(exp, wb$workbook$definedNames)

  tmp <- temp_xlsx()
  expect_silent(wb_save(wb, tmp, overwrite = TRUE))

  # survives write and load
  wb <- wb_load(tmp)
  expect_equal(exp, wb$workbook$definedNames)


})
