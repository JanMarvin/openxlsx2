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

  # writeDataTable example: data frame with various types
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
    T = openxlsx2:::openxlsx2_celltype[["logical"]],
    F = openxlsx2:::openxlsx2_celltype[["logical"]],
    Time = openxlsx2:::openxlsx2_celltype[["long_date"]],
    Cash = openxlsx2:::openxlsx2_celltype[["character"]],
    Cash2 = openxlsx2:::openxlsx2_celltype[["accounting"]],
    hLink = openxlsx2:::openxlsx2_celltype[["hyperlink"]],
    Percentage = openxlsx2:::openxlsx2_celltype[["percentage"]],
    TinyNumbers = openxlsx2:::openxlsx2_celltype[["scientific"]]
  )


  expect_equal(exp, got)

})
