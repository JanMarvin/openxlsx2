
test_that("paste_c() works", {
  res <- paste_c("a", "", "b", NULL, "c", " ")
  exp <- "a b c  "
  expect_identical(res, exp)

  expect_identical(paste_c(character(), "", NULL), "")
})


test_that("rbindlist", {

  if (.Platform$r_arch == "i386" && .Platform$OS.type == "windows")
    skip("Skip test on Windows i386")

  expect_equal(data.frame(), rbindlist(character()))

})
