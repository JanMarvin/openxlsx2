

test_that("Workbook properties", {
  # TODO check datetimeCreated field

  ## check creator
  wb <- wb_workbook(
    creator  = "Alex",
    title    = "title here",
    subject  = "this & that",
    category = "some category"
  )

  expect_match(wb$core, "<dc:creator>Alex</dc:creator>")
  expect_match(wb$core, "<dc:title>title here</dc:title>")
  expect_match(wb$core, "<dc:subject>this &amp; that</dc:subject>")
  expect_match(wb$core, "<cp:category>some category</cp:category>")

  temp_file <- temp_xlsx()
  wb <- write.xlsx(
    x        = iris,
    file     = temp_file,
    creator  = "Alex 2",
    title    = "title here 2",
    subject  = "this & that 2",
    category = "some category 2"
  )

  expect_match(wb$core, "<dc:creator>Alex 2</dc:creator>")
  expect_match(wb$core, "<dc:title>title here 2</dc:title>")
  expect_match(wb$core, "<dc:subject>this &amp; that 2</dc:subject>")
  expect_match(wb$core, "<cp:category>some category 2</cp:category>")

  ## maintain on load
  wb_loaded <- loadWorkbook(temp_file)
  expect_equal(wb_loaded$core, paste0(wb$core, collapse = ""))

  wb <- wb_workbook(
    creator  = "Philipp",
    title    = "title here",
    subject  = "this & that",
    category = "some category"
  )

  addCreator(wb, "test")
  expect_match(wb$core, "<dc:creator>Philipp;test</dc:creator>")

  expect_identical(getCreators(wb), c("Philipp", "test"))
  setLastModifiedBy(wb, "Philipp 2")
  expect_match(wb$core, "<cp:lastModifiedBy>Philipp 2</cp:lastModifiedBy>")

  file.remove(temp_file)
})
