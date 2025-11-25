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
  wb <- write_xlsx(
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
  wb_loaded <- wb_load(temp_file)
  expect_equal(wb_loaded$core, paste0(wb$core, collapse = ""))

  wb <- wb_workbook(
    creator  = "Philipp",
    title    = "title here",
    subject  = "this & that",
    category = "some category"
  )

  wb <- wb_add_creators(wb, "test")
  expect_match(wb$core, "<dc:creator>Philipp;test</dc:creator>")

  res <- wb_get_creators(wb)
  exp <- c("Philipp", "test")
  expect_identical(res, exp)

  res <- wb_get_creators(wb_remove_creators(wb, "test"))
  expect_identical(res, "Philipp")


  wb <- wb_set_last_modified_by(wb, "Philipp 2")
  expect_match(wb$core, "<cp:lastModifiedBy>Philipp 2</cp:lastModifiedBy>")

  creators <- c("person", "place", "thing")
  wb <- wb_workbook(creator = creators)
  res <- wb_get_creators(wb)
  expect_identical(res, creators)

})

test_that("escaping in wbWorkbooks genBaseCore works as expected", {

  got <- "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dc:creator>crea &amp; tor</dc:creator><cp:lastModifiedBy>crea &amp; tor</cp:lastModifiedBy><dcterms:created xsi:type=\"dcterms:W3CDTF\">2023-08-31T23:13:43Z</dcterms:created><dc:title>ti &amp; tle</dc:title><dc:subject>sub &amp; ject</dc:subject><cp:category>cate &amp; gory</cp:category></cp:coreProperties>"

  wb <- wb_workbook(
    creator = "crea & tor",
    title = "ti & tle",
    subject = "sub & ject",
    category = "cate & gory"
  )

  nms <- xml_node_name(got, "cp:coreProperties")

  for (nm in nms[!nms %in% c("cp:lastModifiedBy", "dcterms:created")]) {
    expect_equal(
      xml_node(got, "cp:coreProperties", nm),
      xml_node(wb$core, "cp:coreProperties", nm)
    )
  }

})
