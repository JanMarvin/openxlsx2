# loads character string to pugixml and returns an externalptr

loads character string to pugixml and returns an externalptr

## Usage

``` r
as_xml(x, ...)
```

## Arguments

- x:

  input as xml

- ...:

  additional arguments passed to
  [`read_xml()`](https://janmarvin.github.io/openxlsx2/reference/read_xml.md)

## Details

might be useful for larger documents where single nodes are shortened
and otherwise the full tree has to be reimported. unsure where we have
such a case. is useful, for printing nodes from a larger tree, that have
been exported as characters (at some point in time we have to convert
the xml to R)

## Examples

``` r
tmp_xlsx <- tempfile()
xlsxFile <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
unzip(xlsxFile, exdir = tmp_xlsx)

wb <- wb_load(xlsxFile)
styles_xml <- sprintf("%s/xl/styles.xml", tmp_xlsx)

# is external pointer
sxml <- read_xml(styles_xml)

# is character
font <- xml_node(sxml, "styleSheet", "fonts", "font")

# is again external pointer
as_xml(font)
#> <font>
#>  <sz val="12" />
#>  <color theme="1" />
#>  <name val="Calibri" />
#>  <family val="2" />
#>  <scheme val="minor" />
#> </font>
#> <font>
#>  <sz val="12" />
#>  <color theme="1" />
#>  <name val="Helvetica" />
#>  <family val="2" />
#> </font>
#> <font>
#>  <sz val="10" />
#>  <color rgb="FF000000" />
#>  <name val="Helvetica Neue" />
#>  <family val="2" />
#> </font>
```
