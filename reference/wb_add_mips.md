# wb get and apply MIP section

Read sensitivity labels from files and apply them to workbooks

## Usage

``` r
wb_add_mips(wb, xml = NULL)

wb_get_mips(wb, single_xml = TRUE, quiet = TRUE)
```

## Arguments

- wb:

  a workbook

- xml:

  a mips string obtained from `wb_get_mips()` or a global option
  "openxlsx2.mips_xml_string"

- single_xml:

  option to define if the string should be exported as single string.
  helpful if storing as option is desired.

- quiet:

  option to print a MIP section name. This is not always a human
  readable string.

## Value

the workbook invisible (`wb_add_mips()`) or the xml string
(`wb_get_mips()`)

## Details

The MIP section is a special user-defined XML section that is used to
create sensitivity labels in workbooks. It consists of a series of XML
property nodes that define the sensitivity label. This XML string cannot
be created and it is necessary to first load a workbook with a suitable
sensitivity label. Once the workbook is loaded, the string
`fmips <- wb_get_mips(wb)` can be extracted. This xml string can later
be assigned to an `options("openxlsx2.mips_xml_string" = fmips)` option.

The sensitivity label can then be assigned with `wb_add_mips(wb)`. If no
xml string is passed, the MIP section is taken from the option. This
should make it easier for users to read the section from a specific
workbook, save it to a file or string and copy it to an option via the
.Rprofile.
