openxlsx2
========

![R-CMD-check](https://github.com/JanMarvin/openxlsx2/workflows/R-CMD-check/badge.svg?branch=main) [![codecov](https://codecov.io/gh/JanMarvin/openxlsx2/branch/main/graph/badge.svg?token=HEZ7rXcZNq)](https://app.codecov.io/gh/JanMarvin/openxlsx2)

This R package is a modern reinterpretation of the widely used popular `openxlsx` package. Similar to its predecessor, it simplifies the creation of xlsx files by providing a clean interface for writing, designing and editing worksheets. Based on a powerful XML library and focusing on modern programming flows in pipes or chains, `openxlsx2` allows to break many new ground.

Even though the project is already well progressed and supports most of the features known and appreciated from the predecessor, there may still be open gaps in one or the other place. A quick warning: Until the stable version 1.0 there may still be some changes to the API.

### Introduction
Many examples are in our manual pages and in our vignettes. You can find them under:

```R
vignette(package = "openxlsx2")
```

For a quick introduction to the package, you can try the following:

```R
# read xlsx or xlsm files
df <- read_xlsx("file.xlsx")
# write xlsx files
write.xlsx(df, "new.xlsx")

# or import workbooks
wb <- wb_load("file.xlsx")
# read a data frame
wb_to_df(wb)
# and save it
wb_save(wb, "new_wb.xlsx")

## or create one yourself
wb <- wb_workbook()
# add a worksheet
wb$addWorksheet("sheet")
# add some data
writeData(df, "sheet")
# open it in your default spreadsheet software
wb_open(wb)
```


### Development version
The development version with all the latest code can be installed from inside R (it requires a compiler):

```R
remotes::install_github("JanMarvin/openxlsx2")
```

### Authors and contributions
For a full list of all authors that have made this package possible and for whom we are greatful, please see
```R
system.file("AUTHORS", package = "openxlsx2")
```
If you feel like you should be included on this list, please let us know. If you have something to contribute, you are welcome. If something is not working as expected, open issues or if you have solved an issue, open a pull request. Please be respectful and be aware that we are volunteers doing this for fun in our unpaid free time. We will work on problems when we have time or need.


### License
This package is licensed under the MIT license and is based on `openxlsx` (by Alexander Walker and Philipp Schauberger; COPYRIGHT 2014-2022) and `pugixml` (by Arseny Kapoulkine; COPYRIGHT 2006-2022). Both released under the MIT license.
