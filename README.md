openxlsx2
========

![R-CMD-check](https://github.com/JanMarvin/openxlsx2/workflows/R-CMD-check/badge.svg?branch=main) [![codecov](https://codecov.io/gh/JanMarvin/openxlsx2/branch/main/graph/badge.svg?token=HEZ7rXcZNq)](https://app.codecov.io/gh/JanMarvin/openxlsx2)

This R package is a modern reinterpretation of the widely used popular `openxlsx` package. Similar to its predecessor, it simplifies the creation of xlsx files by providing a clean interface for writing, designing and editing worksheets. Based on a powerful XML library and focusing on modern programming flows in pipes or chains, `openxlsx2` allows to break many new ground.

Even though the project is already well progressed and supports most of the features known and appreciated from the predecessor, there may still be open gaps in one or the other place. A quick warning: Until the stable version 1.0 there may still be some changes to the API.

### Introduction
```R
vignette(package = "openxlsx2")
```


### Development version
```R
remotes::install_github("JanMarvin/openxlsx2")
```

### Authors
For a full list of all authors that have made this package possible, please see
```R
system.file("AUTHORS", package = "openxlsx2")
```
If you feel like you should be included on this list, please let us know.

### License
This package is licensed under the MIT license and is based on `openxlsx` (by Alexander Walker and Philipp Schauberger; COPYRIGHT 2014-2021) and `pugixml` (by Arseny Kapoulkine; COPYRIGHT 2006-2021). Both released under the MIT license.
