
.onAttach <- function(libname, pkgname) {
  # add options that have no already been assigned
  options(op.openxlsx[names(op.openxlsx) %out% names(options())])
}

.onUnload <- function(libpath) {
  library.dynam.unload("openxlsx2", libpath)
}


