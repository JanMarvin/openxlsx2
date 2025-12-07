# Protect a workbook from modifications

Protect or unprotect a workbook from modifications by the user in the
graphical user interface. Replaces an existing protection.

## Usage

``` r
wb_protect(
  wb,
  protect = TRUE,
  password = NULL,
  lock_structure = FALSE,
  lock_windows = FALSE,
  type = 1,
  file_sharing = FALSE,
  username = unname(Sys.info()["user"]),
  read_only_recommended = FALSE,
  ...
)
```

## Arguments

- wb:

  A Workbook object

- protect:

  Whether to protect or unprotect the sheet (default `TRUE`)

- password:

  (optional) password required to unprotect the workbook

- lock_structure:

  Whether the workbook structure should be locked

- lock_windows:

  Whether the window position of the spreadsheet should be locked

- type:

  Lock type (see **Details**)

- file_sharing:

  Whether to enable a popup requesting the unlock password is prompted

- username:

  The username for the `file_sharing` popup

- read_only_recommended:

  Whether or not a post unlock message appears stating that the workbook
  is recommended to be opened in read-only mode.

- ...:

  additional arguments

## Details

Lock types:

- `1` xlsx with password (default)

- `2` xlsx recommends read-only

- `4` xlsx enforces read-only

- `8` xlsx is locked for annotation

## Examples

``` r
wb <- wb_workbook()
wb$add_worksheet("S1")
wb_protect(wb, protect = TRUE, password = "Password", lock_structure = TRUE)

# Remove the protection
wb_protect(wb, protect = FALSE)

wb <- wb_protect(
  wb,
  protect = TRUE,
  password = "Password",
  lock_structure = TRUE,
  type = 2L,
  file_sharing = TRUE,
  username = "Test",
  read_only_recommended = TRUE
)
```
