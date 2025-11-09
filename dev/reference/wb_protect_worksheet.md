# Protect a worksheet from modifications

Protect or unprotect a worksheet from modifications by the user in the
graphical user interface. Replaces an existing protection. Certain
features require applying unlocking of initialized cells in the
worksheet and across columns and/or rows.

## Usage

``` r
wb_protect_worksheet(
  wb,
  sheet = current_sheet(),
  protect = TRUE,
  password = NULL,
  properties = NULL
)
```

## Arguments

- wb:

  A workbook object

- sheet:

  A name or index of a worksheet

- protect:

  Whether to protect or unprotect the sheet (default=TRUE)

- password:

  (optional) password required to unprotect the worksheet

- properties:

  A character vector of properties to lock. Can be one or more of the
  following: `"selectLockedCells"`, `"selectUnlockedCells"`,
  `"formatCells"`, `"formatColumns"`, `"formatRows"`, `"insertColumns"`,
  `"insertRows"`, `"insertHyperlinks"`, `"deleteColumns"`,
  `"deleteRows"`, `"sort"`, `"autoFilter"`, `"pivotTables"`,
  `"objects"`, `"scenarios"`

## Examples

``` r
wb <- wb_workbook()
wb$add_worksheet("S1")
wb$add_data_table(1, x = iris[1:30, ])

wb$protect_worksheet(
  "S1",
  protect = TRUE,
  properties = c("formatCells", "formatColumns", "insertColumns", "deleteColumns")
)

# Formatting cells / columns is allowed , but inserting / deleting columns is protected:
wb$protect_worksheet(
  "S1",
  protect = TRUE,
   c(formatCells = FALSE, formatColumns = FALSE,
                 insertColumns = TRUE, deleteColumns = TRUE)
)

# Remove the protection
wb$protect_worksheet("S1", protect = FALSE)
```
