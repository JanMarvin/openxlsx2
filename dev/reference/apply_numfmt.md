# Format Values using OOXML (Spreadsheet) Number Format Codes

This function emulates a spreadsheet formatting engine. It takes
numeric, date, or character values and applies an OOXML format code to
produce a formatted string. It supports standard number formatting,
date/time, elapsed durations, fractions, and conditional formatting
sections.

## Usage

``` r
apply_numfmt(value, format_code)
```

## Arguments

- value:

  A vector of values to format. Supports `numeric`, `Date`, `POSIXct`,
  `character` (ISO dates or times), `hms`, or `difftime`.

- format_code:

  A character vector of format strings (e.g., `"#,##0.00"`,
  `"yyyy-mm-dd"`, or `"[h]:mm:ss"`).

## Value

A character vector of formatted strings.

## Details

The function splits the `format_code` into up to four sections separated
by semicolons (Positive; Negative; Zero; Text).

- **Date and Time:** Supports standard tokens (`yyyy`, `mm`, `dd`, `hh`,
  `ss`) and AM/PM toggles.

- **Durations:** Supports elapsed time tokens in square brackets, such
  as `[h]`, `[m]`, and `[s]`, calculating total units.

- **HMS Handling:** Character strings in "HH:MM:SS" format are
  automatically coerced to a date-time object using a base date to allow
  clock and duration formatting.

- **Numbers:** Supports thousands separators (`,`), precision,
  percentage conversion, and scientific notation (`E+`).

- **Fractions:** Uses the Farey algorithm to approximate decimals as
  fractions (e.g., `?/?`, `# ?/?`).

- **Conditionals:** Evaluates bracketed conditions like `[<1000]` to
  select the appropriate formatting section.

- **Literals:** Handles escaped characters (`\\`), quoted text
  (`"text"`), and the text placeholder (`@`).

## Examples

``` r
# Numeric formatting
apply_numfmt(1234.5678, "#,##0.00") # "1,234.57"
#> [1] "1,234.57"

# Date and Time
apply_numfmt("2025-01-05", "dddd, mmm dd") # "Sunday, Jan 05"
#> [1] "Sunday, Jan 05"
apply_numfmt("13:45:30", "hh:mm AM/PM")    # "01:45 PM"
#> [1] "01:45 PM"

# Durations
apply_numfmt("1900-01-12 08:17:47", "[h]:mm:ss") # "296:17:47"
#> [1] "296:17:47"

# Fractions
apply_numfmt(1.75, "# ?/?") # "1 3/4"
#> [1] "1 3/4"
```
