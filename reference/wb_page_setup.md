# Set page margins, orientation and print scaling of a worksheet

Set page margins, orientation and print scaling.

## Usage

``` r
wb_set_page_setup(
  wb,
  sheet = current_sheet(),
  black_and_white = NULL,
  cell_comments = NULL,
  copies = NULL,
  draft = NULL,
  errors = NULL,
  first_page_number = NULL,
  id = NULL,
  page_order = NULL,
  paper_height = NULL,
  paper_width = NULL,
  hdpi = NULL,
  vdpi = NULL,
  use_first_page_number = NULL,
  use_printer_defaults = NULL,
  orientation = NULL,
  scale = NULL,
  left = 0.7,
  right = 0.7,
  top = 0.75,
  bottom = 0.75,
  header = 0.3,
  footer = 0.3,
  fit_to_width = FALSE,
  fit_to_height = FALSE,
  paper_size = NULL,
  print_title_rows = NULL,
  print_title_cols = NULL,
  summary_row = NULL,
  summary_col = NULL,
  tab_color = NULL,
  horizontal_centered = NULL,
  vertical_centered = NULL,
  print_headings = NULL,
  ...
)

wb_page_setup(
  wb,
  sheet = current_sheet(),
  orientation = NULL,
  scale = 100,
  left = 0.7,
  right = 0.7,
  top = 0.75,
  bottom = 0.75,
  header = 0.3,
  footer = 0.3,
  fit_to_width = FALSE,
  fit_to_height = FALSE,
  paper_size = NULL,
  print_title_rows = NULL,
  print_title_cols = NULL,
  summary_row = NULL,
  summary_col = NULL,
  ...
)
```

## Arguments

- wb:

  A workbook object

- sheet:

  A name or index of a worksheet

- black_and_white:

  black and white mode

- cell_comments:

  show cell comments

- copies:

  Amount of copies

- draft:

  Draft mode

- errors:

  Show errors

- first_page_number:

  The first page number

- id:

  id (unknown)

- page_order:

  Page order

- paper_height, paper_width:

  paper size

- hdpi, vdpi:

  horizontal and vertical dpi

- use_first_page_number:

  Number on first page

- use_printer_defaults:

  Use printer defaults

- orientation:

  Page orientation. One of "portrait" or "landscape"

- scale:

  Print scaling. Numeric value between 10 and 400

- left, right, top, bottom:

  Page margin in inches

- header, footer:

  Margin in inches

- fit_to_width, fit_to_height:

  An integer that tells the spreadsheet software on how many pages the
  scaling should fit. This does not actually scale the sheet.

- paper_size:

  See details. Default value is 9 (A4 paper).

- print_title_rows, print_title_cols:

  Rows / columns to repeat at top of page when printing. Integer vector.

- summary_row:

  Location of summary rows in groupings. One of "Above" or "Below".

- summary_col:

  Location of summary columns in groupings. One of "Right" or "Left".

- tab_color:

  The tab color

- horizontal_centered, vertical_centered:

  center print output vertical or horizontal

- print_headings:

  print headings

- ...:

  additional arguments

## Details

When adding fitting to width and height manual adjustment of the scaling
factor is required. Setting `fit_to_width` and `fit_to_height` only
tells spreadsheet software that the scaling was applied, but not which
scaling was applied.

`wb_page_setup()` provides a subset of `wb_set_page_setup()`. The former
will soon become deprecated.

`paper_size` is an integer corresponding to:

|      |                                                                                      |
|------|--------------------------------------------------------------------------------------|
| size | "paper type"                                                                         |
| 1    | Letter paper (8.5 in. by 11 in.)                                                     |
| 2    | Letter small paper (8.5 in. by 11 in.)                                               |
| 3    | Tabloid paper (11 in. by 17 in.)                                                     |
| 4    | Ledger paper (17 in. by 11 in.)                                                      |
| 5    | Legal paper (8.5 in. by 14 in.)                                                      |
| 6    | Statement paper (5.5 in. by 8.5 in.)                                                 |
| 7    | Executive paper (7.25 in. by 10.5 in.)                                               |
| 8    | A3 paper (297 mm by 420 mm)                                                          |
| 9    | A4 paper (210 mm by 297 mm)                                                          |
| 10   | A4 small paper (210 mm by 297 mm)                                                    |
| 11   | A5 paper (148 mm by 210 mm)                                                          |
| 12   | B4 paper (250 mm by 353 mm)                                                          |
| 13   | B5 paper (176 mm by 250 mm)                                                          |
| 14   | Folio paper (8.5 in. by 13 in.)                                                      |
| 15   | Quarto paper (215 mm by 275 mm)                                                      |
| 16   | Standard paper (10 in. by 14 in.)                                                    |
| 17   | Standard paper (11 in. by 17 in.)                                                    |
| 18   | Note paper (8.5 in. by 11 in.)                                                       |
| 19   | \#9 envelope (3.875 in. by 8.875 in.)                                                |
| 20   | \#10 envelope (4.125 in. by 9.5 in.)                                                 |
| 21   | \#11 envelope (4.5 in. by 10.375 in.)                                                |
| 22   | \#12 envelope (4.75 in. by 11 in.)                                                   |
| 23   | \#14 envelope (5 in. by 11.5 in.)                                                    |
| 24   | C paper (17 in. by 22 in.)                                                           |
| 25   | D paper (22 in. by 34 in.)                                                           |
| 26   | E paper (34 in. by 44 in.)                                                           |
| 27   | DL envelope (110 mm by 220 mm)                                                       |
| 28   | C5 envelope (162 mm by 229 mm)                                                       |
| 29   | C3 envelope (324 mm by 458 mm)                                                       |
| 30   | C4 envelope (229 mm by 324 mm)                                                       |
| 31   | C6 envelope (114 mm by 162 mm)                                                       |
| 32   | C65 envelope (114 mm by 229 mm)                                                      |
| 33   | B4 envelope (250 mm by 353 mm)                                                       |
| 34   | B5 envelope (176 mm by 250 mm)                                                       |
| 35   | B6 envelope (176 mm by 125 mm)                                                       |
| 36   | Italy envelope (110 mm by 230 mm)                                                    |
| 37   | Monarch envelope (3.875 in. by 7.5 in.)                                              |
| 38   | 6 3/4 envelope (3.625 in. by 6.5 in.)                                                |
| 39   | US standard fanfold (14.875 in. by 11 in.)                                           |
| 40   | German standard fanfold (8.5 in. by 12 in.)                                          |
| 41   | German legal fanfold (8.5 in. by 13 in.)                                             |
| 42   | ISO B4 (250 mm by 353 mm)                                                            |
| 43   | Japanese double postcard (200 mm by 148 mm)                                          |
| 44   | Standard paper (9 in. by 11 in.)                                                     |
| 45   | Standard paper (10 in. by 11 in.)                                                    |
| 46   | Standard paper (15 in. by 11 in.)                                                    |
| 47   | Invite envelope (220 mm by 220 mm)                                                   |
| 50   | Letter extra paper (9.275 in. by 12 in.)                                             |
| 51   | Legal extra paper (9.275 in. by 15 in.)                                              |
| 52   | Tabloid extra paper (11.69 in. by 18 in.)                                            |
| 53   | A4 extra paper (236 mm by 322 mm)                                                    |
| 54   | Letter transverse paper (8.275 in. by 11 in.)                                        |
| 55   | A4 transverse paper (210 mm by 297 mm)                                               |
| 56   | Letter extra transverse paper (9.275 in. by 12 in.)                                  |
| 57   | SuperA/SuperA/A4 paper (227 mm by 356 mm)                                            |
| 58   | SuperB/SuperB/A3 paper (305 mm by 487 mm)                                            |
| 59   | Letter plus paper (8.5 in. by 12.69 in.)                                             |
| 60   | A4 plus paper (210 mm by 330 mm)                                                     |
| 61   | A5 transverse paper (148 mm by 210 mm)                                               |
| 62   | JIS B5 transverse paper (182 mm by 257 mm)                                           |
| 63   | A3 extra paper (322 mm by 445 mm)                                                    |
| 64   | A5 extra paper (174 mm by 235 mm)                                                    |
| 65   | ISO B5 extra paper (201 mm by 276 mm)                                                |
| 66   | A2 paper (420 mm by 594 mm)                                                          |
| 67   | A3 transverse paper (297 mm by 420 mm)                                               |
| 68   | A3 extra transverse paper (322 mm by 445 mm)                                         |
| 69   | Japanese Double Postcard (200 mm x 148 mm) 70=A6(105mm x 148mm)                      |
| 71   | Japanese Envelope Kaku \#2                                                           |
| 72   | Japanese Envelope Kaku \#3                                                           |
| 73   | Japanese Envelope Chou \#3                                                           |
| 74   | Japanese Envelope Chou \#4                                                           |
| 75   | Letter Rotated (11in x 8 1/2 11 in)                                                  |
| 76   | A3 Rotated (420 mm x 297 mm)                                                         |
| 77   | A4 Rotated (297 mm x 210 mm)                                                         |
| 78   | A5 Rotated (210 mm x 148 mm)                                                         |
| 79   | B4 (JIS) Rotated (364 mm x 257 mm)                                                   |
| 80   | B5 (JIS) Rotated (257 mm x 182 mm)                                                   |
| 81   | Japanese Postcard Rotated (148 mm x 100 mm)                                          |
| 82   | Double Japanese Postcard Rotated (148 mm x 200 mm) 83 = A6 Rotated (148 mm x 105 mm) |
| 84   | Japanese Envelope Kaku \#2 Rotated                                                   |
| 85   | Japanese Envelope Kaku \#3 Rotated                                                   |
| 86   | Japanese Envelope Chou \#3 Rotated                                                   |
| 87   | Japanese Envelope Chou \#4 Rotated 88=B6(JIS)(128mm x 182mm)                         |
| 89   | B6 (JIS) Rotated (182 mm x 128 mm)                                                   |
| 90   | (12 in x 11 in)                                                                      |
| 91   | Japanese Envelope You \#4                                                            |
| 92   | Japanese Envelope You \#4 Rotated 93=PRC16K(146mm x 215mm) 94=PRC32K(97mm x 151mm)   |
| 95   | PRC 32K(Big) (97 mm x 151 mm)                                                        |
| 96   | PRC Envelope \#1 (102 mm x 165 mm)                                                   |
| 97   | PRC Envelope \#2 (102 mm x 176 mm)                                                   |
| 98   | PRC Envelope \#3 (125 mm x 176 mm)                                                   |
| 99   | PRC Envelope \#4 (110 mm x 208 mm)                                                   |
| 100  | PRC Envelope \#5 (110 mm x 220 mm)                                                   |
| 101  | PRC Envelope \#6 (120 mm x 230 mm)                                                   |
| 102  | PRC Envelope \#7 (160 mm x 230 mm)                                                   |
| 103  | PRC Envelope \#8 (120 mm x 309 mm)                                                   |
| 104  | PRC Envelope \#9 (229 mm x 324 mm)                                                   |
| 105  | PRC Envelope \#10 (324 mm x 458 mm)                                                  |
| 106  | PRC 16K Rotated                                                                      |
| 107  | PRC 32K Rotated                                                                      |
| 108  | PRC 32K(Big) Rotated                                                                 |
| 109  | PRC Envelope \#1 Rotated (165 mm x 102 mm)                                           |
| 110  | PRC Envelope \#2 Rotated (176 mm x 102 mm)                                           |
| 111  | PRC Envelope \#3 Rotated (176 mm x 125 mm)                                           |
| 112  | PRC Envelope \#4 Rotated (208 mm x 110 mm)                                           |
| 113  | PRC Envelope \#5 Rotated (220 mm x 110 mm)                                           |
| 114  | PRC Envelope \#6 Rotated (230 mm x 120 mm)                                           |
| 115  | PRC Envelope \#7 Rotated (230 mm x 160 mm)                                           |
| 116  | PRC Envelope \#8 Rotated (309 mm x 120 mm)                                           |
| 117  | PRC Envelope \#9 Rotated (324 mm x 229 mm)                                           |
| 118  | PRC Envelope \#10 Rotated (458 mm x 324 mm)                                          |

## Examples

``` r
wb <- wb_workbook()
wb$add_worksheet("S1")
wb$add_worksheet("S2")
wb$add_data_table(1, x = iris[1:30, ])
wb$add_data_table(2, x = iris[1:30, ], dims = c("C5"))

## landscape page scaled to 50%
wb$set_page_setup(sheet = 1, orientation = "landscape", scale = 50)

## portrait page scales to 300% with 0.5in left and right margins
wb$set_page_setup(sheet = 2, orientation = "portrait", scale = 300, left = 0.5, right = 0.5)


## print titles
wb$add_worksheet("print_title_rows")
wb$add_worksheet("print_title_cols")

wb$add_data("print_title_rows", rbind(iris, iris, iris, iris))
wb$add_data("print_title_cols", x = rbind(mtcars, mtcars, mtcars), row_names = TRUE)

wb$set_page_setup(sheet = "print_title_rows", print_title_rows = 1) ## first row
wb$set_page_setup(sheet = "print_title_cols", print_title_cols = 1, print_title_rows = 1)
```
