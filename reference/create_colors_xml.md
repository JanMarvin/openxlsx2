# Create custom color xml schemes

Create custom color themes that can be used with
[`wb_set_base_colors()`](https://janmarvin.github.io/openxlsx2/reference/wb_base_colors.md).
The color input will be checked with
[`wb_color()`](https://janmarvin.github.io/openxlsx2/reference/wb_color.md),
so it must be either a color R from
[`grDevices::colors()`](https://rdrr.io/r/grDevices/colors.html) or a
hex value. Default values for the dark argument are: `black`, `white`,
`darkblue` and `lightgray`. For the accent argument, the six inner
values of
[`grDevices::palette()`](https://rdrr.io/r/grDevices/palette.html). The
link argument uses `blue` and `purple` by default for active and visited
links.

## Usage

``` r
create_colors_xml(name = "Base R", dark = NULL, accent = NULL, link = NULL)
```

## Arguments

- name:

  the color name

- dark:

  four colors: dark, light, brighter dark, darker light

- accent:

  six accent colors

- link:

  two link colors: link and visited link

## See also

Other style creating functions:
[`create_border()`](https://janmarvin.github.io/openxlsx2/reference/create_border.md),
[`create_cell_style()`](https://janmarvin.github.io/openxlsx2/reference/create_cell_style.md),
[`create_dxfs_style()`](https://janmarvin.github.io/openxlsx2/reference/create_dxfs_style.md),
[`create_fill()`](https://janmarvin.github.io/openxlsx2/reference/create_fill.md),
[`create_font()`](https://janmarvin.github.io/openxlsx2/reference/create_font.md),
[`create_numfmt()`](https://janmarvin.github.io/openxlsx2/reference/create_numfmt.md),
[`create_tablestyle()`](https://janmarvin.github.io/openxlsx2/reference/create_tablestyle.md)

## Examples

``` r
colors <- create_colors_xml()
wb <- wb_workbook()$add_worksheet()$set_base_colors(xml = colors)
```
