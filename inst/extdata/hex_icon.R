# install.packages(c("hexSticker", "showtext"))
library(openxlsx2)
library(ggplot2)
library(hexSticker)
library(showtext)

## create the background in a spreadsheet matrix
wb <- wb_workbook()$add_worksheet(grid_lines = FALSE)$
  set_row_heights(rows = 1:5 + 1L, heights = 15)$
  set_col_widths(cols = 1:5 + 1L, width = 2.15)

mm <- matrix(0, 5, 5)
# Fill the secondary diagonals with unique values
for (k in 1:10) { # Loop over all diagonals relative to the main diagonal
  idx <- which(row(mm) + col(mm) - 1 == k, arr.ind = TRUE)
  mm[idx] <- k # Assign the value k or any sequence
}

tm <- mm - 4
tm[tm < 1] <- 0

# Palette from https://loading.io/color/feature/Greens-5/
# Loading.io Free License
#
# With Loading.io Free license ( LD-FREE / FREE / Free License ), items are
# dedicated to the public domain by waiving all our right worldwide under
# copyright law. You can use items under LD-FREE freely for any purpose. No
# attribution is required.
greens <- c("#edf8e9", "#bae4b3", "#74c476", "#31a354", "#006d2c")

for (i in 1:5) {
  coords <- as.data.frame(which(tm == i, arr.ind = TRUE))
  dims <- paste0(
    mapply(coords$row, coords$col, FUN = function(x, y) {
      wb_dims(rows = x, cols = y, from_dims = "B2")
    }), collapse = ",")
  wb$add_fill(dims = dims, color = wb_color(greens[i]))
}

# screenshot the area and save to file
if (interactive()) wb$open()

df <- as.data.frame(as.table(tm))
colnames(df) <- c("x", "y", "value")

# Define the green color palette, with white for 0
greens <- c("0" = "white", setNames(greens, 1:5)) # or "transparent"

# Create the ggplot
gg <- ggplot(df, aes(x = x, y = y, fill = as.character(value))) +
  geom_tile() + # Add black borders to tiles
  scale_y_discrete(limits = rev) + # Reverse the y-axis to align with bottom-left orientation
  scale_fill_manual(values = greens, guide = "none") + # Use custom color mapping
  coord_fixed() + # Keep tiles square
  theme_void() + # Minimalist theme
  theme(legend.position = "none") + # Remove legend
  labs(x = NULL, y = NULL) # Remove axis labels

## Loading Google fonts (http://www.google.com/fonts)
# All Noto fonts are licensed under the Open Font License.
# You can use them in all your products & projects — print or digital,
# commercial or otherwise.
font_add_google("Noto sans", "noto")

s <- sticker(gg, package = "openxlsx2",
             s_x = 1, s_y = 1, s_width = 1, s_height = 1,
             h_fill = greens[1], h_color = greens[length(greens)], h_size = 1,
             p_size = 25, p_color = greens[length(greens)], p_x = 1, p_y = 1.1,
             p_family = "noto", p_fontface = "bold",
             filename = "man/figures/logo.png")

s

# imgurl <- "inst/img/openxlsx2.png"
# s <- sticker(imgurl, package = "openxlsx2",
#              s_x = 1, s_y = 1, s_width = .6, spotlight = FALSE,
#              h_fill = "#FFF", h_color = greens[5], h_size = 1,
#              p_size = 25, p_color = greens[5], p_x = 1, p_y = 1.1,
#              p_family = "noto", p_fontface = "bold",
#              filename = "man/figures/logo.png")
# s
