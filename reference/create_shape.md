# Helper to create a shape

Helper to create a shape

## Usage

``` r
create_shape(
  shape = "rect",
  name = "shape 1",
  text = "",
  fill_color = NULL,
  fill_transparency = 0,
  text_color = NULL,
  text_transparency = 0,
  line_color = fill_color,
  line_transparency = 0,
  text_align = "left",
  rotation = 0,
  id = 1,
  ...
)
```

## Arguments

- shape:

  a shape (see details)

- name:

  a name for the shape

- text:

  a text written into the object. This can be a simple character or a
  [`fmt_txt()`](https://janmarvin.github.io/openxlsx2/reference/fmt_txt.md)

- fill_color, text_color, line_color:

  a color for each, accepts only theme and rgb colors passed with
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/reference/wb_color.md)

- fill_transparency, text_transparency, line_transparency:

  sets the alpha value of the shape, an integer value in the range 0 to
  100

- text_align:

  sets the alignment of the text. Can be 'left', 'center', 'right',
  'justify', 'justifyLow', 'distributed', or 'thaiDistributed'

- rotation:

  the rotation of the shape in degrees

- id:

  an integer id (effect is unknown)

- ...:

  additional arguments

## Value

a character containing the XML

## Details

Possible shapes are (from ST_ShapeType - Preset Shape Types): "line",
"lineInv", "triangle", "rtTriangle", "rect", "diamond", "parallelogram",
"trapezoid", "nonIsoscelesTrapezoid", "pentagon", "hexagon", "heptagon",
"octagon", "decagon", "dodecagon", "star4", "star5", "star6", "star7",
"star8", "star10", "star12", "star16", "star24", "star32", "roundRect",
"round1Rect", "round2SameRect", "round2DiagRect", "snipRoundRect",
"snip1Rect", "snip2SameRect", "snip2DiagRect", "plaque", "ellipse",
"teardrop", "homePlate", "chevron", "pieWedge", "pie", "blockArc",
"donut", "noSmoking", "rightArrow", "leftArrow", "upArrow", "downArrow",
"stripedRightArrow", "notchedRightArrow", "bentUpArrow",
"leftRightArrow", "upDownArrow", "leftUpArrow", "leftRightUpArrow",
"quadArrow", "leftArrowCallout", "rightArrowCallout", "upArrowCallout",
"downArrowCallout", "leftRightArrowCallout", "upDownArrowCallout",
"quadArrowCallout", "bentArrow", "uturnArrow", "circularArrow",
"leftCircularArrow", "leftRightCircularArrow", "curvedRightArrow",
"curvedLeftArrow", "curvedUpArrow", "curvedDownArrow", "swooshArrow",
"cube", "can", "lightningBolt", "heart", "sun", "moon", "smileyFace",
"irregularSeal1", "irregularSeal2", "foldedCorner", "bevel", "frame",
"halfFrame", "corner", "diagStripe", "chord", "arc", "leftBracket",
"rightBracket", "leftBrace", "rightBrace", "bracketPair", "bracePair",
"straightConnector1", "bentConnector2", "bentConnector3",
"bentConnector4", "bentConnector5", "curvedConnector2",
"curvedConnector3", "curvedConnector4", "curvedConnector5", "callout1",
"callout2", "callout3", "accentCallout1", "accentCallout2",
"accentCallout3", "borderCallout1", "borderCallout2", "borderCallout3",
"accentBorderCallout1", "accentBorderCallout2", "accentBorderCallout3",
"wedgeRectCallout", "wedgeRoundRectCallout", "wedgeEllipseCallout",
"cloudCallout", "cloud", "ribbon", "ribbon2", "ellipseRibbon",
"ellipseRibbon2", "leftRightRibbon", "verticalScroll",
"horizontalScroll", "wave", "doubleWave", "plus", "flowChartProcess",
"flowChartDecision", "flowChartInputOutput",
"flowChartPredefinedProcess", "flowChartInternalStorage",
"flowChartDocument", "flowChartMultidocument", "flowChartTerminator",
"flowChartPreparation", "flowChartManualInput",
"flowChartManualOperation", "flowChartConnector",
"flowChartPunchedCard", "flowChartPunchedTape",
"flowChartSummingJunction", "flowChartOr", "flowChartCollate",
"flowChartSort", "flowChartExtract", "flowChartMerge",
"flowChartOfflineStorage", "flowChartOnlineStorage",
"flowChartMagneticTape", "flowChartMagneticDisk",
"flowChartMagneticDrum", "flowChartDisplay", "flowChartDelay",
"flowChartAlternateProcess", "flowChartOffpageConnector",
"actionButtonBlank", "actionButtonHome", "actionButtonHelp",
"actionButtonInformation", "actionButtonForwardNext",
"actionButtonBackPrevious", "actionButtonEnd", "actionButtonBeginning",
"actionButtonReturn", "actionButtonDocument", "actionButtonSound",
"actionButtonMovie", "gear6", "gear9", "funnel", "mathPlus",
"mathMinus", "mathMultiply", "mathDivide", "mathEqual", "mathNotEqual",
"cornerTabs", "squareTabs", "plaqueTabs", "chartX", "chartStar",
"chartPlus"

## See also

[`wb_add_drawing()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_drawing.md)

## Examples

``` r
 wb <- wb_workbook()$add_worksheet()$
   add_drawing(xml = create_shape())
```
