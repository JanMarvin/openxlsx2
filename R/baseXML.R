# functions to create basic xml lists

genBaseContent_Type <- function() {
  c(
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
  )
}

genBaseShapeVML <- function(clientData, id, fillcolor, rID) {
  if (grepl("visible", clientData, ignore.case = TRUE)) {
    visible <- "visible"
  } else {
    visible <- "hidden"
  }

  if (is.null(rID))
    fill <- '<v:fill color2="#ffffe1"/>'
  else
    fill <- sprintf('<v:fill type="frame" on="t" color2="#FFFFFF" focussize="0,0" recolor="t" o:relid="%s"/>', rID)

  paste0(
    sprintf('<v:shape id="_x0000_s%s" type="#_x0000_t202" style=\'position:absolute;', id),
    sprintf('margin-left:107.25pt;margin-top:172.5pt;width:147pt;height:96pt;z-index:1;
          visibility:%s;mso-wrap-style:tight\' fillcolor="%s" o:insetmode="auto">', visible, fillcolor),
    fill,
    '<v:shadow color="black" obscured="t"/>
    <v:path o:connecttype="none"/>
    <v:textbox style=\'mso-direction-alt:auto\'>
    <div style=\'text-align:left\'/>
    </v:textbox>', clientData, "</v:shape>"
  )
}

genClientData <- function(col, row, visible, height, width) {
  txt <- sprintf(
    '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>%s, 15, %s, 10, %s, 147, %s, 18</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>%s</x:Row><x:Column>%s</x:Column>',
    col, row - 2L, col + width - 1L, row + height - 1L, row - 1L, col - 1L
  )

  if (visible) {
    txt <- paste0(txt, "<x:Visible/>")
  }

  txt <- paste0(txt, "</x:ClientData>")

  return(txt)
}

# # TODO this should be merged with the one above for type Note
#' Generates Client data xml string for use in wb_add_form_control
#' @param left left
#' @param top top
#' @param right right
#' @param bottom bottom
#' @param link link (links this cell to the state to the form control)
#' @param range range (input cell range)
#' @param type type (Checkbox, Radio, Drop)
#' @param checked checked (bool)
#' @noRd
genClientDataFC <- function(left, top, right, bottom, link, range, type, checked) {

  if (is.null(link)) {
    link <- ""
  } else {
    link <- sprintf("<x:FmlaLink>%s</x:FmlaLink>", link)
  }

  if (is.null(range) || type != "Drop") {
    range <- ""
  } else {
    range <- sprintf("<x:FmlaRange>%s</x:FmlaRange>", range)
  }

  if (type == "Drop") {
    drop <- '<x:LCT>Normal</x:LCT>
    <x:DropStyle>Combo</x:DropStyle>
    <x:DropLines>8</x:DropLines>'
  } else {
    drop <- ""
  }

  # Anchor:
  # - LeftColumn: Index column starts at 0
  # - LeftOffset: Offset in pxl
  # - TopRow: Index column starts at 0
  # - TopOffset
  # - RightColumn
  # - RightOffset
  # - BottomRow
  # - BottomOffset

  txt <- sprintf(
    '<x:ClientData ObjectType="%s">
    <x:MoveWithCells/><x:SizeWithCells/>
    <x:Anchor>%s, 0, %s, 0, %s, 0, %s, 0</x:Anchor>
    <x:AutoFill>False</x:AutoFill>
    <x:AutoLine>False</x:AutoLine>
    <x:TextVAlign>Center</x:TextVAlign>
    %s %s
    <x:Checked>%s</x:Checked>
    <x:NoThreeD />
    %s
    </x:ClientData>',
    type,
    max(0, left),
    max(0, top),
    max(1, right),
    max(1, bottom),
    link,
    range,
    as_binary(checked),
    drop
  )

  return(txt)
}

# This is just a list to remember possible combinations
genBaseApp <- function() {
  list(
    Application = "<Application>Microsoft Excel</Application>",
    AppVersion = NULL,
    Characters = NULL,
    CharactersWithSpaces = NULL,
    Company = NULL,
    DigSig = NULL,
    DocSecurity = NULL,
    HeadingPairs = NULL,
    HiddenSlides = NULL,
    HLinks = NULL,
    HyperlinkBase = NULL,
    HyperlinksChanged = NULL,
    Lines = NULL,
    LinksUpToDate = NULL,
    Manager = NULL,
    MMClips = NULL,
    Notes = NULL,
    Pages = NULL,
    Paragraphs = NULL,
    PresentationFormat = NULL,
    Properties = NULL,
    ScaleCrop = NULL,
    SharedDoc = NULL,
    Slides = NULL,
    Template = NULL,
    TitlesOfParts = NULL,
    TotalTime = NULL,
    VectorVariantType = NULL,
    Words = NULL
  )
}

# All relationships here must be matched with files inside of the xlsx files.
# Otherwise certain third party tools might not accept the file as valid xlsx file.
genBaseWorkbook.xml.rels <- function() {
  c(
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
  )
}

genBaseWorkbook <- function() {

  # bookViews
  # calcPr
  # customWorkbookViews
  # definedNames
  # externalReferences
  # extLst
  # fileRecoveryPr
  # fileSharing
  # fileVersion
  # functionGroups
  # oleSize
  # pivotCaches
  # sheets
  # smartTagPr
  # smartTagTypes
  # webPublishing
  # webPublishObjects
  # workbookPr
  # workbookProtection

  list(
    fileVersion         = NULL,
    fileSharing         = NULL,
    workbookPr          = '<workbookPr date1904="false"/>',
    alternateContent    = NULL,
    revisionPtr         = NULL,
    absPath             = NULL, # "x15ac:absPath"
    workbookProtection  = NULL,
    bookViews           = NULL,
    sheets              = NULL,
    functionGroups      = NULL,
    externalReferences  = NULL,
    definedNames        = NULL,
    calcPr              = NULL,
    oleSize             = NULL,
    customWorkbookViews = NULL,
    pivotCaches         = NULL,
    smartTagPr          = NULL,
    smartTagTypes       = NULL,
    webPublishing       = NULL,
    fileRecoveryPr      = NULL,
    webPublishObjects   = NULL,
    extLst              = NULL
  )
}

# dummy should be removed soon
genBaseSheetRels <- function(sheetInd) {
  character()
}

genBaseStyleSheet <- function(dxfs = NULL, tableStyles = NULL, extLst = NULL) {
  list(
    numFmts = NULL,

    fonts = c('<font><sz val="11"/><color theme="1"/><name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font>'),

    fills = c(
      '<fill><patternFill patternType="none"/></fill>',
      '<fill><patternFill patternType="gray125"/></fill>'
    ),

    borders = c("<border><left/><right/><top/><bottom/><diagonal/></border>"),

    cellStyleXfs = c('<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'),

    cellXfs = NULL, # c('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'),

    cellStyles = c('<cellStyle name="Normal" xfId="0" builtinId="0"/>'),

    dxfs = dxfs,

    tableStyles = tableStyles,

    indexedColors = NULL,

    extLst = extLst
  )
}

genBasePic <- function(imageNo, next_id) {
  sprintf('<xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="%s" name="Picture %s"/>
        <xdr:cNvPicPr>
          <a:picLocks noChangeAspect="1"/>
        </xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="%s">
        </a:blip>
        <a:stretch>
          <a:fillRect/>
        </a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </xdr:spPr>
    </xdr:pic>', imageNo, imageNo, next_id)
}

# TODO: make genBaseTheme() a more flexible function that allows further
# customization, not only the base font, but also colors and theme name.

# Not sure how the panose values below were created. The second byte looks
# wrong or at least it differs from our panose table.
# "\x0B", should be written as as "0B" and not as its decimal value "11".
# For Calibri it was "\x0F" and for this "0F" was the panose value and
# not "15". Probably it does not matter.
genBaseTheme <- function() {
  read_xml(
  stringi::stri_unescape_unicode(
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
  <a:clrScheme name="Office">
  <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
  <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
  <a:dk2><a:srgbClr val="0E2841"/></a:dk2>
  <a:lt2><a:srgbClr val="E8E8E8"/></a:lt2>
  <a:accent1><a:srgbClr val="156082"/></a:accent1>
  <a:accent2><a:srgbClr val="E97132"/></a:accent2>
  <a:accent3><a:srgbClr val="196B24"/></a:accent3>
  <a:accent4><a:srgbClr val="0F9ED5"/></a:accent4>
  <a:accent5><a:srgbClr val="A02B93"/></a:accent5>
  <a:accent6><a:srgbClr val="4EA72E"/></a:accent6>
  <a:hlink><a:srgbClr val="467886"/></a:hlink>
  <a:folHlink><a:srgbClr val="96607D"/></a:folHlink>
  </a:clrScheme>
  <a:fontScheme name="Office">
  <a:majorFont>
  <a:latin typeface="Aptos Display" panose="02110004020202020204"/>
  <a:ea typeface=""/>
  <a:cs typeface=""/>
  <a:font script="Jpan" typeface="\\u6e38\\u30b4\\u30b7\\u30c3\\u30af Light"/>
  <a:font script="Hang" typeface="\\ub9d1\\uc740 \\uace0\\ub515"/>
  <a:font script="Hans" typeface="\\u7b49\\u7ebf Light"/>
  <a:font script="Hant" typeface="\\u65b0\\u7d30\\u660e\\u9ad4"/>
  <a:font script="Arab" typeface="Times New Roman"/>
  <a:font script="Hebr" typeface="Times New Roman"/>
  <a:font script="Thai" typeface="Tahoma"/>
  <a:font script="Ethi" typeface="Nyala"/>
  <a:font script="Beng" typeface="Vrinda"/>
  <a:font script="Gujr" typeface="Shruti"/>
  <a:font script="Khmr" typeface="MoolBoran"/>
  <a:font script="Knda" typeface="Tunga"/>
  <a:font script="Guru" typeface="Raavi"/>
  <a:font script="Cans" typeface="Euphemia"/>
  <a:font script="Cher" typeface="Plantagenet Cherokee"/>
  <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
  <a:font script="Tibt" typeface="Microsoft Himalaya"/>
  <a:font script="Thaa" typeface="MV Boli"/>
  <a:font script="Deva" typeface="Mangal"/>
  <a:font script="Telu" typeface="Gautami"/>
  <a:font script="Taml" typeface="Latha"/>
  <a:font script="Syrc" typeface="Estrangelo Edessa"/>
  <a:font script="Orya" typeface="Kalinga"/>
  <a:font script="Mlym" typeface="Kartika"/>
  <a:font script="Laoo" typeface="DokChampa"/>
  <a:font script="Sinh" typeface="Iskoola Pota"/>
  <a:font script="Mong" typeface="Mongolian Baiti"/>
  <a:font script="Viet" typeface="Times New Roman"/>
  <a:font script="Uigh" typeface="Microsoft Uighur"/>
  <a:font script="Geor" typeface="Sylfaen"/>
  <a:font script="Armn" typeface="Arial"/>
  <a:font script="Bugi" typeface="Leelawadee UI"/>
  <a:font script="Bopo" typeface="Microsoft JhengHei"/>
  <a:font script="Java" typeface="Javanese Text"/>
  <a:font script="Lisu" typeface="Segoe UI"/>
  <a:font script="Mymr" typeface="Myanmar Text"/>
  <a:font script="Nkoo" typeface="Ebrima"/>
  <a:font script="Olck" typeface="Nirmala UI"/>
  <a:font script="Osma" typeface="Ebrima"/>
  <a:font script="Phag" typeface="Phagspa"/>
  <a:font script="Syrn" typeface="Estrangelo Edessa"/>
  <a:font script="Syrj" typeface="Estrangelo Edessa"/>
  <a:font script="Syre" typeface="Estrangelo Edessa"/>
  <a:font script="Sora" typeface="Nirmala UI"/>
  <a:font script="Tale" typeface="Microsoft Tai Le"/>
  <a:font script="Talu" typeface="Microsoft New Tai Lue"/>
  <a:font script="Tfng" typeface="Ebrima"/>
  </a:majorFont>
  <a:minorFont>
  <a:latin typeface="Aptos Narrow" panose="02110004020202020204"/>
  <a:ea typeface=""/>
  <a:cs typeface=""/>
  <a:font script="Jpan" typeface="\\u6e38\\u30b4\\u30b7\\u30c3\\u30af"/>
  <a:font script="Hang" typeface="\\ub9d1\\uc740 \\uace0\\ub515"/>
  <a:font script="Hans" typeface="\\u7b49\\u7ebf"/>
  <a:font script="Hant" typeface="\\u65b0\\u7d30\\u660e\\u9ad4"/>
  <a:font script="Arab" typeface="Arial"/>
  <a:font script="Hebr" typeface="Arial"/>
  <a:font script="Thai" typeface="Tahoma"/>
  <a:font script="Ethi" typeface="Nyala"/>
  <a:font script="Beng" typeface="Vrinda"/>
  <a:font script="Gujr" typeface="Shruti"/>
  <a:font script="Khmr" typeface="DaunPenh"/>
  <a:font script="Knda" typeface="Tunga"/>
  <a:font script="Guru" typeface="Raavi"/>
  <a:font script="Cans" typeface="Euphemia"/>
  <a:font script="Cher" typeface="Plantagenet Cherokee"/>
  <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
  <a:font script="Tibt" typeface="Microsoft Himalaya"/>
  <a:font script="Thaa" typeface="MV Boli"/>
  <a:font script="Deva" typeface="Mangal"/>
  <a:font script="Telu" typeface="Gautami"/>
  <a:font script="Taml" typeface="Latha"/>
  <a:font script="Syrc" typeface="Estrangelo Edessa"/>
  <a:font script="Orya" typeface="Kalinga"/>
  <a:font script="Mlym" typeface="Kartika"/>
  <a:font script="Laoo" typeface="DokChampa"/>
  <a:font script="Sinh" typeface="Iskoola Pota"/>
  <a:font script="Mong" typeface="Mongolian Baiti"/>
  <a:font script="Viet" typeface="Arial"/>
  <a:font script="Uigh" typeface="Microsoft Uighur"/>
  <a:font script="Geor" typeface="Sylfaen"/>
  <a:font script="Armn" typeface="Arial"/>
  <a:font script="Bugi" typeface="Leelawadee UI"/>
  <a:font script="Bopo" typeface="Microsoft JhengHei"/>
  <a:font script="Java" typeface="Javanese Text"/>
  <a:font script="Lisu" typeface="Segoe UI"/>
  <a:font script="Mymr" typeface="Myanmar Text"/>
  <a:font script="Nkoo" typeface="Ebrima"/>
  <a:font script="Olck" typeface="Nirmala UI"/>
  <a:font script="Osma" typeface="Ebrima"/>
  <a:font script="Phag" typeface="Phagspa"/>
  <a:font script="Syrn" typeface="Estrangelo Edessa"/>
  <a:font script="Syrj" typeface="Estrangelo Edessa"/>
  <a:font script="Syre" typeface="Estrangelo Edessa"/>
  <a:font script="Sora" typeface="Nirmala UI"/>
  <a:font script="Tale" typeface="Microsoft Tai Le"/>
  <a:font script="Talu" typeface="Microsoft New Tai Lue"/>
  <a:font script="Tfng" typeface="Ebrima"/>
  </a:minorFont>
  </a:fontScheme>
  <a:fmtScheme name="Office">
  <a:fillStyleLst>
  <a:solidFill>
  <a:schemeClr val="phClr"/>
  </a:solidFill>
  <a:gradFill rotWithShape="1">
  <a:gsLst>
  <a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs>
  <a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs>
  <a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs>
  </a:gsLst>
  <a:lin ang="5400000" scaled="0"/>
  </a:gradFill>
  <a:gradFill rotWithShape="1">
  <a:gsLst>
  <a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs>
  <a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs>
  <a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs>
  </a:gsLst>
  <a:lin ang="5400000" scaled="0"/>
  </a:gradFill>
  </a:fillStyleLst>
  <a:lnStyleLst>
  <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>
  <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>
  <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>
  </a:lnStyleLst>
  <a:effectStyleLst>
  <a:effectStyle><a:effectLst/></a:effectStyle>
  <a:effectStyle><a:effectLst/></a:effectStyle>
  <a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>
  </a:effectStyleLst>
  <a:bgFillStyleLst>
  <a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>
  <a:gradFill rotWithShape="1">
  <a:gsLst>
  <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs>
  <a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs>
  <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs>
  </a:gsLst>
  <a:lin ang="5400000" scaled="0"/>
  </a:gradFill>
  </a:bgFillStyleLst>
  </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults>
  <a:lnDef>
  <a:spPr/>
  <a:bodyPr/>
  <a:lstStyle/>
  <a:style>
  <a:lnRef idx="2">
  <a:schemeClr val="accent1"/>
  </a:lnRef>
  <a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>
  <a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef>
  <a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef>
  </a:style>
  </a:lnDef>
  </a:objectDefaults>
  <a:extraClrSchemeLst/>
  <a:extLst>
  <a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}">
  <thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{2E142A2C-CD16-42D6-873A-C26D2A0506FA}" vid="{1BDDFF52-6CD6-40A5-AB3C-68EB2F1E4D0A}"/>
  </a:ext>
  </a:extLst>
  </a:theme>'
  ), # stri_unescape_unicode
  pointer = FALSE
  ) # read_xml
}

gen_databar_extlst <- function(guid, sqref, posColor, negColor, values, border, gradient) {
  xml <- sprintf('<x14:cfRule type="dataBar" id="{%s}"><x14:dataBar minLength="0" maxLength="100" border="%s" gradient = "%s" negativeBarBorderColorSameAsPositive="0">', guid, border, gradient)

  if (is.null(values)) {
    xml <- sprintf('<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      %s
                      <x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:borderColor rgb="%s"/><x14:negativeFillColor rgb="%s"/><x14:negativeBorderColor rgb="%s"/><x14:axisColor rgb="FF000000"/>
                      </x14:dataBar></x14:cfRule><xm:sqref>%s</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>', xml, posColor, negColor, negColor, sqref)
  } else {
    xml <- sprintf('<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      %s
                      <x14:cfvo type="num"><xm:f>%s</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>%s</xm:f></x14:cfvo>
                      <x14:borderColor rgb="%s"/><x14:negativeFillColor rgb="%s"/><x14:negativeBorderColor rgb="%s"/><x14:axisColor rgb="FF000000"/>
                      </x14:dataBar></x14:cfRule><xm:sqref>%s</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>', xml, values[[1]], values[[2]], posColor, negColor, negColor, sqref)
  }

  return(xml)
}

genSlicerCachesExtLst <- function(i) {
  paste0(
    '<ext uri=\"{BBE1A952-AA13-448e-AADC-164F8A28A991}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">
    <x14:slicerCaches>',
    paste(sprintf('<x14:slicerCache r:id="rId%s"/>', i), collapse = ""),
    "</x14:slicerCaches></ext>"
  )
}


colors1_xml <- "<cs:colorStyle xmlns:cs=\"http://schemas.microsoft.com/office/drawing/2012/chartStyle\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" meth=\"cycle\" id=\"10\">
<a:schemeClr val=\"accent1\"/>
<a:schemeClr val=\"accent2\"/>
<a:schemeClr val=\"accent3\"/>
<a:schemeClr val=\"accent4\"/>
<a:schemeClr val=\"accent5\"/>
<a:schemeClr val=\"accent6\"/>
<cs:variation/>
<cs:variation><a:lumMod val=\"60000\"/></cs:variation>
<cs:variation><a:lumMod val=\"80000\"/><a:lumOff val=\"20000\"/></cs:variation>
<cs:variation><a:lumMod val=\"80000\"/></cs:variation>
<cs:variation><a:lumMod val=\"60000\"/><a:lumOff val=\"40000\"/></cs:variation>
<cs:variation><a:lumMod val=\"50000\"/></cs:variation>
<cs:variation><a:lumMod val=\"70000\"/><a:lumOff val=\"30000\"/></cs:variation>
<cs:variation><a:lumMod val=\"70000\"/></cs:variation>
<cs:variation><a:lumMod val=\"50000\"/><a:lumOff val=\"50000\"/></cs:variation>
</cs:colorStyle>"

styleplot_xml <- paste0('<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="201">
 <cs:axisTitle>
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1">
    <a:lumMod val="65000" />
    <a:lumOff val="35000" />
   </a:schemeClr>
  </cs:fontRef>
  <cs:defRPr sz="1000" kern="1200" />
 </cs:axisTitle>
 <cs:categoryAxis>
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1">
    <a:lumMod val="65000" />
    <a:lumOff val="35000" />
   </a:schemeClr>
  </cs:fontRef>
  <cs:spPr>
   <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
    <a:solidFill>
     <a:schemeClr val="tx1">
      <a:lumMod val="15000" />
      <a:lumOff val="85000" />
     </a:schemeClr>
    </a:solidFill>
    <a:round />
   </a:ln>
  </cs:spPr>
  <cs:defRPr sz="900" kern="1200" />
 </cs:categoryAxis>
 <cs:chartArea mods="allowNoFillOverride allowNoLineOverride">
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1" />
  </cs:fontRef>
  <cs:spPr>
   <a:solidFill>
    <a:schemeClr val="bg1" />
   </a:solidFill>
   <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
    <a:solidFill>
     <a:schemeClr val="tx1">
      <a:lumMod val="15000" />
      <a:lumOff val="85000" />
     </a:schemeClr>
    </a:solidFill>
    <a:round />
   </a:ln>
  </cs:spPr>
  <cs:defRPr sz="1000" kern="1200" />
 </cs:chartArea>
 <cs:dataLabel>
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1">
    <a:lumMod val="75000" />
    <a:lumOff val="25000" />
   </a:schemeClr>
  </cs:fontRef>
  <cs:defRPr sz="900" kern="1200" />
 </cs:dataLabel>
 <cs:dataLabelCallout>
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="dk1">
    <a:lumMod val="65000" />
    <a:lumOff val="35000" />
   </a:schemeClr>
  </cs:fontRef>
  <cs:spPr>
   <a:solidFill>
    <a:schemeClr val="lt1" />
   </a:solidFill>
   <a:ln>
    <a:solidFill>
     <a:schemeClr val="dk1">
      <a:lumMod val="25000" />
      <a:lumOff val="75000" />
     </a:schemeClr>
    </a:solidFill>
   </a:ln>
  </cs:spPr>
  <cs:defRPr sz="900" kern="1200" />
  <cs:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="clip" horzOverflow="clip" vert="horz" wrap="square" lIns="36576" tIns="18288" rIns="36576" bIns="18288" anchor="ctr" anchorCtr="1">
   <a:spAutoFit />
  </cs:bodyPr>
 </cs:dataLabelCallout>
 <cs:dataPoint>
  <cs:lnRef idx="0" />
  <cs:fillRef idx="1">
   <cs:styleClr val="auto" />
  </cs:fillRef>
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1" />
  </cs:fontRef>
 </cs:dataPoint>
 <cs:dataPoint3D>
  <cs:lnRef idx="0" />
  <cs:fillRef idx="1">
   <cs:styleClr val="auto" />
  </cs:fillRef>
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1" />
  </cs:fontRef>
 </cs:dataPoint3D>
 <cs:dataPointLine>
  <cs:lnRef idx="0">
   <cs:styleClr val="auto" />
  </cs:lnRef>
  <cs:fillRef idx="1" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1" />
  </cs:fontRef>
  <cs:spPr>
   <a:ln w="28575" cap="rnd">
    <a:solidFill>
     <a:schemeClr val="phClr" />
    </a:solidFill>
    <a:round />
   </a:ln>
  </cs:spPr>
 </cs:dataPointLine>
 <cs:dataPointMarker>
  <cs:lnRef idx="0">
   <cs:styleClr val="auto" />
  </cs:lnRef>
  <cs:fillRef idx="1">
   <cs:styleClr val="auto" />
  </cs:fillRef>
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1" />
  </cs:fontRef>
  <cs:spPr>
   <a:ln w="9525">
    <a:solidFill>
     <a:schemeClr val="phClr" />
    </a:solidFill>
   </a:ln>
  </cs:spPr>
 </cs:dataPointMarker>
 <cs:dataPointMarkerLayout symbol="circle" size="5" />
 <cs:dataPointWireframe>
  <cs:lnRef idx="0">
   <cs:styleClr val="auto" />
  </cs:lnRef>
  <cs:fillRef idx="1" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1" />
  </cs:fontRef>
  <cs:spPr>
   <a:ln w="9525" cap="rnd">
    <a:solidFill>
     <a:schemeClr val="phClr" />
    </a:solidFill>
    <a:round />
   </a:ln>
  </cs:spPr>
 </cs:dataPointWireframe>
 <cs:dataTable>
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1">
    <a:lumMod val="65000" />
    <a:lumOff val="35000" />
   </a:schemeClr>
  </cs:fontRef>
  <cs:spPr>
   <a:noFill />
   <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
    <a:solidFill>
     <a:schemeClr val="tx1">
      <a:lumMod val="15000" />
      <a:lumOff val="85000" />
     </a:schemeClr>
    </a:solidFill>
    <a:round />
   </a:ln>
  </cs:spPr>
  <cs:defRPr sz="900" kern="1200" />
 </cs:dataTable>
 <cs:downBar>
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="dk1" />
  </cs:fontRef>
  <cs:spPr>
   <a:solidFill>
    <a:schemeClr val="dk1">
     <a:lumMod val="65000" />
     <a:lumOff val="35000" />
    </a:schemeClr>
   </a:solidFill>
   <a:ln w="9525">
    <a:solidFill>
     <a:schemeClr val="tx1">
      <a:lumMod val="65000" />
      <a:lumOff val="35000" />
     </a:schemeClr>
    </a:solidFill>
   </a:ln>
  </cs:spPr>
 </cs:downBar>
 <cs:dropLine>
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1" />
  </cs:fontRef>
  <cs:spPr>
   <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
    <a:solidFill>
     <a:schemeClr val="tx1">
      <a:lumMod val="35000" />
      <a:lumOff val="65000" />
     </a:schemeClr>
    </a:solidFill>
    <a:round />
   </a:ln>
  </cs:spPr>
 </cs:dropLine>
 <cs:errorBar>
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1" />
  </cs:fontRef>
  <cs:spPr>
   <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
    <a:solidFill>
     <a:schemeClr val="tx1">
      <a:lumMod val="65000" />
      <a:lumOff val="35000" />
     </a:schemeClr>
    </a:solidFill>
    <a:round />
   </a:ln>
  </cs:spPr>
 </cs:errorBar>
 <cs:floor>
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1" />
  </cs:fontRef>
  <cs:spPr>
   <a:noFill />
   <a:ln>
    <a:noFill />
   </a:ln>
  </cs:spPr>
 </cs:floor>
 <cs:gridlineMajor>
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1" />
  </cs:fontRef>
  <cs:spPr>
   <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
    <a:solidFill>
     <a:schemeClr val="tx1">
      <a:lumMod val="15000" />
      <a:lumOff val="85000" />
     </a:schemeClr>
    </a:solidFill>
    <a:round />
   </a:ln>
  </cs:spPr>
 </cs:gridlineMajor>
 <cs:gridlineMinor>
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1" />
  </cs:fontRef>
  <cs:spPr>
   <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
    <a:solidFill>
     <a:schemeClr val="tx1">
      <a:lumMod val="5000" />
      <a:lumOff val="95000" />
     </a:schemeClr>
    </a:solidFill>
    <a:round />
   </a:ln>
  </cs:spPr>
 </cs:gridlineMinor>
 <cs:hiLoLine>
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1" />
  </cs:fontRef>
  <cs:spPr>
   <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
    <a:solidFill>
     <a:schemeClr val="tx1">
      <a:lumMod val="75000" />
      <a:lumOff val="25000" />
     </a:schemeClr>
    </a:solidFill>
    <a:round />
   </a:ln>
  </cs:spPr>
 </cs:hiLoLine>
 <cs:leaderLine>
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1" />
  </cs:fontRef>
  <cs:spPr>
   <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
    <a:solidFill>
     <a:schemeClr val="tx1">
      <a:lumMod val="35000" />
      <a:lumOff val="65000" />
     </a:schemeClr>
    </a:solidFill>
    <a:round />
   </a:ln>
  </cs:spPr>
 </cs:leaderLine>
 <cs:legend>
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1">
    <a:lumMod val="65000" />
    <a:lumOff val="35000" />
   </a:schemeClr>
  </cs:fontRef>
  <cs:defRPr sz="900" kern="1200" />
 </cs:legend>
 <cs:plotArea mods="allowNoFillOverride allowNoLineOverride">
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1" />
  </cs:fontRef>
 </cs:plotArea>
 <cs:plotArea3D mods="allowNoFillOverride allowNoLineOverride">
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1" />
  </cs:fontRef>
 </cs:plotArea3D>
 <cs:seriesAxis>
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1">
    <a:lumMod val="65000" />
    <a:lumOff val="35000" />
   </a:schemeClr>
  </cs:fontRef>
  <cs:defRPr sz="900" kern="1200" />
 </cs:seriesAxis>
 <cs:seriesLine>
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1" />
  </cs:fontRef>
  <cs:spPr>
   <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
    <a:solidFill>
     <a:schemeClr val="tx1">
      <a:lumMod val="35000" />
      <a:lumOff val="65000" />
     </a:schemeClr>
    </a:solidFill>
    <a:round />
   </a:ln>
  </cs:spPr>
 </cs:seriesLine>
 <cs:title>
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1">
    <a:lumMod val="65000" />
    <a:lumOff val="35000" />
   </a:schemeClr>
  </cs:fontRef>
  <cs:defRPr sz="1400" b="0" kern="1200" spc="0" baseline="0" />
 </cs:title>
 <cs:trendline>
  <cs:lnRef idx="0">
   <cs:styleClr val="auto" />
  </cs:lnRef>
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1" />
  </cs:fontRef>
  <cs:spPr>
   <a:ln w="19050" cap="rnd">
    <a:solidFill>
     <a:schemeClr val="phClr" />
    </a:solidFill>
    <a:prstDash val="sysDot" />
   </a:ln>
  </cs:spPr>
 </cs:trendline>
 <cs:trendlineLabel>
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1">
    <a:lumMod val="65000" />
    <a:lumOff val="35000" />
   </a:schemeClr>
  </cs:fontRef>
  <cs:defRPr sz="900" kern="1200" />
 </cs:trendlineLabel>
 <cs:upBar>
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="dk1" />
  </cs:fontRef>
  <cs:spPr>
   <a:solidFill>
    <a:schemeClr val="lt1" />
   </a:solidFill>
   <a:ln w="9525">
    <a:solidFill>
     <a:schemeClr val="tx1">
      <a:lumMod val="15000" />
      <a:lumOff val="85000" />
     </a:schemeClr>
    </a:solidFill>
   </a:ln>
  </cs:spPr>
 </cs:upBar>
 <cs:valueAxis>
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1">
    <a:lumMod val="65000" />
    <a:lumOff val="35000" />
   </a:schemeClr>
  </cs:fontRef>
  <cs:defRPr sz="900" kern="1200" />
 </cs:valueAxis>
 <cs:wall>
  <cs:lnRef idx="0" />
  <cs:fillRef idx="0" />
  <cs:effectRef idx="0" />
  <cs:fontRef idx="minor">
   <a:schemeClr val="tx1" />
  </cs:fontRef>
  <cs:spPr>
   <a:noFill />
   <a:ln>
    <a:noFill />
   </a:ln>
  </cs:spPr>
 </cs:wall>
</cs:chartStyle>')


drawings <- function(drawings, drawing_id) {


  rel_len <- length(xml_node(drawings, "Relationship"))

  drawings <- xml_node_create(
    xml_name = "xdr:wsDr",
    xml_attributes = c(
      `xmlns:xdr` = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
      `xmlns:a`   = "http://schemas.openxmlformats.org/drawingml/2006/main"
    )
  )

  drawing <- paste0(
    sprintf(
     '<xdr:absoluteAnchor>
        <xdr:pos x="0" y="0" />
        <xdr:ext cx="9313333" cy="6070985" />
        <xdr:graphicFrame macro="">
        <xdr:nvGraphicFramePr>
          <xdr:cNvPr id="2" name="Chart %s">
          <a:extLst>
            <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
            <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" />
            </a:ext>
          </a:extLst>
          </xdr:cNvPr>
          <xdr:cNvGraphicFramePr>
          <a:graphicFrameLocks noGrp="1" />
          </xdr:cNvGraphicFramePr>
        </xdr:nvGraphicFramePr>
        <xdr:xfrm>
          <a:off x="0" y="0" />
          <a:ext cx="0" cy="0" />
        </xdr:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId%s" />
          </a:graphicData>
        </a:graphic>
        </xdr:graphicFrame>
        <xdr:clientData />
      </xdr:absoluteAnchor>',
      drawing_id,
      rel_len + 1L
    )
  )

  return(
    xml_append_child1(
      node    = read_xml(drawings),
      child   = read_xml(drawing),
      pointer = FALSE)
  )
}

chart1_rels_xml <- function(x) {
  sprintf("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
<Relationship Id=\"rId2\" Type=\"http://schemas.microsoft.com/office/2011/relationships/chartColorStyle\" Target=\"colors%s.xml\"/>
<Relationship Id=\"rId1\" Type=\"http://schemas.microsoft.com/office/2011/relationships/chartStyle\" Target=\"style%s.xml\"/>
</Relationships>",
          x,
          x
  )
}

drawings_rels <- function(drawings, x) {

  # ignore default case ""
  if (all(drawings == "")) {
    drawings <- NULL
  }

  rel_len <- length(xml_node(drawings, "Relationship"))

  drawings <- c(
    drawings,
    sprintf(
      "<Relationship Id=\"rId%s\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" Target=\"../charts/chart%s.xml\"/>",
      rel_len + 1,
      x
    )
  )

  return(drawings)
}

formCntrlDrawing <- function(type, len) {
  if (type == "Checkbox") {
      drawing <- read_xml(
        sprintf(
          '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
          <mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">
          <xdr:twoCellAnchor editAs="oneCell">
          <xdr:from>
          <xdr:col>0</xdr:col>
          <xdr:colOff>0</xdr:colOff>
          <xdr:row>2</xdr:row>
          <xdr:rowOff>0</xdr:rowOff>
          </xdr:from>
          <xdr:to>
          <xdr:col>2</xdr:col>
          <xdr:colOff>0</xdr:colOff>
          <xdr:row>4</xdr:row>
          <xdr:rowOff>0</xdr:rowOff>
          </xdr:to>
          <xdr:sp macro="" textlink="">
          <xdr:nvSpPr>
          <xdr:cNvPr id="%s" name="Check Box 1" hidden="1">
          <a:extLst>
          <a:ext uri="{63B3BB69-23CF-44E3-9099-%s}">
          <a14:compatExt spid="_x0000_s1025"/>
          </a:ext>
          <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-%s}">
          <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{1EEBA38D-B81F-E55E-AFDB-%s}"/>
          </a:ext>
          </a:extLst>
          </xdr:cNvPr>
          <xdr:cNvSpPr/>
          </xdr:nvSpPr>
          <xdr:spPr bwMode="auto">
          <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          </a:xfrm>
          <a:prstGeom prst="rect">
          <a:avLst/>
          </a:prstGeom>
          <a:noFill/>
          <a:ln>
          <a:noFill/>
          </a:ln>
          <a:extLst>
          <a:ext uri="{909E8E84-426E-40DD-AFC4-%s}">
          <a14:hiddenFill>
          <a:solidFill>
          <a:srgbClr val="FFFFFF" mc:Ignorable="a14" a14:legacySpreadsheetColorIndex="65"/>
          </a:solidFill>
          </a14:hiddenFill>
          </a:ext>
          <a:ext uri="{91240B29-F687-4F45-9708-%s}">
          <a14:hiddenLine w="9525">
          <a:solidFill>
          <a:srgbClr val="000000" mc:Ignorable="a14" a14:legacySpreadsheetColorIndex="64"/>
          </a:solidFill>
          <a:miter lim="800000"/>
          <a:headEnd/>
          <a:tailEnd/>
          </a14:hiddenLine>
          </a:ext>
          </a:extLst>
          </xdr:spPr>
          <xdr:txBody>
          <a:bodyPr vertOverflow="clip" wrap="square" lIns="27432" tIns="22860" rIns="0" bIns="22860" anchor="ctr" upright="1"/>
          <a:lstStyle/>
          <a:p>
          <a:pPr algn="l" rtl="0">
          <a:defRPr sz="1000"/>
          </a:pPr>
          <a:r>
          <a:rPr lang="en-GB" sz="1300" b="0" i="0" u="none" strike="noStrike" baseline="0">
          <a:solidFill>
          <a:srgbClr val="000000"/>
          </a:solidFill>
          <a:latin typeface="Lucida Grande" charset="0"/>
          <a:cs typeface="Lucida Grande" charset="0"/>
          </a:rPr>
          <a:t>Check Box 1</a:t>
          </a:r>
          </a:p>
          </xdr:txBody>
          </xdr:sp>
          <xdr:clientData/>
          </xdr:twoCellAnchor>
          </mc:Choice>
          <mc:Fallback/>
          </mc:AlternateContent>
          </xdr:wsDr>',
          len,
          random_string(length = 12),
          random_string(length = 12),
          random_string(length = 12),
          random_string(length = 12),
          random_string(length = 12)
          ), pointer = FALSE
      )
      } else if (type == "Radio") {
        drawing <- read_xml(
          sprintf(
            '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
            <mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">
            <xdr:twoCellAnchor editAs="oneCell">
            <xdr:from>
            <xdr:col>0</xdr:col>
            <xdr:colOff>800100</xdr:colOff>
            <xdr:row>7</xdr:row>
            <xdr:rowOff>114300</xdr:rowOff>
            </xdr:from>
            <xdr:to>
            <xdr:col>2</xdr:col>
            <xdr:colOff>673100</xdr:colOff>
            <xdr:row>9</xdr:row>
            <xdr:rowOff>88900</xdr:rowOff>
            </xdr:to>
            <xdr:sp macro="" textlink="">
            <xdr:nvSpPr>
            <xdr:cNvPr id="%s" name="Option Button 2" hidden="1">
            <a:extLst>
            <a:ext uri="{63B3BB69-23CF-44E3-9099-%s}">
            <a14:compatExt spid="_x0000_s1026" />
            </a:ext>
            <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-%s}">
            <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{BF281B22-8B3D-E7A6-DB69-%s}" />
            </a:ext>
            </a:extLst>
            </xdr:cNvPr>
            <xdr:cNvSpPr />
            </xdr:nvSpPr>
            <xdr:spPr bwMode="auto">
            <a:xfrm>
            <a:off x="0" y="0" />
            <a:ext cx="0" cy="0" />
            </a:xfrm>
            <a:prstGeom prst="rect">
            <a:avLst />
            </a:prstGeom>
            <a:noFill />
            <a:ln>
            <a:noFill />
            </a:ln>
            <a:extLst>
            <a:ext uri="{909E8E84-426E-40DD-AFC4-%s}">
            <a14:hiddenFill>
            <a:solidFill>
            <a:srgbClr val="FFFFFF" mc:Ignorable="a14" a14:legacySpreadsheetColorIndex="65" />
            </a:solidFill>
            </a14:hiddenFill>
            </a:ext>
            <a:ext uri="{91240B29-F687-4F45-9708-%s}">
            <a14:hiddenLine w="9525">
            <a:solidFill>
            <a:srgbClr val="000000" mc:Ignorable="a14" a14:legacySpreadsheetColorIndex="64" />
            </a:solidFill>
            <a:miter lim="800000" />
            <a:headEnd />
            <a:tailEnd />
            </a14:hiddenLine>
            </a:ext>
            </a:extLst>
            </xdr:spPr>
            <xdr:txBody>
            <a:bodyPr vertOverflow="clip" wrap="square" lIns="27432" tIns="22860" rIns="0" bIns="22860" anchor="ctr" upright="1" />
            <a:lstStyle />
            <a:p>
            <a:pPr algn="l" rtl="0">
            <a:defRPr sz="1000" />
            </a:pPr>
            <a:r>
            <a:rPr lang="en-GB" sz="1300" b="0" i="0" u="none" strike="noStrike" baseline="0">
            <a:solidFill>
            <a:srgbClr val="000000" />
            </a:solidFill>
            <a:latin typeface="Lucida Grande" charset="0" />
            <a:cs typeface="Lucida Grande" charset="0" />
            </a:rPr>
            <a:t>Option Button 2</a:t>
            </a:r>
            </a:p>
            </xdr:txBody>
            </xdr:sp>
            <xdr:clientData />
            </xdr:twoCellAnchor>
            </mc:Choice>
            <mc:Fallback />
            </mc:AlternateContent>
            </xdr:wsDr>',
          len,
          random_string(length = 12),
          random_string(length = 12),
          random_string(length = 12),
          random_string(length = 12),
          random_string(length = 12)
          ), pointer = FALSE
      )
      } else if (type == "Drop") {
        drawing <- read_xml(
          sprintf(
            '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
            <mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">
            <xdr:twoCellAnchor editAs="oneCell">
            <xdr:from>
            <xdr:col>5</xdr:col>
            <xdr:colOff>139700</xdr:colOff>
            <xdr:row>3</xdr:row>
            <xdr:rowOff>76200</xdr:rowOff>
            </xdr:from>
            <xdr:to>
            <xdr:col>7</xdr:col>
            <xdr:colOff>774700</xdr:colOff>
            <xdr:row>7</xdr:row>
            <xdr:rowOff>25400</xdr:rowOff>
            </xdr:to>
            <xdr:sp macro="" textlink="">
            <xdr:nvSpPr>
            <xdr:cNvPr id="%s" name="Drop Down 5" hidden="1">
            <a:extLst>
            <a:ext uri="{63B3BB69-23CF-44E3-9099-%s}">
            <a14:compatExt spid="_x0000_s1029"/>
            </a:ext>
            <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-%s}">
            <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{AAC003F9-B5A3-DB5D-22B2-%s}"/>
            </a:ext>
            </a:extLst>
            </xdr:cNvPr>
            <xdr:cNvSpPr/>
            </xdr:nvSpPr>
            <xdr:spPr bwMode="auto">
            <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="0" cy="0"/>
            </a:xfrm>
            <a:prstGeom prst="rect">
            <a:avLst/>
            </a:prstGeom>
            <a:noFill/>
            <a:ln>
            <a:noFill/>
            </a:ln>
            <a:extLst>
            <a:ext uri="{91240B29-F687-4F45-9708-%s}">
            <a14:hiddenLine w="9525">
            <a:noFill/>
            <a:miter lim="800000"/>
            <a:headEnd/>
            <a:tailEnd/>
            </a14:hiddenLine>
            </a:ext>
            </a:extLst>
            </xdr:spPr>
            </xdr:sp>
            <xdr:clientData/>
            </xdr:twoCellAnchor>
            </mc:Choice>
            <mc:Fallback/>
            </mc:AlternateContent>
            </xdr:wsDr>',
          len,
          random_string(length = 12),
          random_string(length = 12),
          random_string(length = 12),
          random_string(length = 12)
          ), pointer = FALSE
      )
      }

  drawing
}
