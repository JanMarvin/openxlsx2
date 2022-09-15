

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


genBaseShapeVML <- function(clientData, id) {
  if (grepl("visible", clientData, ignore.case = TRUE)) {
    visible <- "visible"
  } else {
    visible <- "hidden"
  }

  paste0(
    sprintf('<v:shape id="_x0000_s%s" type="#_x0000_t202" style=\'position:absolute;', id),
    sprintf('margin-left:107.25pt;margin-top:172.5pt;width:147pt;height:96pt;z-index:1;
          visibility:%s;mso-wrap-style:tight\' fillcolor="#ffffe1" o:insetmode="auto">', visible),
    '<v:fill color2="#ffffe1"/>
            <v:shadow color="black" obscured="t"/>
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


# genBaseRels <- function() {
#
#   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
#    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
#    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
#
# }
#
#
# genBaseApp <- function() {
#   list('<Application>Microsoft Excel</Application>')
# }


genBaseCore <- function(creator = "", title = NULL, subject = NULL, category = NULL) {
  core <- '<coreProperties xmlns="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'

  core <- stri_c(core, sprintf("<dc:creator>%s</dc:creator>", creator))
  core <- stri_c(core, sprintf("<cp:lastModifiedBy>%s</cp:lastModifiedBy>", creator))
  core <- stri_c(core, sprintf('<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>', format(as_POSIXct_utc(Sys.time()), "%Y-%m-%dT%H:%M:%SZ")))

  if (!is.null(title)) {
    core <- stri_c(core, sprintf("<dc:title>%s</dc:title>", replace_legal_chars(title)))
  }

  if (!is.null(subject)) {
    core <- stri_c(core, sprintf("<dc:subject>%s</dc:subject>", replace_legal_chars(subject)))
  }

  if (!is.null(category)) {
    core <- stri_c(core, sprintf("<cp:category>%s</cp:category>", replace_legal_chars(category)))
  }

  core <- stri_c(core, "</coreProperties>")

  return(core)
}

#
# addAuthor <- function(wb,Author = NULL) {
#
#   if (!is.null(Author)) {
#     current_creator <-
#       stri_match(wb$core, regex = "<dc:creator>(.*?)</dc:creator>")[1, 2]
#     wb$core <-
#       stri_replace_all_fixed(
#         wb$core,
#         pattern = current_creator,
#         replacement = stri_c(current_creator, Author, sep = ";")
#       )
#   }
#
#
# }
#
#
# setAuthor <- function(wb,Author = NULL) {
#
#   if (!is.null(Author)) {
#     current_creator <-
#       stri_match(wb$core, regex = "<dc:creator>(.*?)</dc:creator>")[1, 2]
#     wb$core <-
#       stri_replace_all_fixed(
#         wb$core,
#         pattern = current_creator,
#         replacement =  Author
#       )
#   }
#
#
# }
#
# wb_set_last_modified_by <- function(wb,ModifiedBy=NULL) {
#
#   if (!is.null(addmodifier)) {
#     current_lastmodifier <-
#       stri_match(wb$core, regex = "<cp:lastModifiedBy>(.*?)</cp:lastModifiedBy>")[1, 2]
#     wb$core <-
#       stri_replace_all_fixed(
#         wb$core,
#         pattern = current_lastmodifier,
#         replacement = ModifiedBy
#       )
#   }
#
#
# }
#
#

#
#
# setBaseCore <- function(core,setcreator="",setmodifier="",
#                         title = NULL, subject = NULL, category = NULL) {
#
#
#   core <- c(core, sprintf('<dc:creator>%s</dc:creator>', setcreator))
#   core <- c(core, sprintf('<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>', format(Sys.time(), "%Y-%m-%dT%H:%M:%SZ")))
#
#   if (!is.null(title))
#     core <- c(core, sprintf('<dc:title>%s</dc:title>', replace_legal_chars(title)))
#
#   if (!is.null(subject))
#     core <- c(core, sprintf('<dc:subject>%s</dc:subject>', replace_legal_chars(subject)))
#
#   if (!is.null(category))
#     core <- c(core, sprintf('<cp:category>%s</cp:category>', replace_legal_chars(category)))
#
#   core <- c(core, '</coreProperties>')
#
#   return(core)
#
# }




genBaseWorkbook.xml.rels <- function() {
  c(
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>',
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
    fileVersion = NULL,
    fileSharing = NULL,
    workbookPr = '<workbookPr date1904="false"/>',
    alternateContent = NULL,
    absPath = NULL, # "x15ac:absPath"
    workbookProtection = NULL,
    bookViews = NULL,
    sheets = NULL,
    functionGroups = NULL,
    externalReferences = NULL,
    definedNames = NULL,
    calcPr = NULL,
    oleSize = NULL,
    customWorkbookViews = NULL,
    pivotCaches = NULL,
    smartTagPr = NULL,
    smartTagTypes = NULL,
    webPublishing = NULL,
    fileRecoveryPr = NULL,
    webPublishObjects = NULL,
    extLst = NULL
  )
}

# dummy should be removed soon
genBaseSheetRels <- function(sheetInd) {
  character()
}

genBaseStyleSheet <- function(dxfs = NULL, tableStyles = NULL, extLst = NULL) {
  list(
    numFmts = NULL,

    fonts = c('<font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>'),

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


genBasePic <- function(imageNo) {
  sprintf('<xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="%s" name="Picture %s"/>
        <xdr:cNvPicPr>
          <a:picLocks noChangeAspect="1"/>
        </xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId%s">
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
    </xdr:pic>', imageNo, imageNo, imageNo)
}










genBaseTheme <- function() {
  stringi::stri_unescape_unicode(
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
  <a:clrScheme name="Office">
  <a:dk1>
  <a:sysClr val="windowText" lastClr="000000"/>
  </a:dk1>
  <a:lt1>
  <a:sysClr val="window" lastClr="FFFFFF"/>
  </a:lt1>
  <a:dk2>
  <a:srgbClr val="44546A"/>
  </a:dk2>
  <a:lt2>
  <a:srgbClr val="E7E6E6"/>
  </a:lt2>
  <a:accent1>
  <a:srgbClr val="4472C4"/>
  </a:accent1>
  <a:accent2>
  <a:srgbClr val="ED7D31"/>
  </a:accent2>
  <a:accent3>
  <a:srgbClr val="A5A5A5"/>
  </a:accent3>
  <a:accent4>
  <a:srgbClr val="FFC000"/>
  </a:accent4>
  <a:accent5>
  <a:srgbClr val="5B9BD5"/>
  </a:accent5>
  <a:accent6>
  <a:srgbClr val="70AD47"/>
  </a:accent6>
  <a:hlink>
  <a:srgbClr val="0563C1"/>
  </a:hlink>
  <a:folHlink>
  <a:srgbClr val="954F72"/>
  </a:folHlink>
  </a:clrScheme>
  <a:fontScheme name="Office">
  <a:majorFont>
  <a:latin typeface="Calibri Light" panose="020F0302020204030204"/>
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
  <a:latin typeface="Calibri" panose="020F0502020204030204"/>
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
  <a:gs pos="0">
  <a:schemeClr val="phClr">
  <a:lumMod val="110000"/>
  <a:satMod val="105000"/>
  <a:tint val="67000"/>
  </a:schemeClr>
  </a:gs>
  <a:gs pos="50000">
  <a:schemeClr val="phClr">
  <a:lumMod val="105000"/>
  <a:satMod val="103000"/>
  <a:tint val="73000"/>
  </a:schemeClr>
  </a:gs>
  <a:gs pos="100000">
  <a:schemeClr val="phClr">
  <a:lumMod val="105000"/>
  <a:satMod val="109000"/>
  <a:tint val="81000"/>
  </a:schemeClr>
  </a:gs>
  </a:gsLst>
  <a:lin ang="5400000" scaled="0"/>
  </a:gradFill>
  <a:gradFill rotWithShape="1">
  <a:gsLst>
  <a:gs pos="0">
  <a:schemeClr val="phClr">
  <a:satMod val="103000"/>
  <a:lumMod val="102000"/>
  <a:tint val="94000"/>
  </a:schemeClr>
  </a:gs>
  <a:gs pos="50000">
  <a:schemeClr val="phClr">
  <a:satMod val="110000"/>
  <a:lumMod val="100000"/>
  <a:shade val="100000"/>
  </a:schemeClr>
  </a:gs>
  <a:gs pos="100000">
  <a:schemeClr val="phClr">
  <a:lumMod val="99000"/>
  <a:satMod val="120000"/>
  <a:shade val="78000"/>
  </a:schemeClr>
  </a:gs>
  </a:gsLst>
  <a:lin ang="5400000" scaled="0"/>
  </a:gradFill>
  </a:fillStyleLst>
  <a:lnStyleLst>
  <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">
  <a:solidFill>
  <a:schemeClr val="phClr"/>
  </a:solidFill>
  <a:prstDash val="solid"/>
  <a:miter lim="800000"/>
  </a:ln>
  <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
  <a:solidFill>
  <a:schemeClr val="phClr"/>
  </a:solidFill>
  <a:prstDash val="solid"/>
  <a:miter lim="800000"/>
  </a:ln>
  <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">
  <a:solidFill>
  <a:schemeClr val="phClr"/>
  </a:solidFill>
  <a:prstDash val="solid"/>
  <a:miter lim="800000"/>
  </a:ln>
  </a:lnStyleLst>
  <a:effectStyleLst>
  <a:effectStyle>
  <a:effectLst/>
  </a:effectStyle>
  <a:effectStyle>
  <a:effectLst/>
  </a:effectStyle>
  <a:effectStyle>
  <a:effectLst>
  <a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
  <a:srgbClr val="000000">
  <a:alpha val="63000"/>
  </a:srgbClr>
  </a:outerShdw>
  </a:effectLst>
  </a:effectStyle>
  </a:effectStyleLst>
  <a:bgFillStyleLst>
  <a:solidFill>
  <a:schemeClr val="phClr"/>
  </a:solidFill>
  <a:solidFill>
  <a:schemeClr val="phClr">
  <a:tint val="95000"/>
  <a:satMod val="170000"/>
  </a:schemeClr>
  </a:solidFill>
  <a:gradFill rotWithShape="1">
  <a:gsLst>
  <a:gs pos="0">
  <a:schemeClr val="phClr">
  <a:tint val="93000"/>
  <a:satMod val="150000"/>
  <a:shade val="98000"/>
  <a:lumMod val="102000"/>
  </a:schemeClr>
  </a:gs>
  <a:gs pos="50000">
  <a:schemeClr val="phClr">
  <a:tint val="98000"/>
  <a:satMod val="130000"/>
  <a:shade val="90000"/>
  <a:lumMod val="103000"/>
  </a:schemeClr>
  </a:gs>
  <a:gs pos="100000">
  <a:schemeClr val="phClr">
  <a:shade val="63000"/>
  <a:satMod val="120000"/>
  </a:schemeClr>
  </a:gs>
  </a:gsLst>
  <a:lin ang="5400000" scaled="0"/>
  </a:gradFill>
  </a:bgFillStyleLst>
  </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
  <a:extLst>
  <a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}">
  <thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/>
  </a:ext>
  </a:extLst>
  </a:theme>'
  )
}

gen_databar_extlst <- function(guid, sqref, posColour, negColour, values, border, gradient) {
  xml <- sprintf('<x14:cfRule type="dataBar" id="{%s}"><x14:dataBar minLength="0" maxLength="100" border="%s" gradient = "%s" negativeBarBorderColorSameAsPositive="0">', guid, border, gradient)

  if (is.null(values)) {
    xml <- sprintf('<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      %s
                      <x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:borderColor rgb="%s"/><x14:negativeFillColor rgb="%s"/><x14:negativeBorderColor rgb="%s"/><x14:axisColor rgb="FF000000"/>
                      </x14:dataBar></x14:cfRule><xm:sqref>%s</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>', xml, posColour, negColour, negColour, sqref)
  } else {
    xml <- sprintf('<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      %s
                      <x14:cfvo type="num"><xm:f>%s</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>%s</xm:f></x14:cfvo>
                      <x14:borderColor rgb="%s"/><x14:negativeFillColor rgb="%s"/><x14:negativeBorderColor rgb="%s"/><x14:axisColor rgb="FF000000"/>
                      </x14:dataBar></x14:cfRule><xm:sqref>%s</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>', xml, values[[1]], values[[2]], posColour, negColour, negColour, sqref)
  }

  return(xml)
}



contentTypePivotXML <- function(i) {
  c(
    sprintf('<Override PartName="/xl/pivotCache/pivotCacheDefinition%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>', i),
    sprintf('<Override PartName="/xl/pivotCache/pivotCacheRecords%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>', i),
    sprintf('<Override PartName="/xl/pivotTables/pivotTable%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>', i)
  )
}

contentTypeSlicerCacheXML <- function(i) {
  c(
    sprintf('<Override PartName="/xl/slicerCaches/slicerCache%s.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>', i),
    sprintf('<Override PartName="/xl/slicers/slicer%s.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>', i)
  )
}


genBaseSlicerXML <- function() {
  '<ext uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
    <x14:slicerList>
    <x14:slicer r:id="rId0"/>
      </x14:slicerList>
      </ext>'
}


genSlicerCachesExtLst <- function(i) {
  paste0(
    '<ext uri=\"{BBE1A952-AA13-448e-AADC-164F8A28A991}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">
    <x14:slicerCaches>',
    paste(sprintf('<x14:slicerCache r:id="rId%s"/>', i), collapse = ""),
    "</x14:slicerCaches></ext>"
  )
}
