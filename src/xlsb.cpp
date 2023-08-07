// #include <Rcpp/Lightest>

#include "openxlsx2.h"
#include "xlsb_defines.h"
#include "xlsb_funs.h"

#include <iomanip>

// [[Rcpp::export]]
int styles(std::string filePath, std::string outPath, bool debug) {

  std::ofstream out(outPath);
  std::ifstream bin(filePath, std::ios::in | std::ios::binary | std::ios::ate);

  // auto sas_size = bin.tellg();
  if (bin) {
    bin.seekg(0, std::ios_base::beg);
    bool end_of_style_sheet = false;

    while(!end_of_style_sheet) {
      Rcpp::checkUserInterrupt();

      int32_t x = 0, size = 0;

      if (debug) Rcpp::Rcout << "." << std::endl;
      RECORD(x, size, bin);
      if (debug) Rcpp::Rcout << x << ": " << size << std::endl;

      switch(x) {

      case BrtBeginStyleSheet:
      {
        out << "<styleSheet>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

        // fills
      case BrtBeginFmts:
      {
        out << "<numFmts>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtFmt:
      {
        out << "<numFmt";
        // bin.seekg(size, bin.cur);
        uint16_t ifmt = 0;
        ifmt = readbin(ifmt, bin, 0);

        std::string stFmtCode = XLWideString(bin);

        out << " numFmtId=\"" << ifmt << "\" formatCode=\"" << escape_xml(stFmtCode);
        out << "\" />" << std::endl;
        break;
      }

      case BrtEndFmts:
      {
        out << "</numFmts>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtBeginFonts:
      {
        if (debug) Rcpp::Rcout << "<fonts>" << std::endl;
        // bin.seekg(size, bin.cur);
        uint32_t cfonts = 0;
        cfonts = readbin(cfonts, bin, 0);

        out << "<fonts " <<
          "count=\"" << cfonts <<
            "\">" << std::endl;
        break;
      }

      case BrtFont:
      {
        out << "<font>" << std::endl;
        uint8_t uls = 0, bFamily = 0, bCharSet = 0, unused = 0, bFontScheme = 0;
        uint16_t dyHeight = 0, grbit = 0, bls = 0, sss = 0;

        dyHeight = readbin(dyHeight, bin , 0) / 20; // twip
        grbit = readbin(grbit, bin , 0);
        FontFlagsFields *fields = (FontFlagsFields *)&grbit;

        bls = readbin(bls, bin , 0);
        // 0x0190 normal
        // 0x02BC bold
        sss = readbin(sss, bin , 0);
        // 0x0000 None
        // 0x0001 Superscript
        // 0x0002 Subscript
        uls = readbin(uls, bin , 0);
        // 0x00 None
        // 0x01 Single
        // 0x02 Double
        // 0x21 Single accounting
        // 0x22 Double accounting
        bFamily = readbin(bFamily, bin , 0);
        // 0x00 Not applicable
        // 0x01 Roman
        // 0x02 Swiss
        // 0x03 Modern
        // 0x04 Script
        // 0x05 Decorative
        bCharSet = readbin(bCharSet, bin , 0);
        // 0x00 ANSI_CHARSET English
        // 0x01 DEFAULT_CHARSET System locale based
        // 0x02 SYMBOL_CHARSET Symbol
        // 0x4D MAC_CHARSET Macintosh
        // 0x80 SHIFTJIS_CHARSET Japanese
        // 0x81 HANGUL_CHARSET / HANGEUL_CHARSET Hangul (Hangeul) Korean
        // 0x82 JOHAB_CHARSET Johab Korean
        // 0x86 GB2312_CHARSET Simplified Chinese
        // 0x88 CHINESEBIG5_CHARSET Traditional Chinese
        // 0xA1 GREEK_CHARSET Greek
        // 0xA2 TURKISH_CHARSET Turkish
        // 0xA3 VIETNAMESE_CHARSET Vietnamese
        // 0xB1 HEBREW_CHARSET Hebrew
        // 0xB2 ARABIC_CHARSET Arabic
        // 0xBA BALTIC_CHARSET Baltic
        // 0xCC RUSSIAN_CHARSET Russian
        // 0xDE THAI_CHARSET Thai
        // 0xEE EASTEUROPE_CHARSET Eastern European
        // 0xFF OEM_CHARSET OEM code page (based on system locale)
        unused = readbin(unused, bin , 0);

        std::vector<int> color = brtColor(bin);
        // needs handling for rgb/hex colors
        if (debug) Rf_PrintValue(Rcpp::wrap(color));

        bFontScheme = readbin(bFontScheme, bin , 0);
        // 0x00 None
        // 0x01 Major font scheme
        // 0x02 Minor font scheme

        std::string name = XLWideString(bin);

        if (bls == 0x02BC)  out << "<b/>" << std::endl;
        if (fields->fItalic) out << "<i/>" << std::endl;
        if (fields->fStrikeout) out << "<strike/>" << std::endl;
        // if (fields->fOutline) out << "<outline/>" << std::endl;

        if (uls > 0) {
          if (uls == 0x01) out << "<u val=\"single\" />" << std::endl;
          if (uls == 0x02) out << "<u val=\"double\" />" << std::endl;
          if (uls == 0x21) out << "<u val=\"singleAccounting\" />" << std::endl;
          if (uls == 0x22) out << "<u val=\"doubleAccounting\" />" << std::endl;
        }

        out << "<sz val=\"" << dyHeight << "\" />" << std::endl;

        // if (color[0] == 0x01) { // wrong but right?
        //   Rcpp::Rcout << "<color theme=\"" << color[1] << "\" />" << std::endl;
        // }

        if (color[0] == 0x01) {
          out << "<color indexed=\"" << color[1] << "\" />" << std::endl;
        }

        if (color[0] == 0x02) {
          out << "<color rgb=\"" << to_argb(color[6], color[3], color[4], color[5]) << "\" />" << std::endl;
        }

        if (color[0] == 0x03) {
          out << "<color theme=\"" << color[1] << "\" />" << std::endl;
        }

        if (color[1])

        out << "<name val=\"" << name << "\" />" << std::endl;
        if (bFamily > 0)
          out << "<family val=\"" << (uint16_t)bFamily << "\" />" << std::endl;

        if (bFontScheme == 1) {
          out << "<scheme val=\"major\" />" << std::endl;
        }
        if (bFontScheme == 2) {
          out << "<scheme val=\"minor\" />" << std::endl;
        }

        if (bCharSet > 0)
          out << "<charset val=\"" << (uint16_t)bCharSet << "\" />" << std::endl;

        out << "</font>" << std::endl;
        break;
      }

      case BrtEndFonts:
      {
        if (debug) Rcpp::Rcout << "</fonts>" << std::endl;
        bin.seekg(size, bin.cur);

        out << "</fonts>" << std::endl;
        break;
      }


        // fills
      case BrtBeginFills:
      {
        out << "<fills>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtFill:
      {
        uint32_t fls = 0, iGradientType = 0, cNumStop = 0;
        double brtColorFore = 0.0,
          brtColorBack = 0.0,
          xnumDegree = 0.0,
          xnumFillToLeft = 0.0,
          xnumFillToRight = 0.0,
          xnumFillToTop = 0.0,
          xnumFillToBottom = 0.0,
          aPos = 0.0 // not used
        ;
        std::vector<int> aColor; // not used

        fls = readbin(fls, bin, 0);
        std::vector<int> fgColor = brtColor(bin);
        std::vector<int> bgColor = brtColor(bin);
        iGradientType = readbin(iGradientType, bin, 0);
        xnumDegree = Xnum(bin);
        xnumFillToLeft = Xnum(bin);
        xnumFillToRight = Xnum(bin);
        xnumFillToTop = Xnum(bin);
        xnumFillToBottom = Xnum(bin);
        cNumStop = readbin(cNumStop, bin, 0);

        for (uint32_t i = 0; i < cNumStop; ++i) {
          // gradient fill colors and positions
          aColor = brtColor(bin);
          aPos   = Xnum(bin);
        }

        out << "<fill>" << std::endl;

        if (fls != 0x00000028) {
          out << "<patternFill";
          out << " patternType=\"";
          if (fls == 0) out << "none";
          if (fls == 1) out << "solid";
          if (fls == 2) out << "mediumGray";
          if (fls == 3) out << "darkGray";
          if (fls == 4) out << "lightGray";
          if (fls == 5) out << "darkHorizontal";
          if (fls == 6) out << "darkVertical";
          if (fls == 7) out << "darkDown";
          if (fls == 8) out << "darkUp";
          if (fls == 9) out << "darkGrid";
          if (fls == 10) out << "darkTrellis";
          if (fls == 11) out << "lightHorizontal";
          if (fls == 12) out << "lightVertical";
          if (fls == 13) out << "lightDown";
          if (fls == 14) out << "lightUp";
          if (fls == 15) out << "lightGrid";
          if (fls == 16) out << "lightTrellis";
          if (fls == 17) out << "gray125";
          if (fls == 18) out << "gray0625";
          out << "\">" << std::endl;
          // if FF000000 & FFFFFFFF they can be omitted
          out << "<fgColor rgb=\"" << to_argb(fgColor[6], fgColor[3], fgColor[4], fgColor[5]) << std::dec << "\" />" << std::endl;
          out << "<bgColor rgb=\"" << to_argb(bgColor[6], bgColor[3], bgColor[4], bgColor[5]) << std::dec << "\" />" << std::endl;
          out << "</patternFill>" << std::endl;
          out << "</fill>" << std::endl;
        } else {
          Rcpp::Rcout << "gradient fill not implemented" << std::endl;
        }

        break;
      }

      case BrtEndFills:
      {
        out << "</fills>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

        // borders
      case BrtBeginBorders:
      {
        out << "<borders>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtBorder:
      {
        out << "<border>" << std::endl;
        bool fBdrDiagDown = 0, fBdrDiagUp = 0;
        uint8_t borderFlags = 0;
        borderFlags = readbin(borderFlags, bin, 0);

        std::string top      = brtBorder("top", bin);
        std::string bottom   = brtBorder("bottom", bin);
        std::string left     = brtBorder("left", bin);
        std::string right    = brtBorder("right", bin);
        std::string diagonal = brtBorder("diagonal", bin);

        // xml requires a different order a certain spreadsheet software
        // is quite picky in this regard
        out << left << right << top << bottom << diagonal;

        out << "</border>" << std::endl;
        // bin.seekg(size, bin.cur);
        break;
      }

      case BrtEndBorders:
      {
        out << "</borders>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

        // cellStyleXfs
      case BrtBeginCellStyleXFs:
      {
        out << "<cellStyleXfs>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtEndCellStyleXFs:
      {
        out << "</cellStyleXfs>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

        // cellXfs
      case BrtBeginCellXFs:
      {
        out << "<cellXfs>" << std::endl;
        bin.seekg(size, bin.cur);
        // uint32_t cxfs = 0;
        // cxfs = readbin(cxfs, bin, 0);
        break;
      }

      case BrtXF:
      {
        out << "<xf";

        bool is_cell_style_xf = false;

        uint8_t trot = 0, indent = 0;
        uint16_t ixfeParent = 0, iFmt = 0, iFont = 0, iFill = 0, ixBorder = 0;
        uint32_t brtxf = 0;

        ixfeParent = readbin(ixfeParent, bin, 0);
        iFmt = readbin(iFmt, bin, 0);
        iFont = readbin(iFont, bin, 0);
        iFill = readbin(iFill, bin, 0);
        ixBorder = readbin(ixBorder, bin, 0);
        trot = readbin(trot, bin, 0);
        indent = readbin(indent, bin, 0);

        if (indent > 250) {
          Rcpp::stop("indent to big");
        }

        if (ixfeParent == 0xFFFF) {
          is_cell_style_xf = true;
        } else {
          out << " xfId=\"" << ixfeParent << "\"";
        }

        brtxf = readbin(brtxf, bin, 0);
        XFFields *fields = (XFFields *)&brtxf;
        uint8_t xfgbit = fields->xfGrbitAtr;

        xfGrbitAtrFields *xfGrbitAtr = (xfGrbitAtrFields *)&xfgbit;

        out << " numFmtId=\"" << iFmt << "\"";
          out << " fontId=\"" << iFont << "\"";
          out << " fillId=\"" << iFill << "\"";
          out << " borderId=\"" << ixBorder << "\"";

          // if (!is_cell_style_xf) {

            // if (iFmt > 0) {
            out << " applyNumberFormat=\""<< !xfGrbitAtr->bit1 <<"\"";
            // }
            // if (iFont > 0) {
            out << " applyFont=\""<< !xfGrbitAtr->bit2 << "\"";
            // }
            // if (ixBorder > 0) {
            out << " applyBorder=\"" << !xfGrbitAtr->bit4 << "\"";
            // }
            // if (iFill > 0) {
            out << " applyFill=\"" << !xfGrbitAtr->bit5 << "\"";
            // }

            out << " applyAlignment=\""<< !xfGrbitAtr->bit3 << "\"";

            out << " applyProtection=\""<< !xfGrbitAtr->bit6 << "\"";

          // } else {
          //
          //   out << " applyNumberFormat=\""<< "0" <<"\"";
          //
          //   out << " applyFont=\""<< "0" << "\"";
          //
          //   out << " applyBorder=\"" << "0" << "\"";
          //
          //   out << " applyFill=\"" << "0" << "\"";
          //
          //   out << " applyAlignment=\""<< "0" << "\"";
          //
          //   out << " applyProtection=\""<< "0" << "\"";
          // }


          if (fields->alc > 0 || fields->alcv > 0 || indent || trot || fields->iReadingOrder || fields->fShrinkToFit || fields->fWrap || fields->fJustLast) {
            out << "><alignment";
            out << halign(fields->alc);
            out << valign(fields->alcv);

            if (fields->fWrap)
              out << " wrapText=\"" << fields->fWrap <<"\"";
            if (fields->fShrinkToFit)
              out << " shrinkToFit=\"" << fields->fWrap <<"\"";
            // something is not quite right here.
            // I have a file which is left to right, but this here returns 2: right to left
            if (fields->iReadingOrder)
              out << " readingOrder=\"" << (uint16_t)fields->iReadingOrder <<"\"";
            if (indent)
              out << " indent=\"" << (uint16_t)indent/3 <<"\"";
            if (fields->fJustLast)
              out << " justifyLastLine=\"" << fields->fJustLast <<"\"";
            if (trot)
              out << " textRotation=\"" << trot <<"\"";

            out << "/>";
            out << "</xf>" << std::endl;
          } else {
            out << " />" << std::endl;
          }

          // bin.seekg(size, bin.cur);
          break;
      }

      case BrtEndCellXFs:
      {
        out << "</cellXfs>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      // cellStyles
      case BrtBeginStyles:
      {
        out << "<cellStyles>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtStyle:
      {

        uint8_t iStyBuiltIn = 0, iLevel = 0;
        uint16_t grbitObj1 = 0;
        uint32_t ixf = 0;
        ixf = readbin(ixf, bin, 0);
        grbitObj1 = readbin(grbitObj1, bin, 0);
        iStyBuiltIn = readbin(iStyBuiltIn, bin, 0);
        iLevel = readbin(iLevel, bin, 0);
        std::string stName = XLWideString(bin);

        StyleFlagsFields *fields = (StyleFlagsFields *)&grbitObj1;
        out << "<cellStyle";
        out << " name=\"" << stName << "\"";
        out << " xfId=\"" << ixf << "\"";
        out << " builtinId=\"" << (int32_t)iStyBuiltIn << "\"";
        out << " />" << std::endl;


        // bin.seekg(size, bin.cur);
        break;
      }

      case BrtEndStyles:
      {
        out << "</cellStyles>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtBeginColorPalette:
      {
        out << "<colors>" <<std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtBeginIndexedColors:
      {
        out << "<indexedColors>" <<std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtIndexedColor:
      {
        uint8_t bRed = 0, bGreen = 0, bBlue = 0, reserved = 0;
        bRed = readbin(bRed, bin, 0);
        bGreen = readbin(bGreen, bin, 0);
        bBlue = readbin(bBlue, bin, 0);
        reserved = readbin(reserved, bin, 0);

        // std::vector<int> color = brtColor(bin);
        // if (color[0] == 0x02) {
        out << "<rgbColor rgb=\"" <<
          to_argb(reserved, bRed, bGreen, bRed) <<
            std::dec << "\" />" << std::endl;
        // }
        break;
      }

      case BrtEndIndexedColors:
      {
        out << "</indexedColors>" <<std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtEndColorPalette:
      {
        out << "</colors>" <<std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtBeginDXFs:
      {
        uint32_t cdxfs = 0;
        cdxfs = readbin(cdxfs, bin, 0);

        out << "<dxfs>" << std::endl;
        break;
      }

        // This is a todo. the xfPropDataBlob is almost the entire styles part again
        // case BrtDXF:
        // {
        //   uint16_t flags = 0, reserved = 0, cprops = 0;
        //   flags = readbin(flags, bin, 0);
        //   reserved = readbin(reserved, bin, 0);
        //   cprops = readbin(cprops, bin, 0);
        //
        //   out << "<dxfs>" << std::endl;
        //  break;
        // }

      case BrtEndDXFs:
      {
        out << "</dxfs>" << std::endl;
        break;
      }

      case BrtEndStyleSheet:
      {
        end_of_style_sheet = true;
        out << "</styleSheet>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      default:
      {
        if (debug) {
        Rcpp::Rcout << "Unhandled Style: " << std::to_string(x) <<
          ": " << std::to_string(size) <<
            " @ " << bin.tellg() << std::endl;
      }
        bin.seekg(size, bin.cur);
        break;
      }

      }

    }

    out.close();
    bin.close();
    return 1;
  } else {
    return -1;
  };

}

// [[Rcpp::export]]
int bin_table(std::string filePath, std::string outPath, bool debug) {

  std::ofstream out(outPath);
  std::ifstream bin(filePath, std::ios::in | std::ios::binary | std::ios::ate);

  // auto sas_size = bin.tellg();
  if (bin) {
    bin.seekg(0, std::ios_base::beg);
    bool end_of_table = false;

    while(!end_of_table) {
      Rcpp::checkUserInterrupt();

      int32_t x = 0, size = 0;

      if (debug) Rcpp::Rcout << "." << std::endl;
      RECORD(x, size, bin);
      if (debug) Rcpp::Rcout << x << ": " << size << std::endl;

      switch(x) {

      case BrtUID:
      {
        if (debug) Rcpp::Rcout << "<xr:uid>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtBeginList:
      {
        if (debug) Rcpp::Rcout << "<table>" << std::endl;

        std::vector<int> rfxList = UncheckedRfX(bin);
        std::string ref = int_to_col(rfxList[2] + 1) + std::to_string(rfxList[0] + 1) + ":" + int_to_col(rfxList[3] + 1) + std::to_string(rfxList[1] + 1);

        Rcpp::Rcout << "table ref: " << ref << std::endl;

        uint32_t lt = 0, idList = 0, crwHeader = 0, crwTotals = 0, flags = 0;
        lt = readbin(lt, bin, 0);
        idList = readbin(idList, bin, 0);
        crwHeader = readbin(crwHeader, bin, 0);
        crwTotals = readbin(crwTotals, bin, 0);
        flags = readbin(flags, bin, 0);

        uint32_t nDxfHeader = 0, nDxfData = 0, nDxfAgg = 0, nDxfBorder = 0, nDxfHeaderBorder = 0, nDxfAggBorder = 0, dwConnID = 0;

        nDxfHeader = readbin(nDxfHeader, bin, 0);
        nDxfData = readbin(nDxfData, bin, 0);
        nDxfAgg = readbin(nDxfAgg, bin, 0);
        nDxfBorder = readbin(nDxfBorder, bin, 0);
        nDxfHeaderBorder = readbin(nDxfHeaderBorder, bin, 0);
        nDxfAggBorder = readbin(nDxfAggBorder, bin, 0);
        dwConnID = readbin(dwConnID, bin, 0);

        std::string stName = XLNullableWideString(bin);
        Rcpp::Rcout << stName << std::endl;
        std::string stDisplayName = XLNullableWideString(bin);
        Rcpp::Rcout << stDisplayName << std::endl;
        std::string stComment = XLNullableWideString(bin);
        Rcpp::Rcout << stComment << std::endl;
        std::string stStyleHeader = XLNullableWideString(bin);
        Rcpp::Rcout << stStyleHeader << std::endl;
        std::string stStyleData = XLNullableWideString(bin);
        Rcpp::Rcout << stStyleData << std::endl;
        std::string stStyleAgg = XLNullableWideString(bin);
        Rcpp::Rcout << stStyleAgg << std::endl;

        out << "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\" mc:Ignorable=\"xr xr3\"" <<
          " id=\"" << idList <<
          "\" name=\"" << stName <<
            "\" displayName=\"" << stDisplayName <<
              "\" ref=\"" << ref <<
                "\" totalsRowShown=\"" << crwTotals <<
              "\">" << std::endl;

        Rcpp::Rcout << bin.tellg() << std::endl;
        break;
      }

      // this is used in worksheet as well. move it to function?
      case BrtBeginAFilter:
      {
        if (debug) Rcpp::Rcout << "<autofilter>" << std::endl;
        std::vector<int> rfx = UncheckedRfX(bin);

        std::string lref = int_to_col(rfx[2] + 1) + std::to_string(rfx[0] + 1);
        std::string rref = int_to_col(rfx[3] + 1) + std::to_string(rfx[1] + 1);

        std::string ref;
        if (lref.compare(rref) == 0) {
          ref = lref;
        } else {
          ref =  lref + ":" + rref;
        }

        // ignoring filterColumn for now
        // autofilter can consist of filterColumn, filters and filter
        // maybe customFilter, dynamicFilter too
        out << "<autoFilter ref=\"" << ref << "\">" << std::endl;

        break;
      }

      case BrtBeginFilterColumn:
      {
        if (debug) Rcpp::Rcout << "<filterColumn>" << std::endl;

        uint16_t flags = 0;
        uint32_t dwCol = 0;
        dwCol = readbin(dwCol, bin, 0);
        flags = readbin(flags, bin, 0);
        // fHideArrow
        // fNoBtn

        out << "<filterColumn colId=\"" << dwCol <<"\">" << std::endl;

        break;
      }

      case BrtBeginFilters:
      {
        if (debug) Rcpp::Rcout << "<filters>" << std::endl;
        // bin.seekg(size, bin.cur);

        uint32_t fBlank = 0, unused = 0;
        fBlank = readbin(fBlank, bin, 0); // a 32bit flag, after all ... why not?
        unused = readbin(unused, bin, 0); // maybe calendarType?
        out << "<filters blank=\"" << fBlank <<"\">" << std::endl;
        break;
      }

      case BrtFilter:
      {
        std::string rgch = XLWideString(bin);
        out << "<filter val=\"" << rgch << "\" />" << std::endl;
        // bin.seekg(size, bin.cur);
        break;
      }

      case BrtEndFilters:
      {
        if (debug) Rcpp::Rcout << "</filters>" << std::endl;
        out << "</filters>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtColorFilter:
      {
        uint32_t dxfid = 0, fCellColor = 0;
        dxfid = readbin(dxfid, bin, 0);
        fCellColor = readbin(fCellColor, bin, 0);

        out << "<colorFilter dxfId=\"" << dxfid << "\" cellColor=\""<< fCellColor<< "\"/>" << std::endl;

        break;
      }

      case BrtBeginCustomFilters:
      case BrtBeginCustomFilters14:
      case BrtBeginCustomRichFilters:
      {
        Rcpp::warning("Custom Filter found. This is not handled.");
        bin.seekg(size, bin.cur);

        break;
      }

      case BrtDynamicFilter:
      {
        Rcpp::warning("Dynamic Filter found. This is not handled.");
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtIconFilter:
      case BrtIconFilter14:
      {
        uint32_t iIconSet = 0, iIcon = 0;
        iIconSet = readbin(iIconSet, bin, 0);
        iIcon = readbin(iIcon, bin, 0);

        std::string iconSet;
        if (iIconSet) iconSet = to_iconset(iIconSet);

        out << "<iconFilter iconSet=\"" << iconSet << "\" iconId=\""<< iIcon<< "\"/>" << std::endl;

        break;
      }

      case BrtEndFilterColumn:
      {
        if (debug) Rcpp::Rcout << "</filterColumn>" << std::endl;
        out << "</filterColumn>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtEndAFilter:
      {
        if (debug) Rcpp::Rcout << "</autofilter>" << std::endl;
        out << "</autoFilter>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtBeginListCols:
      {
        uint32_t nCols = 0;
        nCols = readbin(nCols, bin, 0);
        out << "<tableColumns count=\"" << nCols << "\">" << std::endl;
        // bin.seekg(size, bin.cur);
        break;
      }

      case BrtBeginListCol:
      {
        if (debug) Rcpp::Rcout << "<tableColumn>" << std::endl;
        // bin.seekg(size, bin.cur);
        // break;

        uint32_t idField = 0, ilta = 0, nDxfHdr = 0, nDxfInsertRow = 0, nDxfAgg = 0, idqsif = 0;

        idField = readbin(idField, bin, 0);
        ilta = readbin(ilta, bin, 0);
        nDxfHdr = readbin(nDxfHdr, bin, 0);
        nDxfInsertRow = readbin(nDxfInsertRow, bin, 0);
        nDxfAgg = readbin(nDxfAgg, bin, 0);
        idqsif = readbin(idqsif, bin, 0);
        std::string stName = XLNullableWideString(bin);
        std::string stCaption = XLNullableWideString(bin);
        std::string stTotal = XLNullableWideString(bin);
        Rcpp::Rcout << stName << ": " << stCaption << ": " << stTotal<< std::endl;
        std::string stStyleHeader = XLNullableWideString(bin);
        std::string stStyleInsertRow = XLNullableWideString(bin);
        std::string stStyleAgg = XLNullableWideString(bin);
        Rcpp::Rcout << stStyleHeader << ": " << stStyleInsertRow << ": " << stStyleAgg << std::endl;

        out << "<tableColumn id=\"" << idField << "\" name=\"" << stCaption << "\">" << std::endl;

        break;
      }

      case BrtEndListCol:
      {
        out << "</tableColumn>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtEndListCols:
      {
        out << "</tableColumns>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtEndList:
      {
        end_of_table = true;
        out << "</table>" << std::endl;
        break;
      }

      case BrtListCCFmla:
      {
        // calculated column formula
        // uint8_t flags = 0;
        // std::string fml;
        // flags = readbin(flags, bin, 0);
        // fml = CellParsedFormula(bin);
        Rcpp::warning("Table formulas are not implemented.");
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtTableStyleClient:
      {
        uint16_t flags = 0;
        flags = readbin(flags, bin, 0);
        std::string stStyleName = XLNullableWideString(bin);

        BrtTableStyleClientFields *fields = (BrtTableStyleClientFields*)&flags;

        out << "<tableStyleInfo name=\""<< stStyleName << "\"";
        out <<" showColHeaders=\"" << fields->fColumnHeaders << "\"";
        out <<" showFirstColumn=\"" << fields->fFirstColumn << "\"";
        out <<" showLastColumn=\"" << fields->fLastColumn << "\"";
        out <<" showRowHeaders=\"" << fields->fRowHeaders << "\"";
        out <<" showRowStripes=\"" << fields->fRowStripes << "\"";
        out <<" showColumnStripes=\"" << fields->fColumnStripes << "\"";
        out << " />" << std::endl;
        break;
      }

      default:
      {
        if (debug) {
        Rcpp::Rcout << std::to_string(x) <<
          ": " << std::to_string(size) <<
            " @ " << bin.tellg() << std::endl;
      }
        bin.seekg(size, bin.cur);
        break;
      }
      }
    }

    out.close();
    bin.close();
    return 1;
  } else {
    return -1;
  };

}

// [[Rcpp::export]]
int sst(std::string filePath, std::string outPath, bool debug) {

  std::ofstream out(outPath);
  std::ifstream bin(filePath, std::ios::in | std::ios::binary | std::ios::ate);

  // auto sas_size = bin.tellg();
  if (bin) {
    bin.seekg(0, std::ios_base::beg);
    bool end_of_shared_strings = false;

    while(!end_of_shared_strings) {
      Rcpp::checkUserInterrupt();

      int32_t x = 0, size = 0;

      if (debug) Rcpp::Rcout << "." << std::endl;
      RECORD(x, size, bin);
      if (debug) Rcpp::Rcout << x << ": " << size << std::endl;

      switch(x) {

      case BrtBeginSst:
      {
        uint32_t count= 0, uniqueCount= 0;
        count = readbin(count, bin, 0);
        uniqueCount = readbin(uniqueCount, bin, 0);
        out << "<sst " <<
          "count=\"" << count <<
            "\" uniqueCount=\"" << uniqueCount <<
              "\">" << std::endl;
        break;
      }

      case BrtSSTItem:
      {
        // Rcpp::Rcout << bin.tellg() << std::endl;
        std::string val = RichStr(bin);
        if (debug)
          Rcpp::Rcout << val << std::endl;
        out << "<si><t>" << escape_xml(val) <<
          "</t></si>" << std::endl;
        break;
      }

      case BrtEndSst:
      {
        end_of_shared_strings = true;
        out << "</sst>" << std::endl;
        break;
      }

      default:
      {
        if (debug) {
        Rcpp::Rcout << std::to_string(x) <<
          ": " << std::to_string(size) <<
            " @ " << bin.tellg() << std::endl;
      }
        bin.seekg(size, bin.cur);
        break;
      }
      }
    }

    out.close();
    bin.close();
    return 1;
  } else {
    return -1;
  };

}

// [[Rcpp::export]]
int workbook(std::string filePath, std::string outPath, bool debug) {

  // WORKBOOK = BrtBeginBook [BrtFileVersion] [[BrtFileSharingIso] BrtFileSharing] [BrtWbProp]
  // [ACABSPATH] [ACREVISIONPTR] [[BrtBookProtectionIso] BrtBookProtection] [BOOKVIEWS]
  // BUNDLESHS [FNGROUP] [EXTERNALS] *BrtName [BrtCalcProp] [BrtOleSize] *(BrtUserBookView
  // *FRT) [PIVOTCACHEIDS] [BrtWbFactoid] [SMARTTAGTYPES] [BrtWebOpt] *BrtFileRecover
  // [WEBPUBITEMS] [CRERRS] FRTWORKBOOK BrtEndBook

  std::ofstream out(outPath);
  std::ifstream bin(filePath, std::ios::in | std::ios::binary | std::ios::ate);

  // auto sas_size = bin.tellg();
  if (bin) {
    bin.seekg(0, std::ios_base::beg);
    bool end_of_workbook = false;

    std::vector<std::string> defNams;
    defNams.push_back("<definedNames>");

    while(!end_of_workbook) {
      Rcpp::checkUserInterrupt();

      int32_t x = 0, size = 0;

      if (debug) Rcpp::Rcout << "." << std::endl;
      RECORD(x, size, bin);

      switch(x) {

      case BrtBeginBook:
      {
        if (debug) Rcpp::Rcout << "<workbook>" << std::endl;
        out << "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr6=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision6\" xmlns:xr10=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision10\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\" mc:Ignorable=\"x15 xr xr6 xr10 xr2\">" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtFileVersion:
      {
        if (debug) Rcpp::Rcout << "<fileVersion>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtWbProp:
      {
        if (debug) Rcpp::Rcout << "<workbookProperties>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtACBegin:
      {
        if (debug) Rcpp::Rcout << "<BrtACBegin>" << std::endl;
        // bin.seekg(size, bin.cur);

        uint16_t cver = 0;
        cver = readbin(cver, bin, 0);

        for (uint16_t i = 0; i < cver; ++i) {
          ProductVersion(bin, debug);
        }
        break;

      }

      case BrtAbsPath15:
      {
        std::string absPath = XLWideString(bin);
        if (debug) Rcpp::Rcout << absPath << std::endl;
        break;
      }

      case BrtACEnd:
      {
        if (debug) Rcpp::Rcout << "<BrtACEnd>" << std::endl;
        break;
      }

      case BrtRevisionPtr:
      {
        if (debug) Rcpp::Rcout << "<BrtRevisionPtr>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtUID:
      {
        if (debug) Rcpp::Rcout << "<BrtUID>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtBeginBookViews:
      {
        if (debug) Rcpp::Rcout << "<workbookViews>" << std::endl;
        out << "<bookViews>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtBookView:
      {
        if (debug) Rcpp::Rcout << "<workbookView>" << std::endl;
        // bin.seekg(size, bin.cur);
        uint8_t flags = 0;
        int32_t xWn = 0, yWn = 0;
        uint32_t dxWn = 0, dyWn = 0, iTabRatio = 0, itabFirst= 0, itabCur = 0;
        xWn       = readbin(xWn, bin, 0);
        yWn       = readbin(yWn, bin, 0);
        dxWn      = readbin(dxWn, bin, 0);
        dyWn      = readbin(dyWn, bin, 0);
        iTabRatio = readbin(iTabRatio, bin, 0);
        itabFirst = readbin(itabFirst, bin, 0);
        itabCur   = readbin(itabCur, bin, 0);
        flags     = readbin(flags, bin, 0);

        out << "<workbookView xWindow=\"" << xWn << "\" yWindow=\"" << yWn <<
          "\" windowWidth=\"" << dxWn<<"\" windowHeight=\"" << dyWn <<
            "\" activeTab=\"" <<  itabCur << "\" />" << std::endl;
        break;
      }

      case BrtEndBookViews:
      {
        if (debug) Rcpp::Rcout << "</workbookViews>" << std::endl;
        out << "</bookViews>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtBeginBundleShs:
      {
        if (debug) Rcpp::Rcout << "<sheets>" << std::endl;
        // unk = readbin(unk, bin, 0);
        // unk = readbin(unk, bin, 0);
        // uint32_t count, uniqueCount;
        // count = readbin(count, bin, 0);
        // uniqueCount = readbin(uniqueCount, bin, 0);
        out << "<sheets>" << std::endl;
        break;
      }

      case BrtBundleSh:
      {
        if (debug) Rcpp::Rcout << "<sheet>" << std::endl;

        uint32_t hsState = 0, iTabID = 0; //  strRelID ???

        hsState = readbin(hsState, bin, 0);
        iTabID = readbin(iTabID, bin, 0);
        std::string rid = XLNullableWideString(bin);

        if (debug)
          Rcpp::Rcout << "sheet vis: " << hsState << ": " << iTabID  << ": " << rid << std::endl;

        std::string val = XLWideString(bin);

        std::string visible;
        if (hsState == 0) visible = "visible";
        if (hsState == 1) visible = "hidden";
        if (hsState == 2) visible = "veryHidden";

        out << "<sheet r:id=\"" << rid << "\" state=\"" << visible<< "\" sheetId=\"" << iTabID<< "\" name=\"" << val << "\"/>" << std::endl;
        break;
      }

      case BrtEndBundleShs:
      {
        if (debug) Rcpp::Rcout << "</sheets>" << std::endl;
        out << "</sheets>" << std::endl;
        break;
      }

      case BrtCalcProp:
      {
        if (debug) Rcpp::Rcout << "<calcPr>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtName:
      {
        if (debug)  Rcpp::Rcout << "<BrtName>" << std::endl;
        // bin.seekg(size, bin.cur);

        uint8_t chKey = 0;
        uint32_t itab = 0, BrtNameUint = 0;
        BrtNameUint = readbin(BrtNameUint, bin, 0);

        BrtNameFields *fields = (BrtNameFields *)&BrtNameUint;

        // fHidden    - visible
        // fFunc      - xml macro
        // fOB        - vba macro
        // fProc      - represents a macro
        // fCalcExp   - rgce returns array
        // fBuiltin   - builtin name
        // fgrp       - FnGroupID if fProc == 0 then 0
        // fPublished - published name
        // fWorkbookParam  - DefinedName is workbook paramenter
        // fFutureFunction - future function
        // reserved        - 0

        // if (debug)
        //   Rprintf(
        //     "%d, %d, %d, %d, %d, %d, %d, %d, %d, %d, %d\n",
        //     fields->fHidden,
        //     fields->fFunc,
        //     fields->fOB,
        //     fields->fProc,
        //     fields->fCalcExp,
        //     fields->fBuiltin,
        //     fields->fgrp,
        //     fields->fPublished,
        //     fields->fWorkbookParam,
        //     fields->fFutureFunction,
        //     fields->reserved
        //   );

        chKey = readbin(chKey, bin, 0);
        // ascii key (0 if fFunc = 1 or fProc = 0 else >= 0x20)

        itab = readbin(itab, bin, 0);

        // XLNameWideString: XLWideString <= 255 characters
        std::string name = XLWideString(bin);

        std::string fml = CellParsedFormula(bin, debug, 0);

        std::string comment = XLNullableWideString(bin);

        // if (debug)
          Rcpp::Rcout << name << comment << std::endl;

        if (fields->fProc) {
          // must be NULL
          std::string unusedstring1 = XLNullableWideString(bin);
          // must be < 32768 characters
          std::string description = XLNullableWideString(bin);
          std::string helpTopic = XLNullableWideString(bin);
          // must be NULL
          std::string unusedstring2 = XLNullableWideString(bin);
        }

        std::string defNam = "<definedName name=\"" + name;

        if (comment.size() > 0)
          defNam += "\" comment=\"" + comment;

        if (itab != 0xFFFFFFFF)
          defNam += "\" localSheetId=\"" + std::to_string(itab);

        if (fields->fHidden)
          defNam += "\" hidden=\"" + std::to_string(fields->fHidden);

        // lacks the formula for the defined name
        defNam = defNam + "\">" + fml +"</definedName>";

        defNams.push_back(defNam);

        break;
      }

      case BrtBeginExternals:
      {
        if (debug) Rcpp::Rcout << "<BrtBeginExternals>" << std::endl;
        // bin.seekg(size, bin.cur);
        break;
      }

        // part of Xti
        // * BrtSupSelf self reference
        // * BrtSupSame same sheet
        // * BrtSupAddin addin XLL
        // * BrtSupBookSrc external link
      case BrtSupSelf:
      {
        if (debug) Rcpp::Rcout << "BrtSupSelf @"<< bin.tellg() << std::endl;
        // Rcpp::stop("BrtSupSelf");
        break;
      }

      case BrtSupSame:
      {
        if (debug) Rcpp::Rcout << "BrtSupSame @"<< bin.tellg() << std::endl;
        // Rcpp::stop("BrtSupSelf");
        break;
      }

      case BrtPlaceholderName:
      {
        std::string name = XLWideString(bin);
        Rcpp::Rcout << name << std::endl;
        break;
      }

      case BrtSupAddin:
      {
        if (debug) Rcpp::Rcout << "BrtSupAddin @"<< bin.tellg() << std::endl;
        // Rcpp::stop("BrtSupSelf");
        break;
      }

      case BrtSupBookSrc:
      {
        if (debug) Rcpp::Rcout << "BrtSupBookSrc @"<< bin.tellg() << std::endl;
        // Rcpp::stop("BrtSupSelf");
        std::string strRelID = XLNullableWideString(bin);
        Rcpp::Rcout << strRelID << std::endl;
        break;
      }

      case BrtSupTabs:
      {
        if (debug) Rcpp::Rcout << "<BrtSupTabs>" << std::endl;
        uint32_t cTab = 0;

        cTab = readbin(cTab, bin, 0);
        if (cTab > 65535) Rcpp::stop("cTab to large");

        for (uint32_t i = 0; i < cTab; ++i) {
          std::string sheetName = XLWideString(bin);
          Rcpp::Rcout << sheetName << std::endl;
        }

        break;
      }

      case BrtExternSheet:
      {
        if (debug) Rcpp::Rcout << "<BrtExternSheet>" << std::endl;
        uint32_t cXti = 0;

        cXti = readbin(cXti, bin, 0);
        if (cXti > 65535) Rcpp::stop("cXti to large");

        std::vector<uint32_t> rgXti(cXti);
        for (uint32_t i = 0; i < cXti; ++i) {
          std::vector<int> xti = Xti(bin);
          Rf_PrintValue(Rcpp::wrap(xti));
        }

        break;
      }

      case BrtEndExternals:
      {
        if (debug) Rcpp::Rcout << "</BrtEndExternals>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtFRTBegin:
      {
        if (debug) Rcpp::Rcout << "<ext>" << std::endl;

        ProductVersion(bin, debug);
        break;
      }

      case BrtWorkBookPr15:
      {
        if (debug) Rcpp::Rcout << "<BrtWorkBookPr15>" << std::endl;
        bin.seekg(size, bin.cur);
        // uint8_t fChartTrackingRefBased = 0;
        // uint32_t FRTHeader = 0;
        // FRTHeader = readbin(FRTHeader, bin, 0);
        // fChartTrackingRefBased = readbin(fChartTrackingRefBased, bin, 0) & 0x01;
        break;
      }

      case BrtBeginCalcFeatures:
      {
        if (debug) Rcpp::Rcout << "<calcs>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtCalcFeature:
      {
        if (debug) Rcpp::Rcout << "<calc>" << std::endl;
        // bin.seekg(size, bin.cur);
        uint32_t FRTHeader = 0;
        FRTHeader = readbin(FRTHeader, bin, 0);
        std::string szName = XLWideString(bin);

        if (debug) Rcpp::Rcout << FRTHeader << ": " << szName << std::endl;
        break;
      }

      case BrtEndCalcFeatures:
      {
        if (debug) Rcpp::Rcout << "</calcs>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtFRTEnd:
      {
        if (debug) Rcpp::Rcout << "</ext>" << std::endl;
        break;
      }

      case BrtWbFactoid:
      { // ???
        if (debug) Rcpp::Rcout << "<BrtWbFactoid>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtFileRecover:
      {
        if (debug) Rcpp::Rcout << "<fileRecovery>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtEndBook:
      {
        end_of_workbook = true;

        if (defNams.size() > 1) {
          defNams.push_back("</definedNames>");

          for (size_t i = 0; i < defNams.size(); ++i) {
            // if (debug)
              Rcpp::Rcout << defNams[i] << std::endl;
            out << defNams[i] << std::endl;
          }
        }

        if (debug) Rcpp::Rcout << "</workbook>" << std::endl;
        out << "</workbook>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      default:
      {
        if (debug) {
        Rcpp::Rcout << std::to_string(x) <<
          ": " << std::to_string(size) <<
            " @ " << bin.tellg() << std::endl;
      }
        bin.seekg(size, bin.cur);
        break;
      }
      }

      if (debug) Rcpp::Rcout << "wb-loop: " << x << ": " << size << ": " << bin.tellg() << std::endl;
    }

    out.close();
    bin.close();
    return 1;
  } else {
    return -1;
  };

}



// [[Rcpp::export]]
int worksheet(std::string filePath, bool chartsheet, std::string outPath, bool debug) {

  // WORKBOOK = BrtBeginBook [BrtFileVersion] [[BrtFileSharingIso] BrtFileSharing] [BrtWbProp]
  // [ACABSPATH] [ACREVISIONPTR] [[BrtBookProtectionIso] BrtBookProtection] [BOOKVIEWS]
  // BUNDLESHS [FNGROUP] [EXTERNALS] *BrtName [BrtCalcProp] [BrtOleSize] *(BrtUserBookView
  // *FRT) [PIVOTCACHEIDS] [BrtWbFactoid] [SMARTTAGTYPES] [BrtWebOpt] *BrtFileRecover
  // [WEBPUBITEMS] [CRERRS] FRTWORKBOOK BrtEndBook

  std::ofstream out(outPath);
  std::ifstream bin(filePath, std::ios::in | std::ios::binary | std::ios::ate);

  // auto sas_size = bin.tellg();
  if (bin) {
    bin.seekg(0, std::ios_base::beg);

    bool first_row = true;

    bool end_of_worksheet = false;

    uint64_t row = 0;
    uint32_t col = 0;
    std::vector<std::string> hlinks;
    hlinks.push_back("<hyperlinks>");

    // auto itr = 0;
    while(!end_of_worksheet) {
      Rcpp::checkUserInterrupt();

      // uint8_t unk = 0, high = 0, low = 0;
      // uint16_t tmp = 0;
      int32_t x = 0, size = 0;

      if (debug) Rcpp::Rcout << "." << std::endl;

      RECORD(x, size, bin);

      switch(x) {

      case BrtBeginSheet:
      {
        if (chartsheet)
          out << "<chartsheet>" << std::endl;
        else
          out << "<worksheet>" << std::endl;

        if (debug) Rcpp::Rcout << "Begin of <worksheet>: " << bin.tellg() << std::endl;
        break;
      }

      case BrtWsProp:
      {
        if (debug) Rcpp::Rcout << "WsProp: " << bin.tellg() << std::endl;

        // read uint24_t as 16 + 8
        uint8_t brtwsprop_sec = 0;
        uint16_t brtwsprop_fst = 0;
        uint32_t rwSync = 0, colSync = 0;

        brtwsprop_fst = readbin(brtwsprop_fst, bin, 0);
        BrtWsPropFields1 *fields1 = (BrtWsPropFields1 *)&brtwsprop_fst;
        brtwsprop_sec = readbin(brtwsprop_sec, bin, 0);
        BrtWsPropFields2 *fields2 = (BrtWsPropFields2 *)&brtwsprop_sec;

        std::vector<int> color = brtColor(bin);
        rwSync = readbin(rwSync, bin, 0);
        colSync = readbin(colSync, bin, 0);
        std::string strName = XLWideString(bin);

        if (debug) {
          Rcpp::Rcout << "sheetPr tabColor:" << std::endl;
          Rf_PrintValue(Rcpp::wrap(color));
        }

        // for now we only handle color in sheetPR
        if (color[0] >= 1 && color[0] <= 3) {
          out << "<sheetPr>" << std::endl;

          if (color[0] == 0x01) {
            out << "<tabColor indexed=\"" << color[1] << "\" />" << std::endl;
          }

          if (color[0] == 0x02) {
            out << "<tabColor rgb=\"" << to_argb(color[6], color[3], color[4], color[5]) << "\" />" << std::endl;
          }

          if (color[0] == 0x03) {
            out << "<tabColor theme=\"" << color[1] << "\" />" << std::endl; //  << "\" tint=\""<< color[2]
          }

          out << "</sheetPr>" << std::endl;
        }

        break;
      }

      case BrtWsDim:
      {
        if (debug) Rcpp::Rcout << "WsDim: " << bin.tellg() << std::endl;

        // 16 bit
        std::vector<int> dims;
        // 0 index vectors
        // first row, last row, first col, last col
        dims = UncheckedRfX(bin);
        if (debug) {
          Rcpp::Rcout << "dimension: ";
          Rf_PrintValue(Rcpp::wrap(dims));
        }

        out << "<dimension ref=\"" <<
          int_to_col(dims[2] + 1) << dims[0] + 1 <<
            ":" <<
              int_to_col(dims[3] + 1) << dims[1] + 1 <<
                "\"/>" << std::endl;

        break;
      }

      case BrtBeginCsViews:
      case BrtBeginWsViews:
      {
        if (debug) Rcpp::Rcout << "<sheetViews>: " << bin.tellg() << std::endl;
        out << "<sheetViews>" << std::endl;
        bin.seekg(size, bin.cur);

        break;
      }

      case BrtBeginCsView:
      {
        if (debug) Rcpp::Rcout << "<sheetView>: " << bin.tellg() << std::endl;
        bool fSelected = false;
        uint16_t flags = 0;
        uint32_t wScale = 0, iWbkView = 0;

        flags = readbin(flags, bin, 0);
        wScale = readbin(wScale, bin, 0);
        iWbkView = readbin(iWbkView, bin, 0);

        fSelected = flags & 0x8000;

        out << "<sheetView";
        // careful, without the following nothing goes
        out << " workbookViewId=\"" << iWbkView << "\"";
        if (wScale)
          out << " zoomScale=\"" << wScale << "\"";
        if (fSelected)
          out << " tabSelected=\"" << fSelected << "\"";
        out << " zoomToFit=\"1\"";

        out << ">" << std::endl;

        break;
      }

      case BrtBeginWsView:
      {
        if (debug) Rcpp::Rcout << "<sheetView>: " << bin.tellg() << std::endl;
        // out << "<sheetView>" << std::endl;
        // bin.seekg(size, bin.cur);

        uint8_t icvHdr = 0, reserved2 = 0;
        uint16_t flags = 0, reserved3 = 0, wScale = 0, wScaleNormal = 0, wScaleSLV = 0, wScalePLV = 0;
        uint32_t xlView = 0, rwTop = 0, colLeft = 0, iWbkView = 0;

        flags = readbin(flags, bin, 0);
        xlView = readbin(xlView, bin, 0);
        rwTop = UncheckedRw(bin);
        colLeft = UncheckedCol(bin);
        icvHdr = readbin(icvHdr, bin, 0);
        reserved2 = readbin(reserved2, bin, 0);
        reserved3 = readbin(reserved3, bin, 0);
        wScale = readbin(wScale, bin, 0);
        wScaleNormal = readbin(wScaleNormal, bin, 0);
        wScaleSLV = readbin(wScaleSLV, bin, 0);
        wScalePLV = readbin(wScalePLV, bin, 0);
        iWbkView = readbin(iWbkView, bin, 0);


        BrtBeginWsViewFields *fields = (BrtBeginWsViewFields *)&flags;

        out << "<sheetView";
        // careful, without the following nothing goes
        out << " workbookViewId=\"" << iWbkView << "\"";
        if (icvHdr)
          out << " colorId=\"" << (int32_t)icvHdr << "\"";
        if (!fields->fDefaultHdr)
          out << " defaultGridColor=\"" << fields->fDefaultHdr << "\"";
        if (fields->fRightToLeft) {
          Rcpp::Rcout << "rightToLeft=\"" << fields->fRightToLeft << "\"" << std::endl;
          out << " rightToLeft=\"" << fields->fRightToLeft << "\"";
        }
        if (fields->fDspFmla)
          out << " showFormulas=\"" << fields->fDspFmla << "\"";
        if (!fields->fDspGrid)
          out << " showGridLines=\"" << fields->fDspGrid << "\"";
        if (!fields->fDspGuts)
          out << " showOutlineSymbols=\"" << fields->fDspGuts << "\"";
        if (!fields->fDspRwCol)
          out << " showRowColHeaders=\"" << fields->fDspRwCol << "\"";
        if (fields->fDspRuler)
          out << " showRuler=\"" << fields->fDspRuler << "\"";
        if (fields->fWhitespaceHidden)
          out << " showWhiteSpace=\"" << fields->fWhitespaceHidden << "\"";
        if (fields->fDspZeros)
          out << " showZeros=\"" << fields->fDspZeros << "\"";
        if (fields->fSelected)
          out << " tabSelected=\"" << fields->fSelected << "\"";
        if (colLeft > 0 || rwTop > 0)
          out << " topLeftCell=\"" << int_to_col(colLeft + 1) << std::to_string(rwTop + 1) << "\"";
        if (xlView)
          out << " view=\"" << xlView << "\"";
        if (fields->fWnProt)
          out << " windowProtection=\"" << fields->fWnProt << "\"";
        if (wScale)
          out << " zoomScale=\"" << wScale << "\"";
        if (wScaleNormal)
          out << " zoomScaleNormal=\"" << wScaleNormal << "\"";
        if (wScalePLV)
          out << " zoomScalePageLayoutView=\"" << wScalePLV<< "\"";
        if (wScaleSLV)
          out << " zoomScaleSheetLayoutView=\"" << wScaleSLV<< "\"";

        out << ">" << std::endl;

        break;
      }

        // whatever this is
      case BrtSel:
      {
        if (debug) Rcpp::Rcout << "BrtSel: " << bin.tellg() << std::endl;
        uint32_t pnn = 0, rwAct = 0, colAct = 0, dwRfxAct = 0;
        std::vector<int> sqrfx;

        pnn = readbin(pnn, bin, 0);
        rwAct = readbin(rwAct, bin, 0);
        colAct = readbin(colAct, bin, 0);
        dwRfxAct = readbin(dwRfxAct, bin, 0);

        // sqrfx[0] is the number of rfx following
        sqrfx = UncheckedSqRfX(bin);

        std::string ac = int_to_col(colAct + 1) + std::to_string(rwAct + 1);
        std::string lref = int_to_col(sqrfx[3] + 1) + std::to_string(sqrfx[1] + 1);
        std::string rref = int_to_col(sqrfx[4] + 1) + std::to_string(sqrfx[2] + 1);

        std::string sqref;
        if (lref.compare(rref) == 0) {
          sqref = lref;
        } else {
          sqref =  lref + ":" + rref;
        }

        out << "<selection activeCell=\"" << ac << "\" sqref=\""<< sqref << "\" />" << std::endl;

        break;
      }

      case BrtEndCsView:
      case BrtEndWsView:
      {
        if (debug) Rcpp::Rcout << "</sheetView>: " << bin.tellg() << std::endl;
        out << "</sheetView>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtEndCsViews:
      case BrtEndWsViews:
      {
        if (debug) Rcpp::Rcout << "</sheetViews>: " << bin.tellg() << std::endl;
        out << "</sheetViews>" << std::endl;
        break;
      }

      case BrtBeginColInfos:
      {
        if (debug) Rcpp::Rcout << "<cols>: " << bin.tellg() << std::endl;
        out << "<cols>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtColInfo:
      {
        if (debug) Rcpp::Rcout << "<col/>: " << bin.tellg() << std::endl;
        uint16_t colinfo = 0;
        uint32_t colFirst = 0, colLast = 0, coldx = 0, ixfe = 0;

        colFirst = UncheckedCol(bin) + 1;
        colLast = UncheckedCol(bin) + 1;
        coldx = readbin(coldx, bin, 0);
        ixfe = readbin(ixfe, bin, 0);
        colinfo = readbin(colinfo, bin, 0);

        BrtColInfoFields *fields = (BrtColInfoFields *)&colinfo;

        out << "<col" << " min=\"" << colFirst << "\" max =\"" << colLast << "\"";

        if (ixfe > 0)
          out << " s=\"" <<  ixfe << "\"";

        out << " width=\"" <<  (double)coldx/256 << "\"";
        if (fields->fHidden)
          out << " hidden=\"" <<  fields->fHidden << "\"";
        if (fields->fUserSet)
          out << " customWidth=\"" <<  fields->fUserSet << "\"";
        if (fields->fBestFit)
          out << " bestFit=\"" <<  fields->fBestFit << "\"";
        if (fields->iOutLevel>0)
          out << " outlineLevel=\"" <<  fields->iOutLevel << "\"";
        if (fields->fCollapsed)
          out << " collapsed=\"" <<  fields->fCollapsed << "\"";

        out << " />" << std::endl;
        break;
      }

      case BrtEndColInfos:
      {
        if (debug) Rcpp::Rcout << "</cols>: " << bin.tellg() << std::endl;
        out << "</cols>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtBeginSheetData:
      {
        if (debug) Rcpp::Rcout << "<sheetData>" << bin.tellg() << std::endl;
        out << "<sheetData>" << std::endl; //  << bin.tellg()
        break;
      }

        // prelude to row entry
      case BrtACBegin:
      {
        if (debug) Rcpp::Rcout << "BrtACBegin: " << bin.tellg() << std::endl;
        bin.seekg(size, bin.cur);
        // int16_t nversions, version;
        // nversions = readbin(nversions, bin , 0);
        // std::vector<int> versions;
        //
        // for (int i = 0; i < nversions; ++i) {
        //   version = readbin(version, bin, 0);
        //   versions.push_back(version);
        // }
        //
        // Rf_PrintValue(Rcpp::wrap(versions));

        break;
      }

      case BrtWsFmtInfoEx14:
      {
        if (debug) Rcpp::Rcout << "BrtWsFmtInfoEx14: " << bin.tellg() << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtACEnd:
      {
        if (debug) Rcpp::Rcout << "BrtACEnd: " << bin.tellg() << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtWsFmtInfo:
      {
        if (debug) Rcpp::Rcout << "BrtWsFmtInfo: " << bin.tellg() << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtRwDescent:
      {
        if (debug) Rcpp::Rcout << "BrtRwDescent: " << bin.tellg() << std::endl;
        // bin.seekg(size, bin.cur);
        uint16_t dyDescent = 0;
        dyDescent = readbin(dyDescent, bin, 0);
        break;
      }

      case BrtRowHdr:
      {
        if (debug) Rcpp::Rcout << "<row/>: " << bin.tellg() << std::endl;

        // close open rows
        if (!first_row) {
        out << "</row>" <<std::endl;
      } else {
        first_row = false;
      }

      uint8_t bits1 = 0, bits2 = 0, bits3 = 0, fExtraAsc = 0, fExtraDsc = 0,
        fCollapsed = 0, fDyZero = 0, fUnsynced = 0, fGhostDirty = 0,
        fReserved = 0, fPhShow = 0;
      uint16_t miyRw = 0;

      // uint24_t;
      int32_t rw = 0;
      uint32_t ixfe = 0, ccolspan = 0, unk32 = 0, colMic = 0, colLast = 0;

      rw = readbin(rw, bin, 0);

      if (rw > 0x00100000 || rw < row)
        Rcpp::stop("row either decreasing or to large");

      ixfe = readbin(ixfe, bin, 0);
      miyRw = readbin(miyRw, bin, 0);

      if (miyRw > 0x2000) Rcpp::stop("miyRw to big");

      uint16_t rwoheaderfields = 0;
      rwoheaderfields = readbin(rwoheaderfields, bin, 0);

      BrtRowHdrFields *fields = (BrtRowHdrFields *)&rwoheaderfields;

      // fExtraAsc = 1
      // fExtraDsc = 1
      // reserved1 = 6
      // iOutLevel   = 3
      // fCollapsed  = 1
      // fDyZero     = 1
      // fUnsynced   = 1
      // fGhostDirty = 1
      // fReserved   = 1

      bits3 = readbin(bits3, bin, 0);
      // fPhShow   = 1
      // fReserved = 7

      ccolspan = readbin(ccolspan, bin, 0);
      std::string spans;
      if (ccolspan) {
        // rgBrtColspan
        // not sure if these are alway around. maybe ccolspan is a counter
        colMic = readbin(colMic, bin, 0);
        colLast = readbin(colLast, bin, 0);

        spans = std::to_string(colMic + 1) + ":" + std::to_string(colLast + 1);
      }

      out << "<row r=\""; //  << bin.tellg()
      out << rw + 1 << "\"";

      if (fields->fUnsynced != 0) {
        // Rcpp::Rcout << " ht=\"" << miyRw/20 << "\"" << std::endl;
        out << " ht=\"" << miyRw/20 << "\"";
        out << " customHeight=\"" << fields->fUnsynced << "\"";
      }

      if (ccolspan)
        out << " spans=\"" << spans << "\"";

      if (fields->iOutLevel > 0) {
        out << " outlineLevel=\"" << fields->iOutLevel << "\"";
      }

      if (fields->fCollapsed) {
        out << " collapsed=\"" << fields->fCollapsed << "\"";
      }

      if (fields->fDyZero) {
        out << " hidden=\"" << fields->fDyZero << "\"";
      }

      if (fields->fGhostDirty) {
        out << " customFormat=\"" << fields->fGhostDirty << "\"";
      }

      if (ixfe > 0) {
        out << " s=\"" << ixfe << "\"";
      }

      out << ">" << std::endl;

      row = rw;

      if (debug)
        Rcpp::Rcout << (rw) << " : " << ixfe << " : " << miyRw << " : " << (int32_t)fExtraAsc << " : " <<
          (int32_t)fExtraDsc << " : " << unk32 << " : " << (int32_t)fCollapsed << " : " << (int32_t)fDyZero << " : " <<
            (int32_t)fUnsynced << " : " << (int32_t)fGhostDirty << " : " << (int32_t)fReserved << " : " <<
              (int32_t)fPhShow << " : " << ccolspan << "; " << bin.tellg() << std::endl;

      break;
      }

      case BrtCellIsst:
      { // shared string
        if (debug) Rcpp::Rcout << "BrtCellIsst: " << bin.tellg() << std::endl;

        int32_t val1 = 0, val2 = 0, val3 = 0;
        val1 = readbin(val1, bin, 0);
        if (debug) Rcpp::Rcout << val1 << std::endl;
        val2 = readbin(val2, bin, 0);
        if (debug) Rcpp::Rcout << val2 << std::endl;
        val3 = readbin(val3, bin, 0);
        if (debug) Rcpp::Rcout << val3 << std::endl;

        out << "<c r=\"" << int_to_col(val1 + 1) << row + 1 << "\"" << cell_style(val2) << " t=\"s\">" << std::endl;
        out << "<v>" << val3 << "</v>" << std::endl;
        out << "</c>" << std::endl;

        break;
      }

      case BrtCellBool:
      { // bool
        if (debug) Rcpp::Rcout << "BrtCellBool: " << bin.tellg() << std::endl;

        int32_t val1 = 0, val2 = 0;
        uint8_t val3 = 0;
        val1 = readbin(val1, bin, 0);
        // out << val1 << std::endl;
        val2 = readbin(val2, bin, 0);
        // out << val2 << std::endl;
        val3 = readbin(val3, bin, 0);
        // out << val3 << std::endl;

        out << "<c r=\"" << int_to_col(val1 + 1) << row + 1<< "\"" << cell_style(val2) << " t=\"b\">" << std::endl;
        out << "<v>" << (int32_t)val3 << "</v>" << std::endl;
        out << "</c>" << std::endl;

        break;
      }

      case BrtCellRk:
      { // integer?
        if (debug) Rcpp::Rcout << "BrtCellRk: " << bin.tellg() << std::endl;

        int32_t val1 = 0, val2= 0, val3= 0;
        val1 = readbin(val1, bin, 0);
        // Rcpp::Rcout << val << std::endl;
        val2 = readbin(val2, bin, 0);
        // Rcpp::Rcout << val << std::endl;
        // wrong?
        val3 = readbin(val3, bin, 0);
        // Rcpp::Rcout << RkNumber(val) << std::endl;

        out << "<c r=\"" << int_to_col(val1 + 1) << row + 1<< "\"" << cell_style(val2) << ">" << std::endl;
        out << "<v>" << std::setprecision(16) << RkNumber(val3) << "</v>" << std::endl;
        out << "</c>" << std::endl;

        break;
      }


      case BrtCellReal:
      {
        if (debug) Rcpp::Rcout << "BrtCellReal: " << bin.tellg() << std::endl;
        int32_t val1= 0, val2= 0;
        val1 = readbin(val1, bin, 0);
        // Rcpp::Rcout << val << std::endl;
        val2 = readbin(val2, bin, 0);
        // Rcpp::Rcout << val << std::endl;

        double dbl = 0.0;
        dbl = readbin(dbl, bin, 0);
        // Rcpp::Rcout << dbl << std::endl;

        out << "<c r=\"" << int_to_col(val1 + 1) << row + 1 << "\"" << cell_style(val2) << ">" << std::endl;
        out << "<v>" << std::setprecision(16) << dbl << "</v>" << std::endl; // << std::fixed
        out << "</c>" << std::endl;

        break;
      }

        // 0 ?
      case BrtCellBlank:
      {
        if (debug) Rcpp::Rcout << "BrtCellBlank: " << bin.tellg() << std::endl;

        std::vector<int> blank = Cell(bin);
        col = blank[0];
        if (debug) Rf_PrintValue(Rcpp::wrap(blank));

        out << "<c r=\"" << int_to_col(blank[0] + 1) << row + 1 << "\"" << cell_style(blank[1]) << "/>" << std::endl;

        break;
      }

      case BrtCellError:
      { // t="e" & <v>#NUM!</v>
        if (debug) Rcpp::Rcout << "BrtCellError: " << bin.tellg() << std::endl;

        // uint8_t val8= 0;

        int32_t val1= 0, val2= 0;
        val1 = readbin(val1, bin, 0);
        // out << val << std::endl;
        val2 = readbin(val2, bin, 0);
        // out << val << std::endl;

        out << "<c r=\"" << int_to_col(val1 + 1) << row + 1<< "\"" << cell_style(val2) << " t=\"e\">" << std::endl;
        out << "<v>" << BErr(bin) << "</v>" << std::endl;
        out << "</c>" << std::endl;

        break;
      }

      case BrtFmlaBool:
      {
        if (debug) Rcpp::Rcout << "BrtFmlaBool: " << bin.tellg() << std::endl;
        // bin.seekg(size, bin.cur);

        std::vector<int> cell;
        cell = Cell(bin);
        col = cell[0];
        if (debug) Rf_PrintValue(Rcpp::wrap(cell));

        bool val = 0;
        val = readbin(val, bin, 0);
        if (debug) Rcpp::Rcout << val << std::endl;

        uint16_t grbitFlags = 0;
        grbitFlags = readbin(grbitFlags, bin, 0);

        GrbitFmlaFields *fields = (GrbitFmlaFields *)&grbitFlags;

        std::string fml = CellParsedFormula(bin, debug, 0);

        out << "<c r=\"" << int_to_col(cell[0] + 1) << row + 1 << "\"" << cell_style(cell[1]) << " t=\"b\">" << std::endl;
        out << "<f>" << fml << "</f>" << std::endl;
        out << "<v>" << val << "</v>" << std::endl;
        out << "</c>" << std::endl;

        break;
      }

      case BrtFmlaError:
      { // t="e" & <f>
        if (debug) Rcpp::Rcout << "BrtFmlaError: " << bin.tellg() << std::endl;
        // bin.seekg(size, bin.cur);
        std::vector<int> cell;
        cell = Cell(bin);
        col = cell[0];
        if (debug) Rf_PrintValue(Rcpp::wrap(cell));

        std::string fErr;
        fErr = BErr(bin);

        uint16_t grbitFlags = 0;
        grbitFlags = readbin(grbitFlags, bin, 0);

        GrbitFmlaFields *fields = (GrbitFmlaFields *)&grbitFlags;

        // int32_t len = size - 4 * 32 - 2 * 8;
        // std::string fml(len, '\0');

        std::string fml = CellParsedFormula(bin, debug, 0);

        out << "<c r=\"" << int_to_col(cell[0] + 1) << row + 1 << "\"" << cell_style(cell[1]) << " t=\"e\">" << std::endl;
        out << "<f>" << fml << "</f>" << std::endl;
        out << "<v>" << fErr << "</v>" << std::endl;
        out << "</c>" << std::endl;

        break;
      }

      case BrtFmlaNum:
      {
        if (debug) Rcpp::Rcout << "BrtFmlaNum: " << bin.tellg() << std::endl;
        // bin.seekg(size, bin.cur);

        std::vector<int> cell;
        cell = Cell(bin);
        col = cell[0];
        if (debug) Rf_PrintValue(Rcpp::wrap(cell));

        double xnum = Xnum(bin);
        if (debug) Rcpp::Rcout << xnum << std::endl;

        uint16_t grbitFlags = 0;
        grbitFlags = readbin(grbitFlags, bin, 0);

        GrbitFmlaFields *fields = (GrbitFmlaFields *)&grbitFlags;

        // Rprintf("%d, %d, %d\n",
        //         fields->reserved,
        //         fields->fAlwaysCalc,
        //         fields->unused);

        std::string fml = CellParsedFormula(bin, debug, 0);

        out << "<c r=\"" << int_to_col(cell[0] + 1) << row + 1 << "\"" << cell_style(cell[1]) << ">" << std::endl;
        out << "<f>" << fml << "</f>" << std::endl;
        out << "<v>" << xnum << "</v>" << std::endl;
        out << "</c>" << std::endl;

        break;
      }

      case BrtFmlaString:
      {
        if (debug) Rcpp::Rcout << "BrtFmlaString: " << bin.tellg() << std::endl;
        // bin.seekg(size, bin.cur);

        std::vector<int> cell;
        cell = Cell(bin);
        col = cell[0];
        if (debug) Rf_PrintValue(Rcpp::wrap(cell));

        std::string val = XLWideString(bin);
        if (debug) Rcpp::Rcout << val << std::endl;

        uint16_t grbitFlags = 0;
        grbitFlags = readbin(grbitFlags, bin, 0);

        std::string fml = CellParsedFormula(bin, debug, 0);

        out << "<c r=\"" << int_to_col(cell[0] + 1) << row + 1 << "\"" << cell_style(cell[1]) << " t=\"str\">" << std::endl;
        out << "<f>" << fml << "</f>" << std::endl;
        out << "<v>" << val << "</v>" << std::endl;
        out << "</c>" << std::endl;

        break;
      }

      case BrtShrFmla:
      {
        if (debug) Rcpp::Rcout << "BrtShrFmla: " << bin.tellg() << std::endl;

        uint32_t rwFirst = 0, rwLast = 0, colFirst = 0, colLast = 0;
        rwFirst  = UncheckedRw(bin) +1;
        rwLast   = UncheckedRw(bin) +1;
        colFirst = UncheckedCol(bin) +1;
        colLast  = UncheckedCol(bin) +1;

        Rcpp::Rcout << int_to_col(colFirst) << rwFirst << ":" << int_to_col(colLast) << rwLast << std::endl;

        std::string fml = CellParsedFormula(bin, debug, col);
        Rcpp::Rcout << fml << std::endl;

        // out << "<f>" << fml << "</f>" << std::endl;
        // out << "<v>" << val << "</v>" << std::endl;
        // out << "</c>" << std::endl;

        break;
      }

      case BrtEndSheetData:
      {
        if (debug) Rcpp::Rcout << "</sheetData>" << bin.tellg() << std::endl;

        if (!first_row) {
          out << "</row>" << std::endl;
        }
        out << "</sheetData>" << std::endl;

        break;
      }

      case BrtSheetProtection:
      {
        if (debug) Rcpp::Rcout << "BrtSheetProtection: " << bin.tellg() << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtBeginMergeCells:
      {
        if (debug) Rcpp::Rcout << "BrtBeginMergeCells: " << bin.tellg() << std::endl;
        uint32_t cmcs = 0;
        cmcs = readbin(cmcs, bin, 0);
        if (debug) Rcpp::Rcout << "count: " << cmcs << std::endl;

        out << "<mergeCells count=\"" << cmcs << "\">"<< std::endl;
          break;
      }

      case BrtMergeCell:
      {
        if (debug) Rcpp::Rcout << "BrtMergeCell: " << bin.tellg() << std::endl;

        uint32_t rwFirst = 0, rwLast = 0, colFirst = 0, colLast = 0;

        rwFirst  = UncheckedRw(bin) + 1L;
        rwLast   = UncheckedRw(bin) + 1L;
        colFirst = UncheckedCol(bin) + 1L;
        colLast  = UncheckedCol(bin) + 1L;

        if (debug)
          Rprintf("MergeCell: %d %d %d %d\n",
                  rwFirst, rwLast, colFirst, colLast);

        out << "<mergeCell ref=\"" <<
          int_to_col(colFirst) << rwFirst << ":" <<
            int_to_col(colLast) << rwLast <<
              "\" />" << std::endl;

        break;
      }

      case BrtEndMergeCells:
      {
        if (debug) Rcpp::Rcout << "BrtEndMergeCells: " << bin.tellg() << std::endl;
        out << "</mergeCells>" << std::endl;
        break;
      }

      case BrtPrintOptions:
      {
        if (debug) Rcpp::Rcout << "BrtPrintOptions: " << bin.tellg() << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtMargins:
      {
        if (debug)  Rcpp::Rcout << "BrtMargins: " << bin.tellg() << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtUID:
      {
        if (debug)  Rcpp::Rcout << "BrtUID: " << bin.tellg() << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtDrawing:
      {
        std::string stRelId = XLNullableWideString(bin);
        out << "<drawing r:id=\"" << stRelId << "\" />" << std::endl;
        break;
      }

      case BrtLegacyDrawing:
      {
        std::string stRelId = XLNullableWideString(bin);
        out << "<legacyDrawing r:id=\"" << stRelId << "\" />" << std::endl;
        break;
      }

      case BrtLegacyDrawingHF:
      {
        std::string stRelId = XLNullableWideString(bin);
        out << "<legacyDrawingHF r:id=\"" << stRelId << "\" />" << std::endl;
        break;
      }

      case BrtHLink:
      {
        std::vector<int> rfx = UncheckedRfX(bin);
        std::string relId = XLNullableWideString(bin);
        std::string location = XLWideString(bin);
        std::string tooltip = XLWideString(bin);
        std::string display = XLWideString(bin);


        std::string lref = int_to_col(rfx[2] + 1) + std::to_string(rfx[0] + 1);
        std::string rref = int_to_col(rfx[3] + 1) + std::to_string(rfx[1] + 1);

        std::string ref;
        if (lref.compare(rref) == 0) {
          ref = lref;
        } else {
          ref =  lref + ":" + rref;
        }

        std::string hlink = "<hyperlink display=\"" +  display + "\" r:id=\"" + relId + "\" location=\"" + location + "\" ref=\"" + ref + "\" tooltip=\"" + tooltip + "\" />";
        hlinks.push_back(hlink);

        break;
      }

      case BrtBeginAFilter:
      {
        if (debug) Rcpp::Rcout << "<autofilter>" << std::endl;
        std::vector<int> rfx = UncheckedRfX(bin);

        std::string lref = int_to_col(rfx[2] + 1) + std::to_string(rfx[0] + 1);
        std::string rref = int_to_col(rfx[3] + 1) + std::to_string(rfx[1] + 1);

        std::string ref;
        if (lref.compare(rref) == 0) {
          ref = lref;
        } else {
          ref =  lref + ":" + rref;
        }

        // ignoring filterColumn for now
        // autofilter can consist of filterColumn, filters and filter
        // maybe customFilter, dynamicFilter too
        out << "<autoFilter ref=\"" << ref << "\">" << std::endl;

        break;
      }

      case BrtBeginFilterColumn:
      {
        if (debug) Rcpp::Rcout << "<filterColumn>" << std::endl;

        uint16_t flags = 0;
        uint32_t dwCol = 0;
        dwCol = readbin(dwCol, bin, 0);
        flags = readbin(flags, bin, 0);
        // fHideArrow
        // fNoBtn

        out << "<filterColumn colId=\"" << dwCol <<"\">" << std::endl;

        break;
      }

      case BrtBeginFilters:
      {
        if (debug) Rcpp::Rcout << "<filters>" << std::endl;
        // bin.seekg(size, bin.cur);

        uint32_t fBlank = 0, unused = 0;
        fBlank = readbin(fBlank, bin, 0); // a 32bit flag, after all ... why not?
        unused = readbin(unused, bin, 0); // maybe calendarType?
        out << "<filters blank=\"" << fBlank <<"\">" << std::endl;
        break;
      }

      case BrtFilter:
      {
        std::string rgch = XLWideString(bin);
        out << "<filter val=\"" << rgch << "\" />" << std::endl;
        // bin.seekg(size, bin.cur);
        break;
      }

      case BrtEndFilters:
      {
        if (debug) Rcpp::Rcout << "</filters>" << std::endl;
        out << "</filters>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtColorFilter:
      {
        uint32_t dxfid = 0, fCellColor = 0;
        dxfid = readbin(dxfid, bin, 0);
        fCellColor = readbin(fCellColor, bin, 0);

        out << "<colorFilter dxfId=\"" << dxfid << "\" cellColor=\""<< fCellColor<< "\"/>" << std::endl;

        break;
      }

      case BrtBeginCustomFilters:
      case BrtBeginCustomFilters14:
      case BrtBeginCustomRichFilters:
      {
        Rcpp::warning("Custom Filter found. This is not handled.");
        bin.seekg(size, bin.cur);

        break;
      }

      case BrtDynamicFilter:
      {
        Rcpp::warning("Dynamic Filter found. This is not handled.");
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtIconFilter:
      case BrtIconFilter14:
      {
        uint32_t iIconSet = 0, iIcon = 0;
        iIconSet = readbin(iIconSet, bin, 0);
        iIcon = readbin(iIcon, bin, 0);

        std::string iconSet;
        if (iIconSet) iconSet = to_iconset(iIconSet);

        out << "<iconFilter iconSet=\"" << iconSet << "\" iconId=\""<< iIcon<< "\"/>" << std::endl;

        break;
      }

      case BrtEndFilterColumn:
      {
        if (debug) Rcpp::Rcout << "</filterColumn>" << std::endl;
        out << "</filterColumn>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtEndAFilter:
      {
        if (debug) Rcpp::Rcout << "</autofilter>" << std::endl;
        out << "</autoFilter>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtBeginListParts:
      {
        if (debug) Rcpp::Rcout << "<tableParts>" << std::endl;
        uint32_t cParts = 0;
        cParts = readbin(cParts, bin, 0);
        out << "<tableParts count=\"" << cParts << "\">" << std::endl;
        break;
      }

      case BrtListPart:
      {
        if (debug) Rcpp::Rcout << "<tablePart/>" << std::endl;
        std::string stRelID = XLNullableWideString(bin);
        out << "<tablePart r:id=\"" << stRelID << "\" />" << std::endl;
        break;
      }

      case BrtEndListParts:
      {
        if (debug) Rcpp::Rcout << "</tableParts>" << std::endl;
        out << "</tableParts>" << std::endl;
        break;
      }

      case BrtBeginConditionalFormatting:
      {

        Rcpp::warning("Worksheet contains unhandled conditional formatting.");

        if (debug) Rcpp::Rcout << "<conditionalFormatting>" << std::endl;

        uint32_t ccf = 0, fPivot = 0;
        std::vector<int> sqrfx;

        ccf = readbin(ccf, bin, 0); // not needed?
        fPivot = readbin(fPivot, bin, 0);

        sqrfx = UncheckedSqRfX(bin);
        std::string lref = int_to_col(sqrfx[3] + 1) + std::to_string(sqrfx[1] + 1);
        std::string rref = int_to_col(sqrfx[4] + 1) + std::to_string(sqrfx[2] + 1);

        std::string sqref;
        if (lref.compare(rref) == 0) {
          sqref = lref;
        } else {
          sqref =  lref + ":" + rref;
        }

        out << "<conditionalFormatting";
        if (fPivot) out << " pivot=\"" << fPivot << "\"";
        out << " sqref=\"" << sqref << "\"";
        out << ">" << std::endl;

        break;
      }

      case BrtBeginCFRule:
      {
        if (debug) Rcpp::Rcout << "<cfRule>" << std::endl;

        uint16_t flags = 0;
        uint32_t iType = 0, iTemplate = 0, dxfId = 0, iPri = 0, iParam = 0, reserved1 = 0, reserved2 = 0;
        uint32_t cbFmla1 = 0, cbFmla2 = 0, cbFmla3 = 0;

        iType = readbin(iType, bin, 0);
        iTemplate = readbin(iTemplate, bin, 0);
        dxfId = readbin(dxfId, bin, 0);
        iPri = readbin(iPri, bin, 0);
        iParam = readbin(iParam, bin, 0);
        reserved1 = readbin(reserved1, bin, 0);
        reserved2 = readbin(reserved2, bin, 0);
        flags = readbin(flags, bin, 0);
        cbFmla1 = readbin(cbFmla1, bin, 0);
        cbFmla2 = readbin(cbFmla2, bin, 0);
        cbFmla3 = readbin(cbFmla3, bin, 0);

        std::string strParam = XLNullableWideString(bin);

        std::string rgce1, rgce2, rgce3;
        if (cbFmla1 != 0x00000000) {
          rgce1 = CellParsedFormula(bin, debug, 0);
        }
        if (cbFmla2 != 0x00000000) {
          rgce2 = CellParsedFormula(bin, debug, 0);
        }
        if (cbFmla3 != 0x00000000) {
          rgce3 = CellParsedFormula(bin, debug, 0);
        }

        BrtBeginCFRuleFields *fields = (BrtBeginCFRuleFields *)&flags;

        Rcpp::Rcout << "<cfRule";
        // the type is defined by iType and iTemplate eg:
        // CF_TYPE_EXPRIS & CF_TEMPLATE_CONTAINSNOBLANKS
        // (but then again, not sure why iType is needed)
        Rcpp::Rcout << " type=\"" << iType << "/" << iTemplate << "\"";
        Rcpp::Rcout << " dxfId=\"" << dxfId << "\"";
        Rcpp::Rcout << " priority=\"" << iPri << "\"";
        Rcpp::Rcout << " operator=\"" << iParam << "\"";
        Rcpp::Rcout << " stopIfTrue=\"" << fields->fStopTrue << "\"";
        Rcpp::Rcout << " percent=\"" << fields->fPercent << "\"";
        Rcpp::Rcout << " aboveAverage=\"" << fields->fAbove << "\"";
        Rcpp::Rcout << " bottom=\"" << fields->fBottom << "\"";
        // rank="" ???
        // stdDev="" ??
        // timePeriod="" ??
        Rcpp::Rcout << " text=\"" << strParam << "\"";
        Rcpp::Rcout << " >" << std::endl;

        if (debug) {
          Rcpp::Rcout << rgce1 << std::endl;
          Rcpp::Rcout << rgce2 << std::endl;
          Rcpp::Rcout << rgce3 << std::endl;
        }


        break;
      }

      case BrtEndCFRule:
      {

        if (debug) Rcpp::Rcout << "</cfRule>" << std::endl;
        Rcpp::Rcout << "</cfRule>" << std::endl;

        break;
      }

      case BrtEndConditionalFormatting:
      {

        if (debug) Rcpp::Rcout << "</conditionalFormatting>" << std::endl;
        out << "</conditionalFormatting>" << std::endl;

        break;
      }

      case BrtEndSheet:
      {

        if (hlinks.size() > 1) {
        // did not see BrtBeginHL or BrtEndHL, likely I'm just blind
        hlinks.push_back("</hyperlinks>");

        for (int i = 0; i < hlinks.size(); ++i) {
          if (debug) Rcpp::Rcout << hlinks[i] << std::endl;
          out << hlinks[i] << std::endl;
        }
      }

        if (debug)  Rcpp::Rcout << "</worksheet>" << bin.tellg() << std::endl;

        if (chartsheet)
          out << "</chartsheet>" << std::endl;
        else
          out << "</worksheet>" << std::endl;
        end_of_worksheet = true;
        break;
      }

      default:
      {
        if (debug) {
        Rcpp::Rcout << "Unhandled: " << std::to_string(x) <<
          ": " << std::to_string(size) <<
            " @ " << bin.tellg() << std::endl;
      }
        bin.seekg(size, bin.cur);
        break;
      }
      }

      // itr++;
      // if (itr == 20) {
      //   Rcpp::stop("stop");
      // }
    }

    out.close();
    bin.close();
    return 1;
  } else {
    return -1;
  };

}
