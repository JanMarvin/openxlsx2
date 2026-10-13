// #include <Rcpp/Lightest>

#include <cstdio>
#include <cstring>
#include <fstream>
#include <iomanip>
#include <iterator>

#include "openxlsx2.h"
#include "xlsb_defines.h"
#include "xlsb_flags.h"
#include "xlsb_funs.h"
#include "xls_defines.h"
#include "xls_funs.h"
#include "xls_pivot.h"

// [[Rcpp::export]]
int32_t styles_bin(std::string filePath, std::string outPath, bool debug) {
  std::ofstream out(outPath);
  std::ifstream bin(filePath, std::ios::in | std::ios::binary | std::ios::ate);

  bool swapit = is_big_endian();

  // auto sas_size = bin.tellg();
  if (bin) {
    bin.seekg(0, std::ios_base::beg);
    bool end_of_style_sheet = false;

    R_xlen_t idx = 0;
    while (!end_of_style_sheet) {
      checkInterrupt(idx);

      int32_t x = 0, size = 0;

      if (debug) Rcpp::Rcout << "." << std::endl;
      RECORD(x, size, bin, swapit);
      if (debug) Rcpp::Rcout << x << ": " << size << std::endl;

      switch (x) {
        case BrtBeginStyleSheet: {
          out << "<styleSheet>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

          // fills
        case BrtBeginFmts: {
          out << "<numFmts>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtFmt: {
          out << "<numFmt";
          // bin.seekg(size, bin.cur);
          uint16_t ifmt = 0;
          ifmt = readbin(ifmt, bin, swapit);

          std::string stFmtCode = XLWideString(bin, swapit);

          out << " numFmtId=\"" << ifmt << "\" formatCode=\"" << escape_xml(stFmtCode);
          out << "\" />" << std::endl;
          break;
        }

        case BrtEndFmts: {
          out << "</numFmts>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtBeginFonts: {
          if (debug) Rcpp::Rcout << "<fonts>" << std::endl;
          // bin.seekg(size, bin.cur);
          uint32_t cfonts = 0;
          cfonts = readbin(cfonts, bin, swapit);

          out << "<fonts " << "count=\"" << cfonts << "\">" << std::endl;
          break;
        }

        case BrtFont: {
          out << "<font>" << std::endl;
          uint8_t uls = 0, bFamily = 0, bCharSet = 0, unused = 0, bFontScheme = 0;
          uint16_t dyHeight = 0, grbit = 0, bls = 0, sss = 0;

          dyHeight = readbin(dyHeight, bin, swapit) / 20;  // twip
          grbit = readbin(grbit, bin, swapit);

          FontFlags view_flags(grbit);

          bls = readbin(bls, bin, swapit);
          // 0x0190 normal
          // 0x02BC bold
          sss = readbin(sss, bin, swapit);
          // 0x0000 None
          // 0x0001 Superscript
          // 0x0002 Subscript
          uls = readbin(uls, bin, swapit);
          // 0x00 None
          // 0x01 Single
          // 0x02 Double
          // 0x21 Single accounting
          // 0x22 Double accounting
          bFamily = readbin(bFamily, bin, swapit);
          // 0x00 Not applicable
          // 0x01 Roman
          // 0x02 Swiss
          // 0x03 Modern
          // 0x04 Script
          // 0x05 Decorative
          bCharSet = readbin(bCharSet, bin, swapit);
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
          unused = readbin(unused, bin, swapit);

          std::vector<int32_t> color = brtColor(bin, swapit);
          // needs handling for rgb/hex colors
          if (debug) Rf_PrintValue(Rcpp::wrap(color));

          bFontScheme = readbin(bFontScheme, bin, swapit);
          // 0x00 None
          // 0x01 Major font scheme
          // 0x02 Minor font scheme

          std::string name = XLWideString(bin, swapit);

          if (bls == 0x02BC)  out << "<b/>" << std::endl;
          if (view_flags.fItalic()) out << "<i/>" << std::endl;
          if (view_flags.fStrikeout()) out << "<strike/>" << std::endl;
          if (view_flags.fOutline()) out << "<outline/>" << std::endl;
          if (view_flags.fShadow()) out << "<shadow/>" << std::endl;
          if (view_flags.fCondense()) out << "<condense/>" << std::endl;
          if (view_flags.fExtend()) out << "<extend/>" << std::endl;

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
            double tint = 0.0;
            if (color[2] != 0) tint = (double)color[2] / 32767;

            std::stringstream stream;
            stream << std::setprecision(16) << tint;

            out << "<color theme=\"" << color[1] << "\" tint=\"" << stream.str() << "\" />";
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

        case BrtEndFonts: {
          if (debug) Rcpp::Rcout << "</fonts>" << std::endl;
          bin.seekg(size, bin.cur);

          out << "</fonts>" << std::endl;
          break;
        }

        // fills
        case BrtBeginFills: {
          out << "<fills>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtFill: {
          uint32_t fls = 0, iGradientType = 0, cNumStop = 0;
          double
            xnumDegree = 0.0,
            xnumFillToLeft = 0.0,
            xnumFillToRight = 0.0,
            xnumFillToTop = 0.0,
            xnumFillToBottom = 0.0,
            aPos = 0.0 // not used
          ;
          std::vector<int32_t> aColor;  // not used

          fls = readbin(fls, bin, swapit);
          std::vector<int32_t> fgColor = brtColor(bin, swapit);
          std::vector<int32_t> bgColor = brtColor(bin, swapit);
          iGradientType = readbin(iGradientType, bin, swapit);
          xnumDegree = Xnum(bin, swapit);
          xnumFillToLeft = Xnum(bin, swapit);
          xnumFillToRight = Xnum(bin, swapit);
          xnumFillToTop = Xnum(bin, swapit);
          xnumFillToBottom = Xnum(bin, swapit);
          cNumStop = readbin(cNumStop, bin, swapit);

          if (debug) {  // #nocov start
            Rcpp::Rcout << "unused gradinet filll colors" <<
              xnumDegree << ": " << xnumFillToLeft << ": " << ": " <<
              xnumFillToRight << ": " << xnumFillToTop << "; " <<
              xnumFillToBottom << std::endl;
          }  // #nocov end

          for (uint32_t i = 0; i < cNumStop; ++i) {
            // gradient fill colors and positions
            aColor = brtColor(bin, swapit);
            aPos   = Xnum(bin, swapit);
            if (debug) {  // #nocov start
              Rcpp::Rcout << "unused gradient Colors: " <<
                to_argb(fgColor[6], fgColor[3], fgColor[4], fgColor[5]) <<
                  ": " << aPos << std::endl;
            }  // #nocov end
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

            double bgtint = 0.0, fgtint = 0.0;
            if (bgColor[2] != 0) bgtint = (double)bgColor[2] / 32767;
            if (fgColor[2] != 0) fgtint = (double)fgColor[2] / 32767;

            std::stringstream bgstream, fgstream;
            bgstream << std::setprecision(16) << bgtint;
            fgstream << std::setprecision(16) << fgtint;

            if (fgColor[0] == 0x00)
              out << "<fgColor auto=\"1\" />" << std::endl;
            if (fgColor[0] == 0x01)
              out << "<fgColor indexed=\"" << fgColor[1] << "\" />";
            if (fgColor[0] == 0x02)
              out << "<fgColor rgb=\"" << to_argb(fgColor[6], fgColor[3], fgColor[4], fgColor[5]) << "\" />";
            if (fgColor[0] == 0x03)
              out << "<fgColor theme=\"" << fgColor[1] << "\" tint=\"" << fgstream.str() << "\" />";

            if (bgColor[0] == 0x00)
              out << "<bgColor auto=\"1\" />" << std::endl;
            if (bgColor[0] == 0x01)
              out << "<bgColor indexed=\"" << bgColor[1] << "\" />";
            if (bgColor[0] == 0x02)
              out << "<bgColor rgb=\"" << to_argb(bgColor[6], bgColor[3], bgColor[4], bgColor[5]) << "\" />";
            if (bgColor[0] == 0x03)
              out << "<bgColor theme=\"" << bgColor[1] << "\" tint=\"" << fgstream.str() << "\" />";

            out << "</patternFill>" << std::endl;
            out << "</fill>" << std::endl;
          } else {
            Rcpp::Rcout << "gradient fill not implemented" << std::endl;
          }

          break;
        }

        case BrtEndFills: {
          out << "</fills>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

          // borders
        case BrtBeginBorders: {
          out << "<borders>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtBorder: {
          out << "<border>" << std::endl;
          // bool fBdrDiagDown = 0, fBdrDiagUp = 0;
          uint8_t borderFlags = 0;
          borderFlags = readbin(borderFlags, bin, swapit);

          std::string top      = brtBorder("top", bin, swapit);
          std::string bottom   = brtBorder("bottom", bin, swapit);
          std::string left     = brtBorder("left", bin, swapit);
          std::string right    = brtBorder("right", bin, swapit);
          std::string diagonal = brtBorder("diagonal", bin, swapit);

          // xml requires a different order a certain spreadsheet software
          // is quite picky in this regard
          out << left << right << top << bottom << diagonal;

          out << "</border>" << std::endl;
          // bin.seekg(size, bin.cur);
          break;
        }

        case BrtEndBorders: {
          out << "</borders>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

          // cellStyleXfs
        case BrtBeginCellStyleXFs: {
          out << "<cellStyleXfs>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtEndCellStyleXFs: {
          out << "</cellStyleXfs>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

          // cellXfs
        case BrtBeginCellXFs: {
          out << "<cellXfs>" << std::endl;
          bin.seekg(size, bin.cur);
          // uint32_t cxfs = 0;
          // cxfs = readbin(cxfs, bin, swapit);
          break;
        }

        case BrtXF: {
          out << "<xf";

          // bool is_cell_style_xf = false;

          uint8_t trot = 0, indent = 0;
          uint16_t ixfeParent = 0, iFmt = 0, iFont = 0, iFill = 0, ixBorder = 0;
          uint32_t brtxf = 0;

          ixfeParent = readbin(ixfeParent, bin, swapit);
          iFmt = readbin(iFmt, bin, swapit);
          iFont = readbin(iFont, bin, swapit);
          iFill = readbin(iFill, bin, swapit);
          ixBorder = readbin(ixBorder, bin, swapit);
          trot = readbin(trot, bin, swapit);
          indent = readbin(indent, bin, swapit);

          if (indent > 250) {
            Rcpp::stop("indent to big");
          }

          if (ixfeParent == 0xFFFF) {
            // is_cell_style_xf = true;
          } else {
            out << " xfId=\"" << ixfeParent << "\"";
          }

          brtxf = readbin(brtxf, bin, swapit);

          XFFlags view_flags(brtxf);

          uint8_t xfgbit = view_flags.xfGrbitAtr();

          xfGrbitAtrFlags view_flags2(xfgbit);

          out << " numFmtId=\"" << iFmt << "\"";
          out << " fontId=\"" << iFont << "\"";
          out << " fillId=\"" << iFill << "\"";
          out << " borderId=\"" << ixBorder << "\"";
          out << " applyNumberFormat=\"" << !view_flags2.bit1() << "\"";
          out << " applyFont=\"" << !view_flags2.bit2() << "\"";
          out << " applyBorder=\"" << !view_flags2.bit4() << "\"";
          out << " applyFill=\"" << !view_flags2.bit5() << "\"";
          out << " applyAlignment=\"" << !view_flags2.bit3() << "\"";
          out << " applyProtection=\"" << !view_flags2.bit6() << "\"";


          if (view_flags.alc() > 0 || view_flags.alcv() > 0 || indent || trot || view_flags.iReadingOrder() || view_flags.fShrinkToFit() || view_flags.fWrap() || view_flags.fJustLast()) {
            out << "><alignment";
            out << halign(view_flags.alc());
            out << valign(view_flags.alcv());

            if (view_flags.fWrap())
              out << " wrapText=\"" << view_flags.fWrap() << "\"";
            if (view_flags.fShrinkToFit())
              out << " shrinkToFit=\"" << view_flags.fWrap() << "\"";
            // something is not quite right here.
            // I have a file which is left to right, but this here returns 2: right to left
            if (view_flags.iReadingOrder())
              out << " readingOrder=\"" << (uint16_t)view_flags.iReadingOrder() << "\"";
            if (indent)
              out << " indent=\"" << (uint16_t)indent / 3 << "\"";
            if (view_flags.fJustLast())
              out << " justifyLastLine=\"" << view_flags.fJustLast() << "\"";
            if (trot)
              out << " textRotation=\"" << (uint16_t)trot << "\"";

            out << "/>";
            out << "</xf>" << std::endl;
          } else {
            out << " />" << std::endl;
          }

          // bin.seekg(size, bin.cur);
          break;
        }

        case BrtEndCellXFs: {
          out << "</cellXfs>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        // cellStyles
        case BrtBeginStyles: {
          out << "<cellStyles>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtStyle: {
          uint8_t iStyBuiltIn = 0, iLevel = 0;
          uint16_t grbitObj1 = 0;
          uint32_t ixf = 0;
          ixf = readbin(ixf, bin, swapit);
          grbitObj1 = readbin(grbitObj1, bin, swapit);
          iStyBuiltIn = readbin(iStyBuiltIn, bin, swapit);
          iLevel = readbin(iLevel, bin, swapit);
          std::string stName = XLWideString(bin, swapit);

          StyleFlags view_flags(grbitObj1);

          out << "<cellStyle";
          out << " name=\"" << escape_xml(stName) << "\"";
          out << " xfId=\"" << ixf << "\"";
          if (view_flags.fBuiltIn())
            out << " builtinId=\"" << (int32_t)iStyBuiltIn << "\"";
          if (iStyBuiltIn >= 1 && iStyBuiltIn <= 2)
            out << " iLevel=\"" << (int32_t)iLevel << "\"";
          if (view_flags.fHidden())
            out << " hidden=\"" << (int32_t)view_flags.fHidden() << "\"";
          if (view_flags.fCustom())
            out << " customBuiltin=\"" << (int32_t)view_flags.fCustom() << "\"";
          out << " />" << std::endl;

          break;
        }

        case BrtEndStyles: {
          out << "</cellStyles>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtBeginColorPalette: {
          out << "<colors>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtBeginIndexedColors: {
          out << "<indexedColors>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtIndexedColor: {
          uint8_t bRed = 0, bGreen = 0, bBlue = 0, reserved = 0;
          bRed = readbin(bRed, bin, swapit);
          bGreen = readbin(bGreen, bin, swapit);
          bBlue = readbin(bBlue, bin, swapit);
          reserved = readbin(reserved, bin, swapit);

          // std::vector<int32_t> color = brtColor(bin);
          // if (color[0] == 0x02) {
          out << "<rgbColor rgb=\"" << to_argb(reserved, bRed, bGreen, bRed) << std::dec << "\" />" << std::endl;
          // }
          break;
        }

        case BrtEndIndexedColors: {
          out << "</indexedColors>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtEndColorPalette: {
          out << "</colors>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtBeginDXFs: {
          uint32_t cdxfs = 0;
          cdxfs = readbin(cdxfs, bin, swapit);

          out << "<dxfs count=\"" << cdxfs << "\">" << std::endl;
          break;
        }

        case BrtDXF: {
          std::vector<uint8_t> dxf(static_cast<size_t>(size));
          bin.read(reinterpret_cast<char*>(dxf.data()), size);
          out << xlsb_dxf(dxf) << std::endl;
          break;
        }

        case BrtEndDXFs: {
          out << "</dxfs>" << std::endl;
          break;
        }

        case BrtEndStyleSheet: {
          end_of_style_sheet = true;
          out << "</styleSheet>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        default: {
          if (debug) {  // #nocov start
            Rcpp::Rcout << "Unhandled Style: " << std::to_string(x) << ": " << std::to_string(size) << " @ "
                        << bin.tellg() << std::endl;
          }  // #nocov end
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
int32_t table_bin(std::string filePath, std::string outPath, bool debug) {
  std::ofstream out(outPath);
  std::ifstream bin(filePath, std::ios::in | std::ios::binary | std::ios::ate);

  bool swapit = is_big_endian();

  // auto sas_size = bin.tellg();
  if (bin) {
    bin.seekg(0, std::ios_base::beg);
    bool end_of_table = false;
    bool has_revision_record = false;

    R_xlen_t idx = 0;
    while (!end_of_table) {
      checkInterrupt(idx);

      int32_t x = 0, size = 0;

      if (debug) Rcpp::Rcout << "." << std::endl;
      RECORD(x, size, bin, swapit);
      if (debug) Rcpp::Rcout << x << ": " << size << std::endl;

      switch (x) {
        case BrtUID: {
          if (debug) Rcpp::Rcout << "<xr:uid>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtBeginList: {
          if (debug) Rcpp::Rcout << "<table>" << std::endl;

          std::vector<int32_t> rfxList = UncheckedRfX(bin, swapit);
          std::string ref = int_to_col(rfxList[2] + 1) + std::to_string(rfxList[0] + 1) + ":" + int_to_col(rfxList[3] + 1) + std::to_string(rfxList[1] + 1);

          if (debug) Rcpp::Rcout << "table ref: " << ref << std::endl;

          uint32_t lt = 0, idList = 0, crwHeader = 0, crwTotals = 0, flags = 0;
          lt = readbin(lt, bin, swapit);
          idList = readbin(idList, bin, swapit);
          crwHeader = readbin(crwHeader, bin, swapit);
          crwTotals = readbin(crwTotals, bin, swapit);
          flags = readbin(flags, bin, swapit);

          uint32_t nDxfHeader = 0, nDxfData = 0, nDxfAgg = 0, nDxfBorder = 0, nDxfHeaderBorder = 0, nDxfAggBorder = 0, dwConnID = 0;

          nDxfHeader = readbin(nDxfHeader, bin, swapit);
          nDxfData = readbin(nDxfData, bin, swapit);
          nDxfAgg = readbin(nDxfAgg, bin, swapit);
          nDxfBorder = readbin(nDxfBorder, bin, swapit);
          nDxfHeaderBorder = readbin(nDxfHeaderBorder, bin, swapit);
          nDxfAggBorder = readbin(nDxfAggBorder, bin, swapit);
          dwConnID = readbin(dwConnID, bin, swapit);

          std::string stName = XLNullableWideString(bin, swapit);
          std::string stDisplayName = XLNullableWideString(bin, swapit);
          std::string stComment = XLNullableWideString(bin, swapit);
          std::string stStyleHeader = XLNullableWideString(bin, swapit);
          std::string stStyleData = XLNullableWideString(bin, swapit);
          std::string stStyleAgg = XLNullableWideString(bin, swapit);
          if (debug) {  // #nocov start
            Rcpp::Rcout << "table:" << std::endl;
            Rcpp::Rcout << stName << std::endl;
            Rcpp::Rcout << stDisplayName << std::endl;
            Rcpp::Rcout << stComment << std::endl;
            Rcpp::Rcout << stStyleHeader << std::endl;
            Rcpp::Rcout << stStyleData << std::endl;
            Rcpp::Rcout << stStyleAgg << std::endl;
          }  // #nocov end

          out << "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
                 "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" "
                 "xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" "
                 "xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\" mc:Ignorable=\"xr "
                 "xr3\""
              << " id=\"" << idList << "\" name=\"" << stName << "\" displayName=\"" << stDisplayName << "\" ref=\""
              << ref << "\" totalsRowShown=\"" << crwTotals << "\">" << std::endl;

          // if (debug) Rcpp::Rcout << bin.tellg() << std::endl;
          break;
        }

        // this is used in worksheet as well. move it to function?
        case BrtBeginAFilter: {
          if (debug) Rcpp::Rcout << "<autofilter>" << std::endl;
          std::vector<int32_t> rfx = UncheckedRfX(bin, swapit);

          std::string lref = int_to_col(rfx[2] + 1) + std::to_string(rfx[0] + 1);
          std::string rref = int_to_col(rfx[3] + 1) + std::to_string(rfx[1] + 1);

          std::string ref;
          if (lref.compare(rref) == 0) {
            ref = lref;
          } else {
            ref = lref + ":" + rref;
          }

          // ignoring filterColumn for now
          // autofilter can consist of filterColumn, filters and filter
          // maybe customFilter, dynamicFilter too
          out << "<autoFilter ref=\"" << ref << "\">" << std::endl;

          break;
        }

        case BrtBeginFilterColumn: {
          if (debug) Rcpp::Rcout << "<filterColumn>" << std::endl;

          uint16_t flags = 0;
          uint32_t dwCol = 0;
          dwCol = readbin(dwCol, bin, swapit);
          flags = readbin(flags, bin, swapit);
          // fHideArrow
          // fNoBtn

          out << "<filterColumn colId=\"" << dwCol << "\">" << std::endl;

          break;
        }

        case BrtBeginFilters: {
          if (debug) Rcpp::Rcout << "<filters>" << std::endl;
          // bin.seekg(size, bin.cur);

          uint32_t fBlank = 0, unused = 0;
          fBlank = readbin(fBlank, bin, swapit);  // a 32bit flag, after all ... why not?
          unused = readbin(unused, bin, swapit);  // maybe calendarType?
          out << "<filters blank=\"" << fBlank << "\">" << std::endl;
          break;
        }

        case BrtFilter: {
          std::string rgch = XLWideString(bin, swapit);
          out << "<filter val=\"" << rgch << "\" />" << std::endl;
          // bin.seekg(size, bin.cur);
          break;
        }

        case BrtEndFilters: {
          if (debug) Rcpp::Rcout << "</filters>" << std::endl;
          out << "</filters>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtColorFilter: {
          uint32_t dxfid = 0, fCellColor = 0;
          dxfid = readbin(dxfid, bin, swapit);
          fCellColor = readbin(fCellColor, bin, swapit);

          out << "<colorFilter dxfId=\"" << dxfid << "\" cellColor=\"" << fCellColor << "\"/>" << std::endl;

          break;
        }

        case BrtBeginCustomFilters:
        case BrtBeginCustomFilters14:
        case BrtBeginCustomRichFilters: {
          int32_t fAnd = 0;

          // in xlsb it is flipped
          fAnd = readbin(fAnd, bin, swapit) ^ 1;

          out << "<customFilters" << std::endl;
          if (fAnd) out << " and=\"" << fAnd << "\"";  // and="1"
          out << ">";

          break;
        }

        case BrtCustomFilter:
        case BrtCustomFilter14: {
          int8_t vts = 0, grbitSgn = 0;
          double union_val = 0;
          std::string vtsStringXls;

          vts = readbin(vts, bin, swapit);
          grbitSgn = readbin(grbitSgn, bin, swapit);

          if (vts == 4) {
            // a double
            union_val = Xnum(bin, swapit);
          } else if (vts == 8) {
            // a bool
            int8_t boolean = 0;
            boolean = readbin(boolean, bin, swapit);
            union_val = static_cast<double>(readbin(boolean, bin, swapit));
            for (int8_t blk = 0; blk < 7; ++blk) {
              readbin(boolean, bin, swapit);
            }
          } else {
            // ignore
            readbin(union_val, bin, swapit);
            readbin(union_val, bin, swapit);
          }

          if (vts == 6)  // a string
            vtsStringXls = XLWideString(bin, swapit);

          out << "<customFilter" << std::endl;
          out << " operator=\"" << grbitSgnOperator(grbitSgn) << "\"";
          if (vts == 6)
            out << " val=\"" << vtsStringXls << "\"";
          else
            out << " val=\"" << union_val << "\"";
          out << " />";

          break;
        }

        case BrtEndCustomFilters:
        case BrtEndCustomRichFilters: {
          out << "</customFilters>" << std::endl;

          break;
        }

        case BrtDynamicFilter: {
          Rcpp::warning("Dynamic Filter found. This is not handled.");
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtIconFilter:
        case BrtIconFilter14: {
          uint32_t iIconSet = 0, iIcon = 0;
          iIconSet = readbin(iIconSet, bin, swapit);
          iIcon = readbin(iIcon, bin, swapit);

          std::string iconSet;
          if (iIconSet) iconSet = to_iconset(iIconSet);

          out << "<iconFilter iconSet=\"" << iconSet << "\" iconId=\"" << iIcon << "\"/>" << std::endl;

          break;
        }

        case BrtEndFilterColumn: {
          if (debug) Rcpp::Rcout << "</filterColumn>" << std::endl;
          out << "</filterColumn>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtEndAFilter: {
          if (debug) Rcpp::Rcout << "</autofilter>" << std::endl;
          out << "</autoFilter>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtBeginListCols: {
          uint32_t nCols = 0;
          nCols = readbin(nCols, bin, swapit);
          out << "<tableColumns count=\"" << nCols << "\">" << std::endl;
          // bin.seekg(size, bin.cur);
          break;
        }

        case BrtBeginListCol: {
          if (debug) Rcpp::Rcout << "<tableColumn>" << std::endl;
          // bin.seekg(size, bin.cur);
          // break;

          uint32_t idField = 0, ilta = 0, nDxfHdr = 0, nDxfInsertRow = 0, nDxfAgg = 0, idqsif = 0;

          idField = readbin(idField, bin, swapit);
          ilta = readbin(ilta, bin, swapit);
          nDxfHdr = readbin(nDxfHdr, bin, swapit);
          nDxfInsertRow = readbin(nDxfInsertRow, bin, swapit);
          nDxfAgg = readbin(nDxfAgg, bin, swapit);
          idqsif = readbin(idqsif, bin, swapit);
          std::string stName = XLNullableWideString(bin, swapit);
          std::string stCaption = XLNullableWideString(bin, swapit);
          std::string stTotal = XLNullableWideString(bin, swapit);
          if (debug) Rcpp::Rcout << stName << ": " << stCaption << ": " << stTotal<< std::endl;
          std::string stStyleHeader = XLNullableWideString(bin, swapit);
          std::string stStyleInsertRow = XLNullableWideString(bin, swapit);
          std::string stStyleAgg = XLNullableWideString(bin, swapit);
          if (debug) Rcpp::Rcout << stStyleHeader << ": " << stStyleInsertRow << ": " << stStyleAgg << std::endl;

          out << "<tableColumn id=\"" << idField << "\" name=\"" << stCaption << "\">" << std::endl;

          break;
        }

        case BrtEndListCol: {
          out << "</tableColumn>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtEndListCols: {
          out << "</tableColumns>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtEndList: {
          end_of_table = true;
          out << "</table>" << std::endl;
          break;
        }

        case BrtListCCFmla: {
          // calculated column formula
          uint8_t flags = 0;
          std::string fml;
          flags = readbin(flags, bin, swapit);
          int32_t sharedFormula = false;
          fml = CellParsedFormula(bin, swapit, debug, 0, 0, sharedFormula, has_revision_record);

          // need to write this formula somehwere
          // if (debug) Rcpp::Rcout << fml << std::endl;
          if (debug) Rcpp::warning("Table formulas are not fully implemented.");

          out << "<calculatedColumnFormula>";
          out << fml;
          out << "</calculatedColumnFormula>" << std::endl;

          break;
        }

        case BrtTableStyleClient: {
          uint16_t flags = 0;
          flags = readbin(flags, bin, swapit);
          std::string stStyleName = XLNullableWideString(bin, swapit);

          BrtTableStyleClientFlags view_flags(flags);

          out << "<tableStyleInfo name=\"" << stStyleName << "\"";
          // out <<" showColHeaders=\"" << view_flags.fColumnHeaders << "\"";  // not part of tableStyleInfo?
          out << " showFirstColumn=\"" << view_flags.fFirstColumn() << "\"";
          out << " showLastColumn=\"" << view_flags.fLastColumn() << "\"";
          // out <<" showRowHeaders=\"" << view_flags.fRowHeaders << "\"";  // not part of tableStyleInfo?
          out << " showRowStripes=\"" << view_flags.fRowStripes() << "\"";
          out << " showColumnStripes=\"" << view_flags.fColumnStripes() << "\"";
          out << " />" << std::endl;
          break;
        }

        case BrtRRChgCell:
        case BrtRRDefName: {
          has_revision_record = true;
          // -- have not seen this yet. if it appears, treat it as if a revision record was found --
          // rgce.rgce or rgceOld.rgce in BrtRRDefName
          if (debug) Rcpp::Rcout << "BrtRRChgCell or BrtRRDefName" << std::endl;
          Rcpp::warning("Assuming revision record.");
          bin.seekg(size, bin.cur);
          break;
        }

        default: {
          if (debug) {
            Rcpp::Rcout << std::to_string(x) << ": " << std::to_string(size) << " @ " << bin.tellg() << std::endl;  // #nocov
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
int32_t comments_bin(std::string filePath, std::string outPath, bool debug) {
  std::ofstream out(outPath);
  std::ifstream bin(filePath, std::ios::in | std::ios::binary | std::ios::ate);

  bool swapit = is_big_endian();

  // auto sas_size = bin.tellg();
  if (bin) {
    bin.seekg(0, std::ios_base::beg);
    bool end_of_comments = false;

    R_xlen_t idx = 0;
    while (!end_of_comments) {
      checkInterrupt(idx);

      int32_t x = 0, size = 0;

      if (debug) Rcpp::Rcout << "." << std::endl;
      RECORD(x, size, bin, swapit);
      if (debug) Rcpp::Rcout << x << ": " << size << std::endl;

      switch (x) {
        case BrtBeginComments: {
          out << "<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
                 "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" "
                 "xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" mc:Ignorable=\"xr\">"
              << std::endl;
          break;
        }

        case BrtBeginCommentAuthors: {
          out << "<authors>" << std::endl;
          break;
        }

        case BrtCommentAuthor: {
          std::string author = XLWideString(bin, swapit);
          out << "<author>" << escape_xml(author) << "</author>" << std::endl;
          break;
        }

        case BrtEndCommentAuthors: {
          out << "</authors>" << std::endl;
          break;
        }

        case BrtBeginCommentList: {
          out << "<commentList>" << std::endl;
          break;
        }

        case BrtBeginComment: {
          uint32_t iauthor = 0;
          int32_t guid0 = 0, guid1 = 0, guid2 = 0, guid3 = 0;
          iauthor = readbin(iauthor, bin, swapit);
          std::vector<int32_t> rfx = UncheckedRfX(bin, swapit);

          std::string lref = int_to_col(rfx[2] + 1) + std::to_string(rfx[0] + 1);
          std::string rref = int_to_col(rfx[3] + 1) + std::to_string(rfx[1] + 1);

          std::string ref;
          if (lref.compare(rref) == 0) {
            ref = lref;
          } else {
            ref = lref + ":" + rref;
          }

          std::vector<int32_t> guids(4);
          guids[0] = readbin(guid0, bin, 0);
          guids[1] = readbin(guid1, bin, 0);
          guids[2] = readbin(guid2, bin, 0);
          guids[3] = readbin(guid3, bin, 0);

          out << "<comment";
          out << " ref=\"" << ref << "\"";
          out << " authorId=\"" << iauthor << "\"";
          // out << " shapeId=\"0\"";  // always?
          out << " >" << std::endl;

          break;
        }

        case BrtEndComment: {
          out << "</comment>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtEndCommentList: {
          out << "</commentList>" << std::endl;
          break;
        }

        case BrtCommentText: {
          // we do not handle RichStr correctly. Ignore all formatting
          std::string commentText = RichStr(bin, swapit);
          // out << "<text><r><rPr><sz val=\"10\"/><color rgb=\"FF000000\"/><rFont val=\"Tahoma\"/><family val=\"2\"/></rPr><t>" << escape_xml(commentText) << "</t></r></text>" << std::endl;
          out << "<text>" << commentText << "</text>" << std::endl;
          break;
        }

        case BrtEndComments: {
          end_of_comments = true;
          out << "</comments>" << std::endl;
          break;
        }

        default: {
          if (debug) {
            Rcpp::Rcout << std::to_string(x) << ": " << std::to_string(size) << " @ " << bin.tellg() << std::endl;  // #nocov
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
int32_t externalreferences_bin(std::string filePath, std::string outPath, bool debug) {
  std::ofstream out(outPath);
  std::ifstream bin(filePath, std::ios::in | std::ios::binary | std::ios::ate);

  bool swapit = is_big_endian();

  // auto sas_size = bin.tellg();
  if (bin) {
    bin.seekg(0, std::ios_base::beg);
    bool first_row = true;
    int32_t row = 0;
    bool end_of_external_reference = false;

    R_xlen_t idx = 0;
    while (!end_of_external_reference) {
      checkInterrupt(idx);

      int32_t x = 0, size = 0;

      if (debug) Rcpp::Rcout << "." << std::endl;
      RECORD(x, size, bin, swapit);
      if (debug) Rcpp::Rcout << x << ": " << size << std::endl;

      switch (x) {
        case BrtBeginSupBook: {
          uint16_t sbt = 0;
          sbt = readbin(sbt, bin, swapit);
          // wb. 0x0000
          // dde 0x0001
          // ole 0x0002

          std::string string1, string2;
          string1 = XLWideString(bin, swapit);

          if (sbt == 0x0000)
            string2 = XLNullableWideString(bin, swapit);
          else
            string2 = XLWideString(bin, swapit);

          if (debug) Rcpp::Rcout << string1 << ": " << string2 << std::endl;

          // no begin externalSheet?
          out << "<externalLink xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
                 "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" "
                 "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" "
                 "xmlns:xxl21=\"http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021\" "
                 "mc:Ignorable=\"x14 xxl21\">"
              << std::endl;

          // assume for now all external references are workbooks
          out << "<externalBook xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\""
              << string1 << "\">" << std::endl;

          break;
        }

          // looks like this file has case 5108:
          // xxl21:absoluteUrl

        case BrtSupTabs: {
          if (debug) Rcpp::Rcout << "<BrtSupTabs>" << std::endl;
          uint32_t cTab = 0;

          cTab = readbin(cTab, bin, swapit);
          if (cTab > 65535) Rcpp::stop("cTab to large");

          out << "<sheetNames>" << std::endl;

          for (uint32_t i = 0; i < cTab; ++i) {
            std::string sheetName = XLWideString(bin, swapit);
            out << "<sheetName val=\"" << escape_xml(sheetName) << "\" />" << std::endl;
          }

          out << "</sheetNames>" << std::endl;

          // just guessing
          out << "<sheetDataSet>" << std::endl;

          break;
        }

        case BrtExternTableStart: {
          if (debug) Rcpp::Rcout << "<BrtExternTableStart>" << std::endl;
          uint8_t flags = 0;
          uint32_t iTab = 0;
          iTab = readbin(iTab, bin, swapit);
          flags = readbin(flags, bin, swapit);

          first_row = true;
          out << "<sheetData sheetId=\"" << iTab << "\">" << std::endl;

          break;
        }

        case BrtExternRowHdr: {
          if (debug) Rcpp::Rcout << "<BrtExternRowHdr>" << std::endl;

          row = UncheckedRw(bin, swapit);

          // close open rows
          if (!first_row) {
            out << "</row>" << std::endl;
          } else {
            first_row = false;
          }
          out << "<row r=\"" << row + 1 << "\">" << std::endl;

          break;
        }

        case BrtExternCellBlank: {
          if (debug) Rcpp::Rcout << "<BrtExternCellBlank>" << std::endl;

          int32_t col = UncheckedCol(bin, swapit);
          out << "<cell r=\"" << int_to_col(col + 1) << row + 1 << "\" />" << std::endl;

          break;
        }

        case BrtExternCellBool: {
          if (debug) Rcpp::Rcout << "<BrtExternCellBool>" << std::endl;

          uint8_t value = 0;
          int32_t col = UncheckedCol(bin, swapit);
          value = readbin(value, bin, swapit);
          out << "<cell r=\"" << int_to_col(col + 1) << row + 1 << "\" t=\"b\">" << std::endl;
          out << "<v>" << (uint16_t)value << "</v>" << std::endl;
          out << "</cell>" << std::endl;

          break;
        }

        case BrtExternCellError: {
          if (debug) Rcpp::Rcout << "<BrtExternCellError>" << std::endl;

          std::string value;
          int32_t col = UncheckedCol(bin, swapit);
          value = BErr(bin, swapit);
          out << "<cell r=\"" << int_to_col(col + 1) << row + 1 << "\" t=\"e\">" << std::endl;
          out << "<v>" << value << "</v>" << std::endl;
          out << "</cell>" << std::endl;

          break;
        }

        case BrtExternCellReal: {
          if (debug) Rcpp::Rcout << "<BrtExternCellReal>" << std::endl;

          double value = 0;
          int32_t col = UncheckedCol(bin, swapit);
          value = Xnum(bin, swapit);
          out << "<cell r=\"" << int_to_col(col + 1) << row + 1 << "\">" << std::endl;
          out << "<v>" << std::setprecision(16) << value << "</v>" << std::endl;
          out << "</cell>" << std::endl;

          break;
        }

        case BrtExternCellString: {
          if (debug) Rcpp::Rcout << "<BrtExternCellString>" << std::endl;

          int32_t col = UncheckedCol(bin, swapit);
          if (col > 16384) Rcpp::stop("Column out of bounds");
          std::string value = XLWideString(bin, swapit);
          if (value.length() >= 32768) Rcpp::stop("String to long");
          out << "<cell r=\"" << int_to_col(col + 1) << row + 1 << "\" t=\"str\">" << std::endl;
          out << "<v>" << escape_xml(value) << "</v>" << std::endl;
          out << "</cell>" << std::endl;

          break;
        }

        case BrtExternTableEnd: {
          if (debug) Rcpp::Rcout << "<BrtExternTableEnd>" << std::endl;

          if (!first_row) out << "</row>" << std::endl;
          out << "</sheetData>" << std::endl;

          break;
        }

        case BrtEndSupBook: {
          end_of_external_reference = true;
          out << "</sheetDataSet>" << std::endl;
          out << "</externalBook>" << std::endl;
          out << "</externalLink>" << std::endl;
          break;
        }

        case BrtACBegin:
        case BrtExternalLinksAlternateUrls:
        case BrtACEnd: {
          // unhandled
          bin.seekg(size, bin.cur);
          break;
        }

        default: {
          Rcpp::Rcout << "Unhandled ER: " << std::to_string(x) << ": " << std::to_string(size) << " @ " << bin.tellg()
                      << std::endl;
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
int32_t sharedstrings_bin(std::string filePath, std::string outPath, bool debug) {
  std::ofstream out(outPath);
  std::ifstream bin(filePath, std::ios::in | std::ios::binary | std::ios::ate);

  bool swapit = is_big_endian();

  // auto sas_size = bin.tellg();
  if (bin) {
    bin.seekg(0, std::ios_base::beg);
    bool end_of_shared_strings = false;

    R_xlen_t idx = 0;
    while (!end_of_shared_strings) {
      checkInterrupt(idx);

      int32_t x = 0, size = 0;

      if (debug) Rcpp::Rcout << "." << std::endl;
      RECORD(x, size, bin, swapit);
      if (debug) Rcpp::Rcout << x << ": " << size << std::endl;

      switch (x) {
        case BrtBeginSst: {
          uint32_t count = 0, uniqueCount = 0;
          count = readbin(count, bin, swapit);
          uniqueCount = readbin(uniqueCount, bin, swapit);
          out << "<sst " << "count=\"" << count << "\" uniqueCount=\"" << uniqueCount << "\">" << std::endl;
          break;
        }

        case BrtSSTItem: {
          std::string val;
          std::streampos pos = bin.tellg();
          pos += static_cast<std::streampos>(size);
          val += RichStr(bin, swapit);
          if (bin.tellg() < pos) {
            // if (debug) {
            // some RichStr() behave different to what is documented. Not sure
            // if this is padding or something else. 12 bytes seems common
            size_t missing = static_cast<size_t>(pos - bin.tellg());
            Rcpp::Rcout << "BrtSSTItem skipping ahead (bytes): " << missing << std::endl;
            // }
            bin.seekg(pos, bin.beg);
          }
          // if (debug)
          out << "<si>" << val << "</si>" << std::endl;
          break;
        }

        case BrtEndSst: {
          end_of_shared_strings = true;
          out << "</sst>" << std::endl;
          break;
        }

        default: {
          Rcpp::Rcout << std::to_string(x) << ": " << std::to_string(size) << " @ " << bin.tellg() << std::endl;
          Rcpp::stop("nonsense");
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
int32_t workbook_bin(std::string filePath, std::string outPath, bool debug) {
  std::ofstream out(outPath);
  std::ifstream bin(filePath, std::ios::in | std::ios::binary | std::ios::ate);

  bool swapit = is_big_endian();

  // auto sas_size = bin.tellg();
  if (bin) {
    bin.seekg(0, std::ios_base::beg);
    bool end_of_workbook = false;
    bool first_extern_sheet = true;
    bool has_revision_record = false;

    std::vector<std::string> defNams, xtis, reference_type, customWorkbookView, pivotCaches;
    defNams.push_back("<definedNames>");
    xtis.push_back("<xtis>");

    R_xlen_t idx = 0;
    while (!end_of_workbook) {
      checkInterrupt(idx);

      int32_t x = 0, size = 0;

      if (debug) Rcpp::Rcout << "." << std::endl;
      RECORD(x, size, bin, swapit);

      switch (x) {
        case BrtBeginBook: {
          if (debug) Rcpp::Rcout << "<workbook>" << std::endl;
          out << "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
                 "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
                 "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" "
                 "xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\" "
                 "xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" "
                 "xmlns:xr6=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision6\" "
                 "xmlns:xr10=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision10\" "
                 "xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\" mc:Ignorable=\"x15 xr "
                 "xr6 xr10 xr2\">"
              << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtFileVersion: {
          if (debug) Rcpp::Rcout << "<fileVersion>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtWbProp: {
          if (debug) Rcpp::Rcout << "<workbookProperties>" << std::endl;
          uint32_t flags = 0, dwThemeVersion = 0;

          flags = readbin(flags, bin, swapit);
          dwThemeVersion = readbin(dwThemeVersion, bin, swapit);
          std::string strName = XLWideString(bin, swapit);

          BrtWbPropFlags view_flags(flags);

          std::string showObjects;
          if (view_flags.mdDspObj() == 0) showObjects = "all";
          if (view_flags.mdDspObj() == 1) showObjects = "placeholders";
          if (view_flags.mdDspObj() == 2) showObjects = "none";

          std::string updateLinks;
          if (view_flags.grbitUpdateLinks() == 0) updateLinks = "userSet";
          if (view_flags.grbitUpdateLinks() == 1) updateLinks = "never";
          if (view_flags.grbitUpdateLinks() == 2) updateLinks = "always";

          out <<
            "<workbookPr" <<
            // " allowRefreshQuery=\"" <<  << "\" " <<
            " autoCompressPictures=\"" << view_flags.fAutoCompressPictures() << "\" " <<
            " backupFile=\"" << view_flags.fBackup() << "\" " <<
            " checkCompatibility=\"" << view_flags.fCheckCompat() << "\" " <<
            " codeName=\"" << strName << "\" " <<
            " date1904=\"" << view_flags.f1904() << "\" " <<
            " defaultThemeVersion=\"" << dwThemeVersion << "\" " <<
            " filterPrivacy=\"" << view_flags.fFilterPrivacy() << "\" " <<
            " hidePivotFieldList=\"" << view_flags.fHidePivotTableFList() << "\" " <<
            " promptedSolutions=\"" << view_flags.fBuggedUserAboutSolution() << "\" " <<
            " publishItems=\"" << view_flags.fPublishedBookItems() << "\" " <<
            " refreshAllConnections=\"" << view_flags.fRefreshAll() << "\" " <<
            " saveExternalLinkValues=\"" << view_flags.fNoSaveSup() << "\" " <<
            // " showBorderUnselectedTables=\"" <<  << "\" " <<
            " showInkAnnotation=\"" << view_flags.fShowInkAnnotation() << "\" ";
            if (showObjects != "all") out << " showObjects=\"" << showObjects << "\"";
            out << " showPivotChartFilter=\"" << view_flags.fShowPivotChartFilter() << "\" ";
            if (updateLinks != "userSet") out << " updateLinks=\"" << updateLinks <<  "\" ";
            out << "/>" << std::endl;

          break;
        }

        case BrtACBegin: {
          if (debug) Rcpp::Rcout << "<BrtACBegin>" << std::endl;
          // bin.seekg(size, bin.cur);

          uint16_t cver = 0;
          cver = readbin(cver, bin, swapit);

          for (uint16_t i = 0; i < cver; ++i) {
            ProductVersion(bin, swapit, debug, false);  // reserved can be 0 or 1
          }
          break;
        }

        case BrtAbsPath15: {
          std::string absPath = XLWideString(bin, swapit);
          if (debug) Rcpp::Rcout << absPath << std::endl;
          break;
        }

        case BrtACEnd: {
          if (debug) Rcpp::Rcout << "<BrtACEnd>" << std::endl;
          break;
        }

        case BrtRevisionPtr: {
          if (debug) Rcpp::Rcout << "<BrtRevisionPtr>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtUID: {
          if (debug) Rcpp::Rcout << "<BrtUID>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtBeginBookViews: {
          if (debug) Rcpp::Rcout << "<workbookViews>" << std::endl;
          out << "<bookViews>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtBookView: {
          if (debug) Rcpp::Rcout << "<workbookView>" << std::endl;
          // bin.seekg(size, bin.cur);
          uint8_t flags = 0;
          int32_t xWn = 0, yWn = 0;
          uint32_t dxWn = 0, dyWn = 0, iTabRatio = 0, itabFirst = 0, itabCur = 0;

          xWn       = readbin(xWn, bin, swapit);
          yWn       = readbin(yWn, bin, swapit);
          dxWn      = readbin(dxWn, bin, swapit);
          dyWn      = readbin(dyWn, bin, swapit);
          iTabRatio = readbin(iTabRatio, bin, swapit);
          itabFirst = readbin(itabFirst, bin, swapit);
          itabCur   = readbin(itabCur, bin, swapit);
          flags     = readbin(flags, bin, swapit);

          out << "<workbookView xWindow=\"" << xWn << "\" yWindow=\"" << yWn << "\" windowWidth=\"" << dxWn
              << "\" windowHeight=\"" << dyWn << "\" activeTab=\"" << itabCur << "\" />" << std::endl;
          break;
        }

        case BrtEndBookViews: {
          if (debug) Rcpp::Rcout << "</workbookViews>" << std::endl;
          out << "</bookViews>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtUserBookView: {
          if (debug) Rcpp::Rcout << "<customWorkbookView>" << std::endl;

          std::ostringstream cwv;

          uint32_t flags = 0;
          int32_t xLeft = 0, xRight = 0, yTop = 0, yBot = 0, iTabid = 0, iTabRatio = 0, guid0 = 0, guid1 = 0, guid2 = 0, guid3 = 0;
          int16_t wMergeInterval = 0;
          std::string stName;

          xLeft = readbin(xLeft, bin, swapit);
          xRight = readbin(xRight, bin, swapit);
          yTop = readbin(yTop, bin, swapit);
          yBot = readbin(yBot, bin, swapit);
          iTabid = readbin(iTabid, bin, swapit);
          iTabRatio = readbin(iTabRatio, bin, swapit);

          std::vector<int32_t> guids(4);
          guids[0] = readbin(guid0, bin, 0);
          guids[1] = readbin(guid1, bin, 0);
          guids[2] = readbin(guid2, bin, 0);
          guids[3] = readbin(guid3, bin, 0);

          wMergeInterval = readbin(wMergeInterval, bin, swapit);

          flags = readbin(flags, bin, swapit);
          BrtUserBookViewFlags view_flags(flags);

          stName = XLWideString(bin, swapit);

          std::string showComments;
          if (view_flags.mdDspNote() == 0) showComments = "commNone";
          if (view_flags.mdDspNote() == 1) showComments = "commIndAndComment";
          if (view_flags.mdDspNote() == 2) showComments = "commIndicator";

          std::string showObjects;
          if (view_flags.mdHideObj() == 0) showObjects = "all";
          if (view_flags.mdHideObj() == 1) showObjects = "placeholders";
          if (view_flags.mdHideObj() == 2) showObjects = "none";

          cwv << "<customWorkbookView" << std::endl;
          cwv << " name=\"" << stName << "\"" << std::endl;
          cwv << " guid=\"{" << guid_str(guids) << "}\"" << std::endl;
          if (view_flags.fTimedUpdate()) cwv << " autoUpdate=\"" << view_flags.fTimedUpdate() << "\"" << std::endl;
          if (view_flags.fTimedUpdate()) cwv << " mergeInterval=\"" << wMergeInterval << "\"" << std::endl;
          if (view_flags.fAllMemChanges()) cwv << " changesSavedWin=\"" << view_flags.fAllMemChanges() << "\"" << std::endl;
          if (view_flags.fOnlySync()) cwv << " onlySync=\"" << view_flags.fOnlySync() << "\"" << std::endl;
          if (view_flags.fPersonalView()) cwv << " personalView=\"" << view_flags.fPersonalView() << "\"" << std::endl;
          if (!view_flags.fPrintIncl()) cwv << " includePrintSettings=\"" << view_flags.fPrintIncl() << "\"" << std::endl;
          // wrong?
          // if (!view_flags.fRowColIncl) cwv << " includeHiddenRowCol=\"" << view_flags.fRowColIncl << "\"" << std::endl;
          if (view_flags.fZoom()) cwv << " maximized=\"" << view_flags.fZoom() << "\"" << std::endl;
          if (view_flags.fIconic()) cwv << " minimized=\"" << view_flags.fIconic() << "\"" << std::endl;
          if (!view_flags.fDspHScroll()) cwv << " showHorizontalScroll=\"" << view_flags.fDspHScroll() << "\"" << std::endl;
          if (!view_flags.fDspVScroll()) cwv << " showVerticalScroll=\"" << view_flags.fDspVScroll() << "\"" << std::endl;
          if (!view_flags.fBotAdornment()) cwv << " showSheetTabs=\"" << view_flags.fBotAdornment() << "\"" << std::endl;
          if (xLeft > 0) cwv << " xWindow=\"" << xLeft << "\"" << std::endl;
          if (yTop > 0) cwv << " yWindow=\"" << yTop << "\"" << std::endl;
          if (xRight > 0) cwv << " windowWidth=\"" << xRight << "\"" << std::endl;
          if ((yBot - yTop) > 0) cwv << " windowHeight=\"" << (yBot - yTop) << "\"" << std::endl;
          if (iTabRatio != 600) cwv << " tabRatio=\"" << iTabRatio << "\"" << std::endl;
          cwv << " activeSheetId=\"" << iTabid << "\"" << std::endl;
          if (!view_flags.fDspFmlaBar()) cwv << " showFormulaBar=\"" << view_flags.fDspFmlaBar() << "\"" << std::endl;
          if (!view_flags.fDspStatus()) cwv << " showStatusbar=\"" << view_flags.fDspStatus() << "\"" << std::endl;
          if (showComments != "commIndicator") cwv << " showComments=\"" << showComments << "\"" << std::endl;
          if (showObjects != "all") cwv << " showObjects=\"" << showObjects << "\"" << std::endl;
          cwv << "/>" << std::endl;

          // Rcpp::Rcout << xLeft << ": " << xRight << ": " << yTop << ": " << yBot << std::endl;

          customWorkbookView.push_back(cwv.str());

          break;
        }

        case BrtBeginBundleShs: {
          if (debug) Rcpp::Rcout << "<sheets>" << std::endl;
          // unk = readbin(unk, bin, swapit);
          // unk = readbin(unk, bin, swapit);
          // uint32_t count, uniqueCount;
          // count = readbin(count, bin, swapit);
          // uniqueCount = readbin(uniqueCount, bin, swapit);
          out << "<sheets>" << std::endl;
          break;
        }

        case BrtBundleSh: {
          if (debug) Rcpp::Rcout << "<sheet>" << std::endl;

          uint32_t hsState = 0, iTabID = 0;  //  strRelID ???

          hsState = readbin(hsState, bin, swapit);
          iTabID = readbin(iTabID, bin, swapit);
          std::string rid = XLNullableWideString(bin, swapit);

          if (debug) Rcpp::Rcout << "sheet vis: " << hsState << ": " << iTabID << ": " << rid << std::endl;

          std::string val = XLWideString(bin, swapit);

          out << "<sheet r:id=\"" << rid << "\" state=\"" << HSState(hsState)<< "\" sheetId=\"" << iTabID<< "\" name=\"" << val << "\"/>" << std::endl;
          break;
        }

        case BrtEndBundleShs: {
          if (debug) Rcpp::Rcout << "</sheets>" << std::endl;
          out << "</sheets>" << std::endl;
          break;
        }

        case BrtCalcProp: {
          if (debug) Rcpp::Rcout << "<calcPr>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtName: {
          if (debug)  Rcpp::Rcout << "<BrtName>" << std::endl;
          std::streampos end_pos = bin.tellg();
          end_pos += static_cast<std::streampos>(size);

          if (debug) Rcpp::Rcout << "BrtName endpos: "<< end_pos << std::endl;

          uint8_t chKey = 0;
          uint16_t BrtNameUint = 0, BrtNameUint2 = 0;
          uint32_t itab = 0;
          BrtNameUint = readbin(BrtNameUint, bin, swapit);
          BrtNameUint2 = readbin(BrtNameUint2, bin, swapit);

          BrtNameFlags view_flags(BrtNameUint);
          BrtName2Flags view_flags2(BrtNameUint2);

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

          if (view_flags2.reserved() != 0)
            Rcpp::stop("reserved not 0");

          if (debug) {  // #nocov start
            Rprintf(
              "%d, %d, %d, %d, %d, %d, %d, %d, %d, %d, %d\n",
              (int32_t)view_flags.fHidden(),
              (int32_t)view_flags.fFunc(),
              (int32_t)view_flags.fOB(),
              (int32_t)view_flags.fProc(),
              (int32_t)view_flags.fCalcExp(),
              (int32_t)view_flags.fBuiltin(),
              (int32_t)view_flags.fgrp(),
              (int32_t)view_flags.fPublished(),
              (int32_t)view_flags2.fWorkbookParam(),
              (int32_t)view_flags2.fFutureFunction(),
              (int32_t)view_flags2.reserved()
            );
          }  // #nocov end

          chKey = readbin(chKey, bin, swapit);
          // ascii key (0 if fFunc = 1 or fProc = 0 else >= 0x20)

          itab = readbin(itab, bin, swapit);

          // XLNameWideString: XLWideString <= 255 characters
          std::string name = XLWideString(bin, swapit);

          // if ((size_t)bin.tellg() < end_pos) {
          //   Rcpp::Rcout << "repositioning" << std::endl;
          //   Rcpp::Rcout << end_pos << std::endl;
          //   bin.seekg(end_pos, bin.beg);
          // }

          std::string fml = "", comment = "";

          int32_t sharedFormula = false;
          fml = CellParsedFormula(bin, swapit, debug, 0, 0, sharedFormula, has_revision_record);

          comment = XLNullableWideString(bin, swapit);

          if (debug) Rcpp::Rcout << "definedName: " << name << ": " << comment << std::endl;

          if (view_flags.fProc() && fml.compare("") != 0) {
            /* -- something is wrong. error with some nhs macro xlsb file -- */
            // must be NULL
            if (debug) Rcpp::Rcout << 1 << std::endl;
            std::string unusedstring1 = XLNullableWideString(bin, swapit);

            // must be < 32768 characters
            if (debug) Rcpp::Rcout << 2 << std::endl;
            std::string description = XLNullableWideString(bin, swapit);
            if (debug) Rcpp::Rcout << 3 << std::endl;
            std::string helpTopic = XLNullableWideString(bin, swapit);

            // must be NULL
            if (debug) Rcpp::Rcout << 4 << std::endl;
            std::string unusedstring2 = XLNullableWideString(bin, swapit);
          }

          if (bin.tellg() != end_pos) {
            Rprintf("%d: %d", (int)bin.tellg(), (int)end_pos);
            Rcpp::stop("repositioning");
          }

          if (view_flags.fBuiltin()) {
            // add the builtin xl name namespace
            if (name.find("_xlnm.") == std::string::npos)
              name = "_xlnm." + name;
          }

          std::string defNam = "<definedName name=\"" + name;

          if (comment.size() > 0)
            defNam += "\" comment=\"" + comment;

          if (itab != 0xFFFFFFFF)
            defNam += "\" localSheetId=\"" + std::to_string(itab);

          if (view_flags.fHidden())
            defNam += "\" hidden=\"" + std::to_string(view_flags.fHidden());

          // lacks the formula for the defined name
          defNam = defNam + "\">" + fml + "</definedName>";

          defNams.push_back(defNam);

          break;
        }

        case BrtBeginPivotCacheIDs:
        case BrtEndPivotCacheIDs:
        case BrtEndPivotCacheID:
          break;

        case BrtBeginPivotCacheID: {
          uint32_t idSx = 0;
          idSx = readbin(idSx, bin, swapit);
          std::string relId = XLNullableWideString(bin, swapit);
          pivotCaches.push_back("<pivotCache cacheId=\"" + std::to_string(idSx) + "\" r:id=\"" + relId + "\"/>");
          break;
        }

        case BrtBeginExternals: {
          if (debug) Rcpp::Rcout << "<BrtBeginExternals>" << std::endl;
          // bin.seekg(size, bin.cur);
          break;
        }

          // part of Xti
          // * BrtSupSelf self reference
          // * BrtSupSame same sheet
          // * BrtSupAddin addin XLL
          // * BrtSupBookSrc external link
        case BrtSupSelf: {
          if (debug) Rcpp::Rcout << "BrtSupSelf @" << bin.tellg() << std::endl;
          // Rcpp::Rcout << "<internalReference type=\"0\"/>" << std::endl;
          reference_type.push_back("0");
          break;
        }

        case BrtSupSame: {
          if (debug) Rcpp::Rcout << "BrtSupSame @" << bin.tellg() << std::endl;
          // Rcpp::Rcout << "<internalReference type=\"1\"/>" << std::endl;
          reference_type.push_back("1");
          break;
        }

        case BrtPlaceholderName: {
          if (debug) Rcpp::Rcout << "BrtPlaceholderName @" << bin.tellg() << std::endl;
          std::string name = XLWideString(bin, swapit);
          // Rcpp::Rcout << "<internalReference name=\""<< name << "\"/>" << std::endl;
          break;
        }

        case BrtSupAddin: {
          if (debug) Rcpp::Rcout << "BrtSupAddin @" << bin.tellg() << std::endl;
          // Rcpp::Rcout << "<internalReference type=\"2\"/>" << std::endl;
          reference_type.push_back("2");
          break;
        }

        case BrtSupBookSrc: {
          if (first_extern_sheet) {
            out << "<externalReferences>" << std::endl;
            first_extern_sheet = false;
          }

          if (debug) Rcpp::Rcout << "BrtSupBookSrc @" << bin.tellg() << std::endl;
          // Rcpp::stop("BrtSupSelf");
          std::string strRelID = XLNullableWideString(bin, swapit);
          out << "<externalReference r:id=\"" << strRelID << "\"/>" << std::endl;
          reference_type.push_back(strRelID);
          break;
        }

        case BrtSupTabs: {
          if (debug) Rcpp::Rcout << "<BrtSupTabs>" << std::endl;
          uint32_t cTab = 0;

          cTab = readbin(cTab, bin, swapit);
          if (cTab > 65535) Rcpp::stop("cTab to large");

          for (uint32_t i = 0; i < cTab; ++i) {
            std::string sheetName = XLWideString(bin, swapit);
            Rcpp::Rcout << "BrtSupTabs: " << sheetName << std::endl;
          }

          break;
        }

        case BrtExternSheet: {
          if (debug) Rcpp::Rcout << "<BrtExternSheet>" << std::endl;
          uint32_t cXti = 0;

          cXti = readbin(cXti, bin, swapit);
          if (cXti > 65535) Rcpp::stop("cXti to large");

          std::vector<uint32_t> rgXti(cXti);
          // Rcpp::Rcout << "reference_type: " << reference_type.size() << std::endl;
          // Rcpp::Rcout << "reference_type: " << cXti << std::endl;
          for (uint32_t i = 0; i < cXti; ++i) {
            std::vector<int32_t> xti = Xti(bin, swapit);
            if ((size_t)xti[0] > reference_type.size()) Rcpp::stop("references do not match");
            std::string tmp = "<xti id=\"" + std::to_string(xti[0]) +
              "\" firstSheet=\"" + std::to_string(xti[1]) +
              "\" lastSheet=\"" +  std::to_string(xti[2]) +
              "\" type=\"" +  reference_type[static_cast<size_t>(xti[0])] +
              "\" />";
            xtis.push_back(tmp);
          }

          break;
        }

        case BrtEndExternals: {
          if (debug) Rcpp::Rcout << "</BrtEndExternals>" << std::endl;
          if (!first_extern_sheet)
            out << "</externalReferences>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtFRTBegin: {
          if (debug) Rcpp::Rcout << "<ext>" << std::endl;

          ProductVersion(bin, swapit, debug, true);  // reserved must be 0
          break;
        }

        case BrtWorkBookPr15: {
          if (debug) Rcpp::Rcout << "<BrtWorkBookPr15>" << std::endl;
          bin.seekg(size, bin.cur);
          // uint8_t fChartTrackingRefBased = 0;
          // uint32_t FRTHeader = 0;
          // FRTHeader = readbin(FRTHeader, bin, swapit);
          // fChartTrackingRefBased = readbin(fChartTrackingRefBased, bin, 0) & 0x01;
          break;
        }

        case BrtBeginCalcFeatures: {
          if (debug) Rcpp::Rcout << "<calcs>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtCalcFeature: {
          if (debug) Rcpp::Rcout << "<calc>" << std::endl;
          // bin.seekg(size, bin.cur);
          uint32_t FRTHeader = 0;
          FRTHeader = readbin(FRTHeader, bin, swapit);
          std::string szName = XLWideString(bin, swapit);

          if (debug) Rcpp::Rcout << FRTHeader << ": " << szName << std::endl;
          break;
        }

        case BrtEndCalcFeatures: {
          if (debug) Rcpp::Rcout << "</calcs>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtFRTEnd: {
          if (debug) Rcpp::Rcout << "</ext>" << std::endl;
          break;
        }

        case BrtWbFactoid: {  // ???
          if (debug) Rcpp::Rcout << "<BrtWbFactoid>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtFileRecover: {
          if (debug) Rcpp::Rcout << "<fileRecovery>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtEndBook: {
          end_of_workbook = true;

          if (defNams.size() > 1) {
            defNams.push_back("</definedNames>");

            for (size_t i = 0; i < defNams.size(); ++i) {
              if (debug) Rcpp::Rcout << defNams[i] << std::endl;
              out << defNams[i] << std::endl;
            }
          }

          // custom xti output
          if (xtis.size() > 1) {
            xtis.push_back("</xtis>");

            for (size_t i = 0; i < xtis.size(); ++i) {
              if (debug) Rcpp::Rcout << xtis[i] << std::endl;
              out << xtis[i] << std::endl;
            }
          }

          if (customWorkbookView.size()) {
            out << "<customWorkbookViews>" << std::endl;
            for (size_t i = 0; i < customWorkbookView.size(); ++i) {
              if (debug) Rcpp::Rcout << customWorkbookView[i] << std::endl;
              out << customWorkbookView[i] << std::endl;
            }
            out << "</customWorkbookViews>" << std::endl;
          }


          if (pivotCaches.size()) {
            out << "<pivotCaches>" << std::endl;
            for (size_t i = 0; i < pivotCaches.size(); ++i) out << pivotCaches[i] << std::endl;
            out << "</pivotCaches>" << std::endl;
          }

          if (debug) Rcpp::Rcout << "</workbook>" << std::endl;
          out << "</workbook>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtRRChgCell:
        case BrtRRDefName: {
          has_revision_record = true;
          // -- have not seen this yet. if it appears, treat it as if a revision record was found --
          // rgce.rgce or rgceOld.rgce in BrtRRDefName
          if (debug) Rcpp::Rcout << "BrtRRChgCell or BrtRRDefName" << std::endl;
          Rcpp::warning("Assuming revision record.");
          bin.seekg(size, bin.cur);
          break;
        }

        default: {
          if (debug) {
            Rcpp::Rcout << std::to_string(x) << ": " << std::to_string(size) << " @ " << bin.tellg() << std::endl;  // #nocov
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
int32_t worksheet_bin(std::string filePath, bool chartsheet, std::string outPath, bool debug) {
  std::ofstream out(outPath);
  std::ifstream bin(filePath, std::ios::in | std::ios::binary | std::ios::ate);

  bool swapit = is_big_endian();

  // auto sas_size = bin.tellg();
  if (bin) {
    bin.seekg(0, std::ios_base::beg);

    bool first_row = true;
    uint16_t def_row_twips = 0;
    bool in_sheet_data = false;
    bool end_of_worksheet = false;
    bool extlst_open = false, ext_open = false;
    bool cf_iconset = false;
    int32_t cf_row = 0, cf_col = 0;
    auto open_ext = [&](const char *uri) {
      if (!extlst_open) out << "<extLst>";
      extlst_open = ext_open = true;
      out << "<ext uri=\"" << uri << "\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">" << std::endl;
    };
    bool has_revision_record = false;
    std::string fml_type;

    int32_t row = 0;
    int32_t col = 0;
    std::vector<std::string> hlinks;
    hlinks.push_back("<hyperlinks>");

    // its a bit funny that this is the structure we now use to read from the
    // xlsb file and save this as xml only to read it back into this structure
    // with the xlsx reader :)
    std::vector<xml_col> colvec;
    std::unordered_map<std::string, int> shared_cells;
    int32_t shared_cell_cntr = 0;

    // auto itr = 0;
    R_xlen_t idx = 0;
    while (!end_of_worksheet) {
      checkInterrupt(idx);

      // uint8_t unk = 0, high = 0, low = 0;
      // uint16_t tmp = 0;
      int32_t x = 0, size = 0;

      if (debug) Rcpp::Rcout << "." << std::endl;

      RECORD(x, size, bin, swapit);
      if (debug) Rcpp::Rcout << x << ": " << size << std::endl;

      switch(x) {

        case BrtBeginSheet: {
          if (chartsheet)
            out << "<chartsheet>" << std::endl;
          else
            out << "<worksheet>" << std::endl;

          if (debug) Rcpp::Rcout << "Begin of <worksheet>: " << bin.tellg() << std::endl;
          break;
        }

        case BrtCsProp: {
          if (debug) Rcpp::Rcout << "CsProp: " << bin.tellg() << std::endl;

          uint16_t A = 0;
          A = readbin(A, bin, swapit);

          std::vector<int32_t> color = brtColor(bin, swapit);
          std::string strName = XLWideString(bin, swapit);

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
              out << "<tabColor theme=\"" << color[1] << "\" />" << std::endl;  //  << "\" tint=\""<< color[2]
            }

            out << "</sheetPr>" << std::endl;
          }

          break;
        }

        case BrtWsProp: {
          if (debug) Rcpp::Rcout << "WsProp: " << bin.tellg() << std::endl;

          // read uint24_t as 16 + 8
          uint8_t brtwsprop_sec = 0;
          uint16_t brtwsprop_fst = 0;
          uint32_t rwSync = 0, colSync = 0;

          brtwsprop_fst = readbin(brtwsprop_fst, bin, swapit);
          // BrtWsPropFields1 *fields1 = (BrtWsPropFields1 *)&brtwsprop_fst;
          brtwsprop_sec = readbin(brtwsprop_sec, bin, swapit);
          // BrtWsPropFields2 *fields2 = (BrtWsPropFields2 *)&brtwsprop_sec;

          std::vector<int32_t> color = brtColor(bin, swapit);
          rwSync = readbin(rwSync, bin, swapit);
          colSync = readbin(colSync, bin, swapit);
          std::string strName = XLWideString(bin, swapit);

          if (debug) {  // #nocov start
            Rcpp::Rcout << "sheetPr tabColor:" << std::endl;
            Rf_PrintValue(Rcpp::wrap(color));
          }  // #nocov end

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
              out << "<tabColor theme=\"" << color[1] << "\" />" << std::endl;  //  << "\" tint=\""<< color[2]
            }

            out << "</sheetPr>" << std::endl;
          }

          break;
        }

        case BrtWsDim: {
          if (debug) Rcpp::Rcout << "WsDim: " << bin.tellg() << std::endl;

          // 16 bit
          std::vector<int32_t> dims;
          // 0 index vectors
          // first row, last row, first col, last col
          dims = UncheckedRfX(bin, swapit);
          if (debug) {  // #nocov start
            Rcpp::Rcout << "dimension: ";
            Rf_PrintValue(Rcpp::wrap(dims));
          }  // #nocov end

          std::string lref = int_to_col(dims[2] + 1) + std::to_string(dims[0] + 1);
          std::string rref = int_to_col(dims[3] + 1) + std::to_string(dims[1] + 1);

          std::string ref;
          if (lref.compare(rref) == 0) {
            ref = lref;
          } else {
            ref =  lref + ":" + rref;
          }

          out << "<dimension ref=\"" << ref << "\"/>" << std::endl;

          break;
        }

        case BrtBeginCsViews:
        case BrtBeginWsViews: {
          if (debug) Rcpp::Rcout << "<sheetViews>: " << bin.tellg() << std::endl;
          out << "<sheetViews>" << std::endl;
          bin.seekg(size, bin.cur);

          break;
        }

        case BrtBeginCsView: {
          if (debug) Rcpp::Rcout << "<sheetView>: " << bin.tellg() << std::endl;
          uint16_t flags = 0;
          uint32_t wScale = 0, iWbkView = 0;

          flags = readbin(flags, bin, swapit);
          wScale = readbin(wScale, bin, swapit);
          iWbkView = readbin(iWbkView, bin, swapit);

          BrtBeginCsViewFlags view_flags(flags);

          out << "<sheetView";
          // careful, without the following nothing goes
          out << " workbookViewId=\"" << iWbkView << "\"";
          if (wScale)
            out << " zoomScale=\"" << wScale << "\"";
          if (view_flags.fSelected())
            out << " tabSelected=\"" << view_flags.fSelected() << "\"";
          out << " zoomToFit=\"1\"";

          out << ">" << std::endl;

          break;
        }

        case BrtPane: {
          if (debug) Rcpp::Rcout << "BrtPane: " << std::endl;
          uint8_t flags = 0;
          int32_t rwTop = 0, colLeft = 0;
          uint32_t pnnAct = 0;
          double xnumXSplit = 0.0, xnumYSplit = 0.0;

          xnumXSplit = Xnum(bin, swapit);
          xnumYSplit = Xnum(bin, swapit);
          rwTop = UncheckedRw(bin, swapit);
          colLeft = UncheckedCol(bin, swapit);
          pnnAct = readbin(pnnAct, bin, swapit);
          flags = readbin(flags, bin, swapit);

          std::string topLeft = int_to_col(colLeft + 1) + std::to_string(rwTop + 1);

          std::string pnn;
          if (pnnAct == 0x00000000) pnn = "bottomRight";
          if (pnnAct == 0x00000001) pnn = "topRight";
          if (pnnAct == 0x00000002) pnn = "bottomLeft";
          if (pnnAct == 0x00000003) pnn = "topLeft";

          PaneFlags view_flags(flags);

          std::string state;
          if (view_flags.fFrozen() && view_flags.fFrozenNoSplit())
            state = "frozen";
          else if (view_flags.fFrozen() && !view_flags.fFrozenNoSplit())
            state = "frozenSplit";
          else if (!view_flags.fFrozen() && !view_flags.fFrozenNoSplit())
            state = "split";

          out << "<pane";
          out << " activePane=\"" << pnn << "\"";
          out << " state=\"" << state << "\"";
          out << " xSplit=\"" << xnumXSplit << "\"";
          out << " ySplit=\"" << xnumYSplit << "\"";
          out << " topLeftCell=\"" << topLeft << "\"";
          out << " />" << std::endl;

          break;
        }

        case BrtBeginWsView: {
          if (debug) Rcpp::Rcout << "<sheetView>: " << bin.tellg() << std::endl;
          // out << "<sheetView>" << std::endl;
          // bin.seekg(size, bin.cur);

          uint8_t icvHdr = 0, reserved2 = 0;
          uint16_t flags = 0, reserved3 = 0, wScale = 0, wScaleNormal = 0, wScaleSLV = 0, wScalePLV = 0;
          uint32_t xlView = 0, iWbkView = 0;
          int32_t rwTop = 0, colLeft = 0;

          flags = readbin(flags, bin, swapit);
          xlView = readbin(xlView, bin, swapit);
          rwTop = UncheckedRw(bin, swapit);
          colLeft = UncheckedCol(bin, swapit);
          icvHdr = readbin(icvHdr, bin, swapit);
          reserved2 = readbin(reserved2, bin, swapit);
          reserved3 = readbin(reserved3, bin, swapit);
          wScale = readbin(wScale, bin, swapit);
          wScaleNormal = readbin(wScaleNormal, bin, swapit);
          wScaleSLV = readbin(wScaleSLV, bin, swapit);
          wScalePLV = readbin(wScalePLV, bin, swapit);
          iWbkView = readbin(iWbkView, bin, swapit);

          BrtBeginWsViewFlags view_flags(flags);

          out << "<sheetView";
          // careful, without the following nothing goes
          out << " workbookViewId=\"" << iWbkView << "\"";
          if (icvHdr)
            out << " colorId=\"" << (int32_t)icvHdr << "\"";
          if (!view_flags.fDefaultHdr())
            out << " defaultGridColor=\"" << view_flags.fDefaultHdr() << "\"";
          if (view_flags.fRightToLeft()) {
            Rcpp::Rcout << "rightToLeft=\"" << view_flags.fRightToLeft() << "\"" << std::endl;
            out << " rightToLeft=\"" << view_flags.fRightToLeft() << "\"";
          }
          if (view_flags.fDspFmla())
            out << " showFormulas=\"" << view_flags.fDspFmla() << "\"";
          if (!view_flags.fDspGrid())
            out << " showGridLines=\"" << view_flags.fDspGrid() << "\"";
          if (!view_flags.fDspGuts())
            out << " showOutlineSymbols=\"" << view_flags.fDspGuts() << "\"";
          if (!view_flags.fDspRwCol())
            out << " showRowColHeaders=\"" << view_flags.fDspRwCol() << "\"";
          if (view_flags.fDspRuler())
            out << " showRuler=\"" << view_flags.fDspRuler() << "\"";
          if (view_flags.fWhitespaceHidden())
            out << " showWhiteSpace=\"" << view_flags.fWhitespaceHidden() << "\"";
          if (view_flags.fDspZeros())
            out << " showZeros=\"" << view_flags.fDspZeros() << "\"";
          if (view_flags.fSelected())
            out << " tabSelected=\"" << view_flags.fSelected() << "\"";
          if (colLeft > 0 || rwTop > 0)
            out << " topLeftCell=\"" << int_to_col(colLeft + 1) << std::to_string(rwTop + 1) << "\"";
          if (xlView)
            out << " view=\"" << XLView(xlView) << "\"";
          if (view_flags.fWnProt())
            out << " windowProtection=\"" << view_flags.fWnProt() << "\"";
          if (wScale)
            out << " zoomScale=\"" << wScale << "\"";
          if (wScaleNormal)
            out << " zoomScaleNormal=\"" << wScaleNormal << "\"";
          if (wScalePLV)
            out << " zoomScalePageLayoutView=\"" << wScalePLV<< "\"";
          if (wScaleSLV)
            out << " zoomScaleSheetLayoutView=\"" << wScaleSLV<< "\"";

          out << ">";

          break;
        }

        case BrtPageSetup: {
          if (debug) Rcpp::Rcout << "BrtPageSetup: " << std::endl;
          uint16_t flags = 0;
          uint32_t iPaperSize = 0, iScale = 0, iRes = 0, iVRes = 0, iCopies = 0, iPageStart = 0, iFitWidth = 0, iFitHeight = 0;

          iPaperSize = readbin(iPaperSize, bin, swapit);
          iScale = readbin(iScale, bin, swapit);
          iRes = readbin(iRes, bin, swapit);
          iVRes = readbin(iVRes, bin, swapit);
          iCopies = readbin(iCopies, bin, swapit);
          iPageStart = readbin(iPageStart, bin, swapit);
          iFitWidth = readbin(iFitWidth, bin, swapit);
          iFitHeight = readbin(iFitHeight, bin, swapit);
          flags = readbin(flags, bin, swapit);

          std::string szRelId = XLNullableWideString(bin, swapit);

          out << "<pageSetup" <<
            // "\" blackAndWhite=\"" << <<
            // "\" draft=\""<< <<
            // "\" cellComments=\""<< <<
            " copies=\""<<  iCopies <<
            // "\" errors=\""<<  <<
            "\" firstPageNumber=\""<< iPageStart <<
            "\" fitToHeight=\""<< iFitHeight <<
            "\" fitToWidth=\""<< iFitWidth <<
            "\" horizontalDpi=\""<< iRes <<
            // "\" orientation=\""<<  <<
            // "\" pageOrder=\""<<  <<
            // "\" paperHeight=\""<<  <<
            "\" paperSize=\""<< iPaperSize <<
            // "\" paperWidth=\""<<  <<
            "\" scale=\""<< iScale <<
            // "\" useFirstPageNumber=\""<<  <<
            // "\" usePrinterDefaults=\""<<  <<
            "\" verticalDpi=\"" << iVRes <<
            "\" />" << std::endl;

          break;
        }

        case BrtPhoneticInfo:
        {
          if (debug) Rcpp::Rcout << "BrtPhoneticInfo: " << std::endl;
          uint16_t iFnt = 0;
          uint32_t phType = 0, phAll = 0;
          iFnt = readbin(iFnt, bin, swapit);
          phType = readbin(phType, bin, swapit);
          phAll = readbin(phAll, bin, swapit);

          if (debug) Rprintf("phnetic: %d, %d, %d", iFnt, phType, phAll);

          break;
        }

        case BrtBeginHeaderFooter: {
          if (debug) Rcpp::Rcout << "BrtBeginHeaderFooter: " << std::endl;
          uint16_t flags = 0;
          flags = readbin(flags, bin, swapit);  // unused

          std::string stHeader = XLNullableWideString(bin, swapit);
          std::string stFooter = XLNullableWideString(bin, swapit);
          std::string stHeaderEven = XLNullableWideString(bin, swapit);
          std::string stFooterEven = XLNullableWideString(bin, swapit);
          std::string stHeaderFirst = XLNullableWideString(bin, swapit);
          std::string stFooterFirst = XLNullableWideString(bin, swapit);

          BrtBeginHeaderFooterFlags view_flags(flags);

          if (debug) {  // #nocov start
            Rcpp::Rcout << stHeader<< ": " << stFooter << ": " <<
              stHeaderEven << ": " << stFooterEven << ": " <<
              stHeaderFirst << ": " << stFooterFirst << std::endl;
          }  // #nocov end

          out << "<headerFooter" << std::endl;
          if (view_flags.fHFDiffOddEven()) out << " differentOddEven=\"" << view_flags.fHFDiffOddEven() << "\"" << std::endl;
          if (view_flags.fHFDiffFirst())  out << " differentFirst=\"" << view_flags.fHFDiffFirst() << "\"" << std::endl;
          if (view_flags.fHFScaleWithDoc())  out << " scaleWithDoc=\"" << view_flags.fHFScaleWithDoc() << "\"" << std::endl;
          if (view_flags.fHFAlignMargins())  out << " alignWithMargins=\"" << view_flags.fHFAlignMargins() << "\"" << std::endl;
          out << ">" << std::endl;

          if (!stHeader.empty()) out << "<oddHeader>" << stHeader <<"</oddHeader>" << std::endl;
          if (!stFooter.empty()) out << "<oddFooter>" << stFooter <<"</oddFooter>" << std::endl;
          if (!stHeaderEven.empty()) out << "<evenHeader>" << stHeaderEven <<"</evenHeader>" << std::endl;
          if (!stFooterEven.empty()) out << "<evenFooter>" << stFooterEven <<"</evenFooter>" << std::endl;
          if (!stHeaderFirst.empty()) out << "<firstHeader>" << stHeaderFirst <<"</firstHeader>" << std::endl;
          if (!stFooterFirst.empty()) out << "<firstFooter>" << stFooterFirst <<"</firstFooter>" << std::endl;
          out << "</headerFooter>" << std::endl;

          break;
        }

        case BrtEndHeaderFooter: {
          bin.seekg(size, bin.cur);
          break;
        }

          // whatever this is
        case BrtSel: {
          if (debug) Rcpp::Rcout << "BrtSel: " << bin.tellg() << std::endl;
          uint32_t pnn = 0, rwAct = 0, colAct = 0, dwRfxAct = 0;
          std::vector<int32_t> sqrfx;

          pnn = readbin(pnn, bin, swapit);
          rwAct = readbin(rwAct, bin, swapit);
          colAct = readbin(colAct, bin, swapit);
          dwRfxAct = readbin(dwRfxAct, bin, swapit);

          // sqrfx[0] is the number of rfx following
          sqrfx = UncheckedSqRfX(bin, swapit);

          std::string ac = int_to_col(colAct + 1) + std::to_string(rwAct + 1);
          std::string lref = int_to_col(sqrfx[3] + 1) + std::to_string(sqrfx[1] + 1);
          std::string rref = int_to_col(sqrfx[4] + 1) + std::to_string(sqrfx[2] + 1);

          std::string sqref;
          if (lref.compare(rref) == 0) {
            sqref = lref;
          } else {
            sqref =  lref + ":" + rref;
          }

          std::string spnn;
          if (pnn == 0x00000000) spnn = "bottomRight";
          if (pnn == 0x00000001) spnn = "topRight";
          if (pnn == 0x00000002) spnn = "bottomLeft";
          if (pnn == 0x00000003) spnn = "topLeft";

          out << "<selection pane=\"" << spnn << "\" activeCell=\"" << ac << "\" sqref=\""<< sqref << "\" />" << std::endl;

          break;
        }

        case BrtEndCsView:
        case BrtEndWsView: {
          if (debug) Rcpp::Rcout << "</sheetView>: " << bin.tellg() << std::endl;
          out << "</sheetView>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtEndCsViews:
        case BrtEndWsViews: {
          if (debug) Rcpp::Rcout << "</sheetViews>: " << bin.tellg() << std::endl;
          out << "</sheetViews>" << std::endl;
          break;
        }

        case BrtBeginColInfos: {
          if (debug) Rcpp::Rcout << "<cols>: " << bin.tellg() << std::endl;
          out << "<cols>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtColInfo: {
          if (debug) Rcpp::Rcout << "<col/>: " << bin.tellg() << std::endl;
          uint16_t colinfo = 0;
          uint32_t coldx = 0, ixfe = 0;
          int32_t colFirst = 0, colLast = 0;

          colFirst = UncheckedCol(bin, swapit) + 1;
          colLast = UncheckedCol(bin, swapit) + 1;
          coldx = readbin(coldx, bin, swapit);
          ixfe = readbin(ixfe, bin, swapit);
          colinfo = readbin(colinfo, bin, swapit);

          BrtColInfoFlags view_flags(colinfo);

          out << "<col" << " min=\"" << colFirst << "\" max =\"" << colLast << "\"";

          if (ixfe > 0)
            out << " style=\"" <<  ixfe << "\"";

          out << " width=\"" <<  (double)coldx/256 << "\"";
          if (view_flags.fHidden())
            out << " hidden=\"" <<  view_flags.fHidden() << "\"";
          if (view_flags.fUserSet())
            out << " customWidth=\"" <<  view_flags.fUserSet() << "\"";
          if (view_flags.fBestFit())
            out << " bestFit=\"" <<  view_flags.fBestFit() << "\"";
          if (view_flags.iOutLevel()>0)
            out << " outlineLevel=\"" <<  (uint16_t)view_flags.iOutLevel() << "\"";
          if (view_flags.fCollapsed())
            out << " collapsed=\"" <<  view_flags.fCollapsed() << "\"";

          out << " />" << std::endl;
          break;
        }

        case BrtEndColInfos: {
          if (debug) Rcpp::Rcout << "</cols>: " << bin.tellg() << std::endl;
          out << "</cols>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtBeginSheetData: {
          if (debug) Rcpp::Rcout << "<sheetData>" << bin.tellg() << std::endl;
          out << "<sheetData>" << std::endl;  //  << bin.tellg()
          in_sheet_data = true;
          break;
        }

          // prelude to row entry
        case BrtACBegin: {
          if (debug) Rcpp::Rcout << "BrtACBegin: " << bin.tellg() << std::endl;
          bin.seekg(size, bin.cur);
          // int16_t nversions, version;
          // nversions = readbin(nversions, bin, swapit);
          // std::vector<int32_t> versions;
          //
          // for (int32_t i = 0; i < nversions; ++i) {
          //   version = readbin(version, bin, swapit);
          //   versions.push_back(version);
          // }
          //
          // Rf_PrintValue(Rcpp::wrap(versions));

          break;
        }

        case BrtWsFmtInfoEx14: {
          if (debug) Rcpp::Rcout << "BrtWsFmtInfoEx14: " << bin.tellg() << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtACEnd: {
          if (debug) Rcpp::Rcout << "BrtACEnd: " << bin.tellg() << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtWsFmtInfo: {
          if (debug) Rcpp::Rcout << "BrtWsFmtInfo: " << bin.tellg() << std::endl;
          uint32_t dxGCol = 0;
          uint16_t cchDefColWidth = 0, miyDefRwHeight = 0, fmt_flags = 0;
          uint8_t iOutLevelRw = 0, iOutLevelCol = 0;
          dxGCol = readbin(dxGCol, bin, swapit);
          cchDefColWidth = readbin(cchDefColWidth, bin, swapit);
          miyDefRwHeight = readbin(miyDefRwHeight, bin, swapit);
          fmt_flags = readbin(fmt_flags, bin, swapit);
          def_row_twips = miyDefRwHeight;
          iOutLevelRw = readbin(iOutLevelRw, bin, swapit);
          iOutLevelCol = readbin(iOutLevelCol, bin, swapit);

          out << "<sheetFormatPr";
          if (cchDefColWidth > 0) out << " baseColWidth=\"" << cchDefColWidth << "\"";
          if (dxGCol > 0 && dxGCol <= 255u * 256u) out << " defaultColWidth=\"" << static_cast<double>(dxGCol) / 256 << "\"";
          out << " defaultRowHeight=\"" << static_cast<double>(miyDefRwHeight) / 20 << "\"";
          if (fmt_flags & 0x0001) out << " customHeight=\"1\"";
          if (fmt_flags & 0x0002) out << " zeroHeight=\"1\"";
          if (iOutLevelRw > 0) out << " outlineLevelRow=\"" << static_cast<uint32_t>(iOutLevelRw) << "\"";
          if (iOutLevelCol > 0) out << " outlineLevelCol=\"" << static_cast<uint32_t>(iOutLevelCol) << "\"";
          out << "/>" << std::endl;
          break;
        }

        case BrtRwDescent: {
          if (debug) Rcpp::Rcout << "BrtRwDescent: " << bin.tellg() << std::endl;
          // bin.seekg(size, bin.cur);
          uint16_t dyDescent = 0;
          dyDescent = readbin(dyDescent, bin, swapit);
          break;
        }

        case BrtRowHdr: {
          if (debug) Rcpp::Rcout << "<row/>: " << bin.tellg() << std::endl;

          if (!in_sheet_data) bin.seekg(size, bin.cur);

          // close open rows
          if (!first_row) {

            for (size_t i = 0; i < colvec.size(); ++i) {
              out << "<c r=\"" << colvec[i].c_r << "\" s=\"" << colvec[i].c_s << "\" t=\""<< colvec[i].c_t << "\">" << std::endl;
              out << "<v>" << escape_xml(colvec[i].v) << "</v>" << std::endl;
              out << colvec[i].is  << std::endl;
              out << "<f " << colvec[i].f_attr << ">" << colvec[i].f << "</f>" << std::endl;
              out << "</c>" << std::endl;
            }

            colvec.clear();

            out << "</row>" <<std::endl;
          } else {
            first_row = false;
          }

          uint8_t bits3 = 0, fExtraAsc = 0, fExtraDsc = 0,
            fCollapsed = 0, fDyZero = 0, fUnsynced = 0, fGhostDirty = 0,
            fReserved = 0, fPhShow = 0;
          uint16_t miyRw = 0;

          // uint24_t;
          int32_t rw = 0, colMic = 0, colLast = 0;  // uint32?
          uint32_t ixfe = 0, ccolspan = 0, unk32 = 0;

          rw = readbin(rw, bin, swapit);

          if (rw > 0x00100000 || rw < row) {
            Rcpp::stop("row either decreasing or to large");
          }

          ixfe = readbin(ixfe, bin, swapit);
          miyRw = readbin(miyRw, bin, swapit);

          if (miyRw > 0x2000) Rcpp::stop("miyRw to big");

          uint16_t rwoheaderfields = 0;
          rwoheaderfields = readbin(rwoheaderfields, bin, swapit);

          BrtRowHdrFlags view_flags(rwoheaderfields);

          // fExtraAsc = 1
          // fExtraDsc = 1
          // reserved1 = 6
          // iOutLevel   = 3
          // fCollapsed  = 1
          // fDyZero     = 1
          // fUnsynced   = 1
          // fGhostDirty = 1
          // fReserved   = 1

          bits3 = readbin(bits3, bin, swapit);
          // fPhShow   = 1
          // fReserved = 7

          ccolspan = readbin(ccolspan, bin, swapit);
          if (ccolspan > 16) Rcpp::stop("ccolspan to large");
          std::string spans;

          std::vector<int32_t> spans_int;
          if (ccolspan) {
            // has number of rgBrtColspan elements. For now take max and min value.
            // there are other types of spans (1:1 1024:16384), but unsure how this is
            // supposed to be handled. It's possible to have elements like spans =
            // "1:1 16384:16384" which we will read as spans = "1:16384"
            for (uint8_t clpn = 0; clpn < ccolspan; ++clpn) {
              colMic = readbin(colMic, bin, swapit);
              colLast = readbin(colLast, bin, swapit);

              if (debug) {  // #nocov start
                spans = std::to_string(colMic + 1) + ":" + std::to_string(colLast + 1);
                Rcpp::Rcout << (int32_t)clpn << ": " << spans << std::endl;
              }  // #nocov end

              spans_int.push_back(colMic);
              spans_int.push_back(colLast);
            }

            int32_t min_span = *std::min_element(spans_int.begin(), spans_int.end());
            int32_t max_span = *std::max_element(spans_int.begin(), spans_int.end());

            spans = std::to_string(min_span + 1) + ":" + std::to_string(max_span + 1);
          }

          out << "<row r=\"";  //  << bin.tellg()
          out << rw + 1 << "\"";

          if (view_flags.fUnsynced() != 0) {
            out << " ht=\"" << static_cast<double>(miyRw) / 20 << "\"";
            out << " customHeight=\"" << view_flags.fUnsynced() << "\"";
          } else if (miyRw > 0 && miyRw != def_row_twips) {
            out << " ht=\"" << static_cast<double>(miyRw) / 20 << "\"";
          }

          if (ccolspan)
            out << " spans=\"" << spans << "\"";

          if (view_flags.iOutLevel() > 0) {
            out << " outlineLevel=\"" << (uint16_t)view_flags.iOutLevel() << "\"";
          }

          if (view_flags.fCollapsed()) {
            out << " collapsed=\"" << view_flags.fCollapsed() << "\"";
          }

          if (view_flags.fDyZero()) {
            out << " hidden=\"" << view_flags.fDyZero() << "\"";
          }

          if (view_flags.fGhostDirty()) {
            out << " customFormat=\"" << view_flags.fGhostDirty() << "\"";
          }

          if (ixfe > 0) {
            out << " s=\"" << ixfe << "\"";
          }

          out << ">" << std::endl;

          row = rw;

          if (debug) {  // #nocov start
            Rcpp::Rcout << "ROW:" << (rw + 1) << " : " << ixfe << " : " << miyRw << " : " << (int32_t)fExtraAsc << " : " <<
              (int32_t)fExtraDsc << " : " << unk32 << " : " << (int32_t)fCollapsed << " : " << (int32_t)fDyZero << " : " <<
              (int32_t)fUnsynced << " : " << (int32_t)fGhostDirty << " : " << (int32_t)fReserved << " : " <<
              (int32_t)fPhShow << " : " << ccolspan << "; " << bin.tellg() << std::endl;
          }  // #nocov end

          break;
        }

        case BrtCellIsst: { // shared string
          if (debug) Rcpp::Rcout << "BrtCellIsst: " << bin.tellg() << std::endl;

          int32_t val1 = 0, val2 = 0, val3 = 0;
          val1 = readbin(val1, bin, swapit);
          if (debug) Rcpp::Rcout << val1 << std::endl;
          val2 = readbin(val2, bin, swapit);
          if (debug) Rcpp::Rcout << val2 << std::endl;
          val3 = readbin(val3, bin, swapit);
          if (debug) Rcpp::Rcout << val3 << std::endl;

          xml_col column;
          column.v = std::to_string(val3);
          column.c_t = "s";
          if (val2) column.c_s = std::to_string(val2);
          column.c_r = int_to_col(val1 + 1) + std::to_string(row + 1);
          colvec.push_back(column);

          // out << "<c r=\"" << int_to_col(val1 + 1) << row + 1 << "\"" << cell_style(val2) << " t=\"s\">" << std::endl;
          // out << "<v>" << val3 << "</v>" << std::endl;
          // out << "</c>" << std::endl;

          break;
        }

        case BrtCellSt: {
          if (debug) Rcpp::Rcout << "BrtCellSt (inlineStr):" << bin.tellg() << std::endl;

          std::vector<int32_t> cell;
          cell = Cell(bin, swapit);
          std::string a0 = int_to_col(cell[0] + 1) + std::to_string(row + 1);
          std::string value = XLWideString(bin, swapit); // not a RichStr?

          xml_col column;
          column.is = txt_to_xml(value, false, true, true, "is");
          column.c_r = a0;
          column.c_t = "inlineStr";
          if (cell[1]) column.c_s = std::to_string(cell[1]);
          colvec.push_back(column);

          break;
        }

        case BrtCellBool: { // bool
          if (debug) Rcpp::Rcout << "BrtCellBool: " << bin.tellg() << std::endl;

          int32_t val1 = 0, val2 = 0;
          uint8_t val3 = 0;
          val1 = readbin(val1, bin, swapit);
          // out << val1 << std::endl;
          val2 = readbin(val2, bin, swapit);
          // out << val2 << std::endl;
          val3 = readbin(val3, bin, swapit);
          // out << val3 << std::endl;


          xml_col column;
          column.v = std::to_string((int32_t)val3);
          column.c_t = "b";
          if (val2) column.c_s = std::to_string(val2);
          column.c_r = int_to_col(val1 + 1) + std::to_string(row + 1);
          colvec.push_back(column);

          // out << "<c r=\"" << int_to_col(val1 + 1) << row + 1<< "\"" << cell_style(val2) << " t=\"b\">" << std::endl;
          // out << "<v>" << (int32_t)val3 << "</v>" << std::endl;
          // out << "</c>" << std::endl;

          break;
        }

        case BrtCellRk: { // integer?
          if (debug) Rcpp::Rcout << "BrtCellRk: " << bin.tellg() << std::endl;

          int32_t val1 = 0, val2= 0, val3= 0;
          val1 = readbin(val1, bin, swapit);
          // Rcpp::Rcout << val << std::endl;
          val2 = readbin(val2, bin, swapit);
          // Rcpp::Rcout << val << std::endl;
          // wrong?
          val3 = readbin(val3, bin, swapit);
          // Rcpp::Rcout << RkNumber(val) << std::endl;

          std::stringstream stream;
          stream << std::setprecision(16) << RkNumber(val3);

          xml_col column;
          column.v = stream.str();
          if (val2) column.c_s = std::to_string(val2);
          column.c_r = int_to_col(val1 + 1) + std::to_string(row + 1);
          colvec.push_back(column);

          // out << "<c r=\"" << int_to_col(val1 + 1) << row + 1<< "\"" << cell_style(val2) << ">" << std::endl;
          // out << "<v>" << std::setprecision(16) << RkNumber(val3) << "</v>" << std::endl;
          // out << "</c>" << std::endl;

          break;
        }

        case BrtCellReal: {
          if (debug) Rcpp::Rcout << "BrtCellReal: " << bin.tellg() << std::endl;
          int32_t val1= 0, val2= 0;
          val1 = readbin(val1, bin, swapit);
          // Rcpp::Rcout << val << std::endl;
          val2 = readbin(val2, bin, swapit);
          // Rcpp::Rcout << val << std::endl;

          double dbl = 0.0;
          dbl = readbin(dbl, bin, swapit);
          // Rcpp::Rcout << dbl << std::endl;

          std::stringstream stream;
          stream << std::setprecision(16) << dbl;

          xml_col column;
          column.v = stream.str();
          if (val2) column.c_s = std::to_string(val2);
          column.c_r = int_to_col(val1 + 1) + std::to_string(row + 1);
          colvec.push_back(column);

          // out << "<c r=\"" << int_to_col(val1 + 1) << row + 1 << "\"" << cell_style(val2) << ">" << std::endl;
          // out << "<v>" << std::setprecision(16) << dbl << "</v>" << std::endl;  // << std::fixed
          // out << "</c>" << std::endl;

          break;
        }

          // 0 ?
        case BrtCellBlank: {
          if (debug) Rcpp::Rcout << "BrtCellBlank: " << bin.tellg() << std::endl;

          std::vector<int32_t> blank = Cell(bin, swapit);
          col = blank[0];
          if (debug) Rf_PrintValue(Rcpp::wrap(blank));


          xml_col column;
          if (blank[1]) column.c_s = std::to_string(blank[1]);
          column.c_r = int_to_col(blank[0] + 1) + std::to_string(row + 1);
          colvec.push_back(column);

          // out << "<c r=\"" << int_to_col(blank[0] + 1) << row + 1 << "\"" << cell_style(blank[1]) << "/>" << std::endl;

          break;
        }

        case BrtCellError: { // t="e" & <v>#NUM!</v>
          if (debug) Rcpp::Rcout << "BrtCellError: " << bin.tellg() << std::endl;

          // uint8_t val8= 0;

          int32_t val1= 0, val2= 0;
          val1 = readbin(val1, bin, swapit);
          // out << val << std::endl;
          val2 = readbin(val2, bin, swapit);
          // out << val << std::endl;

          xml_col column;
          column.v = BErr(bin, swapit);
          if (val2) column.c_s = std::to_string(val2);
          column.c_t = "e";
          column.c_r = int_to_col(val1 + 1) + std::to_string(row + 1);
          colvec.push_back(column);

          // out << "<c r=\"" << int_to_col(val1 + 1) << row + 1<< "\"" << cell_style(val2) << " t=\"e\">" << std::endl;
          // out << "<v>" << BErr(bin, swapit) << "</v>" << std::endl;
          // out << "</c>" << std::endl;

          break;
        }

        case BrtCellMeta: {
          if (debug) Rcpp::Rcout << "BrtCellMeta: " << bin.tellg() << std::endl;
          uint32_t icmb = 0;
          icmb = readbin(icmb, bin, swapit);

          if (debug) Rcpp::Rcout << "cell contains unhandled cell metadata entry" << icmb << std::endl;

          break;
        }

        case BrtFmlaBool: {
          if (debug) Rcpp::Rcout << "BrtFmlaBool: " << bin.tellg() << std::endl;
          // bin.seekg(size, bin.cur);
          int32_t is_shared_formula = false;

          std::vector<int32_t> cell;
          cell = Cell(bin, swapit);
          col = cell[0];
          if (debug) Rf_PrintValue(Rcpp::wrap(cell));

          bool val = 0;
          val = readbin(val, bin, swapit);
          if (debug) Rcpp::Rcout << val << std::endl;

          uint16_t grbitFlags = 0;
          grbitFlags = readbin(grbitFlags, bin, swapit);

          GrbitFmlaFlags view_flags(grbitFlags);

          std::string fml = CellParsedFormula(bin, swapit, debug, 0, row, is_shared_formula, has_revision_record);

          xml_col column;
          column.v = std::to_string((int32_t)val);

          std::string a0 = int_to_col(cell[0] + 1) + std::to_string(row);
          std::string a1 = int_to_col(cell[0] + 1) + std::to_string(row + 1);

          auto sfml = shared_cells.find(fml);
          auto sfml1 = shared_cells.find(a0);
          if ((fml.empty() || sfml != shared_cells.end()) && sfml1 != shared_cells.end()) {
            column.f_attr = "t=\"" + fml_type + "\"";
            if (view_flags.fAlwaysCalc()) column.f_attr += " ca=\"" + std::to_string(view_flags.fAlwaysCalc()) + "\"";

            if (sfml != shared_cells.end())
              column.f_attr += " si=\"" + std::to_string(sfml->second) + "\"";
            else
              column.f_attr += " si=\"" + std::to_string(sfml1->second) + "\"";
          } else {
            if (view_flags.fAlwaysCalc()) column.f_attr += "ca=\"" + std::to_string(view_flags.fAlwaysCalc()) + "\"";
            column.f = fml;
          }
          if (cell[1]) column.c_s = std::to_string(cell[1]);
          column.c_t = "b";
          column.c_r = a1;
          colvec.push_back(column);

          break;
        }

        case BrtFmlaError: { // t="e" & <f>
          if (debug) Rcpp::Rcout << "BrtFmlaError: " << bin.tellg() << std::endl;
          // bin.seekg(size, bin.cur);
          int32_t is_shared_formula = false;

          std::vector<int32_t> cell;
          cell = Cell(bin, swapit);
          col = cell[0];
          if (debug) Rf_PrintValue(Rcpp::wrap(cell));

          std::string fErr;
          fErr = BErr(bin, swapit);

          uint16_t grbitFlags = 0;
          grbitFlags = readbin(grbitFlags, bin, swapit);

          GrbitFmlaFlags view_flags(grbitFlags);

          std::string fml = CellParsedFormula(bin, swapit, debug, 0, row, is_shared_formula, has_revision_record);

          xml_col column;
          column.v = fErr;

          std::string a0 = int_to_col(cell[0] + 1) + std::to_string(row);
          std::string a1 = int_to_col(cell[0] + 1) + std::to_string(row + 1);

          auto sfml = shared_cells.find(fml);
          auto sfml1 = shared_cells.find(a0);
          if ((fml.empty() || sfml != shared_cells.end()) && sfml1 != shared_cells.end()) {
            column.f_attr = "t=\"" + fml_type + "\"";
            if (view_flags.fAlwaysCalc()) column.f_attr += " ca=\"" + std::to_string(view_flags.fAlwaysCalc()) + "\"";

            if (sfml != shared_cells.end())
              column.f_attr += " si=\"" + std::to_string(sfml->second) + "\"";
            else
              column.f_attr += " si=\"" + std::to_string(sfml1->second) + "\"";
          } else {
            if (view_flags.fAlwaysCalc()) column.f_attr += "ca=\"" + std::to_string(view_flags.fAlwaysCalc()) + "\"";
            column.f = fml;
          }
          if (cell[1]) column.c_s = std::to_string(cell[1]);
          column.c_t = "e";
          column.c_r = a1;
          colvec.push_back(column);

          break;
        }

        case BrtFmlaNum: {
          if (debug) Rcpp::Rcout << "BrtFmlaNum: " << bin.tellg() << std::endl;
          // bin.seekg(size, bin.cur);
          int32_t is_shared_formula = false;

          std::vector<int32_t> cell;
          cell = Cell(bin, swapit);
          col = cell[0];
          if (debug) Rf_PrintValue(Rcpp::wrap(cell));

          double xnum = Xnum(bin, swapit);
          if (debug) Rcpp::Rcout << xnum << std::endl;

          uint16_t grbitFlags = 0;
          grbitFlags = readbin(grbitFlags, bin, swapit);

          GrbitFmlaFlags view_flags(grbitFlags);

          std::string fml = CellParsedFormula(bin, swapit, debug, 0, row, is_shared_formula, has_revision_record);

          std::stringstream stream;
          stream << std::setprecision(16) << xnum;
          std::string a0 = int_to_col(cell[0] + 1) + std::to_string(row);
          std::string a1 = int_to_col(cell[0] + 1) + std::to_string(row + 1);

          xml_col column;
          column.v = stream.str();

          // fml is a reference to another cell in shared formula case?

          auto sfml = shared_cells.find(fml);
          auto sfml1 = shared_cells.find(a0);
          if ((fml.empty() || sfml != shared_cells.end()) && sfml1 != shared_cells.end()) {
            column.f_attr = "t=\"" + fml_type + "\"";
            if (view_flags.fAlwaysCalc()) column.f_attr += " ca=\"" + std::to_string(view_flags.fAlwaysCalc()) + "\"";

            if (sfml != shared_cells.end())
              column.f_attr += " si=\"" + std::to_string(sfml->second) + "\"";
            else
              column.f_attr += " si=\"" + std::to_string(sfml1->second) + "\"";
          } else {
            if (view_flags.fAlwaysCalc()) column.f_attr += "ca=\"" + std::to_string(view_flags.fAlwaysCalc()) + "\"";
            column.f = fml;
          }
          if (cell[1]) column.c_s = std::to_string(cell[1]);
          // column.c_t = "e";
          column.c_r = a1;
          colvec.push_back(column);

          break;
        }

        case BrtFmlaString: {
          if (debug) Rcpp::Rcout << "BrtFmlaString: " << bin.tellg() << std::endl;
          // bin.seekg(size, bin.cur);
          int32_t is_shared_formula = false;

          std::vector<int32_t> cell;
          cell = Cell(bin, swapit);
          col = cell[0];
          if (debug) Rf_PrintValue(Rcpp::wrap(cell));

          std::string val = XLWideString(bin, swapit);
          if (debug) Rcpp::Rcout << val << std::endl;

          uint16_t grbitFlags = 0;
          grbitFlags = readbin(grbitFlags, bin, swapit);

          GrbitFmlaFlags view_flags(grbitFlags);

          std::string fml = CellParsedFormula(bin, swapit, debug, 0, row, is_shared_formula, has_revision_record);

          std::string a0 = int_to_col(cell[0] + 1) + std::to_string(row);
          std::string a1 = int_to_col(cell[0] + 1) + std::to_string(row + 1);

          xml_col column;

          auto sfml = shared_cells.find(fml);
          auto sfml1 = shared_cells.find(a0);
          if ((fml.empty() || sfml != shared_cells.end()) && sfml1 != shared_cells.end()) {
            column.f_attr = "t=\"" + fml_type + "\"";
            if (view_flags.fAlwaysCalc()) column.f_attr += " ca=\"" + std::to_string(view_flags.fAlwaysCalc()) + "\"";

            if (sfml != shared_cells.end())
              column.f_attr += " si=\"" + std::to_string(sfml->second) + "\"";
            else
              column.f_attr += " si=\"" + std::to_string(sfml1->second) + "\"";
          } else {
            if (view_flags.fAlwaysCalc()) column.f_attr += "ca=\"" + std::to_string(view_flags.fAlwaysCalc()) + "\"";
            column.f = fml;
          }
          column.v = val;
          if (cell[1]) column.c_s = std::to_string(cell[1]);
          column.c_t = "str";
          column.c_r = a1;
          colvec.push_back(column);

          break;
        }

        case BrtArrFmla: {
          if (debug) Rcpp::Rcout << "BrtArrFmla: " << bin.tellg() << std::endl;
          int32_t is_shared_formula = false;

          uint8_t flags = 0;
          int32_t rwFirst = 0, rwLast = 0, colFirst = 0, colLast = 0;
          rwFirst  = UncheckedRw(bin, swapit) + 1;
          rwLast   = UncheckedRw(bin, swapit) + 1;
          colFirst = UncheckedCol(bin, swapit) + 1;
          colLast  = UncheckedCol(bin, swapit) + 1;

          std::string lref = int_to_col(colFirst) + std::to_string(rwFirst);
          std::string rref = int_to_col(colLast) + std::to_string(rwLast);

          std::string ref;
          if (lref.compare(rref) == 0) {
            ref = lref;
          } else {
            ref =  lref + ":" + rref;
          }

          if (debug) Rcpp::Rcout << "ref: " << ref << std::endl;

          flags = readbin(flags, bin, 0);

          std::string fml = CellParsedFormula(bin, swapit, debug, col, row, is_shared_formula, has_revision_record);
          if (debug) Rcpp::Rcout << "BrtArrFmla: " << fml << std::endl;

          // add to the last colvec element
          auto last = colvec.size() - 1L;
          colvec[last].f = fml;
          colvec[last].f_attr = "t=\"array\"";
          colvec[last].f_attr += " ref=\"" + ref + "\"";
          // colvec[last].f_attr += "si=\"\"";

          break;
        }

        case BrtShrFmla: {
          if (debug) Rcpp::Rcout << "BrtShrFmla: " << bin.tellg() << std::endl;
          int32_t is_shared_formula = false;

          int32_t rwFirst = 0, rwLast = 0, colFirst = 0, colLast = 0;
          rwFirst  = UncheckedRw(bin, swapit);  // + 1
          rwLast   = UncheckedRw(bin, swapit);  // + 1
          colFirst = UncheckedCol(bin, swapit) + 1;
          colLast  = UncheckedCol(bin, swapit) + 1;

          std::vector<std::string> cells = dims_to_cells(rwFirst, rwLast, colFirst, colLast);

          // fill the unordered map
          for (size_t i = 0; i < cells.size(); ++i) {
            shared_cells[cells[i]] = shared_cell_cntr;
          }

          // shared_cells needs to compare against fml. fml is (row - 1)
          // reference must be row
          std::string lref = int_to_col(colFirst) + std::to_string(rwFirst + 1);
          std::string rref = int_to_col(colLast) + std::to_string(rwLast + 1);

          std::string ref;
          if (lref.compare(rref) == 0) {
            ref = lref;
          } else {
            ref =  lref + ":" + rref;
          }

          if (debug) Rcpp::Rcout << "ref: " << ref << std::endl;

          std::string fml = CellParsedFormula(bin, swapit, debug, col, row, is_shared_formula, has_revision_record);
          if (debug) Rcpp::Rcout << "BrtShrFmla: " << fml << std::endl;

          fml_type = "shared";

          // add to the last colvec element
          auto last = colvec.size() - 1L;
          colvec[last].f = fml;

          // TODO do we need to store more attributes?
          // Rcpp::Rcout << colvec[last].f_attr << std::endl;
          std::string bak = colvec[last].f_attr;

          colvec[last].f_attr = "t=\""+ fml_type + "\"";  // not sure how many cells share array type
          colvec[last].f_attr += " ref=\"" + ref + "\"";
          if (bak == "ca=\"1\"") colvec[last].f_attr += " " + bak;
          colvec[last].f_attr += " si=\"" + std::to_string(shared_cell_cntr) + "\"";

          ++shared_cell_cntr;

          break;
        }

        case BrtEndSheetData: {
          if (debug) Rcpp::Rcout << "</sheetData>" << bin.tellg() << std::endl;

          if (!first_row) {

            // should be the last row
            for (size_t i = 0; i < colvec.size(); ++i) {
              out << "<c r=\"" << colvec[i].c_r << "\" s=\"" << colvec[i].c_s << "\" t=\""<< colvec[i].c_t << "\">" << std::endl;
              out << "<v>" << escape_xml(colvec[i].v) << "</v>" << std::endl;
              out << "<f " << colvec[i].f_attr << " >" << colvec[i].f << "</f>" << std::endl;
              out << "</c>" << std::endl;
            }

            colvec.clear();
            out << "</row>" << std::endl;
          }
          out << "</sheetData>" << std::endl;

          in_sheet_data = false;

          if (colvec.size() > 0) {
            Rcpp::warning("colved not zero");
          }

          break;
        }

        case BrtSheetProtection: {
          if (debug) Rcpp::Rcout << "BrtSheetProtection: " << bin.tellg() << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtBeginMergeCells: {
          if (debug) Rcpp::Rcout << "BrtBeginMergeCells: " << bin.tellg() << std::endl;
          uint32_t cmcs = 0;
          cmcs = readbin(cmcs, bin, swapit);
          if (debug) Rcpp::Rcout << "count: " << cmcs << std::endl;

          out << "<mergeCells count=\"" << cmcs << "\">"<< std::endl;
            break;
        }

        case BrtMergeCell: {
          if (debug) Rcpp::Rcout << "BrtMergeCell: " << bin.tellg() << std::endl;

          int32_t rwFirst = 0, rwLast = 0, colFirst = 0, colLast = 0;

          rwFirst  = UncheckedRw(bin, swapit) + 1L;
          rwLast   = UncheckedRw(bin, swapit) + 1L;
          colFirst = UncheckedCol(bin, swapit) + 1L;
          colLast  = UncheckedCol(bin, swapit) + 1L;

          if (debug) Rprintf("MergeCell: %d %d %d %d\n", rwFirst, rwLast, colFirst, colLast);

          out << "<mergeCell ref=\"" <<
            int_to_col(colFirst) << rwFirst << ":" <<
              int_to_col(colLast) << rwLast <<
                "\" />" << std::endl;

          break;
        }

        case BrtEndMergeCells: {
          if (debug) Rcpp::Rcout << "BrtEndMergeCells: " << bin.tellg() << std::endl;
          out << "</mergeCells>" << std::endl;
          break;
        }

        case BrtPrintOptions: {
          if (debug) Rcpp::Rcout << "BrtPrintOptions: " << bin.tellg() << std::endl;
          uint16_t flags = 0;
          flags = readbin(flags, bin, swapit);
          break;
        }

        case BrtMargins: {
          if (debug)  Rcpp::Rcout << "BrtMargins: " << bin.tellg() << std::endl;

          double xnumLeft = 0, xnumRight = 0, xnumTop = 0, xnumBottom = 0, xnumHeader = 0, xnumFooter = 0;
          xnumLeft   = Xnum(bin, swapit);
          xnumRight  = Xnum(bin, swapit);
          xnumTop    = Xnum(bin, swapit);
          xnumBottom = Xnum(bin, swapit);
          xnumHeader = Xnum(bin, swapit);
          xnumFooter = Xnum(bin, swapit);

          if (xnumLeft < 0   || xnumLeft > 49)   Rcpp::stop("xnumLeft size out of bounds");
          if (xnumRight < 0  || xnumRight > 49)  Rcpp::stop("xnumRight size out of bounds");
          if (xnumTop < 0    || xnumTop > 49)    Rcpp::stop("xnumTop size out of bounds");
          if (xnumBottom < 0 || xnumBottom > 49) Rcpp::stop("xnumBottom size out of bounds");
          if (xnumHeader < 0 || xnumHeader > 49) Rcpp::stop("xnumHeader size out of bounds");
          if (xnumFooter < 0 || xnumFooter > 49) Rcpp::stop("xnumFooter size out of bounds");

          out << "<pageMargins" <<
            " left=\"" << xnumLeft << "\"" <<
            " right=\"" << xnumRight << "\"" <<
            " top=\"" << xnumTop << "\"" <<
            " bottom=\"" << xnumBottom << "\"" <<
            " header=\"" << xnumHeader << "\"" <<
            " footer=\"" << xnumFooter << "\"" <<
            " />" << std::endl;

          break;
        }

        case BrtUID: {
          if (debug)  Rcpp::Rcout << "BrtUID: " << bin.tellg() << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtDrawing: {
          if (debug) Rcpp::Rcout << "BrtDrawing" << std::endl;
          std::string stRelId = XLNullableWideString(bin, swapit);
          out << "<drawing r:id=\"" << stRelId << "\" />" << std::endl;
          break;
        }

        case BrtLegacyDrawing: {
          if (debug) Rcpp::Rcout << "BrtLegacyDrawing" << std::endl;
          std::string stRelId = XLNullableWideString(bin, swapit);
          out << "<legacyDrawing r:id=\"" << stRelId << "\" />" << std::endl;
          break;
        }

        case BrtLegacyDrawingHF: {
          if (debug) Rcpp::Rcout << "BrtLegacyDrawingHF" << std::endl;
          std::string stRelId = XLNullableWideString(bin, swapit);
          out << "<legacyDrawingHF r:id=\"" << stRelId << "\" />" << std::endl;
          break;
        }

        case BrtHLink: {
          if (debug) Rcpp::Rcout << "BrtHLink" << std::endl;

          std::vector<int32_t> rfx = UncheckedRfX(bin, swapit);
          std::string relId = XLNullableWideString(bin, swapit);
          std::string location = XLWideString(bin, swapit);
          std::string tooltip = XLWideString(bin, swapit);
          std::string display = XLWideString(bin, swapit);


          std::string lref = int_to_col(rfx[2] + 1) + std::to_string(rfx[0] + 1);
          std::string rref = int_to_col(rfx[3] + 1) + std::to_string(rfx[1] + 1);

          std::string ref;
          if (lref.compare(rref) == 0) {
            ref = lref;
          } else {
            ref =  lref + ":" + rref;
          }

          std::string hlink = "<hyperlink display=\"" +  escape_xml(display) + "\" r:id=\"" + relId + "\" location=\"" + escape_xml(location) + "\" ref=\"" + ref + "\" tooltip=\"" + escape_xml(tooltip) + "\" />";
          hlinks.push_back(hlink);

          break;
        }

        case BrtBeginAFilter: {
          if (debug) Rcpp::Rcout << "BrtBeginAFilter" << std::endl;
          std::vector<int32_t> rfx = UncheckedRfX(bin, swapit);

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

        case BrtBeginFilterColumn: {
          if (debug) Rcpp::Rcout << "BrtBeginFilterColumn" << std::endl;

          uint16_t flags = 0;
          uint32_t dwCol = 0;
          dwCol = readbin(dwCol, bin, swapit);
          flags = readbin(flags, bin, swapit);
          // fHideArrow
          // fNoBtn

          out << "<filterColumn colId=\"" << dwCol <<"\">" << std::endl;

          break;
        }

        case BrtBeginFilters: {
          if (debug) Rcpp::Rcout << "BrtBeginFilters" << std::endl;
          // bin.seekg(size, bin.cur);

          uint32_t fBlank = 0, unused = 0;
          fBlank = readbin(fBlank, bin, swapit);  // a 32bit flag, after all ... why not?
          unused = readbin(unused, bin, swapit);  // maybe calendarType?
          out << "<filters blank=\"" << fBlank <<"\">" << std::endl;
          break;
        }

        case BrtFilter: {
          if (debug) Rcpp::Rcout << "BrtFilter" << std::endl;
          std::string rgch = XLWideString(bin, swapit);
          out << "<filter val=\"" << rgch << "\" />" << std::endl;
          // bin.seekg(size, bin.cur);
          break;
        }

        case BrtEndFilters: {
          if (debug) Rcpp::Rcout << "BrtEndFilters" << std::endl;
          out << "</filters>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtColorFilter: {
          if (debug) Rcpp::Rcout << "BrtColorFilter" << std::endl;
          uint32_t dxfid = 0, fCellColor = 0;
          dxfid = readbin(dxfid, bin, swapit);
          fCellColor = readbin(fCellColor, bin, swapit);

          out << "<colorFilter dxfId=\"" << dxfid << "\" cellColor=\""<< fCellColor<< "\"/>" << std::endl;

          break;
        }


        case BrtBeginCustomFilters:
        case BrtBeginCustomFilters14:
        case BrtBeginCustomRichFilters: {
          int32_t fAnd = 0;

          // in xlsb it is flipped
          fAnd = readbin(fAnd, bin, swapit) ^ 1;

          out << "<customFilters" << std::endl;
          if (fAnd) out << " and=\"" << fAnd << "\"";  // and="1"
          out << ">";

          break;
        }

        case BrtCustomFilter:
        case BrtCustomFilter14: {
          int8_t vts = 0, grbitSgn = 0;
          double union_val = 0;
          std::string vtsStringXls;

          vts = readbin(vts, bin, swapit);
          grbitSgn = readbin(grbitSgn, bin, swapit);

          if (vts == 4) {
            // a double
            union_val = Xnum(bin, swapit);
          } else if (vts == 8) {
            // a bool
            int8_t boolean = 0;
            boolean = readbin(boolean, bin, swapit);
            union_val = static_cast<double>(readbin(boolean, bin, swapit));
            for (int8_t blk = 0; blk < 7; ++blk) {
              readbin(boolean, bin, swapit);
            }
          } else {
            // ignore
            readbin(union_val, bin, swapit);
            readbin(union_val, bin, swapit);
          }

          if (vts == 6)  // a string
            vtsStringXls = XLWideString(bin, swapit);

          out << "<customFilter" << std::endl;
          out << " operator=\"" << grbitSgnOperator(grbitSgn) << "\"";
          if (vts == 6)
            out << " val=\"" << vtsStringXls << "\"";
          else
            out << " val=\"" << union_val << "\"";
          out << " />";

          break;
        }

        case BrtEndCustomFilters:
        case BrtEndCustomRichFilters: {
          out << "</customFilters>" << std::endl;

          break;
        }

        case BrtBeginUserCsViews:
        case BrtBeginUserShViews: {
          if (debug) Rcpp::Rcout << "BrtBeginUserXXViews" << std::endl;
          out << "<customSheetViews>" << std::endl;
          break;
        }

        case BrtBeginUserCsView: {
          if (debug) Rcpp::Rcout << "BrtBeginUserCsView" << std::endl;

          int32_t guid0 = 0, guid1 = 0, guid2 = 0, guid3 = 0, iTabId = 0, dwScale = 0, flags = 0;

          std::vector<int32_t> guids(4);
          guids[0] = readbin(guid0, bin, 0);
          guids[1] = readbin(guid1, bin, 0);
          guids[2] = readbin(guid2, bin, 0);
          guids[3] = readbin(guid3, bin, 0);

          iTabId = readbin(iTabId, bin, swapit);
          if (iTabId < 1 || iTabId > 65535)
            Rcpp::stop("iTabId out of range");

          dwScale = readbin(dwScale, bin, swapit);
          if (dwScale < 0 || dwScale > 400)  // dialog sheet 0 else 10
            Rcpp::stop("dwScale out of range");

          flags = readbin(flags, bin, swapit);
          // hsState
          // fZoomToFit

          out << "<customSheetView" << std::endl;
          out << " guid=\"{" << guid_str(guids) << "}\"";
          if (dwScale != 100)
            out << " scale=\"" << dwScale << "\"";
          out << ">" << std::endl;

          break;
        }

        case BrtBeginUserShView: {
          if (debug) Rcpp::Rcout << "BrtBeginUserShView" << std::endl;

          uint32_t flags = 0;
          int32_t guid0 = 0, guid1 = 0, guid2 = 0, guid3 = 0, iTabId = 0, dwScale = 0, icv = 0;

          std::vector<int32_t> guids(4);
          guids[0] = readbin(guid0, bin, 0);
          guids[1] = readbin(guid1, bin, 0);
          guids[2] = readbin(guid2, bin, 0);
          guids[3] = readbin(guid3, bin, 0);

          iTabId = readbin(iTabId, bin, swapit);
          if (iTabId < 1 || iTabId > 65535)
            Rcpp::stop("iTabId out of range");

          dwScale = readbin(dwScale, bin, swapit);
          if (dwScale < 0 || dwScale > 400)  // dialog sheet 0 else 10
            Rcpp::stop("dwScale out of range");

          icv = readbin(icv, bin, swapit);
          if (icv > 64)
            Rcpp::stop("icv out of range");

          flags = readbin(flags, bin, swapit);

          // rfxTopLeft
          std::vector<int32_t> rfx = UncheckedRfX(bin, swapit);

          BrtBeginUserShViewFlags view_flags(flags);

          out << "<customSheetView";
          out << " guid=\"{" << guid_str(guids) << "}\"";
          if (dwScale != 100)
            out << " scale=\"" << dwScale << "\"";
          if (icv != 64)
            out << " colorId=\"" << (int32_t)icv << "\"";
          if (view_flags.fShowBrks())
            out << " showPageBreaks=\"" << (int16_t)view_flags.fShowBrks() << "\"";
          if (view_flags.fDspFmlaSv())
            out << " showFormulas=\"" << (int16_t)view_flags.fDspFmlaSv() << "\"";
          if (!view_flags.fDspGridSv())
            out << " showGridLines=\"" << (int16_t)view_flags.fDspGridSv() << "\"";
          if (!view_flags.fDspRwColSv())
            out << " showRowCol=\"" << (int16_t)view_flags.fDspRwColSv() << "\"";
          if (!view_flags.fDspGutsSv())
            out << " outlineSymbols=\"" << (int16_t)view_flags.fDspGutsSv() << "\"";
          if (!view_flags.fDspZerosSv())
            out << " zeroValues=\"" << (int16_t)view_flags.fDspZerosSv() << "\"";
          if (view_flags.fFitToPage())
            out << " fitToPage=\"" << (int16_t)view_flags.fFitToPage() << "\"";
          if (view_flags.fPrintArea())
            out << " printArea=\"" << (int16_t)view_flags.fPrintArea() << "\"";
          if (view_flags.fFilterMode())
            out << " filter=\"" << (int16_t)view_flags.fFilterMode() << "\"";
          if (view_flags.fEzFilter())
            out << " showAutoFilter=\"" << (int16_t)view_flags.fEzFilter() << "\"";
          // if (iTabId) // not used?
          //   out << " tabSelected=\"" << iTabId << "\"";
          if (view_flags.fHiddenRw())
            out << " hiddenRows=\"" << (int16_t)view_flags.fHiddenRw() << "\"";
          if (view_flags.fHiddenCol())
            out << " hiddenColumns=\"" << (int16_t)view_flags.fHiddenCol() << "\"";
          if (view_flags.hsState())
            out << " state=\"" << HSState(view_flags.hsState()) << "\"";
          if (view_flags.fFilterUnique())
            out << " filterUnique=\"" << (int16_t)view_flags.fFilterUnique() << "\"";
          if (view_flags.fSheetLayoutView())
            out << " view=\"" << "pageBreakPreview" << "\"";
          if (view_flags.fPageLayoutView())
            out << " view=\"" << "pageLayout" << "\"";
          if (!view_flags.fRuler())
            out << " showRuler=\"" << (int16_t)view_flags.fRuler() << "\"";
          if (rfx[0] > 1 && rfx[2] > 0)
            out << " topLeftCell=\"" << int_to_col(rfx[2] + 1) << std::to_string(rfx[0] + 1) << "\"";

          out << ">" << std::endl;

          // order matters for <customSheetViews/>
          out << "<printOptions" << std::endl;
          if (view_flags.fHorizontal())
            out << " horizontalCentered = \"" << (int16_t)view_flags.fHorizontal() << "\"";
          if (view_flags.fVertical())
            out << " verticalCentered = \"" << (int16_t)view_flags.fVertical() << "\"";
          if (view_flags.fPrintRwCol())
            out << " headings = \"" << (int16_t)view_flags.fPrintRwCol() << "\"";
          if (view_flags.fDspGridSv())
            out << " gridLines = \"" << (int16_t)view_flags.fDspGridSv() << "\"";
          if (!view_flags.fPrintGrid())
            out << " gridLinesSet = \"" << (int16_t)view_flags.fPrintGrid() << "\"";
          out << " />" << std::endl;

          break;
        }

        case BrtEndUserCsView:
        case BrtEndUserShView: {
          if (debug) Rcpp::Rcout << "BrtEndUserXXView" << std::endl;
          out << "</customSheetView>" << std::endl;
          break;
        }

        case BrtEndUserCsViews:
        case BrtEndUserShViews: {
          if (debug) Rcpp::Rcout << "BrtEndUserXXViews" << std::endl;
          out << "</customSheetViews>" << std::endl;
          break;
        }

        case BrtDynamicFilter: {
          if (debug) Rcpp::Rcout << "BrtDynamicFilter" << std::endl;
          Rcpp::warning("Dynamic Filter found. This is not handled.");
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtIconFilter:
        case BrtIconFilter14: {
          if (debug) Rcpp::Rcout << "<BrtIconFilter>" << std::endl;
          uint32_t iIconSet = 0, iIcon = 0;
          iIconSet = readbin(iIconSet, bin, swapit);
          iIcon = readbin(iIcon, bin, swapit);

          std::string iconSet;
          if (iIconSet) iconSet = to_iconset(iIconSet);

          out << "<iconFilter iconSet=\"" << iconSet << "\" iconId=\""<< iIcon<< "\"/>" << std::endl;

          break;
        }

        case BrtEndFilterColumn: {
          if (debug) Rcpp::Rcout << "BrtEndFilterColumn" << std::endl;
          out << "</filterColumn>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtEndAFilter: {
          if (debug) Rcpp::Rcout << "</autofilter>" << std::endl;
          out << "</autoFilter>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtBeginListParts: {
          if (debug) Rcpp::Rcout << "<tableParts>" << std::endl;
          uint32_t cParts = 0;
          cParts = readbin(cParts, bin, swapit);
          out << "<tableParts count=\"" << cParts << "\">" << std::endl;
          break;
        }

        case BrtListPart: {
          if (debug) Rcpp::Rcout << "<tablePart/>" << std::endl;
          std::string stRelID = XLNullableWideString(bin, swapit);
          out << "<tablePart r:id=\"" << stRelID << "\" />" << std::endl;
          break;
        }

        case BrtEndListParts: {
          if (debug) Rcpp::Rcout << "</tableParts>" << std::endl;
          out << "</tableParts>" << std::endl;
          break;
        }

        case BrtBeginOleObjects: {
          if (debug) Rcpp::Rcout << "<oleObjects>" << std::endl;
          out << "<oleObjects>" << std::endl;
          break;
        }

        case BrtOleObject: {
          if (debug) Rcpp::Rcout << "<oleObject/>" << std::endl;

          uint16_t flags = 0;
          uint32_t dwAspect = 0, dwOleUpdate = 0, shapeId = 0;
          dwAspect = readbin(dwAspect, bin, swapit);
          dwOleUpdate = readbin(dwOleUpdate, bin, swapit);
          shapeId = readbin(shapeId, bin, swapit);
          flags = readbin(flags, bin, swapit);

          BrtOleObjectFlags view_flags(flags);

          std::string strProgID = XLNullableWideString(bin, swapit);

          int32_t sharedFormula = false;
          if (view_flags.fLinked()) std::string link = CellParsedFormula(bin, swapit, debug, 0, 0, sharedFormula, has_revision_record);

          std::string stRelID = XLNullableWideString(bin, swapit);
          out << "<oleObject progId=\"" << strProgID << "\" shapeId=\"" << shapeId << "\" r:id=\"" << stRelID << "\" />" << std::endl;
          break;
        }

        case BrtEndOleObjects: {
          if (debug) Rcpp::Rcout << "</oleObjects>" << std::endl;

          out << "</oleObjects>" << std::endl;
          break;
        }

        case BrtBeginDVals: {
          if (debug) Rcpp::Rcout << "BrtBeginDVals" << std::endl;

          uint16_t flags = 0;
          uint32_t xLeft = 0, yTop = 0, unused3 = 0, idvMac = 0;

          // openxlsx2 does not handle this?
          flags   = readbin(flags, bin, swapit);
          xLeft   = readbin(xLeft, bin, swapit);
          yTop    = readbin(yTop, bin, swapit);
          unused3 = readbin(unused3, bin, swapit);
          idvMac  = readbin(idvMac, bin, swapit);

          if (idvMac > 65534) Rcpp::stop("idvMac to big");

          out << "<dataValidations>" << std::endl;
          break;
        }

        case BrtDVal: {
          if (debug) Rcpp::Rcout << "BrtDVal" << std::endl;

          uint32_t flags = 0;
          std::vector<int32_t> sqrfx;
          std::string strErrorTitle = "", strError = "", strPromptTitle = "", strPrompt = "";

          flags = readbin(flags, bin, swapit);

          BrtDValFlags view_flags(flags);

          sqrfx          = UncheckedSqRfX(bin, swapit);
          strErrorTitle  = XLNullableWideString(bin, swapit);
          strError       = XLNullableWideString(bin, swapit);
          strPromptTitle = XLNullableWideString(bin, swapit);
          strPrompt      = XLNullableWideString(bin, swapit);

          int32_t sharedFormula = false;
          std::string formula1 = CellParsedFormula(bin, swapit, debug, 0, 0, sharedFormula, has_revision_record);
          std::string formula2 = CellParsedFormula(bin, swapit, debug, 0, 0, sharedFormula, has_revision_record);

          std::string lref = int_to_col(sqrfx[3] + 1) + std::to_string(sqrfx[1] + 1);
          std::string rref = int_to_col(sqrfx[4] + 1) + std::to_string(sqrfx[2] + 1);

          std::string sqref;
          if (lref.compare(rref) == 0) {
            sqref = lref;
          } else {
            sqref =  lref + ":" + rref;
          }

          out << "<dataValidation " <<
          "type=\"" << valType(view_flags.valType()) << "\" " <<
          "operator=\"" << typOperator(view_flags.typOperator()) << "\" " <<
          "allowBlank=\"" << (bool)view_flags.fAllowBlank() << "\" " <<
          "showInputMessage=\"" << (bool)view_flags.fShowInputMsg() << "\" " <<
          "showErrorMessage=\"" << (bool)view_flags.fShowErrorMsg() << "\" " <<
          "sqref=\"" << sqref << "\" >" <<
          "<formula1>" << formula1 << "</formula1>" <<
          "<formula2>" << formula2 << "</formula2>" <<
          "</dataValidation>" << std::endl;

          break;
        }

        case BrtDValList: {
          Rcpp::Rcout << "BrtDValList" << std::endl;

          std::string strDvalList = XLWideString(bin, swapit);
          Rcpp::Rcout << strDvalList << std::endl;

          break;
        }

        case BrtEndDVals: {
          if (debug) Rcpp::Rcout << "BrtEndDVals" << std::endl;
          out << "</dataValidations>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtBeginDVals14: {
          if (debug) Rcpp::Rcout << "BrtBeginDVals14" << std::endl;
          // Rcpp::warning("Worksheet contains unhandled data validation.");

          // FRT blank
          uint32_t reserved = 0;
          reserved = readbin(reserved, bin, swapit);

          // Dvals
          uint16_t flags = 0;
          uint32_t xLeft = 0, yTop = 0, unused3 = 0, idvMac = 0;

          // openxlsx2 does not handle this?
          flags   = readbin(flags, bin, swapit);
          xLeft   = readbin(xLeft, bin, swapit);
          yTop    = readbin(yTop, bin, swapit);
          unused3 = readbin(unused3, bin, swapit);
          idvMac  = readbin(idvMac, bin, swapit);

          if (idvMac > 65534) Rcpp::stop("idvMac to big");

          open_ext("{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}");
          out << "<x14:dataValidations count=\"" <<
            std::to_string(idvMac) <<
            "\" xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\""
            ">" << std::endl;

          break;
        }

        case BrtDVal14: {
          if (debug) Rcpp::Rcout << "BrtDVal14" << std::endl;

          // Start FRTHeader
          uint32_t flags_frt = 0;
          std::vector<std::string> frt_sqrefs;
          std::vector<std::string> frt_fmls;

          flags_frt = readbin(flags_frt, bin, swapit);

          FRTHeaderFlags view_flags(flags_frt);

          if (view_flags.fRef()) { // but must be 0
            // Rcpp::Rcout << "fRef" << std::endl;
            uint32_t cref = 0;
            cref = readbin(cref, bin, swapit);
            // Rcpp::Rcout << "Refs: " << cref << std::endl;

            for (uint32_t crf = 0; crf < cref; ++crf) {
              uint32_t flags_cref = 0;
              flags_cref = readbin(flags_cref, bin, swapit);
              std::vector<int32_t> rfx_cref = UncheckedRfX(bin, swapit);
            }

          }

          if (view_flags.fSqref()) { // must be 1
            // Rcpp::Rcout << "fSqref" << std::endl;
            uint32_t csqref = 0;
            csqref = readbin(csqref, bin, swapit);
            // Rcpp::Rcout << "Sqrefs: " << csqref << std::endl;

            for (uint32_t csqrf = 0; csqrf < csqref; ++csqrf) {
              uint32_t flags_csqrf = 0;
              flags_csqrf = readbin(flags_csqrf, bin, swapit);

              /* beg move this to function -------------------------------------*/
              std::vector<int32_t> rfx_csqrfx = UncheckedSqRfX(bin, swapit);

              std::string lref = int_to_col(rfx_csqrfx[3] + 1) + std::to_string(rfx_csqrfx[1] + 1);
              std::string rref = int_to_col(rfx_csqrfx[4] + 1) + std::to_string(rfx_csqrfx[2] + 1);

              std::string sqref;
              if (lref.compare(rref) == 0) {
                sqref = lref;
              } else {
                sqref =  lref + ":" + rref;
              }
              /* end move this to function -------------------------------------*/

              frt_sqrefs.push_back(sqref);
            }
          }

          if (view_flags.fFormula()) {  // 0 or 1
            // Rcpp::Rcout << "fFormula" << std::endl;
            uint32_t cformula = 0;
            cformula = readbin(cformula, bin, swapit);
            // Rcpp::Rcout << "Formulas: " << cformula << std::endl;

            for (uint32_t cfml = 0; cfml < cformula; ++cfml) {
              uint32_t flags_cfml = 0;
              flags_cfml = readbin(flags_cfml, bin, swapit);
              int32_t sharedFormula = false;
              std::string frt_fml = FRTParsedFormula(bin, swapit, debug, 0, 0, sharedFormula, has_revision_record);
              frt_fmls.push_back(frt_fml);
              // Rcpp::Rcout << frt_fmls[cfml] << std::endl;
            }
          }

          if (view_flags.fRelID()) {  // 0
            // Rcpp::Rcout << "fRelID" << std::endl;
            std::string relId = LPWideString(bin, swapit);
          }
          // End FRTHeader

          uint32_t flags = 0;
          std::string strErrorTitle = "", strError = "", strPromptTitle = "", strPrompt = "";

          flags = readbin(flags, bin, swapit);

          BrtDValFlags view_flags2(flags);

          if (debug) {  // #nocov start
            Rprintf("%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d\n",
              (uint32_t)view_flags2.valType(),
              (uint32_t)view_flags2.errStyle(),
              (uint32_t)view_flags2.unused(),
              (uint32_t)view_flags2.fAllowBlank(),
              (uint32_t)view_flags2.fSuppressCombo(),
              (uint32_t)view_flags2.mdImeMode(),
              (uint32_t)view_flags2.fShowInputMsg(),
              (uint32_t)view_flags2.fShowErrorMsg(),
              (uint32_t)view_flags2.typOperator(),
              (uint32_t)view_flags2.fDVMinFmla(),
              (uint32_t)view_flags2.fDVMaxFmla()
            );
          }  // #nocov end

          strErrorTitle  = XLNullableWideString(bin, swapit);
          strError       = XLNullableWideString(bin, swapit);
          strPromptTitle = XLNullableWideString(bin, swapit);
          strPrompt      = XLNullableWideString(bin, swapit);

          out << "<x14:dataValidation " <<
          "type=\"" << valType(view_flags2.valType()) << "\" " <<
          "operator=\"" << typOperator(view_flags2.typOperator()) << "\" " <<
          "allowBlank=\"" << (bool)view_flags2.fAllowBlank() << "\" " <<
          "showInputMessage=\"" << (bool)view_flags2.fShowInputMsg() << "\" " <<
          "showErrorMessage=\"" << (bool)view_flags2.fShowErrorMsg() << "\" " <<
          // "xr:uid=\"{00000000-0002-0000-0200-000000000000}\" " <<
          ">";

          for (size_t len = 0; len < frt_fmls.size(); ++len) {

            out <<
              "<x14:formula" << std::to_string(len + 1) << ">" <<
              "<xm:f>" <<
              frt_fmls[len] <<
              "</xm:f>" <<
              "</x14:formula" << std::to_string(len + 1) << ">";
          }
          for (size_t len = 0; len < frt_fmls.size(); ++len) {

            out <<
              "<xm:sqref>" <<
              frt_sqrefs[len] <<
              "</xm:sqref>";
          }

          out  << "</x14:dataValidation>" << std::endl;

          break;
        }

        case BrtEndDVals14: {
          if (debug) Rcpp::Rcout << "BrtEndDVals14" << std::endl;
          out << "</x14:dataValidations>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }


        case BrtBeginDCon:
        case BrtEndDCon: {
          if (debug) Rcpp::Rcout << "BrtBegin/EndDCon" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtBeginConditionalFormattings:
        case BrtBeginConditionalFormatting14: {
          if (debug) Rcpp::Rcout << "BrtBeginConditionalFormattings" << std::endl;

          Rcpp::warning("Worksheet contains unhandled conditional formatting.");
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtBeginConditionalFormatting: {
          if (debug) Rcpp::Rcout << "<conditionalFormatting>" << std::endl;

          uint32_t ccf = 0, fPivot = 0;
          ccf = readbin(ccf, bin, swapit);
          fPivot = readbin(fPivot, bin, swapit);

          std::vector<int32_t> sqrfx = UncheckedSqRfX(bin, swapit);
          std::string sqref;
          for (int32_t k = 0; k < sqrfx[0]; ++k) {
            size_t at = 1 + 4 * static_cast<size_t>(k);
            std::string lref = int_to_col(sqrfx[at + 2] + 1) + std::to_string(sqrfx[at] + 1);
            std::string rref = int_to_col(sqrfx[at + 3] + 1) + std::to_string(sqrfx[at + 1] + 1);
            if (k > 0) sqref += " ";
            sqref += lref == rref ? lref : lref + ":" + rref;
          }
          if (sqrfx[0] > 0) {
            cf_row = sqrfx[1];
            cf_col = sqrfx[3];
          }

          out << "<conditionalFormatting";
          if (fPivot) out << " pivot=\"" << fPivot << "\"";
          out << " sqref=\"" << sqref << "\"";
          out << ">" << std::endl;

          break;
        }

        case BrtBeginCFRule: {
          if (debug) Rcpp::Rcout << "<cfRule>" << std::endl;

          std::streampos rec_end = bin.tellg() + static_cast<std::streamoff>(size);

          uint16_t flags = 0;
          uint32_t iType = 0, iTemplate = 0, dxfId = 0, iPri = 0, iParam = 0, reserved1 = 0, reserved2 = 0;
          uint32_t cbFmla1 = 0, cbFmla2 = 0, cbFmla3 = 0;

          iType = readbin(iType, bin, swapit);
          iTemplate = readbin(iTemplate, bin, swapit);
          dxfId = readbin(dxfId, bin, swapit);
          iPri = readbin(iPri, bin, swapit);
          iParam = readbin(iParam, bin, swapit);
          reserved1 = readbin(reserved1, bin, swapit);
          reserved2 = readbin(reserved2, bin, swapit);
          flags = readbin(flags, bin, swapit);
          cbFmla1 = readbin(cbFmla1, bin, swapit);
          cbFmla2 = readbin(cbFmla2, bin, swapit);
          cbFmla3 = readbin(cbFmla3, bin, swapit);

          std::string strParam = XLNullableWideString(bin, swapit);

          int32_t sharedFormula = false;
          std::string rgce1, rgce2, rgce3;
          if (cbFmla1 != 0x00000000) {
            rgce1 = CellParsedFormula(bin, swapit, debug, cf_col, cf_row, sharedFormula, has_revision_record);
          }
          if (cbFmla2 != 0x00000000) {
            rgce2 = CellParsedFormula(bin, swapit, debug, cf_col, cf_row, sharedFormula, has_revision_record);
          }
          if (cbFmla3 != 0x00000000) {
            rgce3 = CellParsedFormula(bin, swapit, debug, cf_col, cf_row, sharedFormula, has_revision_record);
          }

          static const char* const operators[9] = {"", "between", "notBetween", "equal", "notEqual", "greaterThan",
                                                   "lessThan", "greaterThanOrEqual", "lessThanOrEqual"};
          static const char* const periods[10] = {"today", "tomorrow", "yesterday", "last7Days", "lastMonth",
                                                  "nextMonth", "thisWeek", "nextWeek", "lastWeek", "thisMonth"};
          static const char* const text_types[4] = {"containsText", "notContainsText", "beginsWith", "endsWith"};
          static const char* const text_operators[4] = {"containsText", "notContains", "beginsWith", "endsWith"};

          bool has_dxf = false;
          std::stringstream cfRule;
          cfRule << "<cfRule";

          if (iType == 1) {
            cfRule << " type=\"cellIs\"";
            if (iParam >= 1 && iParam <= 8) cfRule << " operator=\"" << operators[iParam] << "\"";
            has_dxf = true;
          } else if (iType == 2) {
            has_dxf = true;
            if (iTemplate == 2) cfRule << " type=\"uniqueValues\"";
            else if (iTemplate == 3) {
              uint32_t k = iParam < 4 ? iParam : 0;
              cfRule << " type=\"" << text_types[k] << "\" operator=\"" << text_operators[k] << "\" text=\"" << escape_xml(strParam) << "\"";
            }
            else if (iTemplate == 4) cfRule << " type=\"containsBlanks\"";
            else if (iTemplate == 5) cfRule << " type=\"notContainsBlanks\"";
            else if (iTemplate == 6) cfRule << " type=\"containsErrors\"";
            else if (iTemplate == 7) cfRule << " type=\"notContainsErrors\"";
            else if (iTemplate >= 8 && iTemplate <= 17) cfRule << " type=\"timePeriod\" timePeriod=\"" << periods[iTemplate - 8] << "\"";
            else if (iTemplate >= 18 && iTemplate <= 19) {
              cfRule << " type=\"aboveAverage\"";
              if (iTemplate == 19) cfRule << " aboveAverage=\"0\"";
              if (iParam > 0) cfRule << " stdDev=\"" << iParam << "\"";
            }
            else if (iTemplate == 20) cfRule << " type=\"duplicateValues\"";
            else if (iTemplate >= 21 && iTemplate <= 22) {
              cfRule << " type=\"aboveAverage\"";
              if (iTemplate == 22) cfRule << " aboveAverage=\"0\"";
              cfRule << " equalAverage=\"1\"";
              if (iParam > 0) cfRule << " stdDev=\"" << iParam << "\"";
            }
            else cfRule << " type=\"expression\"";
          } else if (iType == 3) {
            cfRule << " type=\"colorScale\"";
          } else if (iType == 4) {
            cfRule << " type=\"dataBar\"";
          } else if (iType == 5) {
            cfRule << " type=\"top10\" rank=\"" << iParam << "\"";
            if (flags & 0x10) cfRule << " percent=\"1\"";
            if (flags & 0x08) cfRule << " bottom=\"1\"";
            has_dxf = true;
          } else if (iType == 6) {
            cfRule << " type=\"iconSet\"";
          }

          if (has_dxf && dxfId != 0xFFFFFFFFu) cfRule << " dxfId=\"" << dxfId << "\"";
          cfRule << " priority=\"" << iPri << "\"";
          if (flags & 0x02) cfRule << " stopIfTrue=\"1\"";
          cfRule << ">";
          if (!rgce1.empty()) cfRule << "<formula>" << rgce1 << "</formula>";
          if (!rgce2.empty()) cfRule << "<formula>" << rgce2 << "</formula>";
          if (!rgce3.empty()) cfRule << "<formula>" << rgce3 << "</formula>";

          if (debug) Rcpp::Rcout << cfRule.str() << std::endl;
          out << cfRule.str() << std::endl;

          bin.seekg(rec_end);
          break;
        }

        case BrtBeginColorScale: {
          out << "<colorScale>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtEndColorScale: {
          out << "</colorScale>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtBeginDatabar: {
          uint8_t bLenMin = 0, bLenMax = 0, fShowValue = 0;
          bLenMin = readbin(bLenMin, bin, swapit);
          bLenMax = readbin(bLenMax, bin, swapit);
          fShowValue = readbin(fShowValue, bin, swapit);
          out << "<dataBar minLength=\"" << static_cast<unsigned>(bLenMin) << "\" maxLength=\"" << static_cast<unsigned>(bLenMax) << "\"";
          if (!fShowValue) out << " showValue=\"0\"";
          out << ">" << std::endl;
          bin.seekg(size - 3, bin.cur);
          break;
        }

        case BrtEndDatabar: {
          out << "</dataBar>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtBeginIconSet: {
          static const char* const icons[17] = {"3Arrows", "3ArrowsGray", "3Flags", "3TrafficLights1", "3Signs",
                                                "3TrafficLights2", "3Symbols", "3Symbols2", "4Arrows", "4ArrowsGray",
                                                "4RedToBlack", "4Rating", "4TrafficLights", "5Arrows", "5ArrowsGray",
                                                "5Rating", "5Quarters"};
          uint32_t iSet = 0;
          uint16_t icon_flags = 0;
          iSet = readbin(iSet, bin, swapit);
          icon_flags = readbin(icon_flags, bin, swapit);
          out << "<iconSet iconSet=\"" << icons[iSet < 17 ? iSet : 0] << "\"";
          if (icon_flags & 0x02) out << " showValue=\"0\"";
          if (!(icon_flags & 0x04)) out << " reverse=\"1\"";
          out << ">" << std::endl;
          cf_iconset = true;
          bin.seekg(size - 6, bin.cur);
          break;
        }

        case BrtEndIconSet: {
          out << "</iconSet>" << std::endl;
          cf_iconset = false;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtCFVO: {
          std::streampos rec_end = bin.tellg() + static_cast<std::streamoff>(size);

          uint32_t iType = 0, fSaveGTE = 0, fGTE = 0, cbFmla = 0;
          double numParam = 0;
          iType = readbin(iType, bin, swapit);
          numParam = readbin(numParam, bin, swapit);
          fSaveGTE = readbin(fSaveGTE, bin, swapit);
          fGTE = readbin(fGTE, bin, swapit);
          cbFmla = readbin(cbFmla, bin, swapit);

          std::string fml;
          if (cbFmla != 0) {
            int32_t sharedFormula = false;
            fml = CellParsedFormula(bin, swapit, debug, cf_col, cf_row, sharedFormula, has_revision_record);
          }

          static const char* const names[8] = {"num", "num", "min", "max", "percent", "percentile", "num", "formula"};
          out << "<cfvo type=\"" << names[iType < 8 ? iType : 0] << "\"";
          if (!fml.empty()) {
            out << " val=\"" << fml << "\"";
          } else if (iType != 2 && iType != 3) {
            std::stringstream num;
            num << std::setprecision(15) << numParam;
            out << " val=\"" << num.str() << "\"";
          }
          if (cf_iconset && !fGTE) out << " gte=\"0\"";
          out << "/>" << std::endl;

          bin.seekg(rec_end);
          break;
        }

        case BrtColor: {
          std::vector<uint8_t> color(static_cast<size_t>(size));
          bin.read(reinterpret_cast<char*>(color.data()), size);
          std::string attr = xlsb_color_attr(color, 0);
          out << "<color" << (attr.empty() ? " auto=\"1\"" : attr) << "/>" << std::endl;
          break;
        }

        case BrtEndConditionalFormatting14:
        case BrtBeginCFRule14:
        case BrtEndCFRule14: {
          if (debug) Rcpp::Rcout << "BrtBegin/EndDCon 14" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        // dxf styles for conditional formatting can be (or is always?) embedded in the worksheet
        // the xlsx conversion creates <x14:dxf/>
        case BrtBeginDXFs:
        case BrtDXF:
        case BrtEndDXFs: {
          if (debug) Rcpp::Rcout << "BrtBegin/EndDXFs in worksheet (?)" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtEndCFRule: {
          if (debug) Rcpp::Rcout << "</cfRule>" << std::endl;
          out << "</cfRule>" << std::endl;
          break;
        }

        case BrtEndConditionalFormatting: {
          if (debug) Rcpp::Rcout << "</conditionalFormatting>" << std::endl;
          out << "</conditionalFormatting>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtEndConditionalFormattings: {
          // UNHANDLED
          if (debug) Rcpp::Rcout << "</conditionalFormattings>" << std::endl;
          bin.seekg(size, bin.cur);
          break;
        }

        // unhandled page breaks. differentiates in column and row breaks
        case BrtBeginColBrk:
        case BrtBeginRwBrk: {

          if (debug) Rcpp::Rcout << "<BrtBegin..Brk>" << std::endl;

          uint32_t ibrkMac = 0, ibrkManMac = 0;
          ibrkMac = readbin(ibrkMac, bin, swapit);
          if (ibrkMac > 1024) Rcpp::stop("ibrkMac to big");
          ibrkManMac = readbin(ibrkManMac, bin, swapit);

          break;
        }

        case BrtBrk: {
          if (debug) Rcpp::Rcout << "<BrtBrk>" << std::endl;

          uint32_t unRwCol = 0, unColRwStrt = 0, unColRwEnd = 0, fMan = 0, fPivot = 0;

          unRwCol     = readbin(unRwCol, bin, swapit);  // row or column
          unColRwStrt = readbin(unColRwStrt, bin, swapit);
          unColRwEnd  = readbin(unColRwEnd, bin, swapit);
          fMan        = readbin(fMan, bin, swapit);
          fPivot      = readbin(fPivot, bin, swapit);

          break;
        }

        case BrtEndColBrk:
        case BrtEndRwBrk: {
          if (debug) Rcpp::Rcout << "<BrtEnd..Brk>" << std::endl;

          break;
        }
          // end unhandled page breaks

        case BrtFRTBegin: {
          if (debug) Rcpp::Rcout << "BrtFRTBegin" << std::endl;

          ProductVersion(bin, swapit, debug, true);

          break;
        }

        case BrtBeginSlicersEx: {
          open_ext("{A8765BA9-456A-4dab-B4F3-ACF838C121DE}");
          out << "<x14:slicerList>";
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtBeginSlicerEx: {
          std::streampos start = bin.tellg();
          bin.seekg(4, bin.cur);
          uint16_t cch = 0;
          cch = readbin(cch, bin, swapit);
          std::string rid;
          for (uint16_t i = 0; i < cch; ++i) {
            uint16_t ch = 0;
            ch = readbin(ch, bin, swapit);
            rid += static_cast<char>(ch);
          }
          out << "<x14:slicer r:id=\"" << rid << "\"/>";
          bin.seekg(start + static_cast<std::streamoff>(size), bin.beg);
          break;
        }

        case BrtEndSlicerEx: {
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtEndSlicersEx: {
          out << "</x14:slicerList>";
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtFRTEnd: {
          if (debug) Rcpp::Rcout << "BrtFRTEnd" << std::endl;
          if (ext_open) out << "</ext>" << std::endl;
          ext_open = false;
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtEndSheet: {
          if (hlinks.size() > 1) {
            // did not see BrtBeginHL or BrtEndHL, likely I'm just blind
            hlinks.push_back("</hyperlinks>");

            for (size_t i = 0; i < hlinks.size(); ++i) {
              if (debug) Rcpp::Rcout << hlinks[i] << std::endl;
              out << hlinks[i] << std::endl;
            }
          }

          if (debug)  Rcpp::Rcout << "</worksheet>" << bin.tellg() << std::endl;

          if (extlst_open) out << "</extLst>" << std::endl;

          if (chartsheet)
            out << "</chartsheet>" << std::endl;
          else
            out << "</worksheet>" << std::endl;
          end_of_worksheet = true;
          row = 0;
          break;
        }

        // ignore error section
        case BrtBeginCellIgnoreECs:
        case BrtCellIgnoreEC:
        case BrtEndCellIgnoreECs: {
          if (debug) Rcpp::Rcout << "BrtBeginCell" << std::endl;
          // Rcpp::warning("Worksheet contains unhandled data validation.");
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtRRChgCell:
        case BrtRRDefName: {
          has_revision_record = true;
          // -- have not seen this yet. if it appears, treat it as if a revision record was found --
          // rgce.rgce or rgceOld.rgce in BrtRRDefName
          if (debug) Rcpp::Rcout << "BrtRRChgCell or BrtRRDefName" << std::endl;
          Rcpp::warning("Assuming revision record.");
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtBeginSortState:
        case BrtEndSortState:
        case BrtBeginSortCond:
        case BrtEndSortCond: {
          if (debug) Rcpp::warning("Worksheet contains unhandled sorting.");
          bin.seekg(size, bin.cur);
          break;
        }

        case BrtValueMeta: {
          if (debug) Rcpp::Rcout << "BrtValueMeta" << std::endl;
          int32_t ivmb = 0;
          ivmb = readbin(ivmb, bin, swapit);
          break;
        }

        case BrtSheetProtectionIso:
        case BrtCsProtection: {
          if (debug) Rcpp::Rcout << "protection" << std::endl;
          // uint16_t protpwd = 0;
          // uint32_t fLocked = 0, fObjects = 0;
          // unhandled
          if (debug) Rcpp::warning("Worksheet contains unhandled protection.");
          bin.seekg(size, bin.cur);
          break;
        }

        default: {
          // if (debug) {
            Rcpp::Rcout << "Unhandled: " << std::to_string(x) <<
              ": " << std::to_string(size) <<
                " @ " << bin.tellg() << std::endl;
          // }
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

// Pivot tables and pivot caches of xlsb files. The parts are converted the way
// the xls parts are: the cache gets the layout and the items of its fields, the
// records of the cache are not imported and the cache is refreshed when the
// file is opened.

namespace {

struct BinRecord {
  int32_t id = 0;
  std::vector<uint8_t> data;
};

std::vector<BinRecord> pivot_records(const std::string& path) {
  std::vector<BinRecord> out;
  std::ifstream in(path, std::ios::in | std::ios::binary);
  if (!in) return out;
  std::vector<uint8_t> bytes((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());

  size_t pos = 0, n = bytes.size();
  while (pos < n) {
    BinRecord rec;
    uint32_t id = bytes[pos++];
    if ((id & 0x80) != 0) {
      if (pos >= n) break;
      id = (id & 0x7Fu) | ((bytes[pos++] & 0x7Fu) << 7);
    }
    uint32_t size = 0;
    bool done = false;
    for (unsigned k = 0; k < 4 && pos < n; ++k) {
      uint8_t b = bytes[pos++];
      size |= static_cast<uint32_t>(b & 0x7Fu) << (7 * k);
      if ((b & 0x80) == 0) {
        done = true;
        break;
      }
    }
    if (!done || size > n - pos) break;
    rec.id = static_cast<int32_t>(id);
    rec.data.assign(bytes.begin() + static_cast<std::ptrdiff_t>(pos), bytes.begin() + static_cast<std::ptrdiff_t>(pos + size));
    pos += size;
    out.push_back(rec);
  }
  return out;
}

// XLWideString and XLNullableWideString: the number of characters and the characters
std::string pivot_string(const std::vector<uint8_t>& r, size_t& pos) {
  if (pos + 4 > r.size()) {
    pos = r.size();
    return std::string();
  }
  uint32_t len = xls_pv_u32(r, pos);
  pos += 4;
  if (len == 0xFFFFFFFFu) return std::string();
  if (len > (r.size() - pos) / 2) {
    pos = r.size();
    return std::string();
  }
  std::u16string s;
  for (uint32_t i = 0; i < len; ++i) {
    s.push_back(static_cast<char16_t>(xls_pv_u16(r, pos)));
    pos += 2;
  }
  return to_utf8(s);
}

}  // namespace

// [[Rcpp::export]]
int32_t pivotcache_bin(std::string filePath, std::string outPath, std::string recordsPath, std::string recordsOut, bool debug) {
  std::vector<BinRecord> recs = pivot_records(filePath);
  if (recs.empty()) return -1;

  XlsPivotSource src;
  uint32_t src_type = 0, n_records = 0, slicer_id = 0;
  std::vector<XlsPivotCacheField> fields;
  bool in_items = false;

  for (size_t i = 0; i < recs.size(); ++i) {
    const std::vector<uint8_t>& d = recs[i].data;
    if (debug) Rcpp::Rcout << "pivot cache record " << recs[i].id << ": " << d.size() << std::endl;

    switch (recs[i].id) {
      case BrtBeginPivotCacheDef:
        n_records = xls_pv_u32(d, 17);
        break;

      case BrtBeginPCDSource:
        src_type = xls_pv_u32(d, 0);
        break;

      case BrtBeginPCD14:
        slicer_id = xls_pv_u32(d, 5);
        break;

      case BrtBeginPCDSRange: {
        if (d.size() < 3) break;
        bool is_name = (d[0] & 0x01) != 0, has_rel = (d[2] & 0x01) != 0, has_sheet = (d[2] & 0x02) != 0;
        size_t pos = 3;
        if (has_sheet) src.sheet = pivot_string(d, pos);
        if (has_rel) {
          src_type = 1;
          pivot_string(d, pos);
        }
        if (is_name) {
          src.name = pivot_string(d, pos);
        } else {
          src.ref = int_to_col(xls_pv_u32(d, pos + 8) + 1u) + std::to_string(xls_pv_u32(d, pos) + 1u) + ":" +
                    int_to_col(xls_pv_u32(d, pos + 12) + 1u) + std::to_string(xls_pv_u32(d, pos + 4) + 1u);
        }
        break;
      }

      case BrtBeginPCDField: {
        XlsPivotCacheField f;
        size_t pos = 20;
        f.name = pivot_string(d, pos);
        fields.push_back(f);
        break;
      }

      case BrtEndPCDField:
        in_items = false;
        break;

      case BrtBeginPCDFAtbl: {
        in_items = true;
        if (fields.empty()) break;
        XlsPivotCacheField& f = fields.back();
        uint32_t fl = xls_pv_u16(d, 0);
        f.flags = fl;
        f.shared = d.size() >= 6 && xls_pv_u32(d, 2) > 0;
        // flags: 0x1 semi mixed types, 0x2 non date, 0x4 date, 0x8 string, 0x10 blank, 0x40 number,
        // 0x80 integer, 0x100 minimum and maximum. Fields with items get them from the items.
        if (d.size() >= 6 && xls_pv_u32(d, 2) == 0) {
          std::string a;
          if ((fl & 0x01u) == 0) a += " containsSemiMixedTypes=\"0\"";
          if ((fl & 0x02u) == 0) a += " containsNonDate=\"0\"";
          if ((fl & 0x04u) != 0) a += " containsDate=\"1\"";
          if ((fl & 0x08u) == 0) a += " containsString=\"0\"";
          if ((fl & 0x10u) != 0) a += " containsBlank=\"1\"";
          if ((fl & 0x40u) != 0) a += " containsNumber=\"1\"";
          if ((fl & 0x80u) != 0) a += " containsInteger=\"1\"";
          if ((fl & 0x100u) != 0 && d.size() >= 22) {
            uint64_t b1 = static_cast<uint64_t>(xls_pv_u32(d, 6)) | (static_cast<uint64_t>(xls_pv_u32(d, 10)) << 32);
            uint64_t b2 = static_cast<uint64_t>(xls_pv_u32(d, 14)) | (static_cast<uint64_t>(xls_pv_u32(d, 18)) << 32);
            double serial[2] = {0, 0};
            std::memcpy(&serial[0], &b1, sizeof(double));
            std::memcpy(&serial[1], &b2, sizeof(double));
            if ((fl & 0x04u) == 0) {
              a += " minValue=\"" + xls_num(serial[0]) + "\" maxValue=\"" + xls_num(serial[1]) + "\"";
            } else {
              // dates are serial numbers, converted to year, month and day
              std::string iso[2];
              for (int k = 0; k < 2; ++k) {
                double whole = std::floor(serial[k]);
                long z = static_cast<long>(whole) - 25569 + 719468;
                long era = (z >= 0 ? z : z - 146096) / 146097;
                long doe = z - era * 146097;
                long yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
                long doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
                long mp = (5 * doy + 2) / 153;
                long day = doy - (153 * mp + 2) / 5 + 1;
                long month = mp < 10 ? mp + 3 : mp - 9;
                long year = yoe + era * 400 + (month <= 2 ? 1 : 0);
                long secs = std::lround((serial[k] - whole) * 86400.0);
                char buf[128];
                std::snprintf(buf, sizeof(buf), "%04ld-%02ld-%02ldT%02ld:%02ld:%02ld", year, month, day, secs / 3600, (secs / 60) % 60, secs % 60);
                iso[k] = buf;
              }
              a += " minDate=\"" + iso[0] + "\" maxDate=\"" + iso[1] + "\"";
            }
          }
          f.attrs = a;
        }
        break;
      }

      case BrtEndPCDFAtbl:
        in_items = false;
        break;

      case BrtPCDIMissing:
      case BrtPCDINumber:
      case BrtPCDIBoolean:
      case BrtPCDIError:
      case BrtPCDIString:
      case BrtPCDIDatetime:
      case BrtPCDIAMissing:
      case BrtPCDIANumber:
      case BrtPCDIABoolean:
      case BrtPCDIAError:
      case BrtPCDIAString:
      case BrtPCDIADatetime:
      case BrtBeginPCDIRun: {
        if (!in_items || fields.empty()) break;
        XlsPivotCacheField& f = fields.back();
        f.attrs.clear();

        // items are missing 0, number 1, boolean 2, error 3, string 4, datetime 5. The
        // items with attributes carry six more bytes, they are not imported. A run holds
        // many items of the same type.
        int kind = 0;
        size_t n_items = 1, pos = 0;
        if (recs[i].id == BrtBeginPCDIRun) {
          uint32_t run_type = xls_pv_u16(d, 0);
          n_items = xls_pv_u32(d, 2);
          pos = 6;
          if (run_type == 1) kind = 1;
          else if (run_type == 2) kind = 4;
          else if (run_type == 0x20) kind = 5;
          else return -1;
        } else {
          kind = recs[i].id >= BrtPCDIAMissing ? recs[i].id - BrtPCDIAMissing : recs[i].id - BrtPCDIMissing;
        }

        for (size_t k = 0; k < n_items && pos <= d.size(); ++k) {
          if (kind == 4) {
            f.items += "<s v=\"" + escape_xml(pivot_string(d, pos)) + "\"/>";
            ++f.n_str;
          } else if (kind == 1) {
            uint64_t bits = static_cast<uint64_t>(xls_pv_u32(d, pos)) | (static_cast<uint64_t>(xls_pv_u32(d, pos + 4)) << 32);
            pos += 8;
            double x = 0;
            std::memcpy(&x, &bits, sizeof(x));
            if (f.n_num == 0 || x < f.min) f.min = x;
            if (f.n_num == 0 || x > f.max) f.max = x;
            if (x == std::floor(x)) ++f.n_int;
            f.items += "<n v=\"" + xls_num(x) + "\"/>";
            ++f.n_num;
          } else if (kind == 2) {
            f.items += std::string("<b v=\"") + ((pos < d.size() && d[pos] != 0) ? "1" : "0") + "\"/>";
            ++pos;
            ++f.n_other;
          } else if (kind == 3) {
            f.items += "<e v=\"" + xls_error(pos < d.size() ? d[pos] : static_cast<uint8_t>(0)) + "\"/>";
            ++pos;
            ++f.n_other;
          } else if (kind == 5) {
            char buf[32];
            std::snprintf(buf, sizeof(buf), "%04u-%02u-%02uT%02u:%02u:%02u", xls_pv_u16(d, pos), xls_pv_u16(d, pos + 2),
                          pos + 4 < d.size() ? d[pos + 4] : 0u, pos + 5 < d.size() ? d[pos + 5] : 0u,
                          pos + 6 < d.size() ? d[pos + 6] : 0u, pos + 7 < d.size() ? d[pos + 7] : 0u);
            pos += 8;
            std::string dt = buf;
            if (f.n_date == 0 || dt < f.min_date) f.min_date = dt;
            if (f.n_date == 0 || dt > f.max_date) f.max_date = dt;
            f.items += "<d v=\"" + dt + "\"/>";
            ++f.n_date;
          } else {
            f.items += "<m/>";
            ++f.n_blank;
          }
          ++f.count;
        }
        break;
      }

      default:
        break;
    }
  }

  if (src_type != 0 || (src.name.empty() && src.ref.empty()) || fields.empty()) return -1;

  // The records: fields with items store the index of the item, the others the value. A file in
  // which a record does not end where its fields end is not imported, the cache is refreshed.
  bool refresh = true;
  if (!recordsPath.empty()) {
    std::ifstream rin(recordsPath, std::ios::in | std::ios::binary);
    std::vector<uint8_t> bytes((std::istreambuf_iterator<char>(rin)), std::istreambuf_iterator<char>());
    std::ofstream rout(recordsOut);
    std::vector<uint8_t> rec;
    std::string line;
    uint32_t n_rec = 0;
    bool ok = !bytes.empty();

    rout << "<pivotCacheRecords xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">";

    size_t pos = 0, n = bytes.size();
    while (ok && pos < n) {
      uint32_t id = bytes[pos++];
      if ((id & 0x80) != 0) {
        if (pos >= n) break;
        id = (id & 0x7Fu) | ((bytes[pos++] & 0x7Fu) << 7);
      }
      uint32_t size = 0;
      bool done = false;
      for (unsigned k = 0; k < 4 && pos < n; ++k) {
        uint8_t b = bytes[pos++];
        size |= static_cast<uint32_t>(b & 0x7Fu) << (7 * k);
        if ((b & 0x80) == 0) {
          done = true;
          break;
        }
      }
      if (!done || size > n - pos) {
        ok = false;
        break;
      }
      if (id == BrtPCRRecordDt) {
        ok = false;
        break;
      }
      if (id != BrtPCRRecord) {
        pos += size;
        continue;
      }

      rec.assign(bytes.begin() + static_cast<std::ptrdiff_t>(pos), bytes.begin() + static_cast<std::ptrdiff_t>(pos + size));
      pos += size;

      size_t at = 0;
      line = "<r>";
      for (size_t k = 0; ok && k < fields.size(); ++k) {
        const XlsPivotCacheField& f = fields[k];
        if (f.shared) {
          if (at + 4 > rec.size()) { ok = false; break; }
          line += "<x v=\"" + std::to_string(xls_pv_u32(rec, at)) + "\"/>";
          at += 4;
        } else if ((f.flags & 0x08u) != 0) {
          std::string str = pivot_string(rec, at);
          line += "<s v=\"" + escape_xml(str) + "\"/>";
        } else if ((f.flags & 0x40u) != 0 && (f.flags & 0x04u) == 0) {
          if (at + 8 > rec.size()) { ok = false; break; }
          uint64_t bits = static_cast<uint64_t>(xls_pv_u32(rec, at)) | (static_cast<uint64_t>(xls_pv_u32(rec, at + 4)) << 32);
          double num = 0;
          std::memcpy(&num, &bits, sizeof(double));
          line += "<n v=\"" + xls_num(num) + "\"/>";
          at += 8;
        } else if ((f.flags & 0x04u) != 0 && (f.flags & 0x40u) == 0) {
          if (at + 8 > rec.size()) { ok = false; break; }
          char buf[128];
          std::snprintf(buf, sizeof(buf), "%04u-%02u-%02uT%02u:%02u:%02u", xls_pv_u16(rec, at), xls_pv_u16(rec, at + 2),
                        static_cast<unsigned>(rec[at + 4]), static_cast<unsigned>(rec[at + 5]), static_cast<unsigned>(rec[at + 6]),
                        static_cast<unsigned>(rec[at + 7]));
          line += std::string("<d v=\"") + buf + "\"/>";
          at += 8;
        } else {
          ok = false;
        }
      }
      if (at != rec.size()) ok = false;
      line += "</r>";
      rout << line;
      ++n_rec;
    }

    rout << "</pivotCacheRecords>\n";
    rout.close();
    if (ok && n_rec > 0) {
      refresh = false;
      n_records = n_rec;
    }
    if (refresh) std::remove(recordsOut.c_str());
  }

  std::ofstream out(outPath);
  out << xls_pivot_cache_write(fields, n_records, src, slicer_id, refresh);
  return static_cast<int32_t>(fields.size());
}

// Returns the id of the cache of the table, -1 if the table is not imported
// [[Rcpp::export]]
int32_t pivottable_bin(std::string filePath, std::string outPath, bool debug) {
  std::vector<BinRecord> recs = pivot_records(filePath);
  if (recs.empty()) return -1;

  XlsPivotView v;
  int64_t cache_id = -1;
  int line_axis = -1;
  bool in_format = false;
  v.axes.resize(2);

  for (size_t i = 0; i < recs.size(); ++i) {
    const std::vector<uint8_t>& d = recs[i].data;
    if (debug) Rcpp::Rcout << "pivot table record " << recs[i].id << ": " << d.size() << std::endl;

    switch (recs[i].id) {
      case BrtBeginSXView: {
        cache_id = xls_pv_u32(d, 28);
        if (d.size() > 5) {
          v.row_grand = (d[5] & 0x20) != 0;
          v.col_grand = (d[5] & 0x40) != 0;
        }
        if (d.size() > 15) {
          v.outline = (d[8] & 0x02) != 0;
          v.version = d[14];
        }
        size_t pos = 32;
        v.name = pivot_string(d, pos);
        v.data_name = pivot_string(d, pos);
        break;
      }

      case BrtBeginSXLocation:
        v.row1 = static_cast<uint16_t>(xls_pv_u32(d, 0));
        v.row2 = static_cast<uint16_t>(xls_pv_u32(d, 4));
        v.col1 = static_cast<uint16_t>(xls_pv_u32(d, 8));
        v.col2 = static_cast<uint16_t>(xls_pv_u32(d, 12));
        v.head_row = static_cast<uint16_t>(xls_pv_u32(d, 16));
        v.data_row = static_cast<uint16_t>(xls_pv_u32(d, 20));
        v.data_col = static_cast<uint16_t>(xls_pv_u32(d, 24));
        break;

      case BrtBeginSXVD: {
        XlsPivotField f;
        f.axis = static_cast<uint16_t>(xls_pv_u32(d, 0) & 0x0F);
        f.default_subtotal = d.size() > 1 && (d[1] & 0x01) != 0;
        f.numfmt = xls_pv_u32(d, 4);
        f.multi = d.size() > 10 && (d[10] & 0x08) != 0;
        f.sort = d.size() > 9 ? static_cast<uint8_t>((d[9] & 0x30) == 0x30 ? 1 : ((d[9] & 0x30) == 0x10 ? 2 : 0)) : static_cast<uint8_t>(0);
        size_t pos = 16;
        f.name = pivot_string(d, pos);
        v.fields.push_back(f);
        break;
      }

      case BrtBeginSXVI: {
        if (v.fields.empty() || d.size() < 7) break;
        XlsPivotItem it;
        it.type = d[0];
        it.flags = static_cast<uint16_t>(xls_pv_u16(d, 1));
        uint32_t c = xls_pv_u32(d, 3);
        it.cache = c > 0xFFFEu ? static_cast<uint16_t>(0xFFFF) : static_cast<uint16_t>(c);
        if (d.size() > 7) {
          size_t pos = 7;
          it.name = pivot_string(d, pos);
        }
        v.fields.back().items.push_back(it);
        break;
      }

      case BrtBeginISXVDRws:
      case BrtBeginISXVDCols: {
        std::vector<uint16_t>& axis = v.axes[recs[i].id == BrtBeginISXVDRws ? 0 : 1];
        uint32_t n = xls_pv_u32(d, 0);
        for (uint32_t k = 0; k < n && 4 + 4 * static_cast<size_t>(k) + 4 <= d.size(); ++k) {
          uint32_t x = xls_pv_u32(d, 4 + 4 * static_cast<size_t>(k));
          axis.push_back(x == 0xFFFFFFFEu ? static_cast<uint16_t>(0xFFFE) : static_cast<uint16_t>(x));
        }
        break;
      }

      case BrtBeginSXLIRws:
      case BrtBeginSXLICols:
        line_axis = recs[i].id == BrtBeginSXLIRws ? 0 : 1;
        break;

      case BrtEndSXLIRws:
      case BrtEndSXLICols:
        line_axis = -1;
        break;

      case BrtBeginSXLI:
        if (line_axis >= 0 && d.size() >= 12) {
          XlsPivotLine l;
          l.reuse = static_cast<uint16_t>(xls_pv_u16(d, 0));
          l.type = d[2];
          l.data = xls_pv_u32(d, 8);
          v.lines[line_axis].push_back(l);
        }
        break;

      case BrtBeginISXVIs:
        if (line_axis >= 0 && !v.lines[line_axis].empty()) {
          for (size_t k = 0; k + 4 <= d.size(); k += 4) v.lines[line_axis].back().x.push_back(xls_pv_u32(d, k));
        }
        break;

      case BrtBeginSXFormat:
        if (d.size() >= 6) {
          XlsPivotFormat f;
          f.dxf = xls_pv_u32(d, 2);
          v.formats.push_back(f);
          in_format = true;
        }
        break;

      case BrtEndSXFormat:
        in_format = false;
        break;

      case BrtBeginPRule:
        if (in_format && !v.formats.empty() && d.size() >= 8) {
          XlsPivotFormat& f = v.formats.back();
          f.field = static_cast<int32_t>(xls_pv_u32(d, 0));
          f.type = d[4];
          f.flags = d[5];
          f.axis = static_cast<uint8_t>(xls_pv_u16(d, 6) & 0x0F);
          f.position = static_cast<uint8_t>((xls_pv_u16(d, 6) >> 4) & 0xFF);
          in_format = false;
        }
        break;

      case BrtTableStyleClient:
        if (d.size() >= 6) {
          v.style_flags = static_cast<uint16_t>(xls_pv_u16(d, 0));
          size_t pos = 2;
          std::string style = pivot_string(d, pos);
          if (!style.empty()) v.style = style;
        }
        break;

      case BrtBeginSXPI: {
        XlsPivotPage p;
        p.field = static_cast<uint16_t>(xls_pv_u32(d, 0));
        uint32_t item = xls_pv_u32(d, 4);
        p.item = item <= 0x00100000u ? static_cast<uint16_t>(item) : static_cast<uint16_t>(0x7FFD);
        v.pages.push_back(p);
        break;
      }

      case BrtBeginSXDI: {
        XlsPivotDataField df;
        df.field = static_cast<uint16_t>(xls_pv_u32(d, 0));
        df.function = static_cast<uint16_t>(xls_pv_u32(d, 4));
        df.display = static_cast<uint16_t>(xls_pv_u32(d, 8));
        df.base_field = static_cast<uint16_t>(xls_pv_u32(d, 12));
        df.base_item = static_cast<uint16_t>(xls_pv_u32(d, 16));
        df.numfmt = static_cast<uint16_t>(xls_pv_u32(d, 20));
        if (d.size() > 25) {
          size_t pos = 25;
          df.name = pivot_string(d, pos);
        }
        v.data.push_back(df);
        break;
      }

      default:
        break;
    }
  }

  if (cache_id < 0 || v.fields.empty()) return -1;

  std::ofstream out(outPath);
  out << xls_pivot_view_xml(v, static_cast<int32_t>(cache_id));
  return static_cast<int32_t>(cache_id);
}

// [[Rcpp::export]]
int32_t slicer_bin(std::string filePath, std::string outPath, bool debug) {
  std::vector<BinRecord> recs = pivot_records(filePath);
  if (recs.empty()) return -1;

  std::string slicers;
  int32_t n = 0;
  for (size_t i = 0; i < recs.size(); ++i) {
    const std::vector<uint8_t>& d = recs[i].data;
    if (debug) Rcpp::Rcout << "slicer record " << recs[i].id << ": " << d.size() << std::endl;
    if (recs[i].id != BrtBeginSlicer || d.size() < 17) continue;

    uint8_t flags = d[0];
    uint32_t start = xls_pv_u32(d, 1), columns = xls_pv_u32(d, 5), level = xls_pv_u32(d, 9), height = xls_pv_u32(d, 13);
    size_t pos = 17;
    std::string name = pivot_string(d, pos);
    std::string cache = pivot_string(d, pos);
    std::string caption, style;
    if ((flags & 0x02) != 0) caption = pivot_string(d, pos);
    if ((flags & 0x04) != 0) style = pivot_string(d, pos);

    slicers += "<slicer name=\"" + escape_xml(name) + "\" cache=\"" + escape_xml(cache) + "\"";
    if ((flags & 0x02) != 0) slicers += " caption=\"" + escape_xml(caption) + "\"";
    if ((flags & 0x01) == 0) slicers += " showCaption=\"0\"";
    if (start != 0) slicers += " startItem=\"" + std::to_string(start) + "\"";
    if (columns != 1) slicers += " columnCount=\"" + std::to_string(columns) + "\"";
    if (level != 0) slicers += " level=\"" + std::to_string(level) + "\"";
    if ((flags & 0x04) != 0) slicers += " style=\"" + escape_xml(style) + "\"";
    if ((flags & 0x08) != 0) slicers += " lockedPosition=\"1\"";
    slicers += " rowHeight=\"" + std::to_string(height) + "\"/>";
    ++n;
  }
  if (n == 0) return -1;

  std::ofstream out(outPath);
  out << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      << "<slicers xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" "
         "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x\" "
         "xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      << slicers << "</slicers>\n";
  return n;
}

// [[Rcpp::export]]
int32_t slicercache_bin(std::string filePath, std::string outPath, bool debug) {
  std::vector<BinRecord> recs = pivot_records(filePath);
  if (recs.empty()) return -1;

  std::string name, source, tables, items;
  uint32_t pivot_cache_id = 0;
  size_t n_items = 0;
  bool have_def = false;

  for (size_t i = 0; i < recs.size(); ++i) {
    const std::vector<uint8_t>& d = recs[i].data;
    if (debug) Rcpp::Rcout << "slicer cache record " << recs[i].id << ": " << d.size() << std::endl;

    switch (recs[i].id) {
      case BrtBeginSlicerCacheDef: {
        size_t pos = 0;
        name = pivot_string(d, pos);
        source = pivot_string(d, pos);
        have_def = true;
        break;
      }

      case BrtSlicerCachePivotTables: {
        size_t pos = 4;
        for (uint32_t k = 0, n = xls_pv_u32(d, 0); k < n && pos + 8 <= d.size(); ++k) {
          uint32_t tab = xls_pv_u32(d, pos);
          pos += 4;
          std::string pt = pivot_string(d, pos);
          tables += "<pivotTable tabId=\"" + std::to_string(tab) + "\" name=\"" + escape_xml(pt) + "\"/>";
        }
        break;
      }

      case BrtBeginSlicerCacheNative:
        pivot_cache_id = xls_pv_u32(d, 4);
        break;

      case BrtSlicerCacheNativeItem: {
        for (uint32_t k = 0, n = xls_pv_u32(d, 0); k < n && 4 + 5 * static_cast<size_t>(k) + 5 <= d.size(); ++k) {
          size_t at = 4 + 5 * static_cast<size_t>(k);
          items += "<i x=\"" + std::to_string(xls_pv_u32(d, at)) + "\"" + ((d[at + 4] & 0x01) != 0 ? " s=\"1\"" : "") + "/>";
          ++n_items;
        }
        break;
      }

      default:
        break;
    }
  }
  if (!have_def || pivot_cache_id == 0) return -1;

  std::ofstream out(outPath);
  out << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      << "<slicerCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" "
         "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x\" "
         "xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\""
      << escape_xml(name) << "\" sourceName=\"" << escape_xml(source) << "\">";
  if (!tables.empty()) out << "<pivotTables>" << tables << "</pivotTables>";
  out << "<data><tabular pivotCacheId=\"" << pivot_cache_id << "\"><items count=\"" << n_items << "\">" << items
      << "</items></tabular></data></slicerCacheDefinition>\n";
  return 1;
}
