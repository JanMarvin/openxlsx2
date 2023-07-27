// #include <Rcpp/Lightest>

#include "openxlsx2.h"
#include "xlsb_defines.h"
#include "xlsb_funs.h"

// [[Rcpp::export]]
int styles(std::string filePath, std::string outPath, bool debug) {

  std::ofstream out(outPath);
  std::ifstream bin(filePath, std::ios::in | std::ios::binary | std::ios::ate);

  // auto sas_size = bin.tellg();
  if (bin) {
    bin.seekg(0, std::ios_base::beg);
    bool end_of_style_sheet = false;

    while(!end_of_style_sheet) {

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

        out << "numFmtId=\"" << ifmt << "\" formatCode=\"" << stFmtCode;
        out << " />" << std::endl;
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

        out << "<sz val=\"" << dyHeight << "\" />" << std::endl;

        // if (color[0] == 0x01) { // wrong but right?
        //   Rcpp::Rcout << "<color theme=\"" << color[1] << "\" />" << std::endl;
        // }
        if (color[0] == 0x03) {
          out << "<color theme=\"" << color[1] << "\" />" << std::endl;
        }

        out << "<name val=\"" << name << "\" />" << std::endl;
        if (bFamily > 0)
          out << "<family val=\"" << (uint16_t)bFamily << "\" />" << std::endl;

        if (bFontScheme == 1) {
          out << "<scheme val=\"major\" />" << std::endl;
        }
        if (bFontScheme == 2) {
          out << "<scheme val=\"minor\" />" << std::endl;
        }

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
        Rcpp::Rcout << "<fills>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtFill:
      {
        Rcpp::Rcout << "<fill>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtEndFills:
      {
        Rcpp::Rcout << "</fills>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

        // borders
      case BrtBeginBorders:
      {
        Rcpp::Rcout << "<borders>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtBorder:
      {
        Rcpp::Rcout << "<borders>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtEndBorders:
      {
        Rcpp::Rcout << "</borders>" << std::endl;
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

        brtxf = readbin(brtxf, bin, 0);
        XFFields *fields = (XFFields *)&brtxf;

        out << " numFmtId=\"" << iFmt <<"\"";
          out << " fontId=\"" << iFont <<"\"";
          out << " fillId=\"" << iFill <<"\"";
          out << " borderId=\"" << ixBorder <<"\"";

          if (iFmt > 0) {
            out << " applyNumberFormat=\"1\"";
          }
          if (iFont > 0) {
            out << " applyFont=\"1\"";
          }
          if (ixBorder > 0) {
            out << " applyBorder=\"1\"";
          }
          if (iFill > 0) {
            out << " applyFill=\"1\"";
          }
          out << " />" << std::endl;
          // bin.seekg(size, bin.cur);
          break;
      }

      case BrtEndCellXFs:
      {
        out << "</cellXfs>" << std::endl;
        bin.seekg(size, bin.cur);
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
        if (debug) Rcpp::Rcout << val << std::endl;
        out << "<si><t>" << val <<
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

    while(!end_of_workbook) {
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
          ProductVersion(bin);
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
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtBookView:
      {
        if (debug) Rcpp::Rcout << "<workbookView>" << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtEndBookViews:
      {
        if (debug) Rcpp::Rcout << "</workbookViews>" << std::endl;
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

        if (debug) Rcpp::Rcout << hsState << ": " << iTabID  << ": " << rid << std::endl;

        std::string val = XLWideString(bin);

        out << "<sheet r:id=\"" << rid << "\" sheetId=\""<< iTabID<< "\" name=\"" << val << "\"/>" << std::endl;
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

        if (debug)
          Rprintf(
            "%d, %d, %d, %d, %d, %d, %d, %d , %d, %d, %d\n",
            fields->fHidden,
            fields->fFunc,
            fields->fOB,
            fields->fProc,
            fields->fCalcExp,
            fields->fBuiltin,
            fields->fgrp,
            fields->fPublished,
            fields->fWorkbookParam,
            fields->fFutureFunction,
            fields->reserved
          );

        chKey = readbin(chKey, bin, 0);
        // ascii key (0 if fFunc = 1 or fProc = 0 else >= 0x20)

        itab = readbin(itab, bin, 0);

        // XLNameWideString: XLWideString <= 255 characters
        std::string name = XLWideString(bin);

        CellParsedFormula(bin, debug);

        std::string comment = XLNullableWideString(bin);

        if (debug) Rcpp::Rcout << name << comment << std::endl;

        if (fields->fProc) {
          // must be NULL
          std::string unusedstring1 = XLNullableWideString(bin);
          // must be < 32768 characters
          std::string description = XLNullableWideString(bin);
          std::string helpTopic = XLNullableWideString(bin);
          // must be NULL
          std::string unusedstring2 = XLNullableWideString(bin);
        }

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
        break;
      }

      case BrtSupTabs:
      {
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
        uint32_t cXti = 0;

        cXti = readbin(cXti, bin, 0);
        if (cXti > 65535) Rcpp::stop("cXti to large");

        std::vector<uint32_t> rgXti(cXti);
        for (uint32_t i = 0; i < cXti; ++i) {
          Xti(bin);
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

        ProductVersion(bin);
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

      if (debug) Rcpp::Rcout << x << ": " << size << ": " << bin.tellg() << std::endl;
    }

    out.close();
    bin.close();
    return 1;
  } else {
    return -1;
  };

}



// [[Rcpp::export]]
int worksheet(std::string filePath, std::string outPath, bool debug) {

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

    bool in_worksheet = false;
    bool in_sheet_data = false;

    bool first_row = true;

    bool end_of_worksheet = false;

    uint64_t row = 0;

    // auto itr = 0;
    while(!end_of_worksheet) {

      // uint8_t unk = 0, high = 0, low = 0;
      // uint16_t tmp = 0;
      int32_t x = 0, size = 0;

      if (debug) Rcpp::Rcout << "." << std::endl;

      RECORD(x, size, bin);

      switch(x) {

      case BrtBeginSheet:
      {
        out << "<worksheet>" << std::endl;
        in_worksheet = true;

        // uint8_t A, B;
        //
        // B = readbin(B, bin, 0); // len?

        // if (debug) {
        //   printf("%d : %d\n", A, B);
        // }

        if (debug) Rcpp::Rcout << "Begin of <worksheet>: " << bin.tellg() << std::endl;
        break;
      }

      case BrtWsProp:
      {
        if (debug) Rcpp::Rcout << "WsProp: " << bin.tellg() << std::endl;

        // uint8_t A, B;
        //
        // B = readbin(B, bin, 0); // len?
        //
        // if (debug) {
        //   printf("%d : %d\n", A, B);
        // }

        // uint8_t fShowAutoBreaks = 0, fPublish = 0, fDialog = 0, fApplyStyles = 0, fRowSumsBelow = 0,
        // fColSumsBelow = 0, fColSumsRight = 0, fFitToPage = 0, reserved2 = 0,
        // fShowOutlineSymbols = 0, reserved3 = 0, fSyncHoriz = 0, fSyncVert = 0,
        // fAltExprEval = 0, fAltFormulaEntry = 0, fFilterMode = 0, fCondFmtCalc = 0;
        // uint16_t rserved1 = 0;
        // uint32_t rwSync = 0, colSync = 0;
        // int64_t brtcolorTab = 0;

        // all these are bit, not byte ...
        // fShowAutoBreaks = readbin(fShowAutoBreaks, bin, 0);
        // rserved1 = readbin(rserved1, bin, 0);
        // fPublish = readbin(fPublish, bin, 0);
        // fDialog = readbin(fDialog, bin, 0);
        // fApplyStyles = readbin(fApplyStyles, bin, 0);
        // fRowSumsBelow = readbin(fRowSumsBelow, bin, 0);
        // fColSumsRight = readbin(fColSumsRight, bin, 0);
        // fFitToPage = readbin(fFitToPage, bin, 0);
        // reserved2 = readbin(reserved2, bin, 0);
        // fShowOutlineSymbols = readbin(fShowOutlineSymbols, bin, 0);
        // reserved3 = readbin(reserved3, bin, 0);
        // fSyncHoriz = readbin(fSyncHoriz, bin, 0);
        // fSyncVert = readbin(fSyncVert, bin, 0);
        // fAltExprEval = readbin(fAltExprEval, bin, 0);
        // fAltFormulaEntry = readbin(fAltFormulaEntry, bin, 0);
        // fFilterMode = readbin(fFilterMode, bin, 0);
        // fCondFmtCalc = readbin(fCondFmtCalc, bin, 0);
        // bin.seekg(6, bin.cur);
        // brtcolorTab = readbin(brtcolorTab, bin, 0);
        // rwSync = readbin(rwSync, bin, 0);
        // colSync = readbin(colSync, bin, 0);
        // strName a code name ...

        bin.seekg(size, bin.cur);
        break;
      }

      case BrtWsDim:
      {
        if (debug) Rcpp::Rcout << "WsDim: " << bin.tellg() << std::endl;

        // uint8_t A, B;
        //
        // B = readbin(B, bin, 0); // len?
        //
        // if (debug) {
        //   printf("%d : %d\n", A, B);
        // }

        // 16 bit
        std::vector<int> dims;
        // 0 index vectors
        // first row, last row, first col, last col
        dims = UncheckedRfX(bin);
        Rf_PrintValue(Rcpp::wrap(dims));

        out << "<dimension ref=\"" <<
          int_to_col(dims[2] + 1) << dims[0] + 1 <<
            ":" <<
              int_to_col(dims[3] + 1) << dims[1] + 1 <<
                "\"/>" << std::endl;

        break;
      }

      case BrtBeginWsViews:
      {
        if (debug) Rcpp::Rcout << "<sheetViews>: " << bin.tellg() << std::endl;
        // uint8_t A, B;
        //
        // B = readbin(B, bin, 0); // len?
        //
        // if (debug) {
        //   printf("%d : %d\n", A, B);
        // }
        break;
      }

        // whatever this is
      case BrtSel:
      {
        if (debug) Rcpp::Rcout << "BrtSel: " << bin.tellg() << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtBeginWsView:
      {
        if (debug) Rcpp::Rcout << "<sheetView>: " << bin.tellg() << std::endl;
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtEndWsView:
      {
        if (debug) Rcpp::Rcout << "</sheetView>: " << bin.tellg() << std::endl;
        break;
      }

      case BrtEndWsViews:
      {
        if (debug) Rcpp::Rcout << "</sheetViews>: " << bin.tellg() << std::endl;
        break;
      }

      case BrtBeginSheetData:
      {
        if (debug) Rcpp::Rcout << "<sheetData>" << bin.tellg() << std::endl;
        out << "<sheetData>" << std::endl; //  << bin.tellg()
        in_sheet_data = true;
        break;
      }

        // case BrtRowHdr:
        // {
        //   if (debug) Rcpp::Rcout << "BrtRowHdr: " << bin.tellg() << std::endl;
        //   if (!first_row) {
        //     out << "</row>" <<std::endl;
        //   }
        //   first_row = false;
        //   break;
        // }

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
        bin.seekg(size, bin.cur);
        break;
      }

      case BrtRowHdr:
      {

        // close open rows
        if (!first_row) {
        out << "</row>" <<std::endl;
      } else {
        first_row = false;
      }



      out << "<row r=\""; //  << bin.tellg()
      uint8_t bits1 = 0, bits2 = 0, bits3 = 0, fExtraAsc = 0, fExtraDsc = 0, fCollapsed = 0,
        fDyZero = 0, fUnsynced = 0, fGhostDirty = 0, fReserved = 0, fPhShow = 0;
      uint16_t miyRw = 0;

      // uint24_t;
      int32_t rw = 0;
      uint32_t ixfe = 0, ccolspan = 0, unk32 = 0, colMic = 0, colLast = 0;

      rw = readbin(rw, bin, 0);
      ixfe = readbin(ixfe, bin, 0);
      miyRw = readbin(miyRw, bin, 0);

      bits1 = readbin(bits1, bin, 0);
      // fExtraAsc = 1
      // fExtraDsc = 1
      // reserved1 = 6

      bits2 = readbin(bits2, bin, 0);
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
      // rgBrtColspan
      // not sure if these are alway around. maybe ccolspan is a counter
      colMic = readbin(colMic, bin, 0);
      colLast = readbin(colLast, bin, 0);

      if (debug) Rcpp::Rcout << ccolspan << std::endl;

      out << rw + 1 << "\">" << std::endl;

      row = rw;

      if (debug)
        Rcpp::Rcout << (rw) << " : " << ixfe << " : " << miyRw << " : " << (int32_t)fExtraAsc << " : " <<
          (int32_t)fExtraDsc << " : " << unk32 << " : " << (int32_t)fCollapsed << " : " << (int32_t)fDyZero << " : " <<
            (int32_t)fUnsynced << " : " << (int32_t)fGhostDirty << " : " << (int32_t)fReserved << " : " << (int32_t)fPhShow << " : " <<
              ccolspan << "; " << bin.tellg() << std::endl;

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
        out << "<v>" << RkNumber(val3) << "</v>" << std::endl;
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

#include <iomanip>

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
        if (debug) Rf_PrintValue(Rcpp::wrap(blank));

        out << "<c r=\"" << int_to_col(blank[0] + 1) << row + 1 << "\"" << cell_style(blank[1]) << "/>" << std::endl;

        break;
      }

      case BrtCellError:
      { // t="e" & <v>#NUM!</v>
        Rcpp::Rcout << "BrtCellError: " << bin.tellg() << std::endl;

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
        Rcpp::Rcout << "BrtFmlaBool: " << bin.tellg() << std::endl;
        // bin.seekg(size, bin.cur);

        std::vector<int> cell;
        cell = Cell(bin);
        if (debug) Rf_PrintValue(Rcpp::wrap(cell));

        bool val = 0;
        val = readbin(val, bin, 0);
        if (debug) Rcpp::Rcout << val << std::endl;

        uint16_t grbitFlags = 0;
        grbitFlags = readbin(grbitFlags, bin, 0);

        std::string fml = CellParsedFormula(bin, debug);

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
        if (debug) Rf_PrintValue(Rcpp::wrap(cell));

        std::string fErr;
        fErr = BErr(bin);

        uint16_t frbitfmla = 0;
        // [0] == 0
        // [1] == fAlwaysCalc
        // [2-15] == unused
        frbitfmla = readbin(frbitfmla, bin, 0);

        // int32_t len = size - 4 * 32 - 2 * 8;
        // std::string fml(len, '\0');

        std::string fml = CellParsedFormula(bin, debug);

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
        if (debug) Rf_PrintValue(Rcpp::wrap(cell));

        double xnum = Xnum(bin);
        if (debug) Rcpp::Rcout << xnum << std::endl;

        uint16_t grbitFlags = 0;
        grbitFlags = readbin(grbitFlags, bin, 0);

        std::string fml = CellParsedFormula(bin, debug);

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
        if (debug) Rf_PrintValue(Rcpp::wrap(cell));

        std::string val = XLWideString(bin);
        if (debug) Rcpp::Rcout << val << std::endl;

        uint16_t grbitFlags = 0;
        grbitFlags = readbin(grbitFlags, bin, 0);

        std::string fml = CellParsedFormula(bin, debug);

        out << "<c r=\"" << int_to_col(cell[0] + 1) << row + 1 << "\"" << cell_style(cell[1]) << " t=\"str\">" << std::endl;
        out << "<f>" << fml << "</f>" << std::endl;
        out << "<v>" << val << "</v>" << std::endl;
        out << "</c>" << std::endl;

        break;
      }

      case BrtEndSheetData:
      {
        if (debug) Rcpp::Rcout << "</sheetData>" << bin.tellg() << std::endl;

        if (!first_row) {
          out << "</row>" << std::endl;
        }
        out << "</sheetData>" << std::endl;
        in_sheet_data = false;

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

      case BrtEndSheet:
      {
        if (debug)  Rcpp::Rcout << "</worksheet>" << bin.tellg() << std::endl;
        out << "</worksheet>" << std::endl;
        in_worksheet = false;
        end_of_worksheet = true;
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
