#include <Rcpp/Lightest>

#include "xlsb.h"

std::string RichStr(std::istream& sas) {

  uint8_t A, B;

  // Rcpp::Rcout << sas.tellg() << std::endl;
  A = readbin(A, sas, 0);
  B = readbin(B, sas, 0);

  // sas.seekg(6, sas.cur); ???

  uint32_t len;
  len = readbin(len, sas, 0);

  std::string str(len, '\0');

  // Rcpp::Rcout << len << std::endl;

  str = read_xlwidestring(str, sas);

  return(str);

}

std::vector<int> UncheckedRfX(std::istream& sas) {

  std::vector<int> out;
  int32_t rwFirst, rwLast, colFirst, colLast;

  out.push_back(readbin(rwFirst, sas, 0));
  out.push_back(readbin(rwLast, sas, 0));
  out.push_back(readbin(colFirst, sas, 0));
  out.push_back(readbin(colLast, sas, 0));

  return(out);
}

std::vector<int> Cell(std::istream& sas) {

  std::vector<int> out;
  int32_t rwFirst, rwLast, colFirst, colLast;

  return(out);
}


// [[Rcpp::export]]
int sst(std::string filePath, std::string outPath, bool debug) {

  std::ofstream out(outPath);
  std::ifstream bin(filePath, std::ios::in | std::ios::binary | std::ios::ate);

  auto sas_size = bin.tellg();
  if (bin) {
    bin.seekg(0, std::ios_base::beg);

    while(!bin.eof()) {
      uint8_t x, unk;

      if (debug) Rcpp::Rcout << "." << std::endl;
      x = readbin(x, bin, 0);
      if (x == BrtBeginSst)  {
        unk = readbin(unk, bin, 0);
        unk = readbin(unk, bin, 0);

        uint32_t count, uniqueCount;
        count = readbin(count, bin, 0);
        uniqueCount = readbin(uniqueCount, bin, 0);
        out << "<sst " <<
          "count=\"" << count <<
          "\" uniqueCount=\"" << uniqueCount <<
          "\">" << std::endl;
      }

      if (x == BrtSSTItem) {
        out << "<si><t>" << RichStr(bin) <<
                    "</t></si>" << std::endl;
      }

      if (x == BrtEndSst) {
        out << "</sst>" << std::endl;
        break;
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

  auto sas_size = bin.tellg();
  if (bin) {
    bin.seekg(0, std::ios_base::beg);

    while(!bin.eof()) {
      uint8_t x, unk;

      if (debug) Rcpp::Rcout << "." << std::endl;
      x = readbin(x, bin, 0);
      if (x == BrtBeginBundleShs)  {
        // unk = readbin(unk, bin, 0);
        // unk = readbin(unk, bin, 0);
        // uint32_t count, uniqueCount;
        // count = readbin(count, bin, 0);
        // uniqueCount = readbin(uniqueCount, bin, 0);
        out << "<sheets>" << std::endl;
      }

      if (x == BrtBundleSh) {
        uint32_t hsState, iTabID; //  strRelID ???
        hsState = readbin(hsState, bin, 0);
        iTabID = readbin(iTabID, bin, 0);
        // Rcpp::Rcout << hsState << " : " << iTabID << std::endl;
        out << "<sheet r:id=\"" << RichStr(bin) << "\" name=\"";

        uint32_t len;
        len = readbin(len, bin, 0);
        std::string sheet_name(len, '\0');
        out << read_xlwidestring(sheet_name, bin)  <<
          "\"/>" << std::endl;
      }

      if (x == BrtEndBundleShs) {
        out << "</sheets>" << std::endl;
        break;
      }
    }

    out.close();
    bin.close();
    return 1;
  } else {
    return -1;
  };

}

int32_t RECORD(uint8_t highByte, uint8_t lowByte) {

  if (highByte & 0x80) {
    // If the high bit is 1, then it's a two-byte record type
    int32_t recordType = ((lowByte & 0x7F) << 7) | (highByte & 0x7F);
    if (recordType >= 128 && recordType < 16384) {
      return recordType;
    }
  } else {
    // If the high bit is not 1, then it's a one-byte record type
    int32_t recordType = highByte;
    if (recordType >= 0 && recordType < 128) {
      return recordType;
    }
  }

  // Return -1 if the record type is invalid
  return -1;
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

  auto sas_size = bin.tellg();
  if (bin) {
    bin.seekg(0, std::ios_base::beg);

    bool in_worksheet = false;
    bool in_sheet_data = false;

    auto itr = 0;
    while(!bin.eof()) {

      uint8_t unk, high, low;
      uint16_t tmp;
      int32_t x = 0;

      if (debug) Rcpp::Rcout << "." << std::endl;
      tmp = readbin(tmp, bin, 0);
      // Rcpp::Rcout << tmp << std::endl;
      uint8_t hi = tmp & 0xff;
      uint8_t lo = (tmp >> 8) & 0xff;

      if (debug) Rprintf("high/low = %d : %d\n", hi, lo);

      x = RECORD(hi, lo);
      if (debug) Rcpp::Rcout << "x: " << x << std::endl;


      if (x == BrtBeginSheet) {
        out << "<worksheet>" << std::endl;
        in_worksheet = true;

        uint8_t A, B;

        B = readbin(B, bin, 0); // len?

        if (debug) {
          printf("%d : %d\n", A, B);
        }
      }

      if (in_worksheet && x == BrtWsProp) {
        Rcpp::Rcout << "WsProp: " << bin.tellg() << std::endl;

        uint8_t A, B;

        B = readbin(B, bin, 0); // len?

        if (debug) {
          printf("%d : %d\n", A, B);
        }

        uint8_t fShowAutoBreaks, fPublish, fDialog, fApplyStyles, fRowSumsBelow,
                fColSumsBelow, fColSumsRight, fFitToPage, reserved2,
                fShowOutlineSymbols, reserved3, fSyncHoriz, fSyncVert,
                fAltExprEval, fAltFormulaEntry, fFilterMode, fCondFmtCalc;
        uint16_t rserved1;
        uint32_t rwSync, colSync;
        int64_t brtcolorTab;

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

        bin.seekg(B, bin.cur);
      }

      if (in_worksheet && x == BrtWsDim) {
        Rcpp::Rcout << "WsDim: " << bin.tellg() << std::endl;

        uint8_t A, B;

        B = readbin(B, bin, 0); // len?

        if (debug) {
          printf("%d : %d\n", A, B);
        }

        // 16 bit
        std::vector<int> dims;
        // 0 index vectors
        // first row, last row, first col, last col
        dims = UncheckedRfX(bin);
        Rf_PrintValue(Rcpp::wrap(dims));
      }

      if (in_worksheet && x == BrtBeginWsViews) {
        out << "<sheetViews>: " << bin.tellg() << std::endl;
        uint8_t A, B;

        B = readbin(B, bin, 0); // len?

        if (debug) {
          printf("%d : %d\n", A, B);
        }
      }

      if (in_worksheet && x == BrtBeginWsView) {
        out << "<sheetView>: " << bin.tellg() << std::endl;
      }

      if (in_worksheet && x == BrtEndWsView) {
        out << "</sheetView>: " << bin.tellg() << std::endl;
      }

      if (in_worksheet && x == BrtEndWsViews) {
        out << "</sheetViews>: " << bin.tellg() << std::endl;
      }

      if (!in_sheet_data && x == BrtBeginSheetData)  {
        out << "<sheetData>" << bin.tellg() << std::endl;
        in_sheet_data = true;

        // uint8_t A, B;
        //
        // A = readbin(A, bin, 0); // 1?
        // B = readbin(B, bin, 0); // len?
        //
        // if (debug) {
        //   printf("%d : %d\n", A, B);
        // }

      } // 2.2.1 Cell Table

      if (in_sheet_data && x == BrtRowHdr)  {
        Rcpp::Rcout << "BrtRowHdr: " << bin.tellg() << std::endl;

        uint8_t A, B;

        // A = readbin(A, bin, 0); // 1?
        B = readbin(B, bin, 0); // len?

        if (debug) {
          printf("%d : %d\n", A, B);
        }

          // 17?
          bin.seekg(15, bin.cur); // ??? Not sure what is going on there some pre <r> lines?

          out << "<r> " << bin.tellg() << std::endl;
          uint8_t bits1, bits2, bits3, fExtraAsc, fExtraDsc, fCollapsed,
                  fDyZero, fUnsynced, fGhostDirty, fReserved, fPhShow;
          uint16_t miyRw;

          // uint24_t;
          int32_t rw;
          uint32_t ixfe, ccolspan, unk32, colMic, colLast;

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

          Rcpp::Rcout << ccolspan << std::endl;

          Rcpp::Rcout << rw << " : " << ixfe << " : " << miyRw << " : " << (int32_t)fExtraAsc << " : " <<
            (int32_t)fExtraDsc << " : " << unk32 << " : " << (int32_t)fCollapsed << " : " << (int32_t)fDyZero << " : " <<
              (int32_t)fUnsynced << " : " << (int32_t)fGhostDirty << " : " << (int32_t)fReserved << " : " << (int32_t)fPhShow << " : " <<
                ccolspan << "; " << bin.tellg() << std::endl;

        // }
      }

      if (x == row_description) {
        bin.seekg(15, bin.cur);
      }

      if (x == BrtCellIsst)  {
        Rcpp::Rcout << "BrtCellIsst: " << bin.tellg() << std::endl;

        // bin.seekg(8, bin.cur);

        int32_t val;
        val = readbin(val, bin, 0);
        out << val << std::endl;
        val = readbin(val, bin, 0);
        out << val << std::endl;
        val = readbin(val, bin, 0);
        out << val << std::endl;
        // Rcpp::stop("stop");
      }


      if (in_sheet_data && x == BrtEndSheetData) {
        out << "</sheetData>" << std::endl;
        in_sheet_data = false;
      }

      if (x == BrtEndSheet) {
        out << "</worksheet>" << std::endl;
        in_worksheet = false;
        break;
      }

      itr++;
      // if (itr == 2) Rcpp::stop("stop");
    }

    out.close();
    bin.close();
    return 1;
  } else {
    return -1;
  };

}
