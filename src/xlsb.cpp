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

  std::vector<int> out(3);

  int32_t uint;
  out[0] = readbin(uint, sas, 0);
  uint = readbin(uint, sas, 0);

  out[1] = uint >> 8;   // iStyleRef
  out[2] = uint & 0xFF; // unused

  return(out);
}

static double RkNumber(int32_t val) {

  double out;
  if (val & 0x02) { // integer
    int32_t tmp = (int32_t)val >> 2;
    out = (double)tmp;
  } else { // double
    uint64_t tmp = val & 0xfffffffc;
    tmp <<= 32;
    memcpy(&out, &tmp, sizeof(uint64_t));
  }

  if (val & 0x01) {
    out /= 100.0;
  }
  return out;
}

std::string BErr(std::istream& sas) {
  uint8_t error = 0;
  error = readbin(error, sas, 0);

  if (error == 0x00) return "#NULL!";
  if (error == 0x07) return "#DIV/0";
  if (error == 0x0F) return "#VALUE!";
  if (error == 0x17) return "#REF!";
  if (error == 0x1D) return "#NAME?";
  if (error == 0x24) return "#NUM!";
  if (error == 0x2A) return "#N/A";
  if (error == 0x2B) return "#GETTING_DATA";

  return "unknown_ERROR";
}

void CellParsedFormula(std::istream& sas) {
  uint32_t  cce, cb;

  Rcpp::Rcout << sas.tellg() << std::endl;

  cce = readbin(cce, sas, 0);
  Rcpp::Rcout << "cce: " << cce << std::endl;
  sas.seekg(cce, sas.cur);

  // int32_t val = 0;
  // val = readbin((int8_t)val, sas, 0);
  // Rcpp::Rcout << val << std::endl;
  //
  // // this will be a nightmare ...
  // if (val == 0x44) {
  //
  //   Rcpp::Rcout << "PtgRef" << ": " << sas.tellg() << std::endl;
  //   uint8_t ptg = 0;
  //   ptg = readbin(ptg, sas, 0);
  //
  //   int8_t A, B;
  //
  //   A = (ptg >> 5) & 0b11;
  //   B = (ptg >> 7) & 1;
  //
  //   uint32_t row = 0;
  //   row = readbin(row, sas, 0);
  //
  //   uint16_t col = 0;
  //   col = readbin(col, sas, 0);
  //
  //   Rcpp::Rcout << (uint8_t)A << " : " << (uint8_t)B << " : " << row << " : " << col << std::endl;
  // }
  //
  // val = readbin((int8_t)val, sas, 0);
  // Rcpp::Rcout << val << std::endl;


  // if (cce != 0) {
  //   Rcpp::stop("cce != 0");
  // }

  cb = readbin(cb, sas, 0);
  Rcpp::Rcout << "cb: " << cb << std::endl;
  // if (cb != 0) {
  //   Rcpp::stop("cb != 0");
  // }
  sas.seekg(cb, sas.cur);
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


int32_t RECORD_ID(std::istream& sas) {
  uint8_t var1 = 0, var2 = 0;
  var1 = readbin(var1, sas, 0);

  if (var1 & 0x80) {

    var2 = readbin(var2, sas, 0);

    // If the high bit is 1, then it's a two-byte record type
    int32_t recordType = ((var2 & 0x7F) << 7) | (var1 & 0x7F);
    if (recordType >= 128 && recordType < 16384) {
      return recordType;
    }
  } else {
    // If the high bit is not 1, then it's a one-byte record type
    int32_t recordType = var1;
    if (recordType >= 0 && recordType < 128) {
      return recordType;
    }
  }
  return -1;
}

int32_t RECORD_SIZE(std::istream& sas) {
  int8_t sar1 = 0, sar2 = 0, sar3 = 0, sar4 = 0;

  sar1 = readbin(sar1, sas, 0);
  if (sar1 & 0x80) sar2 = readbin(sar2, sas, 0);
  if (sar2 & 0x80) sar3 = readbin(sar3, sas, 0);
  if (sar3 & 0x80) sar4 = readbin(sar4, sas, 0);

  // Rcpp::Rcout << sar1 << ": " << sar2 << ": " << sar3 << ": " << sar4 << std::endl;


  if (sar2 != 0 & sar3 != 0 && sar4 != 0) {
    int32_t recordType = ((sar4 & 0x7F) << 7) | ((sar3 & 0x7F) << 7) | ((sar2 & 0x7F) << 7) | (sar1 & 0x7F);
    return recordType;
  }

  if (sar2 != 0 & sar3 != 0 && sar4 == 0) {
    int32_t recordType = ((sar3 & 0x7F) << 7) | ((sar2 & 0x7F) << 7) | (sar1 & 0x7F);
    return recordType;
  }

  if (sar2 != 0 & sar3 == 0 && sar4 == 0) {
    int32_t recordType = ((sar2 & 0x7F) << 7) | (sar1 & 0x7F);
    return recordType;
  }

  if (sar2 == 0 & sar3 == 0 && sar4 == 0) {
    // Rcpp::Rcout << "yes" << std::endl;
    int32_t recordType = sar1;
    return recordType;
  }

  // Return -1 if the record type is invalid
  return -1;
}

int32_t RECORD(int &rid, int &rsize, std::istream& sas) {

  /* Record ID ---------------------------------------------------------------*/
  rid = RECORD_ID(sas);

  /* Record Size -------------------------------------------------------------*/
  rsize = RECORD_SIZE(sas);

  return 0;
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

    bool first_row = true;

    auto itr = 0;
    while(!bin.eof()) {

      uint8_t unk, high, low;
      uint16_t tmp;
      int32_t x = 0, size = 0;

      if (debug) Rcpp::Rcout << "." << std::endl;

      RECORD(x, size, bin);

      if (debug) {
        Rcpp::Rcout << "x: " << x << std::endl;
        Rcpp::Rcout << "size: " << size << std::endl;
        Rcpp::Rcout << "@: " << bin.tellg() << std::endl;
      }

      if (x == BrtBeginSheet) {
        out << "<worksheet>" << std::endl;
        in_worksheet = true;

        // uint8_t A, B;
        //
        // B = readbin(B, bin, 0); // len?

        // if (debug) {
        //   printf("%d : %d\n", A, B);
        // }

        Rcpp::Rcout << "End of <worksheet>: " << bin.tellg() << std::endl;
      }

      if (in_worksheet && x == BrtWsProp) {
        Rcpp::Rcout << "WsProp: " << bin.tellg() << std::endl;

        // uint8_t A, B;
        //
        // B = readbin(B, bin, 0); // len?
        //
        // if (debug) {
        //   printf("%d : %d\n", A, B);
        // }

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

        bin.seekg(size, bin.cur);
      }

      if (in_worksheet && x == BrtWsDim) {
        Rcpp::Rcout << "WsDim: " << bin.tellg() << std::endl;

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
      }

      if (in_worksheet && x == BrtBeginWsViews) {
        Rcpp::Rcout << "<sheetViews>: " << bin.tellg() << std::endl;
        // uint8_t A, B;
        //
        // B = readbin(B, bin, 0); // len?
        //
        // if (debug) {
        //   printf("%d : %d\n", A, B);
        // }
      }

      // whatever this is
      if (x == BrtSel) {
        Rcpp::Rcout << "BrtSel: " << bin.tellg() << std::endl;
        bin.seekg(size, bin.cur);
      }

      if (in_worksheet && x == BrtBeginWsView) {
        Rcpp::Rcout << "<sheetView>: " << bin.tellg() << std::endl;
        bin.seekg(size, bin.cur);
      }

      if (in_worksheet && x == BrtEndWsView) {
        Rcpp::Rcout << "</sheetView>: " << bin.tellg() << std::endl;
      }

      if (in_worksheet && x == BrtEndWsViews) {
        Rcpp::Rcout << "</sheetViews>: " << bin.tellg() << std::endl;
      }

      if (x == BrtBeginSheetData)  {
        out << "<sheetData>" << std::endl; //  << bin.tellg()
        in_sheet_data = true;
      } // 2.2.1 Cell Table

      if (in_sheet_data && x == BrtRowHdr)  {
        Rcpp::Rcout << "BrtRowHdr: " << bin.tellg() << std::endl;
        if (!first_row) {
          out << "</row>" <<std::endl;
        }
        first_row = false;
      }

      // prelude to row entry
      if (x == BrtACBegin) {
        Rcpp::Rcout << "BrtACBegin: " << bin.tellg() << std::endl;
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
      }

      if (x == BrtACEnd) {
        Rcpp::Rcout << "BrtACEnd: " << bin.tellg() << std::endl;
        bin.seekg(size, bin.cur);
      }

      if (x == BrtWsFmtInfo) {
        Rcpp::Rcout << "BrtWsFmtInfo: " << bin.tellg() << std::endl;
        bin.seekg(size, bin.cur);
      }

      if (x == BrtRwDescent) {
        Rcpp::Rcout << "BrtRwDescent: " << bin.tellg() << std::endl;
        bin.seekg(size, bin.cur);
      }

      if (x == BrtRowHdr) {

        out << "<row r=\""; //  << bin.tellg()
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

        out << rw << "\">" << std::endl;

        Rcpp::Rcout << rw << " : " << ixfe << " : " << miyRw << " : " << (int32_t)fExtraAsc << " : " <<
          (int32_t)fExtraDsc << " : " << unk32 << " : " << (int32_t)fCollapsed << " : " << (int32_t)fDyZero << " : " <<
            (int32_t)fUnsynced << " : " << (int32_t)fGhostDirty << " : " << (int32_t)fReserved << " : " << (int32_t)fPhShow << " : " <<
              ccolspan << "; " << bin.tellg() << std::endl;

      }

      if (in_sheet_data && x == BrtCellIsst)  { // shared string
        Rcpp::Rcout << "BrtCellIsst: " << bin.tellg() << std::endl;

        int32_t val1, val2, val3;
        val1 = readbin(val1, bin, 0);
        Rcpp::Rcout << val1 << std::endl;
        val2 = readbin(val2, bin, 0);
        Rcpp::Rcout << val2 << std::endl;
        val3 = readbin(val3, bin, 0);
        Rcpp::Rcout << val3 << std::endl;

        out << "<c r=\"" << val1 << "\" t=\"s\">" << std::endl;
        out << "<v>" << val3 << "</v>" << std::endl;
        out << "</c>" << std::endl;

      }

      if (in_sheet_data && x == BrtCellBool) { // bool
        Rcpp::Rcout << "BrtCellBool: " << bin.tellg() << std::endl;


        int32_t val1, val2;
        uint8_t val3;
        val1 = readbin(val1, bin, 0);
        // out << val1 << std::endl;
        val2 = readbin(val2, bin, 0);
        // out << val2 << std::endl;
        val3 = readbin(val3, bin, 0);
        // out << val3 << std::endl;

        out << "<c r=\"" << val1 << "\" t=\"b\">" << std::endl;
        out << "<v>" << (int32_t)val3 << "</v>" << std::endl;
        out << "</c>" << std::endl;
      }

      if (x == BrtCellRk) { // integer?
        Rcpp::Rcout << "BrtCellRk: " << bin.tellg() << std::endl;

        int32_t val1, val2, val3;
        val1 = readbin(val1, bin, 0);
        // Rcpp::Rcout << val << std::endl;
        val2 = readbin(val2, bin, 0);
        // Rcpp::Rcout << val << std::endl;
        // wrong?
        val3 = readbin(val3, bin, 0);
        // Rcpp::Rcout << RkNumber(val) << std::endl;

        out << "<c r=\"" << val1 << "\">" << std::endl;
        out << "<v>" << RkNumber(val3) << "</v>" << std::endl;
        out << "</c>" << std::endl;
      }

      if (x == BrtFmlaError) { // t="e" & <f>
        Rcpp::Rcout << "BrtFmlaError: " << bin.tellg() << std::endl;
        // bin.seekg(size, bin.cur);
        std::vector<int> cell;
        cell = Cell(bin);
        Rf_PrintValue(Rcpp::wrap(cell));

        std::string fErr;
        fErr = BErr(bin);
        out << "<c>" << std::endl;
        out << "<v>" << fErr << "</v>" << std::endl;
        out << "</c>" << std::endl;

        uint16_t frbitfmla = 0;
        // [0] == 0
        // [1] == fAlwaysCalc
        // [2-15] == unused
        frbitfmla = readbin(frbitfmla, bin, 0);

        // int32_t len = size - 4 * 32 - 2 * 8;
        // std::string fml(len, '\0');

        CellParsedFormula(bin);

      }


      if (x == BrtCellReal) {
        Rcpp::Rcout << "BrtCellReal: " << bin.tellg() << std::endl;
        int32_t val1, val2;
        val1 = readbin(val1, bin, 0);
        // Rcpp::Rcout << val << std::endl;
        val2 = readbin(val2, bin, 0);
        // Rcpp::Rcout << val << std::endl;

        double dbl;
        dbl = readbin(dbl, bin, 0);
        // Rcpp::Rcout << dbl << std::endl;

        #include <iomanip>

        out << "<c r=\"" << val1 << "\">" << std::endl;
        out << "<v>" << std::setprecision(16) << dbl << "</v>" << std::endl; // << std::fixed
        out << "</c>" << std::endl;
      }

      // 0 ?
      if (x == BrtCellBlank) {
        Rcpp::Rcout << "BrtCellBlank: " << bin.tellg() << std::endl;
        bin.seekg(size, bin.cur);
      }

      if (x == BrtCellError) { // t="e" & <v>#NUM!</v>
        Rcpp::Rcout << "BrtCellError: " << bin.tellg() << std::endl;

        uint8_t val8;

        int32_t val1, val2, val3;
        val1 = readbin(val1, bin, 0);
        // out << val << std::endl;
        val2 = readbin(val2, bin, 0);
        // out << val << std::endl;
        // val3 = readbin(val8, bin, 0);
        // out << val3 << std::endl;

        out << "<c r=\"" << val1 << "\" t=\"e\">" << std::endl;
        out << "<v>" << BErr(bin) << "</v>" << std::endl;
        out << "</c>" << std::endl;

      }

      if (x == BrtFmlaNum) {
        bin.seekg(size, bin.cur);
      }

      if (in_sheet_data && x == BrtEndSheetData) {
        if (!first_row) {
          out << "</row>" << std::endl;
        }
        out << "</sheetData>" << std::endl;
        in_sheet_data = false;
      // break;
      // }
      //
      // if (x == BrtEndSheet) {
        out << "</worksheet>" << std::endl;
        in_worksheet = false;
        break;
      }

      itr++;
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
