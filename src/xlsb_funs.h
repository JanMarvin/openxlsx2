#include <string>
#include <fstream>
#include <cstdint>
#include <typeinfo>

// for to_utf8
#include <locale>
#include <codecvt>

#define GCC_VERSION (__GNUC__ * 10000 \
+ __GNUC_MINOR__ * 100                \
+ __GNUC_PATCHLEVEL__)

/* Test for GCC < 4.8.0 */
#if GCC_VERSION < 40800 & !__clang__
static inline unsigned short __builtin_bswap16(unsigned short a)
{
  return (a<<8)|(a>>8);
}
#endif

template <typename T>
T swap_endian(T t) {
  if (typeid(T) == typeid(int16_t))
    return __builtin_bswap16(t);
  if (typeid(T) == typeid(uint16_t))
    return __builtin_bswap16(t);

  if (typeid(T)  == typeid(int32_t))
    return __builtin_bswap32(t);
  if (typeid(T)  == typeid(uint32_t))
    return __builtin_bswap32(t);

  if (typeid(T)  == typeid(int64_t))
    return __builtin_bswap64(t);
  if (typeid(T)  == typeid(uint64_t))
    return __builtin_bswap64(t);

  union v {
    double      d;
    float       f;
    uint32_t    i32;
    uint64_t    i64;
  } val;

  if (typeid(T) == typeid(float)){
    val.f = t;
    val.i32 = __builtin_bswap32(val.i32);
    return val.f;
  }

  if (typeid(T) == typeid(double)){
    val.d = t;
    val.i64 = __builtin_bswap64(val.i64);
    return val.d;
  }

  else
    return t;
}

template <typename T>
T readbin( T t , std::istream& sas, bool swapit)
{
  if (!sas.read ((char*)&t, sizeof(t)))
    Rcpp::stop("readbin: a binary read error occurred");
  if (swapit==0)
    return(t);
  else
    return(swap_endian(t));
}

template <typename T>
inline std::string readstring(std::string &mystring, T& sas)
{

  if (!sas.read(&mystring[0], mystring.size()))
    Rcpp::stop("char: a binary read error occurred");

  return(mystring);
}

std::string to_utf8(const std::u16string& utf16String)
{
  std::wstring_convert<std::codecvt_utf8_utf16<char16_t>, char16_t> converter;
  return converter.to_bytes(utf16String);
}

// should really add a reliable function to get all these tricky bits
// bit 1 from uint8_t
// bit 2 from uint8_t
// bits 1-n from uint8_t
// bits 5-7 from uint8_t
// bit 8 from uint8_t

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


  if (sar2 != 0 && sar3 != 0 && sar4 != 0) {
    int32_t recordType = ((sar4 & 0x7F) << 7) | ((sar3 & 0x7F) << 7) | ((sar2 & 0x7F) << 7) | (sar1 & 0x7F);
    return recordType;
  }

  if (sar2 != 0 && sar3 != 0 && sar4 == 0) {
    int32_t recordType = ((sar3 & 0x7F) << 7) | ((sar2 & 0x7F) << 7) | (sar1 & 0x7F);
    return recordType;
  }

  if (sar2 != 0 && sar3 == 0 && sar4 == 0) {
    int32_t recordType = ((sar2 & 0x7F) << 7) | (sar1 & 0x7F);
    return recordType;
  }

  if (sar2 == 0 && sar3 == 0 && sar4 == 0) {
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

std::string read_xlwidestring(std::string &mystring, std::istream& sas) {

  uint8_t size = mystring.size();
  std::u16string str;
  str.resize(size * 2);

  if (!sas.read((char*)&str[0], str.size()))
    Rcpp::stop("char: a binary read error occurred");

  mystring = to_utf8(str);
  mystring.resize(size);

  return(mystring);
}

std::string cell_style(int style) {
  std::string out = "";
  if (style > 0) {
    out = out + " s=\"" + std::to_string(style) + "\"";
  }
  return out;
}


std::string XLWideString(std::istream& sas) {
  uint32_t len = 0;
  len = readbin(len, sas, 0);
  std::string str(len, '\0');
  return read_xlwidestring(str, sas);
}

// same, but can be NULL
std::string XLNullableWideString(std::istream& sas) {
  uint32_t len = 0;
  len = readbin(len, sas, 0);
  std::string str(len, '\0');
  if (len == 0xFFFFFFFF) {
    return str;
  }

  return read_xlwidestring(str, sas);
}

void StrRun(std::istream& sas, uint32_t dwSizeStrRun) {

  uint16_t ich = 0, ifnt = 0;

  // something?
  for (uint8_t i = 0; i < dwSizeStrRun; ++i) {
    ich  = readbin(ich, sas, 0);
    ifnt = readbin(ifnt, sas, 0);
    Rprintf("strRun: %d %d\n", ich, ifnt);
  }
}

void PhRun(std::istream& sas, uint32_t dwPhoneticRun) {

  uint16_t ichFirst = 0, ichMom = 0, cchMom = 0, ifnt = 0;
  // uint32_t phrun = 0;
  for (uint8_t i = 0; i < dwPhoneticRun; ++i) {
    ichFirst = readbin(ichFirst, sas, 0);
    ichMom   = readbin(ichMom, sas, 0);
    cchMom   = readbin(cchMom, sas, 0);
    ifnt     = readbin(ifnt, sas, 0);
  }
}

std::string RichStr(std::istream& sas) {

  uint8_t AB= 0, A= 0, B= 0;
  AB = readbin(AB, sas, 0);

  A = AB & 0x01;
  B = AB >> 1 & 0x01;
  // remaining 6 bits are ignored

  std::string str = XLWideString(sas);

  uint32_t dwSizeStrRun = 0, dwPhoneticRun = 0;

  if (A) {
    // number of runs following
    dwSizeStrRun = readbin(dwSizeStrRun, sas, 0);
    StrRun(sas, dwSizeStrRun);
  }

  if (B) {
    std::string phoneticStr = XLWideString(sas);

    // number of runs following
    dwPhoneticRun = readbin(dwPhoneticRun, sas, 0);
    PhRun(sas, dwPhoneticRun);
  }

  return(str);

}

void ProductVersion(std::istream& sas) {
  uint16_t fileVersion = 0, fileProduct = 0;
  int8_t fileExtension = 0;

  fileVersion = readbin(fileVersion, sas, 0);
  fileProduct = readbin(fileProduct, sas, 0);

  fileExtension = fileProduct & 0x01;
  fileProduct   = fileProduct & ~static_cast<uint16_t>(0x01);
  Rprintf("ProductVersion: %d: %d: %d\n", fileVersion, fileProduct, fileExtension);
}

std::vector<int> UncheckedRfX(std::istream& sas) {

  std::vector<int> out;
  int32_t rwFirst= 0, rwLast= 0, colFirst= 0, colLast= 0;

  out.push_back(readbin(rwFirst, sas, 0));
  out.push_back(readbin(rwLast, sas, 0));
  out.push_back(readbin(colFirst, sas, 0));
  out.push_back(readbin(colLast, sas, 0));

  return(out);
}

int UncheckedCol(std::istream& sas) {
  int32_t col = 0;
  col = readbin(col, sas, 0);
  if (col >= 0 && col <= 16383)
    return col;
  else
    Rcpp::stop("col size bad: %d @ %d", col, sas.tellg());
}

int UncheckedRw(std::istream& sas) {
  int32_t row = 0;
  row = readbin(row, sas, 0);
  if (row >= 0 && row <= 1048575)
    return row;
  else
    Rcpp::stop("row size bad: %d @ %d", row, sas.tellg());
}

int ColRelShort(std::istream& sas) {
  uint16_t col = 0;
  int32_t out = 0;
  col = readbin(col, sas, 0);
  int8_t fColRel= 0, fRwRel= 0;

  out = col & 0x3FFF;
  fColRel = (col >> 14) & 1;
  fRwRel  = (col >> 15) & 1;

  // Rprintf("ColRelShort: %d, %d\n", fColRel, fRwRel);

  return out;
}

std::vector<int> Loc(std::istream& sas) {

  std::vector<int> out(2);
  out[0] = UncheckedRw(sas);
  out[1] = ColRelShort(sas);

  return out;
}

std::vector<int> Area(std::istream& sas) {

  std::vector<int> out(4);
  out[0] = UncheckedRw(sas); // rowFirst
  out[1] = UncheckedRw(sas); // rowLast
  out[2] = ColRelShort(sas); // columnFirst
  out[3] = ColRelShort(sas); // columnLast

  return out;
}

std::vector<int> Cell(std::istream& sas) {

  std::vector<int> out(3);

  out[0] = UncheckedCol(sas);

  int32_t uint = 0;
  uint = readbin(uint, sas, 0);

  out[1] = uint & 0xFFFFFF;   // iStyleRef
  out[2] = (uint & 0x02000000) >> 24; // fPhShow
  // unused

  return(out);
}

std::vector<int> brtColor(std::istream& sas) {

  bool fValidRGB = 0;
  uint8_t AB = 0, xColorType = 0, index = 0,
    bRed = 0, bGreen = 0, bBlue = 0, bAlpha = 0;
  int16_t nTintAndShade = 0;

  AB = readbin(AB, sas, 0);
  fValidRGB = AB & 0x01;
  xColorType = AB >> 1;
  // 0x00 picked by application
  // 0x01 color index value
  // 0x02 argb color
  // 0x03 theme color

  index = readbin(index, sas, 0);
  // 0x00 Undefined
  // 0x01 Icv() ...
  // 0x02 Undefined
  // 0x03 clrScheme subelement

  nTintAndShade = readbin(nTintAndShade, sas, 0);

  bRed   = readbin(bRed, sas, 0);
  bGreen = readbin(bGreen, sas, 0);
  bBlue  = readbin(bBlue, sas, 0);
  bAlpha = readbin(bAlpha, sas, 0);

  if (bRed > 255 || bGreen > 255 || bBlue > 255 || bAlpha > 255) {
    Rcpp::stop("invalid color");
  }

  std::vector<int32_t> out(7);
  out[0] = xColorType;
  out[1] = index;
  out[2] = nTintAndShade;
  out[3] = bRed;
  out[4] = bGreen;
  out[5] = bBlue;
  out[6] = bAlpha;

  return out;
}

static double Xnum(std::istream& sas) {
  double out = 0;
  out = readbin(out, sas, 0);
  return out;
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
  if (error == 0x07) return "#DIV/0!";
  if (error == 0x0F) return "#VALUE!";
  if (error == 0x17) return "#REF!";
  if (error == 0x1D) return "#NAME?";
  if (error == 0x24) return "#NUM!";
  if (error == 0x2A) return "#N/A";
  if (error == 0x2B) return "#GETTING_DATA";

  return "unknown_ERROR";
}


std::vector<int> Xti(std::istream& sas) {
  int32_t firstSheet = 0, lastSheet = 0;
  uint32_t externalLink = 0;
  externalLink = readbin(externalLink, sas, 0);
  // scope
  // -2 workbook
  // -1 heet level
  // >= 0 sheet level
  firstSheet = readbin(firstSheet, sas, 0);
  lastSheet = readbin(lastSheet, sas, 0);

  // Rprintf("Xti: %d %d %d\n", externalLink, firstSheet, lastSheet);
  std::vector<int> out(3);
  out[0] = externalLink;
  out[1] = firstSheet;
  out[2] = lastSheet;

  return out;
}

std::string CellParsedFormula(std::istream& sas, bool debug) {
  uint32_t  cce= 0, cb= 0;

  if (debug) Rcpp::Rcout << "CellParsedFormula: " << sas.tellg() << std::endl;

  cce = readbin(cce, sas, 0);
  if (cce > 16385) Rcpp::stop("wrong cce size");
  if (debug) Rcpp::Rcout << "cce: " << cce << std::endl;
  size_t pos = sas.tellg();
  // sas.seekg(cce, sas.cur);
  pos += cce;

  std::string fml_out;
  while(sas.tellg() < pos) {

    if (debug) Rcpp::Rcout << ".";

    // 1
    int8_t val1 = 0, val2 = 0, controlbit = 0;
    val1 = readbin(val1, sas, 0);

    controlbit = (val1 & 0x80) >> 7;
    if (controlbit != 0) Rcpp::stop("controlbit unexpectedly not 0");
    val1 &= 0x7F; // the remaining 7 bits form ptg
    // for some Ptgs only the first 5 are of interest
    // and 6 and 7 contain DataType information

    switch(val1) {

    case PtgList_PtgSxName:
    case PtgAttr:
    {
      if (debug) Rcpp::Rcout << "reading eptg @ " << sas.tellg() << std::endl;
      val2 = readbin(val2, sas, 0); // full 8 bit forming eptg
      Rcpp::Rcout << std::hex << (int)val1 << ": "<< (int)val2 << std::dec << std::endl;


      if (val2 == PtgAttrSemi) {
        if (debug) Rcpp::Rcout << "PtgAttrSemi" << std::endl;
        // val2[1] == 1
        // val2[2:8] == 0
        uint16_t unused = 0;
        unused = readbin(unused, sas, 0);
        fml_out += ";";
        fml_out += "\n";
      }

      if (val2 == PtgAttrSum) {
        if (debug) Rcpp::Rcout << "PtgAttrSum" << std::endl;
        // val2[1:4] == 0
        // val2[5]   == 1
        // val2[6:8] == 0
        uint16_t unused = 0;
        unused = readbin(unused, sas, 0);
        fml_out += "SUM"; // maybe attr because it is a single cell function?
        fml_out += "\n";
      }

      break;
    }


    case PtgAdd:
    {
      if (debug) Rcpp::Rcout << "+" <<std::endl;
      fml_out += "+";
      fml_out += "\n";
      break;
    }

    case PtgSub:
    {
      if (debug) Rcpp::Rcout << "-" <<std::endl;
      fml_out += "-";
      fml_out += "\n";
      break;
    }

    case PtgMul:
    {
      if (debug) Rcpp::Rcout << "*" <<std::endl;
      fml_out += "*";
      fml_out += "\n";
      break;
    }

    case PtgDiv:
    {
      if (debug) Rcpp::Rcout << "/" <<std::endl;
      fml_out +=  "/";
      fml_out += "\n";
      break;
    }

    case PtgPower:
    {
      if (debug) Rcpp::Rcout << "^" <<std::endl;
      fml_out += "^";
      fml_out += "\n";
      break;
    }

    case PtgConcat:
    {
      if (debug) Rcpp::Rcout << "&" <<std::endl;
      fml_out += "&";
      fml_out += "\n";
      break;
    }

    case PtgGt:
    {
      fml_out += ">";
      fml_out += "\n";
      break;
    }

    case PtgGe:
    {
      fml_out += ">=";
      fml_out += "\n";
      break;
    }

    case PtgLt:
    {
      fml_out += "<";
      fml_out += "\n";
      break;
    }

    case PtgLe:
    {
      fml_out += "<=";
      fml_out += "\n";
      break;
    }

    case PtgNe:
    {
      fml_out += "<>";
      fml_out += "\n";
      break;
    }

    case PtgInt:
    {
      uint16_t integer = 0;
      integer = readbin(integer, sas, 0);
      // Rcpp::Rcout << integer << std::endl;
      fml_out += std::to_string(integer);
      fml_out += "\n";
      break;
    }

    case PtgRef:
    case PtgRef2:
    case PtgRef3:
    {
      // uint8_t ptg8 = 0, ptg = 0, PtgDataType = 0, null = 0;
      //
      // 2
      // Rprintf("PtgRef2: %d, %d, %d\n", ptg, PtgDataType, null);

      // val1[6:7] == PtgDataType

      // 6 & 10
      std::vector<int> out = Loc(sas);

      // A1 notation cell
      fml_out += int_to_col(out[1] + 1L);
      fml_out += std::to_string(out[0] + 1L);
      fml_out += "\n";

      if (debug) Rf_PrintValue(Rcpp::wrap(out));
      if (debug) Rcpp::Rcout << sas.tellg() << std::endl;

      break;
    }

    case PtgRef3d:
    case PtgRef3d2:
    case PtgRef3d3:
    {

      uint16_t ixti = 0;
      ixti = readbin(ixti, sas, 0); // XtiIndex
      if (debug) Rprintf("XtiIndex: %d\n", ixti);

      std::vector<int> out = Loc(sas);

      // A1 notation cell
      fml_out += int_to_col(out[1] + 1L);
      fml_out += std::to_string(out[0] + 1L);
      fml_out += "\n";

      if (debug) Rf_PrintValue(Rcpp::wrap(out));
      if (debug) Rcpp::Rcout << sas.tellg() << std::endl;

      break;
    }

    case PtgArea:
    case PtgArea2:
    case PtgArea3:
    {

      std::vector<int> out = Area(sas);

      // A1 notation cell
      fml_out += int_to_col(out[2] + 1L);
      fml_out += std::to_string(out[0] + 1L);
      fml_out += ":";
      fml_out += int_to_col(out[3] + 1L);
      fml_out += std::to_string(out[1] + 1L);
      fml_out += "\n";

      if (debug) Rf_PrintValue(Rcpp::wrap(out));
      if (debug) Rcpp::Rcout << sas.tellg() << std::endl;

      break;
    }

    case PtgFunc:
    case PtgFunc2:
    case PtgFunc3:
    {

      uint16_t iftab = 0;
      iftab = readbin(iftab, sas, 0);

      fml_out += Ftab(iftab);
      fml_out += "\n";

      break;
    }

    case PtgFuncVar:
    case PtgFuncVar2:
    case PtgFuncVar3:
    {
      uint8_t cparams = 0, fCeFunc = 0; // number of parameters
      cparams = readbin(cparams, sas, 0);

      uint16_t tab = 0;
      tab = readbin(tab, sas, 0);
      // tab[16] == fCeFunc bool
      // tab[1:15] == tab

      fCeFunc = (tab >> 15) & 0x0001;
      tab &= 0x7FFF;
      if (fCeFunc) {
        fml_out += Cetab(tab);
        fml_out += "\n";
      } else {
        fml_out += Ftab(tab);
        fml_out += "\n";
      }

      if (debug) Rprintf("PtgFuncVar: %d %d %d\n", cparams, tab, fCeFunc);

      break;
    }

    case PtgErr:
    {
      fml_out += BErr(sas);
      fml_out += "\n";
      break;
    }

    case PtgBool:
    {
      int8_t boolean = 0;
      boolean = readbin(boolean, sas, 0);
      fml_out += std::to_string(boolean);
      fml_out += "\n";
      break;
    }

    default:
    {
      // if (debug)
      Rprintf("Formula: %d %d\n", val1, val2);

    }
    }

  }

  if (debug) Rcpp::Rcout << "--- formula ---\n" << fml_out << std::endl;


  cb = readbin(cb, sas, 0);
  if (debug) Rcpp::Rcout << "cb: " << cb << std::endl;

  sas.seekg(cb, sas.cur);

  return fml_out;
}
