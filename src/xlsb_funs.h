#ifndef XLSB_FUNS_H
#define XLSB_FUNS_H

#include <string>
#include <fstream>
#include <cstdint>

// for swap_endian
#include <type_traits>

// detect if we need to swap. assuming that there is no big endian xlsb format,
// we only need to swap little endian xlsb files on big endian systems
bool is_big_endian() {
  uint32_t num = 1;
  uint8_t* bytePtr = reinterpret_cast<uint8_t*>(&num);
  return bytePtr[0] == 0;
}

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

// start swap_endian
template <typename T>
typename std::enable_if<std::is_same<T, int16_t>::value || std::is_same<T, uint16_t>::value, T>::type
swap_endian(T t) {
  return __builtin_bswap16(t);
}

template <typename T>
typename std::enable_if<std::is_same<T, int32_t>::value || std::is_same<T, uint32_t>::value, T>::type
swap_endian(T t) {
  return __builtin_bswap32(t);
}

template <typename T>
typename std::enable_if<std::is_same<T, int64_t>::value || std::is_same<T, uint64_t>::value, T>::type
swap_endian(T t) {
  return __builtin_bswap64(t);
}

template <typename T>
typename std::enable_if<std::is_same<T, float>::value, float>::type
swap_endian(T t) {
  union v {
    float       f;
    uint32_t    i32;
  } val;
  val.f = t;
  val.i32 = __builtin_bswap32(val.i32);
  return val.f;
}

template <typename T>
typename std::enable_if<std::is_same<T, double>::value, double>::type
swap_endian(T t) {
  union v {
    double      d;
    uint64_t    i64;
  } val;
  val.d = t;
  val.i64 = __builtin_bswap64(val.i64);
  return val.d;
}

template <typename T>
typename std::enable_if<!std::is_same<T, int16_t>::value &&
!std::is_same<T, uint16_t>::value &&
!std::is_same<T, int32_t>::value &&
!std::is_same<T, uint32_t>::value &&
!std::is_same<T, int64_t>::value &&
!std::is_same<T, uint64_t>::value &&
!std::is_same<T, float>::value &&
!std::is_same<T, double>::value, T>::type
swap_endian(T t) {
  return t;
}
// end swap_endian

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

std::string escape_quote(const std::string& input) {
  std::string result;
  result.reserve(input.length());

  for (char c : input) {
    switch (c) {
    case '\"':
      result += "\"\"";
      break;
    default:
      result += c;
    break;
    }
  }

  return result;
}

std::string escape_xml(const std::string& input) {
  std::string result;
  result.reserve(input.length());

  for (char c : input) {
    switch (c) {
    case '&':
      result += "&amp;";
      break;
    case '\"':
      result += "&quot;";
      break;
    case '\'':
      result += "&apos;";
      break;
    case '<':
      result += "&lt;";
      break;
    case '>':
      result += "&gt;";
      break;
    default:
      result += c;
    break;
    }
  }

  return result;
}

std::string to_utf8(const std::u16string& u16str) {

  std::string utf8str;
  utf8str.reserve(u16str.length() * 3); // Reserve enough space for UTF-8 characters

  for (std::size_t i = 0; i < u16str.length(); ++i) {
    char16_t u16char = u16str[i];

    // Determine the endianness of the system
    bool isLittleEndian = is_big_endian();
    char16_t networkOrderChar = u16char;
    if (isLittleEndian) {
      networkOrderChar = static_cast<char16_t>((u16char << 8) | (u16char >> 8));
    }

    if (networkOrderChar <= 0x7F) {
      // Single-byte UTF-8 character (0x00 - 0x7F)
      utf8str.push_back(static_cast<char>(networkOrderChar));
    } else if (networkOrderChar <= 0x7FF) {
      // Two-byte UTF-8 character (0x80 - 0x07FF)
      utf8str.push_back(static_cast<char>(0xC0 | (networkOrderChar >> 6)));
      utf8str.push_back(static_cast<char>(0x80 | (networkOrderChar & 0x3F)));
    } else {
      // Three-byte or four-byte UTF-8 character
      if (networkOrderChar >= 0xD800 && networkOrderChar <= 0xDBFF &&
          i + 1 < u16str.length()) {
        char16_t high_surrogate = networkOrderChar;
        char16_t low_surrogate = u16str[++i];

        // Calculate the code point from the surrogate pair
        int code_point = ((high_surrogate - 0xD800) << 10) + (low_surrogate - 0xDC00) + 0x10000;

        // Four-byte UTF-8 character (0x10000 - 0x10FFFF)
        utf8str.push_back(static_cast<char>(0xF0 | (code_point >> 18)));
        utf8str.push_back(static_cast<char>(0x80 | ((code_point >> 12) & 0x3F)));
        utf8str.push_back(static_cast<char>(0x80 | ((code_point >> 6) & 0x3F)));
        utf8str.push_back(static_cast<char>(0x80 | (code_point & 0x3F)));
      } else {
        // Three-byte UTF-8 character (0x0800 - 0xFFFF)
        utf8str.push_back(static_cast<char>(0xE0 | (networkOrderChar >> 12)));
        utf8str.push_back(static_cast<char>(0x80 | ((networkOrderChar >> 6) & 0x3F)));
        utf8str.push_back(static_cast<char>(0x80 | (networkOrderChar & 0x3F)));
      }
    }
  }

  return utf8str;
}

std::string read_xlwidestring(std::string &mystring, std::istream& sas) {

  size_t size = mystring.size();
  std::u16string str;
  str.resize(size * 2);

  if (!sas.read((char*)&str[0], str.size()))
    Rcpp::stop("char: a binary read error occurred");

  std::string outstr = to_utf8(str);
  if (str.size()/2 != size) Rcpp::warning("String size unexpected");
  // cannot resize but have to remove '\0' from string
  // mystring.resize(size);
  outstr.erase(std::remove(outstr.begin(), outstr.end(), '\0'), outstr.end());

  return(outstr);
}

std::string PtrStr(std::istream& sas, bool swapit) {
  uint16_t len = 0;
  len = readbin(len, sas, swapit);
  std::string str(len, '\0');
  return read_xlwidestring(str, sas);
}

std::string XLWideString(std::istream& sas, bool swapit) {
  uint32_t len = 0;
  len = readbin(len, sas, swapit);
  std::string str(len, '\0');
  return read_xlwidestring(str, sas);
}

// same, but can be NULL
std::string XLNullableWideString(std::istream& sas, bool swapit) {
  uint32_t len = 0;
  len = readbin(len, sas, swapit);
  if (len == 0xFFFFFFFF) {
    return "";
  }
  std::string str(len, '\0');

  return read_xlwidestring(str, sas);
}


// should really add a reliable function to get all these tricky bits
// bit 1 from uint8_t
// bit 2 from uint8_t
// bits 1-n from uint8_t
// bits 5-7 from uint8_t
// bit 8 from uint8_t

int32_t RECORD_ID(std::istream& sas, bool swapit) {
  uint8_t var1 = 0, var2 = 0;
  var1 = readbin(var1, sas, swapit);

  if (var1 & 0x80) {

    var2 = readbin(var2, sas, swapit);

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

int32_t RECORD_SIZE(std::istream& sas, bool swapit) {
  int8_t sar1 = 0, sar2 = 0, sar3 = 0, sar4 = 0;

  sar1 = readbin(sar1, sas, swapit);
  if (sar1 & 0x80) sar2 = readbin(sar2, sas, swapit);
  if (sar2 & 0x80) sar3 = readbin(sar3, sas, swapit);
  if (sar3 & 0x80) sar4 = readbin(sar4, sas, swapit);

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

int32_t RECORD(int &rid, int &rsize, std::istream& sas, bool swapit) {

  /* Record ID ---------------------------------------------------------------*/
  rid = RECORD_ID(sas, swapit);

  /* Record Size -------------------------------------------------------------*/
  rsize = RECORD_SIZE(sas, swapit);

  return 0;
}

std::string cell_style(int style) {
  std::string out = "";
  if (style > 0) {
    out = out + " s=\"" + std::to_string(style) + "\"";
  }
  return out;
}


std::string halign(int style) {
  std::string out = "";
  std::string hali = "";
  if (style > 0) {
    if (style == 1) hali = "general"; // default
    if (style == 1) hali = "left";
    if (style == 2) hali = "center";
    if (style == 3) hali = "right";
    if (style == 4) hali = "fill";
    if (style == 5) hali = "justify";
    if (style == 6) hali = "center-across";
    if (style == 7) hali = "distributed";

    out = out + " horizontal=\"" + hali + "\"";
  }
  return out;
}

std::string to_iconset(uint32_t style) {
  std::string out = "";
  if (style != 0xFFFFFFFF) {
    if (style == 0) out = "3Arrows";
    if (style == 1) out = "3ArrowsGray";
    if (style == 2) out = "3Flags";
    if (style == 3) out = "3TrafficLights1";
    if (style == 4) out = "3TrafficLights2";
    if (style == 5) out = "3Signs";
    if (style == 6) out = "3Symbols";
    if (style == 7) out = "3Symbols2";
    if (style == 8) out = "4Arrows";
    if (style == 9) out = "4ArrowsGray";
    if (style == 10) out = "4RedToBlack";
    if (style == 11) out = "4Rating";
    if (style == 12) out = "4TrafficLights";
    if (style == 13) out = "5Arrows";
    if (style == 14) out = "5ArrowsGray";
    if (style == 15) out = "5Rating";
    if (style == 16) out = "5Quarters";
    // beg iconset 14
    if (style == 17) out = "3Stars";
    if (style == 18) out = "3Triangles";
    if (style == 19) out = "5Boxes";
    // end iconset 14
  } else {
    out = "";
  }

  return out;
}

std::string valign(int style) {
  std::string out = "";
  std::string vali = "";
  if (style >= 0) {
    if (style == 0) vali = "top";
    if (style == 1) vali = "center";
    if (style == 2) vali = "bottom";
    if (style == 3) vali = "justify";
    if (style == 4) vali = "distributed";

    out = out + " vertical=\"" + vali + "\"";
  }
  return out;
}

void StrRun(std::istream& sas, uint32_t dwSizeStrRun, bool swapit) {

  uint16_t ich = 0, ifnt = 0;

  // something?
  for (uint8_t i = 0; i < dwSizeStrRun; ++i) {
    ich  = readbin(ich, sas, swapit);
    ifnt = readbin(ifnt, sas, swapit);
    // Rprintf("Styled string will be unstyled - strRun: %d %d\n", ich, ifnt);
  }
}

void PhRun(std::istream& sas, uint32_t dwPhoneticRun, bool swapit) {

  uint16_t ichFirst = 0, ichMom = 0, cchMom = 0, ifnt = 0;
  // uint32_t phrun = 0;
  for (uint8_t i = 0; i < dwPhoneticRun; ++i) {
    ichFirst = readbin(ichFirst, sas, swapit);
    ichMom   = readbin(ichMom, sas, swapit);
    cchMom   = readbin(cchMom, sas, swapit);
    ifnt     = readbin(ifnt, sas, swapit);
  }
}

std::string RichStr(std::istream& sas, bool swapit) {

  uint8_t AB = 0; // , unk = 0
  bool A = 0, B = 0;
  AB = readbin(AB, sas, swapit);

  A = AB & 0x01;
  B = (AB >> 1) & 0x01;
  // remaining 6 bits are ignored
  // unk = AB & 0x3F;
  // if (debug) if (unk) Rcpp::Rcout << std::to_string(unk) << std::endl;

  std::string str = XLWideString(sas, swapit);

  uint32_t dwSizeStrRun = 0, dwPhoneticRun = 0;

  if (A) { //  || unk
    // Rcpp::Rcout << "a" << std::endl;
    // number of runs following
    dwSizeStrRun = readbin(dwSizeStrRun, sas, swapit);
    if (dwSizeStrRun > 0x7FFF) Rcpp::stop("dwSizeStrRun to large");
    StrRun(sas, dwSizeStrRun, swapit);
  }

  if (B) {
    // Rcpp::Rcout << "b" << std::endl;
    std::string phoneticStr = XLWideString(sas, swapit);

    // number of runs following
    dwPhoneticRun = readbin(dwPhoneticRun, sas, swapit);
    if (dwPhoneticRun > 0x7FFF) Rcpp::stop("dwPhoneticRun to large");
    PhRun(sas, dwPhoneticRun, swapit);
  }

  return(str);

}

void ProductVersion(std::istream& sas, bool swapit, bool debug) {
  uint16_t fileVersion = 0, fileProduct = 0;
  int8_t fileExtension = 0;

  fileVersion = readbin(fileVersion, sas, swapit);
  fileProduct = readbin(fileProduct, sas, swapit);

  fileExtension = fileProduct & 0x01;
  fileProduct   = fileProduct & ~static_cast<uint16_t>(0x01);
  if (debug) Rprintf("ProductVersion: %d: %d: %d\n", fileVersion, fileProduct, fileExtension);
}

std::vector<int> UncheckedRfX(std::istream& sas, bool swapit) {

  std::vector<int> out;
  int32_t rwFirst= 0, rwLast= 0, colFirst= 0, colLast= 0;

  out.push_back(readbin(rwFirst, sas, swapit));
  out.push_back(readbin(rwLast, sas, swapit));
  out.push_back(readbin(colFirst, sas, swapit));
  out.push_back(readbin(colLast, sas, swapit));

  return(out);
}

std::vector<int> UncheckedSqRfX(std::istream& sas, bool swapit) {

  std::vector<int> out;
  int32_t crfx = 0;
  crfx = readbin(crfx, sas, swapit);
  out.push_back(crfx);

  for (int32_t i = 0; i < crfx; ++i) {
    std::vector<int> ucrfx = UncheckedRfX(sas, swapit);
    out.insert(out.end(), ucrfx.begin(), ucrfx.end());
  }

  return out;
}


int UncheckedCol(std::istream& sas, bool swapit) {
  int32_t col = 0;
  col = readbin(col, sas, swapit);
  if (col >= 0 && col <= 16383)
    return col;
  else
    Rcpp::stop("col size bad: %d @ %d", col, sas.tellg());
}

int UncheckedRw(std::istream& sas, bool swapit) {
  int32_t row = 0;
  row = readbin(row, sas, swapit);
  if (row >= 0 && row <= 1048575)
    return row;
  else
    Rcpp::stop("row size bad: %d @ %d", row, sas.tellg());
}

uint16_t ColShort(std::istream& sas, bool swapit) {
  uint16_t col = 0;
  col = readbin(col, sas, swapit);
  if (col >= 0 && col <= 16383)
    return col;
  else
    Rcpp::stop("col size bad: %d @ %d", col, sas.tellg());
}

std::vector<int> ColRelShort(std::istream& sas, bool swapit) {
  uint16_t tmp = 0;
  tmp = readbin(tmp, sas, swapit);

  int16_t col = 0, fColRel = 0, fRwRel = 0;
  col     = static_cast<int16_t>((tmp & 0x3FFF));
  fColRel = (tmp >> 14) & 0x0001;
  fRwRel  = (tmp >> 15) & 0x0001;

  std::vector<int> out(3);
  out[0] = col;
  out[1] = fColRel;
  out[2] = fRwRel;

  return out;
}

std::string Loc(std::istream& sas, bool swapit) {

  std::vector<int> col;
  uint32_t row = 0;
  row = UncheckedRw(sas, swapit);
  col = ColRelShort(sas, swapit);

  bool fColRel = col[1];
  bool fRwRel  = col[2];

  std::string out;

  if (!fColRel) out += "$";
  out += int_to_col(col[0] + 1);

  if (!fRwRel) out += "$";
  out += std::to_string(row + 1);

  return out;
}

// RgceLocRel
std::string LocRel(std::istream& sas, bool swapit, int col, int row) {

  std::vector<int> col_rel;
  int32_t row_rel = 0;
  row_rel = readbin(row_rel, sas, swapit);
  col_rel = ColRelShort(sas, swapit);

  bool fColRel = col_rel[1];
  bool fRwRel  = col_rel[2];

  std::string out;

  if (fRwRel) {
    row_rel += row;

    if (row_rel < 0x00000000)
      row_rel += 0x00100000;
    else if (row_rel > 0x000FFFFF)
      row_rel -= 0x00100000;
  }

  if (fColRel) {
    col_rel[0] += col;

    if (col_rel[0] < 0x0000)
      col_rel[0] += 0x4000;
    else if (col_rel[0] > 0x3FFF)
      col_rel[0] -= 0x4000;
  }

  // if (!fColRel) out += "$";
  out += int_to_col(col_rel[0] + 1);

  // if (!fRwRel) out += "$";
  out += std::to_string(row_rel + 1);

  return out;
}

std::string Area(std::istream& sas, bool swapit) {

  std::vector<int> col0(3), col1(3);
  uint32_t row0 = 0, row1 = 0;
  row0 = UncheckedRw(sas, swapit); // rowFirst
  row1 = UncheckedRw(sas, swapit); // rowLast
  col0 = ColRelShort(sas, swapit); // columnFirst
  col1 = ColRelShort(sas, swapit); // columnLast

  bool fColRel0 = col0[1];
  bool fRwRel0  = col0[2];
  bool fColRel1 = col1[1];
  bool fRwRel1  = col1[2];

  std::string out;

  if (!fColRel0) out += "$";
  out += int_to_col(col0[0] + 1);

  if (!fRwRel0) out += "$";
  out += std::to_string(row0 + 1);

  out += ":";

  if (!fColRel1) out += "$";
  out += int_to_col(col1[0] + 1);

  if (!fRwRel1) out += "$";
  out += std::to_string(row1 + 1);

  return out;
}

std::string AreaRel(std::istream& sas, bool swapit, int col, int row) {

  std::vector<int> col_rel0(3), col_rel1(3);
  int32_t row_rel0 = 0, row_rel1 = 0;
  row_rel0 = UncheckedRw(sas, swapit); // rowFirst
  row_rel1 = UncheckedRw(sas, swapit); // rowLast
  col_rel0 = ColRelShort(sas, swapit); // columnFirst
  col_rel1 = ColRelShort(sas, swapit); // columnLast

  bool fColRel0 = col_rel0[1];
  bool fRwRel0  = col_rel0[2];
  bool fColRel1 = col_rel1[1];
  bool fRwRel1  = col_rel1[2];

  std::string out;

  if (fRwRel0) {
    row_rel0 += row;

    if (row_rel0 < 0x00000000)
      row_rel0 += 0x00100000;
    else if (row_rel0 > 0x000FFFFF)
      row_rel0 -= 0x00100000;
  }

  if (fColRel0) {
    col_rel0[0] += col;

    if (col_rel0[0] < 0x0000)
      col_rel0[0] += 0x4000;
    else if (col_rel0[0] > 0x3FFF)
      col_rel0[0] -= 0x4000;
  }

  if (fRwRel1) {
    row_rel1 += row;

    if (row_rel1 < 0x00000000)
      row_rel1 += 0x00100000;
    else if (row_rel1 > 0x000FFFFF)
      row_rel1 -= 0x00100000;
  }

  if (fColRel1) {
    col_rel1[0] += col;

    if (col_rel1[0] < 0x0000)
      col_rel1[0] += 0x4000;
    else if (col_rel1[0] > 0x3FFF)
      col_rel1[0] -= 0x4000;
  }

  // if (!fColRel0) out += "$";
  out += int_to_col(col_rel0[0] + 1);

  // if (!fRwRel0) out += "$";
  out += std::to_string(row_rel0 + 1);

  out += ":";

  // if (!fColRel1) out += "$";
  out += int_to_col(col_rel1[0] + 1);

  // if (!fRwRel1) out += "$";
  out += std::to_string(row_rel1 + 1);

  return out;
}

std::vector<int> Cell(std::istream& sas, bool swapit) {

  std::vector<int> out(3);

  out[0] = UncheckedCol(sas, swapit);

  int32_t uint = 0;
  uint = readbin(uint, sas, swapit);

  out[1] = uint & 0xFFFFFF;   // iStyleRef
  out[2] = (uint & 0x02000000) >> 24; // fPhShow
  // unused

  return(out);
}

std::vector<std::string> dims_to_cells(int firstRow, int lastRow, int firstCol, int lastCol) {

  std::vector<int> cols, rows;
  for (int32_t i = firstCol; i <= lastCol; ++i) {
    cols.push_back(i);
  }

  for (int32_t i = firstRow; i <= lastRow; ++i) {
    rows.push_back(i);
  }

  std::vector<std::string> cells;
  for (int32_t col : cols) {
    for (int32_t row : rows) {
      cells.push_back(int_to_col(col) + std::to_string(row));
    }
  }

  return cells;
}

std::vector<int> brtColor(std::istream& sas, bool swapit) {

  uint8_t AB = 0, xColorType = 0, index = 0,
    bRed = 0, bGreen = 0, bBlue = 0, bAlpha = 0;
  int16_t nTintAndShade = 0;

  AB = readbin(AB, sas, swapit);
  // bool fValidRGB = AB & 0x01;
  // Rprintf("valid RGB color: %d", fValidRGB);
  xColorType = AB >> 1;
  // 0x00 picked by application
  // 0x01 color index value
  // 0x02 argb color
  // 0x03 theme color

  index = readbin(index, sas, swapit);
  // 0x00 Undefined
  // 0x01 Icv() ...
  // 0x02 Undefined
  // 0x03 clrScheme subelement

  nTintAndShade = readbin(nTintAndShade, sas, swapit);

  bRed   = readbin(bRed, sas, swapit);
  bGreen = readbin(bGreen, sas, swapit);
  bBlue  = readbin(bBlue, sas, swapit);
  bAlpha = readbin(bAlpha, sas, swapit);

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

std::string as_border_style(int style) {
  if (style == 0) return "none";
  if (style == 1) return "thin";
  if (style == 2) return "medium";
  if (style == 3) return "dashed";
  if (style == 4) return "dotted";
  if (style == 5) return "thick";
  if (style == 6) return "double";
  if (style == 7) return "hair";
  if (style == 8) return "mediumDashed";
  if (style == 9) return "dashDot";
  if (style == 10) return "mediumDashDot";
  if (style == 11) return "dashDotDot";
  if (style == 12) return "mediumDashDotDot";
  if (style == 13) return "slantDashDot";
  return "";
}

std::string to_argb(int a, int r, int g, int b) {
  std::stringstream out;
  out << std::uppercase << std::hex <<
    std::setw(2) << std::setfill('0') << (int32_t)a <<
      std::setw(2) << std::setfill('0') << (int32_t)r <<
        std::setw(2) << std::setfill('0') << (int32_t)g <<
          std::setw(2) << std::setfill('0') << (int32_t)b;
  return out.str();
}

std::string brtBorder(std::string type, std::istream& sas, bool swapit) {
  uint8_t dg = 0, reserved = 0;

  dg = readbin(dg, sas, swapit);
  reserved = readbin(reserved, sas, swapit);

  std::vector<int> color = brtColor(sas, swapit);

  std::stringstream out;

  out << "<" << type << " style = \"" << as_border_style(dg) << "\"";
  if (dg > 0) {

    double tint = 0.0;
    if (color[2] != 0) tint = (double)color[2]/32767;

    std::stringstream stream;
    stream << std::setprecision(16) << tint;

    if (color[0] == 0x00)
      out << "><color auto=\"1\" />" << std::endl;
    if (color[0] == 0x01)
      out << "><color indexed=\"" << color[1] << "\" />";
    if (color[0] == 0x02)
      out << "><color hex=\"" << to_argb(color[6], color[3], color[4], color[5]) << "\" />";
    if (color[0] == 0x03)
      out << "><color theme=\"" << color[1] << "\" tint=\"" << stream.str() << "\" />";

    out << "</" << type << ">" << std::endl;
  } else {
    out << "/>" << std::endl;
  }

  return out.str();
}

static double Xnum(std::istream& sas, bool swapit) {
  double out = 0;
  out = readbin(out, sas, swapit);
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

std::string BErr(std::istream& sas, bool swapit) {
  uint8_t error = 0;
  error = readbin(error, sas, swapit);

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

std::vector<int> Xti(std::istream& sas, bool swapit) {
  int32_t firstSheet = 0, lastSheet = 0;
  uint32_t externalLink = 0;
  externalLink = readbin(externalLink, sas, swapit);
  // scope
  // -2 workbook
  // -1 heet level
  // >= 0 sheet level
  firstSheet = readbin(firstSheet, sas, swapit);
  lastSheet = readbin(lastSheet, sas, swapit);

  // Rprintf("Xti: %d %d %d\n", externalLink, firstSheet, lastSheet);
  std::vector<int> out(3);
  out[0] = externalLink;
  out[1] = firstSheet;
  out[2] = lastSheet;

  return out;
}

// bool isOperator(const std::string& token) {
//   return token == "+" || token == "-" || token == "*" || token == "/" ||
//     token == "^" || token == "%" || token == "="  || token == "&lt;&gt;" ||
//     token == "&lt;" || token == "&gt;" || token == "&le;" || token == "&ge;" ||
//     token == " " || token == "&amp;" || token == ":" || token == "," ||
//     token == "#" || token == "@";
// }

#include <stack>

std::string parseRPN(const std::string& expression) {

  std::istringstream iss(expression);
  std::string line;
  std::stack<std::string> formulaStack;

  while (std::getline(iss, line)) {
    if (line.empty()) {
      break;  // Stop parsing at the first empty line
    }
    std::string token = line;
    // Rcpp::Rcout << token << std::endl;

    // if (isOperator(token)) {
    //
    //   if (formulaStack.size() == 1) {
    //     std::string operand1 = formulaStack.top();
    //     formulaStack.pop();
    //     // uminus & uplus
    //     std::string infixExpression = token + operand1;
    //     formulaStack.push(infixExpression);
    //   } else if (formulaStack.size() >= 2) {
    //     // Rcpp::Rcout << "Formula stacksize is: " << formulaStack.size() << std::endl;
    //     std::string operand2 = formulaStack.top();
    //     formulaStack.pop();
    //     std::string operand1 = formulaStack.top();
    //     formulaStack.pop();
    //     std::string infixExpression = operand1 + token + operand2;
    //     formulaStack.push(infixExpression);
    //   }
    // } else {
      if (token.find("%s") != std::string::npos) {
        size_t pos = token.rfind("%s");
        while (pos != std::string::npos) {
          if (!formulaStack.empty()) {
            token.replace(pos, 2, formulaStack.top());
            formulaStack.pop();
          } else {
            // some functions are actually empty like RAND(). Have to check if
            // argument order is correct in this case
            token.replace(pos, 2, "");
          }
          pos = token.rfind("%s");
        }
        formulaStack.push(token);
      } else {
        // Rcpp::Rcout << "push to stack: " << token << std::endl;
        formulaStack.push(token);
      }
    // }
  }

  std::string parsedFormula;
  while (!formulaStack.empty()) {
    if (parsedFormula.empty()) {
      parsedFormula = formulaStack.top();
    } else {
      parsedFormula = formulaStack.top() + " " + parsedFormula;
    }
    formulaStack.pop();
  }

  return parsedFormula;
}


std::string CellParsedFormula(std::istream& sas, bool swapit, bool debug, int col, int row, int &sharedFml) {
  // bool ptg_extra_array = false;
  uint32_t  cce= 0, cb= 0;

  if (debug) Rcpp::Rcout << "CellParsedFormula: " << sas.tellg() << std::endl;

  cce = readbin(cce, sas, swapit);
  if (cce >= 16385) Rcpp::stop("wrong cce size");
  if (debug) Rcpp::Rcout << "cce: " << cce << std::endl;
  size_t pos = sas.tellg();
  // sas.seekg(cce, sas.cur);
  pos += cce;
  int8_t val1 = 0;

  // row = 0;

  std::vector<int32_t> ptgextra;

  std::string fml_out;
  while((size_t)sas.tellg() < pos) {

    if (debug) Rcpp::Rcout << ".";

    // 1
    uint8_t val2 = 0, controlbit = 0;
    val1 = readbin(val1, sas, swapit);

    controlbit = (val1 & 0x80) >> 7;
    if (controlbit != 0) Rcpp::warning("controlbit unexpectedly not 0");
    val1 &= 0x7F; // the remaining 7 bits form ptg
    // for some Ptgs only the first 5 are of interest
    // and 6 and 7 contain DataType information
    if (debug) Rprintf("Formula: %d %d\n", val1, val2);

    switch(val1) {

    case PtgList_PtgSxName:
    case PtgAttr:
    {
      if (debug) Rcpp::Rcout << "reading eptg @ " << sas.tellg() << std::endl;
      val2 = readbin(val2, sas, swapit); // full 8 bit forming eptg
      if (debug) Rcpp::Rcout << "PtgAttr: " << std::hex << (int)val1 << ": "<< (int)val2 << std::dec << std::endl;


      switch(val2) {

      case PtgList:
      {
        RgbExtra typ = PtgExtraList;
        ptgextra.push_back(typ);
        if (debug) Rcpp::Rcout << "PtgList " << sas.tellg() << std::endl;
        uint16_t ixti = 0, flags = 0;
        uint32_t listIndex = 0;
        int16_t colFirst = 0, colLast = 0;

        // this is a reference to a table column something like "tab[col]"
        ixti = readbin(ixti, sas, swapit);
        flags = readbin(flags, sas, swapit);
        listIndex = readbin(listIndex, sas, swapit);
        colFirst = ColShort(sas, swapit);
        colLast = ColShort(sas, swapit);

        std::stringstream paddedStr;
        paddedStr << std::setw(12) << std::setfill('0') << ixti;

        // A1 notation cell
        fml_out += "openxlsx2xlsb_" + paddedStr.str();
        // maybe [ ]
        // fml_out += "#REF!";
        fml_out += "\n";

        // Do something with this, just ... what?
        if (debug) Rprintf("PtgList: %d, %d, %d, %d\n",
            ixti, listIndex, colFirst, colLast);

        Rcpp::warning("formulas with table references are not implemented.");

        break;
      }

      case PtgAttrSemi:
      {
        // function is volatile and has to be calculated every time the
        // workbook is opened
        if (debug) Rcpp::Rcout << "PtgAttrSemi" << std::endl;

        if ((val2  & 1) != 1) Rcpp::stop("wrong value");
        // val2[1] == 1
        // val2[2:8] == 0
        uint16_t unused = 0;
        unused = readbin(unused, sas, swapit);
        // Rcpp::Rcout << unused << std::endl;
        break;
      }

      // beg control tokens. according to manual they can be ignored
      case PtgAttrIf:
      case PtgAttrIfError:
      case PtgAttrGoTo:
      {
        if (debug) Rcpp::Rcout << "PtgAttrIf" <<std::endl;

        uint16_t offset = 0;

        offset = readbin(offset, sas, swapit);

        break;
      }

      case PtgAttrChoose:
      {
        if (debug) Rcpp::Rcout << "PtgAttrChoose" <<std::endl;

        uint16_t cOffset = 0;
        uint32_t rgOffset0 = 0, rgOffset1 = 0;

        cOffset = readbin(cOffset, sas, swapit);
        rgOffset0 = readbin(rgOffset0, sas, swapit);
        rgOffset1 = readbin(rgOffset1, sas, swapit);

        break;
      }
      // end control tokens

      case PtgAttrSpace:
      {
        uint8_t type = 0, cch = 0;

        if (((val2 >> 6) & 1) != 1) Rcpp::stop("wrong value");

        type = readbin(type, sas, swapit);
        // 0-6 various different types where to add the whitespace
        cch = readbin(cch, sas, swapit);

        // hm, there is also " " as operator. genius move ...
        // fml_out += "%s";
        for (uint8_t i = 0; i < cch; ++i) {
          fml_out += " ";
        }
        // fml_out += "%s";
        // fml_out += "\n";

        if (debug) Rprintf("AttrSpace: %d %d\n", type, cch);
        break;
      }

      case PtgAttrSum:
      {
        if (debug) Rcpp::Rcout << "PtgAttrSum" << std::endl;
        // val2[1:4] == 0
        // val2[5]   == 1
        // val2[6:8] == 0
        uint16_t unused = 0;
        unused = readbin(unused, sas, swapit);
        // Rcpp::Rcout << unused << std::endl;
        fml_out += "SUM(%s)"; // maybe attr because it is a single cell function?
        fml_out += "\n";
        break;
      }

      default:{
        Rprintf("Undefined Formula_TWO: %d %d\n", val1, val2);
        break;
      }
      }

      break;
    }

    case PtgRange:
    {
      if (debug) Rcpp::Rcout << ":" <<std::endl;
      fml_out += "%s:%s";
      fml_out += "\n";
      break;
    }

    case PtgUnion:
    {
      if (debug) Rcpp::Rcout << "," <<std::endl;
      fml_out += "%s,%s";
      fml_out += "\n";
      break;
    }

    case PtgIsect:
    {
      if (debug) Rcpp::Rcout << " " <<std::endl;
      fml_out += "%s %s";
      fml_out += "\n";
      break;
    }

    case PtgAdd:
    {
      if (debug) Rcpp::Rcout << "+" <<std::endl;
      fml_out += "%s+%s";
      fml_out += "\n";
      break;
    }

    case PtgSub:
    {
      if (debug) Rcpp::Rcout << "-" <<std::endl;
      fml_out += "%s-%s";
      fml_out += "\n";
      break;
    }

    case PtgMul:
    {
      if (debug) Rcpp::Rcout << "*" <<std::endl;
      fml_out += "%s*%s";
      fml_out += "\n";
      break;
    }

    case PtgDiv:
    {
      if (debug) Rcpp::Rcout << "/" <<std::endl;
      fml_out +=  "%s/%s";
      fml_out += "\n";
      break;
    }

    case PtgPercent:
    {
      if (debug) Rcpp::Rcout << "%" <<std::endl;
      fml_out +=  "%s%";
      fml_out += "\n";
      break;
    }

    case PtgPower:
    {
      if (debug) Rcpp::Rcout << "^" <<std::endl;
      fml_out += "%s^%s";
      fml_out += "\n";
      break;
    }

    case PtgConcat:
    {
      if (debug) Rcpp::Rcout << "&" <<std::endl;
      fml_out += "%s&amp;%s";
      fml_out += "\n";
      break;
    }

    case PtgEq:
    {
      if (debug) Rcpp::Rcout << "=" <<std::endl;
      fml_out += "%s=%s";
      fml_out += "\n";
      break;
    }

    case PtgGt:
    {
      if (debug) Rcpp::Rcout << ">" <<std::endl;
      fml_out += "%s&gt;%s";
      fml_out += "\n";
      break;
    }

    case PtgGe:
    {
      if (debug) Rcpp::Rcout << ">=" <<std::endl;
      fml_out += "%s&gt;=%s";
      fml_out += "\n";
      break;
    }

    case PtgLt:
    {
      if (debug) Rcpp::Rcout << "<" <<std::endl;
      fml_out += "%s&lt;%s";
      fml_out += "\n";
      break;
    }

    case PtgLe:
    {
      if (debug) Rcpp::Rcout << "<=" <<std::endl;
      fml_out += "%s&lt;=%s";
      fml_out += "\n";
      break;
    }

    case PtgNe:
    {
      if (debug) Rcpp::Rcout << "!=" <<std::endl;
      fml_out += "%s&lt;&gt;%s";
      fml_out += "\n";
      break;
    }

    case PtgUPlus:
    {
      if (debug) Rcpp::Rcout << "+val" <<std::endl;
      fml_out += "+%s";
      fml_out += "\n";
      break;
    }

    case PtgUMinus:
    {
      if (debug) Rcpp::Rcout << "-val" <<std::endl;
      fml_out += "-%s";
      fml_out += "\n";
      break;
    }

    case PtgParen:
    {
      if (debug) Rcpp::Rcout << "()" <<std::endl;
      fml_out += "(%s)";
      fml_out += "\n";
      break;
    }

    case PtgMissArg:
    {
      if (debug) Rcpp::Rcout << "MISSING()" <<std::endl;
      fml_out += "";
      break;
    }

    case PtgInt:
    {
      if (debug) Rcpp::Rcout << "PtgInt" <<std::endl;
      uint16_t integer = 0;
      integer = readbin(integer, sas, swapit);
      // Rcpp::Rcout << integer << std::endl;
      fml_out += std::to_string(integer);
      fml_out += "\n";
      break;
    }

    case PtgNum:
    {
      if (debug) Rcpp::Rcout << "PtgNum" <<std::endl;
      double value = Xnum(sas, swapit);
      fml_out += std::to_string(value);
      fml_out += "\n";
      break;
    }

    case PtgStr:
    {
      if (debug) Rcpp::Rcout << "PtgStr" <<std::endl;

      fml_out += "\"" + escape_quote(PtrStr(sas, swapit)) + "\"";
      fml_out += "\n";

      break;
    }

    case PtgArray:
    case PtgArray2:
    case PtgArray3:
    {
      if (debug) Rcpp::Rcout << "PtgArray" <<std::endl;

      RgbExtra typ = PtgExtraArray;
      ptgextra.push_back(typ);

      // ptg_extra_array = true;
      uint16_t unused2 = 0;
      uint32_t unused1 = 0, unused3 = 0, unused4 = 0;
      unused1 = readbin(unused1, sas, swapit);
      unused2 = readbin(unused2, sas, swapit);
      unused3 = readbin(unused3, sas, swapit);
      unused4 = readbin(unused4, sas, swapit);

      // TODO: this saves the spot for the array. the formula is still broken:
      // in the formula parser the stack will remain in reverse order
      // something like SUM(@array@) {"myarray"} will be returned. *sigh*
      // NOTE: maybe it's better to push the formula to a string vector and find
      // and replace @array@ with array information
      fml_out += "@array@";
      fml_out += "\n";

      break;
    }


    case PtgRef:
    case PtgRef2:
    case PtgRef3:
    {
      if (debug) Rcpp::Rcout << "PtgRef" <<std::endl;
      // uint8_t ptg8 = 0, ptg = 0, PtgDataType = 0, null = 0;
      //
      // 2
      // Rprintf("PtgRef2: %d, %d, %d\n", ptg, PtgDataType, null);

      // val1[6:7] == PtgDataType

      fml_out +=  Loc(sas, swapit);
      fml_out += "\n";

      if (debug) Rcpp::Rcout << sas.tellg() << std::endl;

      break;
    }

    case PtgRef3d:
    case PtgRef3d2:
    case PtgRef3d3:
    {
      if (debug) Rcpp::Rcout << "PtgRef3d" <<std::endl;
      // need_ptg_revextern = true;
      RgbExtra typ = RevExtern;
      ptgextra.push_back(typ);

      uint16_t ixti = 0;
      ixti = readbin(ixti, sas, swapit); // XtiIndex
      if (debug) Rprintf("XtiIndex: %d\n", ixti);

      std::stringstream paddedStr;
      paddedStr << std::setw(12) << std::setfill('0') << ixti;


      // A1 notation cell
      fml_out += "openxlsx2xlsb_" + paddedStr.str() + "!";
      fml_out += Loc(sas, swapit);
      fml_out += "\n";

      if (debug) Rcpp::Rcout << sas.tellg() << std::endl;

      break;
    }

    case PtgRefN:
    case PtgRefN2:
    case PtgRefN3:
    {
      if (debug) Rcpp::Rcout << "PtgRefN" <<std::endl;

      // A1 notation cell
      fml_out += LocRel(sas, swapit, col, row);
      fml_out += "\n";

      break;
    }

    case PtgArea:
    case PtgArea2:
    case PtgArea3:
    {
      if (debug) Rcpp::Rcout << "PtgArea" <<std::endl;

      // A1 notation cell
      fml_out += Area(sas, swapit);
      fml_out += "\n";

      break;
    }

    case PtgArea3d:
    case PtgArea3d2:
    case PtgArea3d3:
    {
      if (debug) Rcpp::Rcout << "PtgArea3d" <<std::endl;

      // need_ptg_revextern = true;
      RgbExtra typ = RevExtern;
      ptgextra.push_back(typ);

      uint16_t ixti = 0;

      ixti = readbin(ixti, sas, swapit);
      if (debug) Rprintf("ixti in PtgArea3d: %d\n", ixti);

      std::stringstream paddedStr;
      paddedStr << std::setw(12) << std::setfill('0') << ixti;

      // A1 notation cell
      fml_out += "openxlsx2xlsb_" + paddedStr.str() + "!";
      fml_out += Area(sas, swapit);
      fml_out += "\n";

      if (debug) Rcpp::Rcout << sas.tellg() << std::endl;

      break;
    }

    case PtgAreaN:
    case PtgAreaN2:
    case PtgAreaN3:
    {
      if (debug) Rcpp::Rcout << "PtgAreaN" <<std::endl;

      // A1 notation cell
      fml_out += AreaRel(sas, swapit, col, row);
      fml_out += "\n";

      break;
    }

    case PtgMemArea:
    case PtgMemArea2:
    case PtgMemArea3:
    {
      if (debug) Rcpp::Rcout << "PtgMemArea" <<std::endl;

      // need_ptg_extra_mem = true;
      RgbExtra typ = PtgExtraMem;
      ptgextra.push_back(typ);

      uint16_t cce = 0;
      uint32_t unused = 0;
      unused = readbin(unused, sas, swapit);
      cce = readbin(cce, sas, swapit);
      // number of bytes for PtgExtraMem section?

      fml_out += "@mem@";
      fml_out += "\n";

      break;
    }

    case PtgName:
    case PtgName2:
    case PtgName3:
    {
      if (debug) Rcpp::Rcout << "PtgName" <<std::endl;

      // need_ptg_revnametabid = true;
      RgbExtra typ = RevNameTabid;
      ptgextra.push_back(typ);
      uint32_t nameindex = 0;
      nameindex = readbin(nameindex, sas, swapit);
      // Rcpp::Rcout << nameindex << std::endl;
      fml_out += std::to_string(nameindex);
      fml_out += "\n";

      break;
    }

    case PtgNameX:
    case PtgNameX2:
    case PtgNameX3:
    {
      if (debug) Rcpp::Rcout << "PtgNameX" <<std::endl;

      // need_ptg_revname = true;
      RgbExtra typ = RevName;
      ptgextra.push_back(typ);
      // not yet found
      uint16_t ixti = 0;
      uint32_t nameindex = 0;
      ixti = readbin(ixti, sas, swapit);
      nameindex = readbin(nameindex, sas, swapit);
      // Rcpp::Rcout << nameindex << std::endl;
      // fml_out += std::to_string(ixti);
      fml_out += std::to_string(nameindex);
      fml_out += "\n";

      break;
    }

    case PtgRefErr:
    case PtgRefErr2:
    case PtgRefErr3:
    {
      if (debug) Rcpp::Rcout << "PtgRefErr" <<std::endl;

      uint16_t unused2 = 0;
      uint32_t unused1 = 0;

      unused1 = readbin(unused1, sas, swapit);
      unused2 = readbin(unused2, sas, swapit);

      fml_out += "#REF!";
      fml_out += "\n";

      break;
    }

    case PtgRefErr3d:
    case PtgRefErr3d2:
    case PtgRefErr3d3:
    {
      if (debug) Rcpp::Rcout << "PtgRefErr3d" <<std::endl;

      // need_ptg_revextern = true;
      RgbExtra typ = RevExtern;
      ptgextra.push_back(typ);
      uint16_t ixti = 0, unused2 = 0;
      uint32_t unused1 = 0;

      ixti = readbin(ixti, sas, swapit);
      unused1 = readbin(unused1, sas, swapit);
      unused2 = readbin(unused2, sas, swapit);

      std::stringstream paddedStr;
      paddedStr << std::setw(12) << std::setfill('0') << ixti;

      // A1 notation cell
      fml_out += "openxlsx2xlsb_" + paddedStr.str() + "!";
      fml_out += "#REF!";
      fml_out += "\n";

      if (debug) Rcpp::Rcout << sas.tellg() << std::endl;

      break;
    }

    case PtgAreaErr3d:
    case PtgAreaErr3d2:
    case PtgAreaErr3d3:
    {
      if (debug) Rcpp::Rcout << "PtgAreaErr3d" <<std::endl;

      // need_ptg_revextern = true;
      RgbExtra typ = RevExtern;
      ptgextra.push_back(typ);
      uint16_t ixti = 0;
      uint32_t unused1 = 0, unused2 = 0, unused3 = 0;

      ixti = readbin(ixti, sas, swapit);
      unused1 = readbin(unused1, sas, swapit);
      unused2 = readbin(unused2, sas, swapit);
      unused3 = readbin(unused3, sas, swapit);

      std::stringstream paddedStr;
      paddedStr << std::setw(12) << std::setfill('0') << ixti;

      // A1 notation cell
      fml_out += "openxlsx2xlsb_" + paddedStr.str() + "!";
      fml_out += "#REF!";
      fml_out += "\n";

      if (debug) Rcpp::Rcout << sas.tellg() << std::endl;

      break;
    }

    case PtgFunc:
    case PtgFunc2:
    case PtgFunc3:
    {
      if (debug) Rcpp::Rcout << "PtgFunc" <<std::endl;

      uint16_t iftab = 0;
      iftab = readbin(iftab, sas, swapit);

      std::string fml = Ftab(iftab);

      fml_out += fml;
      fml_out += "\n";

      break;
    }

    case PtgFuncVar:
    case PtgFuncVar2:
    case PtgFuncVar3:
    {
      if (debug) Rcpp::Rcout << "PtgFuncVar" <<std::endl;

      uint8_t cparams = 0, fCeFunc = 0; // number of parameters
      cparams = readbin(cparams, sas, swapit);

      uint16_t tab = 0;
      tab = readbin(tab, sas, swapit);
      // tab[16] == fCeFunc bool
      // tab[1:15] == tab

      // TODO add a check that Ftab() returns only functions
      // with variable number of  arguments
      fCeFunc = (tab >> 15) & 0x0001;
      tab &= 0x7FFF;
      if (fCeFunc) {
        fml_out += Cetab(tab);
      } else {
        fml_out += Ftab(tab);
      }

      fml_out += "(";
      if (cparams) fml_out += "%s"; // RAND() needs no argument
      for (uint8_t i = 1; i < cparams; ++i) {
        fml_out += ",%s";
      }
      fml_out += ")\n";

      if (debug) Rprintf("PtgFuncVar: %d %d %d\n", cparams, tab, fCeFunc);

      break;
    }

    case PtgErr:
    {
      if (debug) Rcpp::Rcout << "PtgErr" <<std::endl;

      fml_out += BErr(sas, swapit);
      fml_out += "\n";
      break;
    }

    case PtgBool:
    {
      if (debug) Rcpp::Rcout << "PtgBool" <<std::endl;

      int8_t boolean = 0;
      boolean = readbin(boolean, sas, swapit);
      fml_out += std::to_string(boolean);
      fml_out += "\n";
      break;
    }

    case PtgExp:
    {
      if (debug) Rcpp::Rcout << "PtgExp" <<std::endl;

      // this is a reference to the cell that contains the shared formula
      uint32_t row = UncheckedRw(sas, swapit) + 1;
      if (debug) Rcpp::Rcout << "PtgExp: " << row << std::endl;
      sharedFml = row;
      break;
    }

    case PtgMemFunc:
    case PtgMemFunc2:
    case PtgMemFunc3:
    {
      if (debug) Rcpp::Rcout << "PtgMemFunc" <<std::endl;

      uint16_t cce = 0;
      cce = readbin(cce, sas, swapit);
      break;
    }

    default:
    {
      // if (debug)
      Rcpp::warning("Undefined Formula: %d %d\n", val1, val2);
      break;
    }
    }

  }

  if ((size_t)sas.tellg() != pos) {
    // somethings not correct
    sas.seekg(pos, sas.beg);
  }

  if (debug) Rcpp::Rcout << "--- formula ---\n" << fml_out << std::endl;

  cb = readbin(cb, sas, swapit);

  if (debug)
    Rcpp::Rcout << "cb: " << cb << std::endl;

  pos = sas.tellg();
  // sas.seekg(cce, sas.cur);

  // Rcpp::Rcout << "pre " << pos << std::endl;
  pos += cb;
  // Rcpp::Rcout << "post " << pos << std::endl;

  size_t cntr = 0;

  while((size_t)sas.tellg() < pos) {

    if (debug) Rcpp::Rcout << ".";
    if (debug) {
      Rprintf("Formula cb: %d\n", val1);
      Rprintf("%d: %d\n", (int)sas.tellg(), (int)pos);
    }
    // this is a little risky. maybe its some kind of vector indicating the
    // order in which extra elements are going to be selected?

    if (ptgextra.size() > 0 && ptgextra.size() > cntr) {
      if (ptgextra[cntr] == PtgExtraArray) {
        if (debug) Rcpp::Rcout << "need PtgArray" << std::endl;
        val1 = PtgArray;
      } else if  (ptgextra[cntr] == PtgExtraCol) {
        if (debug) Rcpp::Rcout << "need PtgExp" << std::endl;
        val1 = PtgExp;
      } else if (ptgextra[cntr] == RevExtern) {
        if (debug) Rcpp::Rcout << "need RevExtern" << std::endl;
        val1 = RevExtern;
      } else{
        Rcpp::Rcout << ptgextra[cntr] << std::endl;
      }
    } else if (ptgextra.size() < (cntr + 1)) {
      if (debug) Rprintf("ptgextra %d and %d\n", (int)ptgextra.size(),  (int)cntr);
    }


    ++cntr;

    switch(val1) {
    case PtgExp:
    { // PtgExtraCol

      // need_ptg_extra_col = true;
      if (debug) Rcpp::Rcout << "PtgExtraCol" << std::endl;
      int32_t col = UncheckedCol(sas, swapit);
      if (debug) Rcpp::Rcout << "cb PtgExp: " << int_to_col(col+1) << std::endl;

      fml_out += int_to_col(col + 1);
      fml_out += std::to_string(row);
      fml_out += "\n";
      break;
    }

    // TODO: this does not handle {"foo", "bar"}
    case PtgArray:
    case PtgArray2:
    case PtgArray3:
    {
      if (debug) Rcpp::Rcout << "PtgExtraArray" << std::endl;
      // PtgExtraArray

      uint32_t rows = 0, cols = 0;
      rows = readbin(rows, sas, swapit);
      cols = readbin(cols, sas, swapit);
      // blob (it is actually called this way)
      uint8_t reserved = 0;
      reserved = readbin(reserved, sas, swapit);

      std::string array = "";

      if (debug) Rcpp::Rcout << (int32_t)reserved << std::endl;

      // SerBool
      if (reserved == 0x02) {
        if (debug) Rcpp::Rcout << "SerBool" << std::endl;
        uint8_t f = 0;
        f = readbin(f, sas, swapit);
        if (debug) Rcpp::Rcout << (int32_t)f << std::endl;

        array = "{" + std::to_string((int32_t)f) + "}";
        // fml_out += "\n";
      }

      // SerErr
      if (reserved == 0x04) {
        if (debug) Rcpp::Rcout << "SerErr" << std::endl;
        uint8_t reserved2 = 0;
        uint16_t reserved3 = 0;
        std::string strerr = BErr(sas, swapit);
        if (debug) Rcpp::Rcout << strerr << std::endl;
        reserved2 = readbin(reserved2, sas, swapit);
        reserved3 = readbin(reserved3, sas, swapit);

        array = "{" + strerr + "}";
        // fml_out += "\n";
      }

      // SerNum
      if (reserved == 0x00) {
        if (debug) Rcpp::Rcout << "SerNum" << std::endl;
        double xnum = 0.0;
        xnum = Xnum(sas, swapit);

        std::stringstream stream;
        stream << std::setprecision(16) << xnum;

        if (debug) Rcpp::Rcout << xnum << std::endl;
        array = "{" + stream.str() + "}";
        // fml_out += "\n";
      }

      // SerStr
      if (reserved == 0x01) {
        if (debug) Rcpp::Rcout << "SerStr" << std::endl;
        uint16_t cch = 0;
        cch = readbin(cch, sas, swapit);
        std::string rgch(cch, '\0');
        rgch = read_xlwidestring(rgch, sas);
        if (debug)
          Rcpp::Rcout << rgch << std::endl;

        array = "{\"" + rgch + "\"}";
        // fml_out += "\n";
      }

      size_t fi = fml_out.find("@array@");

      if (fi != std::string::npos) {
        if (debug) Rcpp::Rcout << "replacing @array@" << std::endl;
        fml_out.replace(fi, 7, array);
      } else {
        if (debug) Rcpp::Rcout << "no  @array@" <<  fml_out << std::endl;
        fml_out += array;
        fml_out += "\n";
      }

      if (debug) Rcpp::Rcout << fml_out << std::endl;

      break;
    }

      // do i need this?
    case RevExtern:
    {
      if (debug) Rcpp::Rcout << "RevExtern" << std::endl;

      sas.seekg(pos, sas.beg);
      break;

      // I fear that I have to, but on another rainy day
      // The issue I was tackling is an array in a reference

      // uint16_t book = 0;
      // book = readbin(book, sas, swapit);
      // if ((book >> 8) & 0xFF == 0x01) {
      //   if (book & 0xFF != 0x02)
      //     Rcpp::stop("not two");
      // } else {
      //   uint8_t flags = 0;
      //   flags = readbin(flags, sas, swapit);
      //   std::string str(book, '\0');
      //   return read_xlwidestring(str, sas);
      //
      // }
    }

    default :
    {
      // Rcpp::stop("Skip");
      Rcpp::Rcout << "undefined cb: " << cb << std::endl;
      sas.seekg(pos, sas.beg);
      break;
    }
    }
  }

  if ((size_t)sas.tellg() != pos) {
    // somethings not correct
    sas.seekg(pos, sas.beg);
  }

  if (debug) {
    Rcpp::Rcout << "...fml..." << std::endl;
    Rcpp::Rcout << fml_out << std::endl;
  }
  std::string inflix =  parseRPN(fml_out);

  return inflix;
}

#endif
