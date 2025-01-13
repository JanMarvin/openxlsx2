#ifndef XLSB_FUNS_H
#define XLSB_FUNS_H

#include <cstdint>
#include <fstream>
#include <string>

/* We have no real way to test if the big endian stuff works. Some parts might,
 * others not. */

// #nocov start
// for swap_endian
#include <type_traits>

// detect if we need to swap. assuming that there is no big endian xlsb format,
// we only need to swap little endian xlsb files on big endian systems
bool is_big_endian() {
  uint32_t num = 1;
  uint8_t* bytePtr = reinterpret_cast<uint8_t*>(&num);
  return bytePtr[0] == 0;
}
#include <cstdint>
#include <type_traits>

#define GCC_VERSION (__GNUC__ * 10000 + __GNUC_MINOR__ * 100 + __GNUC_PATCHLEVEL__)

/* Test for GCC < 4.8.0 */
#if GCC_VERSION < 40800 && !__clang__
static inline uint16_t __builtin_bswap16(uint16_t a) {
  return static_cast<uint16_t>((a << 8) | (a >> 8));
}
#endif

// Start swap_endian
template <typename T>
typename std::enable_if<
  std::is_same<T, int16_t>::value || std::is_same<T, uint16_t>::value, T>::type
swap_endian(T t) {
  using UnsignedType = typename std::make_unsigned<T>::type;
  UnsignedType swapped = __builtin_bswap16(static_cast<UnsignedType>(t));
  return static_cast<T>(swapped);
}

template <typename T>
typename std::enable_if<
  std::is_same<T, int32_t>::value || std::is_same<T, uint32_t>::value, T>::type
swap_endian(T t) {
  using UnsignedType = typename std::make_unsigned<T>::type;
  UnsignedType swapped = __builtin_bswap32(static_cast<UnsignedType>(t));
  return static_cast<T>(swapped);
}

template <typename T>
typename std::enable_if<
  std::is_same<T, int64_t>::value || std::is_same<T, uint64_t>::value, T>::type
swap_endian(T t) {
  using UnsignedType = typename std::make_unsigned<T>::type;
  UnsignedType swapped = __builtin_bswap64(static_cast<UnsignedType>(t));
  return static_cast<T>(swapped);
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
typename std::enable_if<
  !std::is_same<T, int16_t>::value &&
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
// End swap_endian

// #nocov end

template <typename T>
T readbin(T t, std::istream& sas, bool swapit) {
  if (!sas.read((char*)&t, sizeof(t)))
    Rcpp::stop("readbin: a binary read error occurred");
  if (swapit == 0)
    return (t);
  else
    return (swap_endian(t));
}

template <typename T>
inline std::string readstring(std::string& mystring, T& sas) {
  if (!sas.read(&mystring[0], mystring.size()))
    Rcpp::stop("char: a binary read error occurred");

  return (mystring);
}

std::string escape_quote(const std::string& input) {
  std::string result;
  result.reserve(input.length());

  for (char c : input) {
    switch (c) {
      case '\"':
        result += "\"\"";
        break;
      case '\n':
        result += "&#xA;";
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

std::string wrap_xml(const std::string& str) {
  if (!str.empty() && isspace(str[0])) {
    return "<t xml:space=\"preserve\">" + str + "</t>";
  } else {
    return "<t>" + str + "</t>";
  }
}

std::string to_utf8(const std::u16string& u16str) {
  std::string utf8str;
  utf8str.reserve(u16str.length() * 3);  // Reserve enough space for UTF-8 characters

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
        int32_t code_point = ((high_surrogate - 0xD800) << 10) + (low_surrogate - 0xDC00) + 0x10000;

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

std::string read_xlwidestring(std::string& mystring, std::istream& sas) {
  size_t size = mystring.size();
  std::u16string str;
  str.resize(size * 2);

  if (!sas.read((char*)&str[0], static_cast<uint32_t>(str.size())))
    Rcpp::stop("char: a binary read error occurred");

  std::string outstr = to_utf8(str);
  if (str.size()/2 != size) Rcpp::warning("String size unexpected");
  // cannot resize but have to remove '\0' from string
  // mystring.resize(size);
  outstr.erase(std::remove(outstr.begin(), outstr.end(), '\0'), outstr.end());

  return (outstr);
}

std::string PtrStr(std::istream& sas, bool swapit) {
  uint16_t len = 0;
  len = readbin(len, sas, swapit);
  std::string str(len, '\0');
  return read_xlwidestring(str, sas);
}

std::string LPWideString(std::istream& sas, bool swapit) {
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

int32_t RECORD(int32_t& rid, int32_t& rsize, std::istream& sas, bool swapit) {
  /* Record ID ---------------------------------------------------------------*/
  rid = RECORD_ID(sas, swapit);

  /* Record Size -------------------------------------------------------------*/
  rsize = RECORD_SIZE(sas, swapit);

  return 0;
}

std::string cell_style(int32_t style) {
  std::string out = "";
  if (style > 0) {
    out = out + " s=\"" + std::to_string(style) + "\"";
  }
  return out;
}

std::string halign(int32_t style) {
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

std::string valign(int32_t style) {
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

std::vector<std::pair<int, int>> StrRun(std::istream& sas, uint32_t dwSizeStrRun, bool swapit) {
  uint16_t ich = 0, ifnt = 0;

  // something?
  std::vector<std::pair<int, int>> str_run;
  for (uint8_t i = 0; i < dwSizeStrRun; ++i) {
    ich  = readbin(ich, sas, swapit);
    ifnt = readbin(ifnt, sas, swapit);
    // Rprintf("Styled string will be unstyled - strRun: %d %d\n", ich, ifnt);
    str_run.push_back({ich, ifnt});
  }

  return str_run;
}

// Function to get a safe substring from a UTF-8 string
std::string utf8_substr(const std::string& str, int32_t start, int32_t length) {
  size_t byte_pos = 0;  // Byte position in the original string
  size_t char_pos = 0;  // Character position

  // Find the byte position of the start character
  while (char_pos < static_cast<size_t>(start) && byte_pos < str.size()) {
    if ((str[byte_pos] & 0x80) == 0) {
      // Single-byte character (ASCII)
      byte_pos += 1;
    } else if ((str[byte_pos] & 0xE0) == 0xC0) {
      // Two-byte character
      byte_pos += 2;
    } else if ((str[byte_pos] & 0xF0) == 0xE0) {
      // Three-byte character
      byte_pos += 3;
    } else if ((str[byte_pos] & 0xF8) == 0xF0) {
      // Four-byte character
      byte_pos += 4;
    } else {
      Rcpp::stop("Invalid UTF-8 encoding detected.");
    }
    ++char_pos;
  }

  size_t start_byte_pos = byte_pos;

  // Find the byte position of the end character
  while (char_pos < static_cast<size_t>(start + length) && byte_pos < str.size()) {
    if ((str[byte_pos] & 0x80) == 0) {
      // Single-byte character (ASCII)
      byte_pos += 1;
    } else if ((str[byte_pos] & 0xE0) == 0xC0) {
      // Two-byte character
      byte_pos += 2;
    } else if ((str[byte_pos] & 0xF0) == 0xE0) {
      // Three-byte character
      byte_pos += 3;
    } else if ((str[byte_pos] & 0xF8) == 0xF0) {
      // Four-byte character
      byte_pos += 4;
    } else {
      Rcpp::stop("Invalid UTF-8 encoding detected.");
    }
    ++char_pos;
  }

  return str.substr(start_byte_pos, byte_pos - start_byte_pos);
}

std::string to_rich_text(const std::string& str, const std::vector<std::pair<int, int>>& str_runs) {
  std::string result;
  int32_t start = 0, len = 0;

  for (size_t str_run = 0; str_run < str_runs.size(); ++str_run) {
    if (str_run == 0 && str_runs[0].first > 0) {
      start = 0;
      len = str_runs[str_run].first;

      std::string part = utf8_substr(str, start, len);

      result += "<r><FONT_" + std::to_string(str_runs[str_run].second) + "/>" + wrap_xml(escape_xml(part)) + "</r>";
    }

    start = str_runs[str_run].first;
    if ((str_run + 1) < str_runs.size())
      len = str_runs[str_run + 1].first - start;
    else
      len = static_cast<int32_t>(str.size()) - start;

    std::string part = utf8_substr(str, start, len);

    result += "<r><FONT_" + std::to_string(str_runs[str_run].second) + "/>" + wrap_xml(escape_xml(part)) + "</r>";
  }

  return result;
}

void PhRun(std::istream& sas, uint32_t dwPhoneticRun, bool swapit) {
  uint16_t ichFirst = 0, ichMom = 0, cchMom = 0, ifnt = 0, AB = 0;
  // uint32_t phrun = 0;
  for (uint8_t i = 0; i < dwPhoneticRun; ++i) {
    ichFirst = readbin(ichFirst, sas, swapit);
    ichMom = readbin(ichMom, sas, swapit);
    cchMom = readbin(cchMom, sas, swapit);
    ifnt = readbin(ifnt, sas, swapit);
    AB = readbin(AB, sas, swapit);  // phType (2) alcH (2) unused (12)
  }
}

std::string RichStr(std::istream& sas, bool swapit) {
  uint8_t AB = 0;  // , unk = 0
  bool A = 0, B = 0;
  AB = readbin(AB, sas, swapit);

  A = AB & 0x01;
  B = (AB >> 1) & 0x01;
  // remaining 6 bits are ignored
  // unk = AB & 0x3F;
  // if (debug) if (unk) Rcpp::Rcout << std::to_string(unk) << std::endl;

  std::vector<std::pair<int, int>> str_run;

  std::string str = XLWideString(sas, swapit);

  uint32_t dwSizeStrRun = 0, dwPhoneticRun = 0;

  if (A) {  //  || unk
    // Rcpp::Rcout << "a" << std::endl;
    // number of runs following
    dwSizeStrRun = readbin(dwSizeStrRun, sas, swapit);
    if (dwSizeStrRun > 0x7FFF) Rcpp::stop("dwSizeStrRun to large");
    str_run = StrRun(sas, dwSizeStrRun, swapit);

    str = to_rich_text(str, str_run);
  } else {
    str = wrap_xml(escape_xml(str));
  }

  if (B) {
    // Rcpp::Rcout << "b" << std::endl;
    std::string phoneticStr = XLWideString(sas, swapit);

    // number of runs following
    dwPhoneticRun = readbin(dwPhoneticRun, sas, swapit);
    if (dwPhoneticRun > 0x7FFF) Rcpp::stop("dwPhoneticRun to large");
    PhRun(sas, dwPhoneticRun, swapit);
  }

  return (str);
}

void ProductVersion(std::istream& sas, bool swapit, bool debug) {
  uint16_t version = 0, flags = 0;
  version = readbin(version, sas, swapit);  // 3586 - x14?
  flags = readbin(flags, sas, swapit);      // 0

  /* unused and commented due to a false positive in GCC12 reported on CRAN */
  // FRTVersionFields *fields = (FRTVersionFields *)&flags;
  // if (fields->reserved != 0) Rcpp::stop("product version reserved not 0");
  // if (debug) Rprintf("ProductVersion: %d: %d: %d\n", version, fields->product, fields->reserved);
}

std::vector<int32_t> UncheckedRfX(std::istream& sas, bool swapit) {
  std::vector<int32_t> out;
  int32_t rwFirst = 0, rwLast = 0, colFirst = 0, colLast = 0;

  out.push_back(readbin(rwFirst, sas, swapit));
  out.push_back(readbin(rwLast, sas, swapit));
  out.push_back(readbin(colFirst, sas, swapit));
  out.push_back(readbin(colLast, sas, swapit));

  return (out);
}

std::vector<int32_t> UncheckedSqRfX(std::istream& sas, bool swapit) {
  std::vector<int32_t> out;
  int32_t crfx = 0;
  crfx = readbin(crfx, sas, swapit);
  out.push_back(crfx);

  for (int32_t i = 0; i < crfx; ++i) {
    std::vector<int32_t> ucrfx = UncheckedRfX(sas, swapit);
    out.insert(out.end(), ucrfx.begin(), ucrfx.end());
  }

  return out;
}

int32_t UncheckedCol(std::istream& sas, bool swapit) {
  int32_t col = 0;
  col = readbin(col, sas, swapit);
  if (col >= 0 && col <= 16383)
    return col;
  else
    Rcpp::stop("col size bad: %d @ %d", col, sas.tellg());
}

int32_t UncheckedRw(std::istream& sas, bool swapit) {
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

std::vector<int32_t> ColRelShort(std::istream& sas, bool swapit) {
  uint16_t tmp = 0;
  tmp = readbin(tmp, sas, swapit);

  int16_t col = 0, fColRel = 0, fRwRel = 0;
  col     = static_cast<int16_t>((tmp & 0x3FFF));
  fColRel = (tmp >> 14) & 0x0001;
  fRwRel  = (tmp >> 15) & 0x0001;

  std::vector<int32_t> out(3);
  out[0] = col;
  out[1] = fColRel;
  out[2] = fRwRel;

  return out;
}

std::string Loc(std::istream& sas, bool swapit) {
  std::vector<int32_t> col;
  int32_t row = 0;
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
std::string LocRel(std::istream& sas, bool swapit, int32_t col, int32_t row) {
  std::vector<int32_t> col_rel;
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

  if (!fColRel) out += "$";
  out += int_to_col(col_rel[0] + 1);

  if (!fRwRel) out += "$";
  out += std::to_string(row_rel + 1);

  return out;
}

std::string Area(std::istream& sas, bool swapit) {
  std::vector<int32_t> col0(3), col1(3);
  int32_t row0 = 0, row1 = 0;
  row0 = UncheckedRw(sas, swapit);  // rowFirst
  row1 = UncheckedRw(sas, swapit);  // rowLast
  col0 = ColRelShort(sas, swapit);  // columnFirst
  col1 = ColRelShort(sas, swapit);  // columnLast

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

std::string AreaRel(std::istream& sas, bool swapit, int32_t col, int32_t row) {
  std::vector<int32_t> col_rel0(3), col_rel1(3);
  int32_t row_rel0 = 0, row_rel1 = 0;
  row_rel0 = UncheckedRw(sas, swapit);  // rowFirst
  row_rel1 = UncheckedRw(sas, swapit);  // rowLast
  col_rel0 = ColRelShort(sas, swapit);  // columnFirst
  col_rel1 = ColRelShort(sas, swapit);  // columnLast

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

  if (!fColRel0) out += "$";
  out += int_to_col(col_rel0[0] + 1);

  if (!fRwRel0) out += "$";
  out += std::to_string(row_rel0 + 1);

  out += ":";

  if (!fColRel1) out += "$";
  out += int_to_col(col_rel1[0] + 1);

  if (!fRwRel1) out += "$";
  out += std::to_string(row_rel1 + 1);

  return out;
}

std::vector<int32_t> Cell(std::istream& sas, bool swapit) {
  std::vector<int32_t> out(3);

  out[0] = UncheckedCol(sas, swapit);

  int32_t uint = 0;
  uint = readbin(uint, sas, swapit);

  out[1] = uint & 0xFFFFFF;            // iStyleRef
  out[2] = (uint & 0x02000000) >> 24;  // fPhShow
  // unused

  return (out);
}

std::vector<std::string> dims_to_cells(int32_t firstRow, int32_t lastRow, int32_t firstCol, int32_t lastCol) {
  std::vector<int32_t> cols, rows;
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

std::vector<int32_t> brtColor(std::istream& sas, bool swapit) {
  uint8_t AB = 0, xColorType = 0, index = 0, bRed = 0, bGreen = 0, bBlue = 0, bAlpha = 0;
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

std::string as_border_style(int32_t style) {
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

std::string to_argb(int32_t a, int32_t r, int32_t g, int32_t b) {
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

  std::vector<int32_t> color = brtColor(sas, swapit);

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
  if (val & 0x02) {  // integer
    int32_t tmp = val >> 2;
    out = static_cast<double>(tmp);
  } else {  // double
    uint64_t tmp = static_cast<uint32_t>(val) & 0xfffffffcU;
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

std::string valType(uint8_t type) {

  if (type == 0x0) return "none";
  if (type == 0x1) return "whole";
  if (type == 0x2) return "decimal";
  if (type == 0x3) return "list";
  if (type == 0x4) return "date";
  if (type == 0x5) return "time";
  if (type == 0x6) return "textLength";
  if (type == 0x7) return "custom";

  return "unknown_type";
}

std::string typOperator(uint8_t oprtr) {
  if (oprtr == 0x0) return "between";
  if (oprtr == 0x1) return "notBetween";
  if (oprtr == 0x2) return "equal";
  if (oprtr == 0x3) return "notEqual";
  if (oprtr == 0x4) return "greaterThan";
  if (oprtr == 0x5) return "lessThan";
  if (oprtr == 0x6) return "greaterThanOrEqual";
  if (oprtr == 0x7) return "lessThanOrEqual";

  return "unknown_operator";
}

std::string grbitSgnOperator(int8_t oprtr) {

  if (oprtr == 0x01) return "lessThan";
  if (oprtr == 0x02) return "equal";
  if (oprtr == 0x03) return "lessThanOrEqual";
  if (oprtr == 0x04) return "greaterThan";
  if (oprtr == 0x05) return "notEqual";
  if (oprtr == 0x06) return "greaterThanOrEqual";

  return "unknown_operator";
}

std::vector<int32_t> Xti(std::istream& sas, bool swapit) {
  int32_t firstSheet = 0, lastSheet = 0;
  int32_t externalLink = 0;  // TODO actually uint32?
  externalLink = readbin(externalLink, sas, swapit);
  // scope
  // -2 workbook
  // -1 heet level
  // >= 0 sheet level
  firstSheet = readbin(firstSheet, sas, swapit);
  lastSheet = readbin(lastSheet, sas, swapit);

  // Rprintf("Xti: %d %d %d\n", externalLink, firstSheet, lastSheet);
  std::vector<int32_t> out(3);
  out[0] = externalLink;
  out[1] = firstSheet;
  out[2] = lastSheet;

  return out;
}

// first half little endian, second half big endian
std::string guid_str(const std::vector<int32_t>& guid_ints) {
  std::ostringstream guidStream;

  guidStream << std::uppercase << std::hex << std::setfill('0');

  guidStream << std::setw(2) << ((guid_ints[0] >> 24) & 0xFF)
             << std::setw(2) << ((guid_ints[0] >> 16) & 0xFF)
             << std::setw(2) << ((guid_ints[0] >> 8) & 0xFF)
             << std::setw(2) << (guid_ints[0] & 0xFF) << "-";

  guidStream << std::setw(2) << ((guid_ints[1] >> 8) & 0xFF)
             << std::setw(2) << (guid_ints[1] & 0xFF) << "-";

  guidStream << std::setw(2) << ((guid_ints[1] >> 24) & 0xFF)
             << std::setw(2) << ((guid_ints[1] >> 16) & 0xFF) << "-";

  guidStream << std::setw(2) << (guid_ints[2] & 0xFF)
             << std::setw(2) << ((guid_ints[2] >> 8) & 0xFF) << "-"
             << std::setw(2) << ((guid_ints[2] >> 16) & 0xFF)
             << std::setw(2) << ((guid_ints[2] >> 24) & 0xFF)
             << std::setw(2) << (guid_ints[3] & 0xFF)
             << std::setw(2) << ((guid_ints[3] >> 8) & 0xFF)
             << std::setw(2) << ((guid_ints[3] >> 16) & 0xFF)
             << std::setw(2) << ((guid_ints[3] >> 24) & 0xFF)
  ;

  return guidStream.str();
}

// bool isOperator(const std::string& token) {
//   return token == "+" || token == "-" || token == "*" || token == "/" ||
//     token == "^" || token == "%" || token == "="  || token == "&lt;&gt;" ||
//     token == "&lt;" || token == "&gt;" || token == "&le;" || token == "&ge;" ||
//     token == " " || token == "&amp;" || token == ":" || token == "," ||
//     token == "#" || token == "@";
// }

std::string array_elements(const std::vector<std::string>& elements, int32_t n, int32_t k) {
  std::stringstream ss;
  ss << "{";
  for (int32_t i = 0; i < n; ++i) {
    if (i > 0)
      ss << ";";
    for (int32_t j = 0; j < k; ++j) {
      if (j > 0)
        ss << ",";
      size_t index = static_cast<size_t>(i * k + j);
      if (index < elements.size()) {
        // check if it needs escaping
        ss << "\"";
        ss << escape_quote(elements[index]);
        ss << "\"";
      }
    }
  }
  ss << "}";
  return ss.str();
}

#include <stack>

// undo the newline escaping
std::string replaceXmlEscapesWithNewline(const std::string& input) {
    std::string output = input;
    std::string search1 = "&#xA;";
    std::string search2 = "&amp;#xA;";
    std::string replacement = "\n";

    size_t pos = 0;

    // Replace "&#xA;"
    while ((pos = output.find(search1, pos)) != std::string::npos) {
        output.replace(pos, search1.length(), replacement);
        pos += replacement.length();
    }

    pos = 0;

    // Replace "&amp;#xA;"
    while ((pos = output.find(search2, pos)) != std::string::npos) {
        output.replace(pos, search2.length(), replacement);
        pos += replacement.length();
    }

    return output;
}

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

  return replaceXmlEscapesWithNewline(parsedFormula);
}

std::string rgce(std::string fml_out, std::istream& sas, bool swapit, bool debug, int32_t col, int32_t row, int32_t &sharedFml, bool has_revision_record, std::streampos pos, std::vector<int32_t> &ptgextra) {

  int8_t val1 = 0;
  // std::vector<int32_t> ptgextra;
  if (debug) Rcpp::Rcout << ".";

  while (sas.tellg() < pos) {
    // 1
    uint8_t val2 = 0, controlbit = 0;
    val1 = readbin(val1, sas, swapit);

    controlbit = (val1 & 0x80) >> 7;
    if (controlbit != 0) Rcpp::warning("controlbit unexpectedly not 0");
    val1 &= 0x7F;  // the remaining 7 bits form ptg
    // for some Ptgs only the first 5 are of interest
    // and 6 and 7 contain DataType information
    if (debug) Rprintf("Formula: %d %d\n", val1, val2);

    switch (val1) {
      case PtgList_PtgSxName:
      case PtgAttr: {
        if (debug) Rcpp::Rcout << "reading eptg @ " << sas.tellg() << std::endl;
        val2 = readbin(val2, sas, swapit); // full 8 bit forming eptg
        if (debug) Rcpp::Rcout << "PtgAttr: " << std::hex << (int)val1 << ": "<< (int)val2 << std::dec << std::endl;
        switch(val2) {

          case PtgList: {
            RgbExtra typ = PtgExtraList;

            if (debug) Rcpp::Rcout << "PtgList " << sas.tellg() << std::endl;
            uint16_t ixti = 0, flags = 0;
            uint32_t listIndex = 0;
            uint16_t colFirst = 0, colLast = 0;

            // ixti = location of table
            ixti = readbin(ixti, sas, swapit);

            // B:
            // 0x00 columns consist of all columns in table
            // 0x01 one column wide, only colFirst required
            // 0x02 columns from colFirst to colLast
            // rowType = PtgRowType()
            // squareBracketSpace: spacing?
            // commaSpace: comma space?
            // unused
            // type: PtgDataType()
            // invalid: bool
            // nonresident: bool
            flags = readbin(flags, sas, swapit);

            // table identifier: unused if invalid=1 || nonresident=1
            listIndex = readbin(listIndex, sas, swapit);

            // cols: unused if invalid = 1 || nonresident = 1 || columns = 0
            colFirst  = ColShort(sas, swapit);
            colLast   = ColShort(sas, swapit);

            PtgListFields* fields = (PtgListFields*)&flags;

            // if (debug)
            // Rprintf("%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d\n",
            //   (uint32_t)fields->columns,
            //   (uint32_t)fields->commaSpace,
            //   (uint32_t)fields->invalid,
            //   (uint32_t)fields->nonresident,
            //   (uint32_t)fields->reserved2,
            //   (uint32_t)fields->rowType,
            //   (uint32_t)fields->squareBracketSpace,
            //   (uint32_t)fields->type,
            //   (uint32_t)fields->unused
            // );

            if (fields->nonresident)  // different workbook and invalid == 0
              ptgextra.push_back(typ);

            std::stringstream paddedStr;
            paddedStr << std::setw(12) << std::setfill('0') << listIndex;  // << ixti;

            // A1 notation cell
            // something like this: Table1[[#This Row],[a]]
            fml_out += "openxlsx2tab_" + paddedStr.str();

            bool no_row_type = fields->invalid == 1 || fields->nonresident == 1;

            fml_out += "[";

            bool need_bracket = fields->columns > 0 ||
                (fields->columns == 0 &&
                  (fields->rowType == dataheaders ||
                    fields->rowType == datatotals)
                );


            // if rowType == 0 no #Data etc is added
            if (!no_row_type && fields->rowType) {
              if (need_bracket) fml_out += "[";
              if (fields->rowType == data)        fml_out += "";
              if (fields->rowType == all)         fml_out += "#All";
              if (fields->rowType == headers)     fml_out += "#Headers";
              if (fields->rowType == data2)       fml_out += "#Data";
              if (fields->rowType == dataheaders) fml_out += "#Headers],[#Data";
              if (fields->rowType == totals)      fml_out += "#Totals";
              if (fields->rowType == datatotals)  fml_out += "#Data],[#Totals";
              if (fields->rowType == current)     fml_out += "#This Row";
              if (need_bracket) fml_out += "]";
              if (fields->columns > 0) fml_out += ",";
            }

            // not sure what is supposed to happen in this case?
            // have to replace colFirst with a variable name
            if (!(fields->invalid == 1 || fields->nonresident == 1 || fields->columns == 0)) {
              // Rcpp::Rcout << "colFirst" << std::endl;
              if (fields->columns > 1 || fields->rowType > data) fml_out += "[";
              fml_out += "openxlsx2col_";
              fml_out += std::to_string(listIndex);
              fml_out += "_";
              fml_out += std::to_string(colFirst);
              if (fields->columns > 1 || fields->rowType > data) fml_out  += "]";
            }

            // have to replace colLast with a variable name
            if ((colFirst < colLast) && !(fields->invalid == 1 || fields->nonresident == 1 || fields->columns == 0)) {
              // Rcpp::Rcout << "colLast" << std::endl;
              fml_out += ":[openxlsx2col_";
              fml_out += std::to_string(listIndex);
              fml_out += "_";
              fml_out += std::to_string(colLast);
              if (fields->columns > 1 || fields->rowType > data) fml_out  += "]";
            }

            fml_out += "]";
            fml_out += "\n";

            // Do something with this, just ... what?
            if (debug)
              Rprintf("PtgList: %d, %d, %d, %d\n", ixti, listIndex, colFirst, colLast);

            // if (debug)
            // Rcpp::warning("formulas with table references are not implemented.");

            break;
          }

          case PtgAttrSemi: {
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
          case PtgAttrGoTo: {
            if (debug) Rcpp::Rcout << "PtgAttrIf" <<std::endl;

            uint16_t offset = 0;

            offset = readbin(offset, sas, swapit);

            break;
          }

          case PtgAttrChoose: {
            if (debug) Rcpp::Rcout << "PtgAttrChoose" <<std::endl;

            uint16_t cOffset = 0, rgOffset = 0;

            cOffset = readbin(cOffset, sas, swapit);

            // some offsets, currently unused
            for (int16_t off = 0; off < (cOffset + 1); ++off) {
              rgOffset = readbin(rgOffset, sas, swapit);
              if (debug) Rcpp::Rcout << rgOffset << std::endl;
            }

            break;
          }
            // end control tokens

          case PtgAttrSpace: {
            uint8_t type = 0, cch = 0;

            if (((val2 >> 6) & 1) != 1) Rcpp::stop("wrong value");

            // PtgAttrSpaceType
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

          case PtgAttrSpaceSemi: {
            // type: A PtgAttrSpaceType
            uint8_t type = 0, cch = 0;
            type = readbin(type, sas, swapit);
            cch = readbin(cch, sas, swapit);

            if (debug) Rprintf("PtgAttrSpaceSemi: %d %d\n", type, cch);
            break;
          }

          case PtgAttrSum: {
            if (debug) Rcpp::Rcout << "PtgAttrSum" << std::endl;
            // val2[1:4] == 0
            // val2[5]   == 1
            // val2[6:8] == 0
            uint16_t unused = 0;
            unused = readbin(unused, sas, swapit);
            // Rcpp::Rcout << unused << std::endl;
            fml_out += "SUM(%s)";  // maybe attr because it is a single cell function?
            fml_out += "\n";
            break;
          }

          case PtgAttrBaxcel:
          case PtgAttrBaxcel2: {
            // val1[8]   == bitSemi if Rgce is volatile
            // val2[1:4] == 0
            // val2[5]   == 1
            // val2[6:8] == 0

            uint16_t unused = 0;
            unused = readbin(unused, sas, swapit);
            Rcpp::warning("PtgAttrBaxcel: unhandled formula thing");
            break;
          }

          default: {
            Rprintf("Undefined Formula_TWO: %d %d\n", val1, val2);
            break;
          }
        }

        break;
      }

      case PtgRange: {
        if (debug) Rcpp::Rcout << ":" <<std::endl;
        fml_out += "%s:%s";
        fml_out += "\n";
        break;
      }

      case PtgUnion: {
        if (debug) Rcpp::Rcout << "," <<std::endl;
        fml_out += "%s,%s";
        fml_out += "\n";
        break;
      }

      case PtgIsect: {
        if (debug) Rcpp::Rcout << " " <<std::endl;
        fml_out += "%s %s";
        fml_out += "\n";
        break;
      }

      case PtgAdd: {
        if (debug) Rcpp::Rcout << "+" <<std::endl;
        fml_out += "%s+%s";
        fml_out += "\n";
        break;
      }

      case PtgSub: {
        if (debug) Rcpp::Rcout << "-" <<std::endl;
        fml_out += "%s-%s";
        fml_out += "\n";
        break;
      }

      case PtgMul: {
        if (debug) Rcpp::Rcout << "*" <<std::endl;
        fml_out += "%s*%s";
        fml_out += "\n";
        break;
      }

      case PtgDiv: {
        if (debug) Rcpp::Rcout << "/" <<std::endl;
        fml_out += "%s/%s";
        fml_out += "\n";
        break;
      }

      case PtgPercent: {
        if (debug) Rcpp::Rcout << "%" <<std::endl;
        fml_out += "%s%";
        fml_out += "\n";
        break;
      }

      case PtgPower: {
        if (debug) Rcpp::Rcout << "^" <<std::endl;
        fml_out += "%s^%s";
        fml_out += "\n";
        break;
      }

      case PtgConcat: {
        if (debug) Rcpp::Rcout << "&" <<std::endl;
        fml_out += "%s&amp;%s";
        fml_out += "\n";
        break;
      }

      case PtgEq: {
        if (debug) Rcpp::Rcout << "=" <<std::endl;
        fml_out += "%s=%s";
        fml_out += "\n";
        break;
      }

      case PtgGt: {
        if (debug) Rcpp::Rcout << ">" <<std::endl;
        fml_out += "%s&gt;%s";
        fml_out += "\n";
        break;
      }

      case PtgGe: {
        if (debug) Rcpp::Rcout << ">=" <<std::endl;
        fml_out += "%s&gt;=%s";
        fml_out += "\n";
        break;
      }

      case PtgLt: {
        if (debug) Rcpp::Rcout << "<" <<std::endl;
        fml_out += "%s&lt;%s";
        fml_out += "\n";
        break;
      }

      case PtgLe: {
        if (debug) Rcpp::Rcout << "<=" <<std::endl;
        fml_out += "%s&lt;=%s";
        fml_out += "\n";
        break;
      }

      case PtgNe: {
        if (debug) Rcpp::Rcout << "!=" <<std::endl;
        fml_out += "%s&lt;&gt;%s";
        fml_out += "\n";
        break;
      }

      case PtgUPlus: {
        if (debug) Rcpp::Rcout << "+val" <<std::endl;
        fml_out += "+%s";
        fml_out += "\n";
        break;
      }

      case PtgUMinus: {
        if (debug) Rcpp::Rcout << "-val" <<std::endl;
        fml_out += "-%s";
        fml_out += "\n";
        break;
      }

      case PtgParen: {
        if (debug) Rcpp::Rcout << "()" <<std::endl;
        fml_out += "(%s)";
        fml_out += "\n";
        break;
      }

      case PtgMissArg: {
        if (debug) Rcpp::Rcout << "MISSING()" <<std::endl;
        fml_out += " ";
        fml_out += "\n";
        break;
      }

      case PtgInt: {
        if (debug) Rcpp::Rcout << "PtgInt" <<std::endl;
        uint16_t integer = 0;
        integer = readbin(integer, sas, swapit);
        // Rcpp::Rcout << integer << std::endl;
        fml_out += std::to_string(integer);
        fml_out += "\n";
        break;
      }

      case PtgNum: {
        if (debug) Rcpp::Rcout << "PtgNum" <<std::endl;
        double value = Xnum(sas, swapit);
        fml_out += std::to_string(value);
        fml_out += "\n";
        break;
      }

      case PtgStr: {
        if (debug) Rcpp::Rcout << "PtgStr" <<std::endl;

        std::string esc_quote = PtrStr(sas, swapit);
        // Rcpp::Rcout << esc_quote << std::endl;
        if (esc_quote == "\n") fml_out += '\"'; // + escape_quote(esc_quote) + "\"";
        else fml_out += "\"" + escape_xml(escape_quote(esc_quote)) + "\"";
        fml_out += "\n";

        break;
      }

      case PtgArray:
      case PtgArray2:
      case PtgArray3: {
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
      case PtgRef3: {
        if (debug) Rcpp::Rcout << "PtgRef" <<std::endl;
        // uint8_t ptg8 = 0, ptg = 0, PtgDataType = 0, null = 0;
        //
        // 2
        // Rprintf("PtgRef2: %d, %d, %d\n", ptg, PtgDataType, null);

        // val1[6:7] == PtgDataType

        fml_out += Loc(sas, swapit);
        fml_out += "\n";

        if (debug) Rcpp::Rcout << sas.tellg() << std::endl;

        break;
      }

      case PtgRef3d:
      case PtgRef3d2:
      case PtgRef3d3: {
        if (debug) Rcpp::Rcout << "PtgRef3d" <<std::endl;
        // need_ptg_revextern = true;
        RgbExtra typ = RevExtern;
        if (has_revision_record)
          ptgextra.push_back(typ);

        uint16_t ixti = 0;
        if (debug) Rprintf("XtiIndex: %d\n", ixti);
        ixti = readbin(ixti, sas, swapit);  // XtiIndex

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
      case PtgRefN3: {
        if (debug) Rcpp::Rcout << "PtgRefN" <<std::endl;

        // A1 notation cell
        fml_out += LocRel(sas, swapit, col, row);
        fml_out += "\n";

        break;
      }

      case PtgArea:
      case PtgArea2:
      case PtgArea3: {
        if (debug) Rcpp::Rcout << "PtgArea" <<std::endl;

        // A1 notation cell
        fml_out += Area(sas, swapit);
        fml_out += "\n";

        break;
      }

      case PtgArea3d:
      case PtgArea3d2:
      case PtgArea3d3: {
        if (debug) Rcpp::Rcout << "PtgArea3d" <<std::endl;

        // need_ptg_revextern = true;
        RgbExtra typ = RevExtern;
        if (has_revision_record)
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
      case PtgAreaN3: {
        if (debug) Rcpp::Rcout << "PtgAreaN" <<std::endl;

        // A1 notation cell
        fml_out += AreaRel(sas, swapit, col, row);
        fml_out += "\n";

        break;
      }

      case PtgMemArea:
      case PtgMemArea2:
      case PtgMemArea3: {
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
      case PtgName3: {
        if (debug) Rcpp::Rcout << "PtgName" <<std::endl;

        // need_ptg_revnametabid = true;
        RgbExtra typ = RevNameTabid;
        if (has_revision_record)
          ptgextra.push_back(typ);

        uint32_t nameindex = 0;
        nameindex = readbin(nameindex, sas, swapit);
        // Rcpp::Rcout << nameindex << std::endl;
        // fml_out += std::to_string(nameindex);

        std::stringstream paddedStr;
        paddedStr << std::setw(12) << std::setfill('0') << nameindex;

        fml_out += "openxlsx2defnam_" + paddedStr.str();
        fml_out += "\n";

        break;
      }

      case PtgNameX:
      case PtgNameX2:
      case PtgNameX3: {
        if (debug) Rcpp::Rcout << "PtgNameX" <<std::endl;

        // need_ptg_revname = true;
        RgbExtra typ = RevName;
        if (has_revision_record)
          ptgextra.push_back(typ);

        // not yet found
        uint16_t ixti = 0;
        uint32_t nameindex = 0;
        ixti = readbin(ixti, sas, swapit);
        nameindex = readbin(nameindex, sas, swapit);
        // Rcpp::Rcout << nameindex << std::endl;
        // fml_out += std::to_string(ixti);
        // fml_out += std::to_string(nameindex);

        // copied from above, does this work?
        std::stringstream paddedStr;
        paddedStr << std::setw(12) << std::setfill('0') << nameindex;

        fml_out += "openxlsx2defnam_" + paddedStr.str();
        fml_out += "\n";

        break;
      }

      case PtgRefErr:
      case PtgRefErr2:
      case PtgRefErr3: {
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
      case PtgRefErr3d3: {
        if (debug) Rcpp::Rcout << "PtgRefErr3d" <<std::endl;

        // need_ptg_revextern = true;
        RgbExtra typ = RevExtern;
        if (has_revision_record)
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

      case PtgAreaErr:
      case PtgAreaErr2:
      case PtgAreaErr3: {
        if (debug) Rcpp::Rcout << "PtgAreaErr" <<std::endl;

        uint32_t unused1 = 0, unused2 = 0, unused3 = 0;
        unused1 = readbin(unused1, sas, swapit);
        unused2 = readbin(unused2, sas, swapit);
        unused3 = readbin(unused3, sas, swapit);

        // could not reproduce this locally
        fml_out += "#REF!";
        fml_out += "\n";

        if (debug) Rcpp::Rcout << sas.tellg() << std::endl;

        break;
      }

      case PtgAreaErr3d:
      case PtgAreaErr3d2:
      case PtgAreaErr3d3: {
        if (debug) Rcpp::Rcout << "PtgAreaErr3d" <<std::endl;

        // need_ptg_revextern = true;
        RgbExtra typ = RevExtern;
        if (has_revision_record)
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
      case PtgFunc3: {
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
      case PtgFuncVar3: {
        if (debug) Rcpp::Rcout << "PtgFuncVar" <<std::endl;

        uint8_t cparams = 0, fCeFunc = 0;  // number of parameters
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

      case PtgErr: {
        if (debug) Rcpp::Rcout << "PtgErr" <<std::endl;

        fml_out += BErr(sas, swapit);
        fml_out += "\n";
        break;
      }

      case PtgBool: {
        if (debug) Rcpp::Rcout << "PtgBool" <<std::endl;

        int8_t boolean = 0;
        boolean = readbin(boolean, sas, swapit);
        fml_out += std::to_string(boolean);
        fml_out += "\n";
        break;
      }

      case PtgExp: {
        if (debug) Rcpp::Rcout << "PtgExp" <<std::endl;

        // this is a reference to the cell that contains the shared formula
        int32_t ptg_row = UncheckedRw(sas, swapit) + 1;
        if (debug) Rcpp::Rcout << "PtgExp: " << ptg_row << std::endl;
        sharedFml = ptg_row;
        break;
      }

      case PtgMemFunc:
      case PtgMemFunc2:
      case PtgMemFunc3: {
        if (debug) Rcpp::Rcout << "PtgMemFunc" <<std::endl;

        uint16_t cce = 0;
        cce = readbin(cce, sas, swapit);  // count of bytes in the binary reference expression
        break;
      }

      case PtgMemErr:
      case PtgMemErr2:
      case PtgMemErr3: {
        if (debug) Rcpp::Rcout << "PtgMemErr" <<std::endl;

        uint8_t unused1 = 0;
        uint16_t unused2 = 0, cce = 0;

        // Error txt is already in the value field
        // fml_out += BErr(sas, swapit);
        // fml_out += "\n";

        std::string err_str = "";
        err_str = BErr(sas, swapit);

        if (debug) Rcpp::Rcout << "PtgMemErr: " << err_str << std::endl;

        unused1 = readbin(unused1, sas, swapit);
        unused2 = readbin(unused2, sas, swapit);
        cce = readbin(cce, sas, swapit);  // count of bytes in the binary reference expression

        break;
      }

      case PtgMemNoMem:
      case PtgMemNoMem2:
      case PtgMemNoMem3: {
        if (debug) Rcpp::Rcout << "PtgMemNoMem" <<std::endl;

        uint32_t unused = 0;
        unused = readbin(unused, sas, swapit);

        uint16_t cce = 0;
        cce = readbin(cce, sas, swapit);  // count of bytes in the binary reference expression

        break;
      }

      default: {
        // if (debug)
        Rcpp::warning("Undefined Formula: %d %d\n", val1, val2);
        break;
      }
    }
  }

  if (sas.tellg() != pos) {
    // somethings not correct
    Rcpp::Rcout << "[fml] unexpected position when parsing head" << std::endl;
    sas.seekg(pos, sas.beg);
  }

  return fml_out;
}

std::string rgcb(std::string fml_out, std::istream& sas, bool swapit, bool debug, int32_t col, int32_t row, int32_t &sharedFml, bool has_revision_record, std::streampos pos, std::vector<int32_t> &ptgextra) {

  int8_t val1 = 0;
  // std::vector<int32_t> ptgextra;
  // RgbExtra
  for (size_t cntr = 0; cntr < ptgextra.size(); ++cntr) {
    val1 = static_cast<int8_t>(ptgextra[cntr]);

    if (debug) Rcpp::Rcout << cntr << ": " << (int32_t)val1 << std::endl;

    switch (val1) {
      case PtgExtraCol: {
        // need_ptg_extra_col = true;
        if (debug) Rcpp::Rcout << "PtgExtraCol" << std::endl;
        int32_t ptg_col = UncheckedCol(sas, swapit);
        if (debug) Rcpp::Rcout << "cb PtgExp: " << int_to_col(ptg_col+1) << std::endl;

        fml_out += int_to_col(ptg_col + 1);
        fml_out += std::to_string(row);
        fml_out += "\n";
        break;
      }

      case PtgExtraArray: {
        if (debug) Rcpp::Rcout << "PtgExtraArray" << std::endl;

        int32_t rows = 0, cols = 0;
        // actually its DRw() and DCol(), but it does not matter?
        rows = UncheckedRw(sas, swapit);
        cols = UncheckedCol(sas, swapit);

        std::string array = "";
        std::vector<std::string> array_elems;  // (cols*rows);

        if (debug) Rcpp::Rcout << rows << ": " << cols << std::endl;

        // number of elements in row order: must be equal to rows * cols
        for (int32_t rr = 0; rr < rows; ++rr) {
          for (int32_t cc = 0; cc < cols; ++cc) {
            // blob (it is actually called this way)
            uint8_t reserved = 0;
            reserved = readbin(reserved, sas, swapit);

            if (debug) Rcpp::Rcout << (int32_t)reserved << std::endl;

            // SerBool
            if (reserved == 0x02) {
              if (debug) Rcpp::Rcout << "SerBool" << std::endl;
              uint8_t f = 0;
              f = readbin(f, sas, swapit);

              if (debug) Rcpp::Rcout << (int32_t)f << std::endl;
              array_elems.push_back(std::to_string((int32_t)f));
            }

            // SerErr
            if (reserved == 0x04) {
              if (debug) Rcpp::Rcout << "SerErr" << std::endl;
              uint8_t reserved2 = 0;
              uint16_t reserved3 = 0;
              std::string strerr = BErr(sas, swapit);
              reserved2 = readbin(reserved2, sas, swapit);
              reserved3 = readbin(reserved3, sas, swapit);

              if (debug) Rcpp::Rcout << strerr << std::endl;
              array_elems.push_back(strerr);
            }

            // SerNum
            if (reserved == 0x00) {
              if (debug) Rcpp::Rcout << "SerNum" << std::endl;
              double xnum = 0.0;
              xnum = Xnum(sas, swapit);

              std::stringstream stream;
              stream << std::setprecision(16) << xnum;

              if (debug) Rcpp::Rcout << xnum << std::endl;
              array_elems.push_back(stream.str());
            }

            // SerStr
            if (reserved == 0x01) {
              if (debug) Rcpp::Rcout << "SerStr" << std::endl;
              uint16_t cch = 0;
              cch = readbin(cch, sas, swapit);
              std::string rgch(cch, '\0');
              rgch = read_xlwidestring(rgch, sas);

              if (debug) Rcpp::Rcout << rgch << std::endl;
              array_elems.push_back(rgch);
            }
          }
        }

        array += array_elements(array_elems, rows, cols);

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

      case PtgExtraMem: {
        // not sure what this is good for
        if (debug) Rcpp::Rcout << "PtgExtraMem: " << (int32_t)val1 << std::endl;

        int32_t count = 0;
        count = readbin(count, sas, swapit);

        for (int32_t cnt = 0; cnt < count; ++cnt) {
          std::vector<int32_t> ucrfx = UncheckedRfX(sas, swapit);
        }

        break;
      }

      case RevNameTabid: {
        // Rcpp::stop("Skip");
        if (debug) Rcpp::Rcout << "RevNameTabid: " << (int32_t)val1 << std::endl;
        sas.seekg(pos, sas.beg);
        break;
      }

      case RevName: {
        // Rcpp::stop("Skip");
        if (debug) Rcpp::Rcout << "RevName: " << (int32_t)val1 << std::endl;
        sas.seekg(pos, sas.beg);
        break;
      }

      case PtgExtraList: {
        // Rcpp::stop("Skip");
        if (debug) Rcpp::Rcout << "PtgExtraList: " << (int32_t)val1 << std::endl;
        sas.seekg(pos, sas.beg);
        break;
      }

      case RevExtern: {  // do i need this?
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

      default: {
        // Rcpp::stop("Skip");
        Rcpp::Rcout << "undefined cb: " << (int32_t)val1 << std::endl;
        sas.seekg(pos, sas.beg);
        break;
      }
    }
  }

  if (sas.tellg() != pos) {
    // somethings not correct
    sas.seekg(pos, sas.beg);
  }

  return fml_out;
}

std::string CellParsedFormula(std::istream& sas, bool swapit, bool debug, int32_t col, int32_t row, int32_t &sharedFml, bool has_revision_record) {
  // bool ptg_extra_array = false;
  uint32_t cce = 0, cb = 0;
  std::vector<int32_t> ptgextra;

  if (debug) Rcpp::Rcout << "CellParsedFormula: " << sas.tellg() << std::endl;

  cce = readbin(cce, sas, swapit);
  if (cce >= 16385) Rcpp::stop("wrong cce size");
  if (debug) Rcpp::Rcout << "cce: " << cce << std::endl;

  std::streampos pos = sas.tellg();
  // sas.seekg(cce, sas.cur);
  pos += cce;

  std::string fml_out;

  fml_out = rgce(fml_out, sas, swapit, debug, col, row, sharedFml, has_revision_record, pos, ptgextra);

  cb = readbin(cb, sas, swapit);  // is there a control bit, even if CB is empty?

  if (debug)
    Rcpp::Rcout << "cb: " << cb << std::endl;

  pos = sas.tellg();
  // sas.seekg(cce, sas.cur);

  pos += cb;

  if (debug) Rcpp::Rcout << ".";
  if (debug) {
    // Rprintf("Formula cb: %d\n", val1);
    Rprintf("%d: %d\n", (int)sas.tellg(), (int)pos);
  }

  if (debug) Rcpp::Rcout << "--- formula ---\n" << fml_out << std::endl;

  fml_out = rgcb(fml_out, sas, swapit, debug, col, row, sharedFml, has_revision_record, pos, ptgextra);

  if (debug) {
    Rcpp::Rcout << "...fml..." << std::endl;
    Rcpp::Rcout << fml_out << std::endl;
  }

  std::string inflix = parseRPN(fml_out);

  return inflix;
}

std::string FRTParsedFormula(std::istream& sas, bool swapit, bool debug, int32_t col, int32_t row, int32_t &sharedFml, bool has_revision_record) {
  // bool ptg_extra_array = false;
  uint32_t cce = 0, cb = 0;
  std::vector<int32_t> ptgextra;

  if (debug) Rcpp::Rcout << "CellParsedFormula: " << sas.tellg() << std::endl;

  cce = readbin(cce, sas, swapit);
  if (cce >= 16385) Rcpp::stop("wrong cce size");
  if (debug) Rcpp::Rcout << "cce: " << cce << std::endl;

  cb = readbin(cb, sas, swapit);  // is there a control bit, even if CB is empty?

  std::streampos pos = sas.tellg();
  pos += cce;

  std::string fml_out;

  fml_out = rgce(fml_out, sas, swapit, debug, col, row, sharedFml, has_revision_record, pos, ptgextra);

  if (debug)
    Rcpp::Rcout << "cb: " << cb << std::endl;

  pos = sas.tellg();
  pos += cb;

  if (debug) Rcpp::Rcout << ".";
  if (debug) {
    Rprintf("%d: %d\n", (int)sas.tellg(), (int)pos);
  }

  if (debug) Rcpp::Rcout << "--- formula ---\n" << fml_out << std::endl;

  fml_out = rgcb(fml_out, sas, swapit, debug, col, row, sharedFml, has_revision_record, pos, ptgextra);

  if (debug) {
    Rcpp::Rcout << "...fml..." << std::endl;
    Rcpp::Rcout << fml_out << std::endl;
  }

  std::string inflix = parseRPN(fml_out);

  return inflix;
}

#endif
