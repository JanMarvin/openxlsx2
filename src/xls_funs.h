#ifndef XLS_FUNS_H
#define XLS_FUNS_H

#include <cstdint>
#include <cstdio>
#include <cstring>
#include <iomanip>
#include <limits>
#include <map>
#include <sstream>
#include <string>
#include <utility>
#include <vector>

#include <Rcpp.h>

#include "xlsb_defines.h"
#include "xlsb_flags.h"
#include "xlsb_funs.h"
#include "xls_defines.h"

// Every BIFF8 record is: id (2 bytes), size (2 bytes), data. Returns false at the
// end of the stream. Padding of fewer than 4 bytes behind the last record is the end.
inline uint16_t* xls_last_records() {
  static uint16_t last[2] = {0, 0};
  return last;
}

inline std::string xls_cut_off() {
  char buf[64];
  std::snprintf(buf, sizeof(buf), "xls: string is cut off (record 0x%04X, followed by 0x%04X)",
                static_cast<unsigned>(xls_last_records()[0]), static_cast<unsigned>(xls_last_records()[1]));
  return std::string(buf);
}

inline bool XLS_RECORD(uint16_t& rid, uint16_t& rsize, std::istream& sas, bool swapit) {
  if (sas.peek() == std::char_traits<char>::eof()) return false;
  uint16_t id = 0, size = 0;
  if (!sas.read(reinterpret_cast<char*>(&id), sizeof(id)) || !sas.read(reinterpret_cast<char*>(&size), sizeof(size))) {
    sas.clear();
    sas.seekg(0, std::ios_base::end);
    return false;
  }
  rid = swapit ? swap_endian(id) : id;
  rsize = swapit ? swap_endian(size) : size;
  if (rid != XlsContinue) {
    xls_last_records()[0] = xls_last_records()[1];
    xls_last_records()[1] = rid;
  }
  return true;
}

// Strings in a BIFF8 file can be longer than a record. The rest then continues in
// a Continue record. For these functions rec_end is the position where the
// current record ends. Outside of the SST pass xls_no_end.
static const std::streampos xls_no_end = std::streampos(std::streamoff(1) << 60);

// skip n bytes, step over record borders
inline void xls_skip(std::istream& sas, bool swapit, size_t n, std::streampos& rec_end) {
  while (n > 0) {
    if (sas.tellg() >= rec_end) {
      uint16_t rid = 0, rsize = 0;
      if (!XLS_RECORD(rid, rsize, sas, swapit) || rid != XlsContinue) Rcpp::stop(xls_cut_off());
      rec_end = sas.tellg() + std::streamoff(rsize);
      continue;
    }
    size_t avail = static_cast<size_t>(rec_end - sas.tellg());
    size_t step = std::min(n, avail);
    sas.seekg(static_cast<std::streamoff>(step), std::ios::cur);
    n -= step;
  }
}

// read cch characters. At a record border the next record starts with a new flags
// byte, bit 0 tells if the remaining characters are 8 or 16 bit wide.
inline std::u16string xls_chars(std::istream& sas, bool swapit, size_t cch, bool wide, std::streampos& rec_end) {
  std::u16string out;
  out.reserve(cch);
  while (out.size() < cch) {
    if (sas.tellg() >= rec_end) {
      uint16_t rid = 0, rsize = 0;
      if (!XLS_RECORD(rid, rsize, sas, swapit) || rid != XlsContinue) Rcpp::stop(xls_cut_off());
      rec_end = sas.tellg() + std::streamoff(rsize);
      uint8_t flags = 0;
      flags = readbin(flags, sas, swapit);
      wide = (flags & 0x01) != 0;
      continue;
    }
    size_t avail = static_cast<size_t>(rec_end - sas.tellg());
    size_t fit = wide ? avail / 2 : avail;
    if (fit == 0) Rcpp::stop(xls_cut_off());
    size_t n = std::min(cch - out.size(), fit);
    for (size_t i = 0; i < n; ++i) {
      if (wide) {
        uint16_t c = 0;
        c = readbin(c, sas, swapit);
        out.push_back(static_cast<char16_t>(c));
      } else {
        uint8_t c = 0;
        c = readbin(c, sas, swapit);
        out.push_back(static_cast<char16_t>(c));
      }
    }
  }
  return out;
}

// XLUnicodeString (16 bit length) and ShortXLUnicodeString (8 bit length):
// cch, flags, [runs], [ext size], characters, run data, ext data. The flags tell
// if the characters are 16 bit (0x01), have phonetic data (0x04) or formatting
// runs (0x08). Formatting runs and phonetic data are skipped.
inline std::string XLUnicodeString(std::istream& sas, bool swapit, bool len16, std::streampos& rec_end) {
  size_t cch = 0;
  if (len16) {
    uint16_t n = 0;
    n = readbin(n, sas, swapit);
    cch = n;
  } else {
    uint8_t n = 0;
    n = readbin(n, sas, swapit);
    cch = n;
  }

  uint8_t flags = 0;
  flags = readbin(flags, sas, swapit);

  uint16_t runs = 0;
  uint32_t ext = 0;
  if (flags & 0x08) runs = readbin(runs, sas, swapit);
  if (flags & 0x04) ext = readbin(ext, sas, swapit);

  std::u16string chars = xls_chars(sas, swapit, cch, (flags & 0x01) != 0, rec_end);
  xls_skip(sas, swapit, static_cast<size_t>(runs) * 4 + ext, rec_end);

  return to_utf8(chars);
}

inline std::string XLUnicodeString(std::istream& sas, bool swapit, bool len16) {
  std::streampos rec_end = xls_no_end;
  return XLUnicodeString(sas, swapit, len16, rec_end);
}

// Like XLUnicodeString with a 16 bit length, but keeps the characters and the
// formatting runs: the first character of a run and the font it uses.
inline std::u16string XLUnicodeRichString(std::istream& sas, bool swapit, std::streampos& rec_end,
                                          std::vector<std::pair<uint16_t, uint16_t> >& runs) {
  uint16_t cch = 0;
  cch = readbin(cch, sas, swapit);
  uint8_t flags = 0;
  flags = readbin(flags, sas, swapit);

  uint16_t crun = 0;
  uint32_t ext = 0;
  if (flags & 0x08) crun = readbin(crun, sas, swapit);
  if (flags & 0x04) ext = readbin(ext, sas, swapit);

  std::u16string chars = xls_chars(sas, swapit, cch, (flags & 0x01) != 0, rec_end);

  runs.clear();
  for (uint16_t i = 0; i < crun; ++i) {
    if (sas.tellg() >= rec_end) {
      uint16_t rid = 0, rsize = 0;
      if (!XLS_RECORD(rid, rsize, sas, swapit) || rid != XlsContinue) Rcpp::stop(xls_cut_off());
      rec_end = sas.tellg() + std::streamoff(rsize);
    }
    uint16_t ich = 0, ifnt = 0;
    ich = readbin(ich, sas, swapit);
    ifnt = readbin(ifnt, sas, swapit);
    runs.push_back(std::make_pair(ich, ifnt));
  }
  xls_skip(sas, swapit, ext, rec_end);
  return chars;
}

inline std::string xls_error(uint8_t code) {
  switch (code) {
    case 0x00: return "#NULL!";
    case 0x07: return "#DIV/0!";
    case 0x0F: return "#VALUE!";
    case 0x17: return "#REF!";
    case 0x1D: return "#NAME?";
    case 0x24: return "#NUM!";
    case 0x2A: return "#N/A";
    default: return "#VALUE!";
  }
}

// number as text that reads back as the same double
inline std::string xls_num(double d) {
  std::string s;
  for (int prec = 15; prec <= 17; prec += 2) {
    std::ostringstream os;
    os.imbue(std::locale::classic());
    os << std::setprecision(prec) << d;
    s = os.str();
    std::istringstream is(s);
    is.imbue(std::locale::classic());
    double back = 0;
    is >> back;
    if (back == d) break;
  }
  return s;
}

// Defined names, sheet names and 3D references need to know the sheets of the
// workbook. This is what the formula decoder gets to see.
// The formula of a cell, name or rule: the tokens (rgce, cce bytes) and the data
// that goes with some of them (rgcb). The only data are the constants of arrays.
struct XlsRgce {
  std::vector<uint8_t> data;
  size_t cce = 0;
  XlsRgce() {}
  explicit XlsRgce(const std::vector<uint8_t>& tokens) : data(tokens), cce(tokens.size()) {}
};

struct XlsFormulaContext {
  std::vector<std::string> xti;    // text in front of a 3D reference, e.g. "Sheet1!". Empty: not supported
  std::vector<std::string> names;  // defined names, PtgName counts from 1
  std::vector<std::vector<std::string> > xnames;  // names of external workbooks and add-in functions for each XTI, counting from 1
  uint32_t row = 0;                // the cell of the formula, relevant for shared formulas
  uint32_t col = 0;
  int unsupported = -1;            // set to the token that could not be decoded
};

// cell reference of a BIFF8 formula: row (2 bytes) and column (low 8 bits of 2
// bytes) with the flags column relative (0x4000) and row relative (0x8000)
inline std::string xls_cell(uint16_t row, uint16_t col, bool relative_to_cell, uint32_t cell_row, uint32_t cell_col) {
  bool col_rel = (col & 0x4000) != 0;
  bool row_rel = (col & 0x8000) != 0;

  uint32_t r = row;
  uint32_t c = col & 0x00FF;
  if (relative_to_cell) {
    // shared formulas store offsets to the cell they are used in
    if (row_rel) r = (cell_row + static_cast<uint32_t>(static_cast<int16_t>(row))) & 0xFFFF;
    if (col_rel) c = (cell_col + static_cast<uint32_t>(static_cast<int8_t>(c))) & 0xFF;
  }

  std::string out;
  if (!col_rel) out += "$";
  out += int_to_col(c + 1);
  if (!row_rel) out += "$";
  out += std::to_string(r + 1);
  return out;
}

inline std::string xls_loc(std::istream& sas, bool swapit, bool relative_to_cell, const XlsFormulaContext& ctx) {
  uint16_t row = 0, col = 0;
  row = readbin(row, sas, swapit);
  col = readbin(col, sas, swapit);
  return xls_cell(row, col, relative_to_cell, ctx.row, ctx.col);
}

// area: first row, last row, first column, last column. Whole columns (A:A) and
// whole rows (1:1) are covered by all rows or all columns.
inline std::string xls_area(std::istream& sas, bool swapit, bool relative_to_cell, const XlsFormulaContext& ctx) {
  uint16_t row1 = 0, row2 = 0, col1 = 0, col2 = 0;
  row1 = readbin(row1, sas, swapit);
  row2 = readbin(row2, sas, swapit);
  col1 = readbin(col1, sas, swapit);
  col2 = readbin(col2, sas, swapit);

  if (!relative_to_cell && row1 == 0 && row2 == 0xFFFF) {
    std::string first = std::string((col1 & 0x4000) ? "" : "$") + int_to_col((col1 & 0x00FF) + 1);
    std::string last = std::string((col2 & 0x4000) ? "" : "$") + int_to_col((col2 & 0x00FF) + 1);
    return first + ":" + last;
  }

  if (!relative_to_cell && (col1 & 0x00FF) == 0 && (col2 & 0x00FF) == 0x00FF) {
    std::string first = std::string((col1 & 0x8000) ? "" : "$") + std::to_string(row1 + 1);
    std::string last = std::string((col2 & 0x8000) ? "" : "$") + std::to_string(row2 + 1);
    return first + ":" + last;
  }

  return xls_cell(row1, col1, relative_to_cell, ctx.row, ctx.col) + ":" +
         xls_cell(row2, col2, relative_to_cell, ctx.row, ctx.col);
}

// Decodes the token stream (rgce) of a formula. The tokens are in reverse polish
// notation: operands come before their operator. Like the xlsb decoder this writes
// one line per token and lets parseRPN() put the formula together. %s is where
// an operand from the stack goes. Returns an empty string for tokens that are not
// supported, then ctx.unsupported tells which one.
inline std::string xls_rgce(const XlsRgce& rgce, XlsFormulaContext& ctx, bool swapit, bool debug) {
  std::istringstream sas(std::string(rgce.data.begin(), rgce.data.end()), std::ios::in | std::ios::binary);
  std::string out;
  std::streampos end = static_cast<std::streampos>(rgce.cce);
  // the data of arrays follows the tokens, one array after the other
  std::streampos array_data = end;

  while (sas.tellg() < end) {
    uint8_t ptg = 0;
    ptg = readbin(ptg, sas, swapit);

    // The upper bits of tokens from 0x20 on tell the data class (reference,
    // value, array). This does not matter for the text of the formula.
    if (ptg >= 0x20) ptg = static_cast<uint8_t>((ptg & 0x1F) | 0x20);
    if (debug) Rcpp::Rcout << "Ptg: " << std::hex << static_cast<int>(ptg) << std::dec << std::endl;

    switch (ptg) {
      case PtgAdd: out += "%s+%s\n"; break;
      case PtgSub: out += "%s-%s\n"; break;
      case PtgMul: out += "%s*%s\n"; break;
      case PtgDiv: out += "%s/%s\n"; break;
      case PtgPower: out += "%s^%s\n"; break;
      case PtgConcat: out += "%s&%s\n"; break;
      case PtgLt: out += "%s<%s\n"; break;
      case PtgLe: out += "%s<=%s\n"; break;
      case PtgEq: out += "%s=%s\n"; break;
      case PtgGe: out += "%s>=%s\n"; break;
      case PtgGt: out += "%s>%s\n"; break;
      case PtgNe: out += "%s<>%s\n"; break;
      case PtgIsect: out += "%s %s\n"; break;
      case PtgUnion: out += "%s,%s\n"; break;
      case PtgRange: out += "%s:%s\n"; break;
      case PtgUPlus: out += "+%s\n"; break;
      case PtgUMinus: out += "-%s\n"; break;
      case PtgPercent: out += "%s%\n"; break;
      case PtgParen: out += "(%s)\n"; break;
      case PtgMissArg: out += " \n"; break;

      case PtgStr: {
        std::streampos rec_end = xls_no_end;
        std::string text = XLUnicodeString(sas, swapit, false, rec_end);
        out += "\"" + escape_quote(text) + "\"\n";
        break;
      }

      case PtgErr: {
        out += BErr(sas, swapit) + "\n";
        break;
      }

      case PtgBool: {
        uint8_t value = 0;
        value = readbin(value, sas, swapit);
        out += value ? "TRUE\n" : "FALSE\n";
        break;
      }

      case PtgInt: {
        uint16_t value = 0;
        value = readbin(value, sas, swapit);
        out += std::to_string(value) + "\n";
        break;
      }

      case PtgNum: {
        double value = 0;
        value = readbin(value, sas, swapit);
        std::ostringstream num;
        num.imbue(std::locale::classic());
        num << std::setprecision(15) << value;
        out += num.str() + "\n";
        break;
      }

      case PtgAttr: {
        // Hints for the calculation: how to skip the unused branch of IF and
        // CHOOSE, where to put blanks. The only one that is part of the formula
        // is the sum of a single argument.
        uint8_t attr = 0;
        attr = readbin(attr, sas, swapit);
        if (attr == PtgAttrChoose) {
          uint16_t count = 0;
          count = readbin(count, sas, swapit);
          sas.seekg(static_cast<std::streamoff>(count + 1) * 2, std::ios::cur);
        } else {
          sas.seekg(2, std::ios::cur);
        }
        if (attr == PtgAttrSum) out += "SUM(%s)\n";
        break;
      }

      case PtgFunc: {
        // functions with a fixed number of arguments
        uint16_t iftab = 0;
        iftab = readbin(iftab, sas, swapit);
        out += Ftab(iftab) + "\n";
        break;
      }

      case PtgFuncVar: {
        uint8_t cparams = 0;
        uint16_t iftab = 0;
        cparams = readbin(cparams, sas, swapit);
        iftab = readbin(iftab, sas, swapit);
        cparams &= 0x7F;
        // the top bit of the function index marks a command equivalent (macro) function
        bool fCeFunc = (iftab & 0x8000) != 0;
        iftab &= 0x7FFF;

        std::string fml;
        if (iftab == 0x00FF) {
          // user defined function: the first argument is the name of the function
          fml = "%s(";
          for (uint8_t i = 1; i < cparams; ++i) fml += i == 1 ? "%s" : ",%s";
          fml += ")";
        } else {
          fml = (fCeFunc ? Cetab(iftab) : Ftab(iftab)) + "(";
          if (cparams) fml += "%s";
          for (uint8_t i = 1; i < cparams; ++i) fml += ",%s";
          fml += ")";
        }
        out += fml + "\n";
        break;
      }

      case PtgName: {
        // index of a defined name. 2 bytes, followed by 2 unused bytes
        uint16_t index = 0;
        uint16_t unused = 0;
        index = readbin(index, sas, swapit);
        unused = readbin(unused, sas, swapit);
        if (index == 0 || index > ctx.names.size()) out += "#NAME?\n";
        else out += ctx.names[index - 1] + "\n";
        break;
      }

      case PtgRef: out += xls_loc(sas, swapit, false, ctx) + "\n"; break;
      case PtgRefN: out += xls_loc(sas, swapit, true, ctx) + "\n"; break;
      case PtgArea: out += xls_area(sas, swapit, false, ctx) + "\n"; break;
      case PtgAreaN: out += xls_area(sas, swapit, true, ctx) + "\n"; break;

      case PtgRef3d:
      case PtgArea3d: {
        // the sheet is an index into the XTI table of the ExternSheet record
        uint16_t ixti = 0;
        ixti = readbin(ixti, sas, swapit);
        std::string ref = ptg == PtgRef3d ? xls_loc(sas, swapit, false, ctx) : xls_area(sas, swapit, false, ctx);
        if (ixti < ctx.xti.size() && !ctx.xti[ixti].empty()) out += ctx.xti[ixti] + ref + "\n";
        else out += "#REF!\n";
        break;
      }

      case PtgNameX: {
        uint16_t ixti = 0;
        uint32_t index = 0;
        ixti = readbin(ixti, sas, swapit);
        index = readbin(index, sas, swapit);
        if (ixti < ctx.xnames.size() && index > 0 && index <= ctx.xnames[ixti].size()) out += ctx.xnames[ixti][index - 1] + "\n";
        else out += "#NAME?\n";
        break;
      }

      case PtgRefErr: sas.seekg(4, std::ios::cur); out += "#REF!\n"; break;
      case PtgAreaErr: sas.seekg(8, std::ios::cur); out += "#REF!\n"; break;
      case PtgRefErr3d: sas.seekg(6, std::ios::cur); out += "#REF!\n"; break;
      case PtgAreaErr3d: sas.seekg(10, std::ios::cur); out += "#REF!\n"; break;

      case PtgArray: {
        // the 7 bytes of the token are unused. The values are after the tokens: the
        // number of columns and rows, then the values row by row.
        sas.seekg(7, std::ios::cur);
        std::streampos back = sas.tellg();
        sas.seekg(array_data, std::ios::beg);

        uint8_t cols = 0;
        uint16_t rows = 0;
        cols = readbin(cols, sas, swapit);
        rows = readbin(rows, sas, swapit);
        size_t ncols = static_cast<size_t>(cols) + 1, nrows = static_cast<size_t>(rows) + 1;

        std::string arr = "{";
        for (size_t r = 0; r < nrows; ++r) {
          if (r > 0) arr += ";";
          for (size_t c = 0; c < ncols; ++c) {
            if (c > 0) arr += ",";
            uint8_t kind = 0;
            kind = readbin(kind, sas, swapit);
            if (kind == 0x01) {
              double value = 0;
              value = readbin(value, sas, swapit);
              arr += xls_num(value);
            } else if (kind == 0x02) {
              // a text: characters (16 bit), flags, characters
              std::string text = XLUnicodeString(sas, swapit, true);
              arr += "\"" + escape_quote(text) + "\"";
            } else if (kind == 0x04) {
              uint8_t b = 0;
              b = readbin(b, sas, swapit);
              sas.seekg(7, std::ios::cur);
              arr += b ? "TRUE" : "FALSE";
            } else if (kind == 0x10) {
              arr += BErr(sas, swapit);
              sas.seekg(7, std::ios::cur);
            } else if (kind == 0x00) {
              sas.seekg(8, std::ios::cur);
            } else {
              ctx.unsupported = ptg;
              return "";
            }
          }
        }
        arr += "}";

        array_data = sas.tellg();
        sas.seekg(back, std::ios::beg);
        out += arr + "\n";
        break;
      }

      case PtgMemArea:
      case PtgMemErr:
      case PtgMemNoMem: {
        // a cache for what follows, the formula itself continues after it
        sas.seekg(6, std::ios::cur);
        break;
      }

      case PtgMemFunc: {
        sas.seekg(2, std::ios::cur);
        break;
      }

      default: {
        ctx.unsupported = ptg;
        return "";
      }
    }
  }

  std::string formula = parseRPN(out);

  // Excel saves functions that were new in Excel 2007 with a prefix in xls files.
  // In xlsx files these functions have no prefix.
  static const char* const functions2007[5] = {"IFERROR", "SUMIFS", "COUNTIFS", "AVERAGEIF", "AVERAGEIFS"};
  for (size_t i = 0; i < 5; ++i) {
    std::string prefixed = std::string("_xlfn.") + functions2007[i] + "(";
    size_t pos = 0;
    while ((pos = formula.find(prefixed, pos)) != std::string::npos) formula.erase(pos, 6);
  }

  return formula;
}

// ---------------------------------------------------------------------------
// conditional formatting
// ---------------------------------------------------------------------------

// 0xAARRGGBB as text. The color is an index into the palette: 0 to 7 are fixed,
// 8 to 63 are the palette, 65 is the window background.
inline std::string xls_argb(const std::vector<uint32_t>& palette, uint32_t icv) {
  uint32_t rgb = 0;
  if (icv < 8) rgb = XlsDefaultPalette[icv];
  else if (icv < 64) rgb = palette[icv - 8];
  else if (icv == 65) rgb = 0xFFFFFF;
  std::ostringstream os;
  os << "FF" << std::uppercase << std::hex << std::setw(6) << std::setfill('0') << rgb;
  return os.str();
}

// formats that are not stored in the file ([MS-XLS] 2.5.165, same as ECMA-376 18.8.30)
inline const char* xls_builtin_format(uint32_t id) {
  switch (id) {
    case 1: return "0";
    case 2: return "0.00";
    case 3: return "#,##0";
    case 4: return "#,##0.00";
    case 9: return "0%";
    case 10: return "0.00%";
    case 11: return "0.00E+00";
    case 12: return "# ?/?";
    case 13: return "# ?\?/?\?";
    case 14: return "mm-dd-yy";
    case 15: return "d-mmm-yy";
    case 16: return "d-mmm";
    case 17: return "mmm-yy";
    case 18: return "h:mm AM/PM";
    case 19: return "h:mm:ss AM/PM";
    case 20: return "h:mm";
    case 21: return "h:mm:ss";
    case 22: return "m/d/yy h:mm";
    case 37: return "#,##0 ;(#,##0)";
    case 38: return "#,##0 ;[Red](#,##0)";
    case 39: return "#,##0.00;(#,##0.00)";
    case 40: return "#,##0.00;[Red](#,##0.00)";
    case 45: return "mm:ss";
    case 46: return "[h]:mm:ss";
    case 47: return "mmss.0";
    case 48: return "##0.0E+0";
    case 49: return "@";
    default: return nullptr;
  }
}

// CFColor: how the color is referenced (1 palette, 2 RGB, 3 theme), the value and the tint
struct XlsCfColor {
  uint32_t type = 0, value = 0;
  double tint = 0;
};

// The attributes of a color element: the color of a palette, RGB (r, g, b, a) or a theme color.
// Colors that are not set or automatic return an empty string.
inline std::string xls_color_attr(const XlsCfColor& col, const std::vector<uint32_t>& palette) {
  std::string x;
  if (col.type == 1) {
    x = " rgb=\"" + xls_argb(palette, col.value) + "\"";
  } else if (col.type == 2) {
    char hex[16];
    std::snprintf(hex, sizeof(hex), "FF%02X%02X%02X", col.value & 0xFFu, (col.value >> 8) & 0xFFu, (col.value >> 16) & 0xFFu);
    x = std::string(" rgb=\"") + hex + "\"";
  } else if (col.type == 3) {
    x = " theme=\"" + std::to_string(col.value) + "\"";
  } else {
    return x;
  }
  if (col.tint != 0) x += " tint=\"" + xls_num(col.tint) + "\"";
  return x;
}

// ExtProp structures of XFExt and CFEx: type, size (including these 4 bytes) and the
// data. Colors (FullColorExt) are the type of the color (2 bytes), the tint (2 bytes,
// -32768 to 32767 for -1 to 1) and the value (4 bytes), followed by 8 unused bytes.
// Fills the colors by type and returns the indent (type 0x0F), 0 if not set.
inline void xls_extprops(const std::vector<uint8_t>& rec, size_t pos, size_t count,
                         std::map<uint16_t, XlsCfColor>& colors, uint16_t& indent) {
  auto u16 = [&](size_t at) -> uint32_t { return at + 1 < rec.size() ? static_cast<uint32_t>(rec[at] | (rec[at + 1] << 8)) : 0u; };
  for (size_t i = 0; i < count; ++i) {
    if (pos + 4 > rec.size()) return;
    uint16_t type = static_cast<uint16_t>(u16(pos));
    size_t cb = u16(pos + 2);
    if (cb < 4 || pos + cb > rec.size()) return;
    if ((type == 0x04 || type == 0x05 || (type >= 0x07 && type <= 0x0B) || type == 0x0D) && cb >= 12) {
      XlsCfColor c;
      c.type = u16(pos + 4);
      c.tint = static_cast<int16_t>(u16(pos + 6)) / 32767.0;
      c.value = u16(pos + 8) | (u16(pos + 10) << 16);
      colors[type] = c;
    } else if (type == 0x0F && cb >= 6) {
      indent = static_cast<uint16_t>(u16(pos + 4));
    }
    pos += cb;
  }
}

// The formatting of a conditional format (DXFN, [MS-XLS] 2.5.95) as a dxf. A DXFN
// starts with flags: for the values "ninch" (ignore) bits, and bits that tell which
// of the parts follow. The parts have a fixed order: number format, font,
// alignment, border, pattern, protection.
// rec holds the record, pos is the start of the DXFN and moves to its end.
// Returns false if the record is too short.
inline bool xls_dxfn(const std::vector<uint8_t>& rec, size_t& pos, const std::vector<uint32_t>& palette,
                     const std::map<uint16_t, std::string>& formats, int numfmt_id, std::string& dxf,
                     const std::map<uint16_t, XlsCfColor>* ext = nullptr) {
  auto u8 = [&](size_t at) -> uint32_t { return at < rec.size() ? rec[at] : 0u; };
  auto u16 = [&](size_t at) -> uint32_t { return u8(at) | (u8(at + 1) << 8); };
  auto u32 = [&](size_t at) -> uint32_t { return u16(at) | (u16(at + 2) << 16); };

  if (pos + 6 > rec.size()) return false;
  uint32_t flags = u32(pos);
  uint32_t flags2 = u16(pos + 4);
  pos += 6;

  // flags: ninch bits 0 to 20, then which parts follow
  bool has_num = (flags >> 25) & 1;
  bool has_font = (flags >> 26) & 1;
  bool has_align = (flags >> 27) & 1;
  bool has_border = (flags >> 28) & 1;
  bool has_pattern = (flags >> 29) & 1;
  bool has_protect = (flags >> 30) & 1;
  bool num_is_text = (flags2 & 0x0001) != 0;

  std::string font, num, fill, border;
  auto full_color = [&](uint16_t type, const std::string& palette_attr) {
    if (ext != nullptr) {
      auto it = ext->find(type);
      if (it != ext->end()) {
        std::string attr = xls_color_attr(it->second, palette);
        if (!attr.empty()) return attr;
      }
    }
    return palette_attr;
  };

  if (has_num) {
    if (num_is_text) {
      // size, then the format as text
      size_t cch = u16(pos + 2);
      uint8_t wide = static_cast<uint8_t>(u8(pos + 4));
      std::u16string s;
      size_t at = pos + 5;
      for (size_t i = 0; i < cch; ++i) {
        if (wide & 1) {
          s.push_back(static_cast<char16_t>(u16(at)));
          at += 2;
        } else {
          s.push_back(static_cast<char16_t>(u8(at)));
          at += 1;
        }
      }
      pos = at;
      if (!((flags >> 19) & 1)) {
        num = "<numFmt numFmtId=\"" + std::to_string(numfmt_id) + "\" formatCode=\"" + escape_xml(to_utf8(s)) + "\"/>";
      }
    } else {
      uint32_t ifmt = u8(pos + 1);
      pos += 2;
      if (!((flags >> 19) & 1)) {
        std::string code;
        auto it = formats.find(static_cast<uint16_t>(ifmt));
        if (it != formats.end()) code = it->second;
        else if (xls_builtin_format(ifmt) != nullptr) code = xls_builtin_format(ifmt);
        if (!code.empty()) num = "<numFmt numFmtId=\"" + std::to_string(ifmt) + "\" formatCode=\"" + escape_xml(code) + "\"/>";
      }
    }
  }

  if (has_font) {
    // name (1 + 63 bytes), then 16 bytes of attributes, then 9 values of 4 bytes and the index
    size_t f = pos;
    uint32_t ts = u32(f + 64 + 4);
    int bls = static_cast<int16_t>(u16(f + 64 + 8));
    int sss = static_cast<int16_t>(u16(f + 64 + 10));
    uint32_t uls = u8(f + 64 + 12);
    int32_t icv = static_cast<int32_t>(u32(f + 64 + 16));
    uint32_t ts_ninch = u32(f + 64 + 24);
    uint32_t sss_ninch = u32(f + 64 + 28);
    uint32_t uls_ninch = u32(f + 64 + 32);
    uint32_t bls_ninch = u32(f + 64 + 36);
    pos += 118;

    if (!bls_ninch && bls == 0x02BC) font += "<b/>";
    if (!(ts_ninch & 0x02) && (ts & 0x02)) font += "<i/>";
    if (!(ts_ninch & 0x80) && (ts & 0x80)) font += "<strike/>";
    if (!uls_ninch && uls != 0xFF && uls != 0) {
      if (uls == 0x01) font += "<u/>";
      else if (uls == 0x02) font += "<u val=\"double\"/>";
      else if (uls == 0x21) font += "<u val=\"singleAccounting\"/>";
      else if (uls == 0x22) font += "<u val=\"doubleAccounting\"/>";
    }
    if (!sss_ninch && sss == 1) font += "<vertAlign val=\"superscript\"/>";
    if (!sss_ninch && sss == 2) font += "<vertAlign val=\"subscript\"/>";
    if (icv >= 0 && icv != 0x7FFF && icv < 66) font += "<color" + full_color(0x0D, " rgb=\"" + xls_argb(palette, static_cast<uint32_t>(icv)) + "\"") + "/>";
  }

  if (has_align) pos += 8;

  if (has_border) {
    uint32_t b0 = u16(pos), b1 = u16(pos + 2), b2 = u32(pos + 4);
    pos += 8;

    // style (4 bits each) and color (7 bits each) of the four sides; a side is
    // only written if its ninch bit (10 to 13) is not set
    uint32_t style[4] = {b0 & 0x0F, (b0 >> 4) & 0x0F, (b0 >> 8) & 0x0F, (b0 >> 12) & 0x0F};
    uint32_t color[4] = {b1 & 0x7F, (b1 >> 7) & 0x7F, b2 & 0x7F, (b2 >> 7) & 0x7F};
    static const char* const sides[4] = {"left", "right", "top", "bottom"};
    static const uint16_t side_ext[4] = {0x09, 0x0A, 0x07, 0x08};
    for (int k = 0; k < 4; ++k) {
      if ((flags >> (10 + k)) & 1) continue;
      border += std::string("<") + sides[k];
      if (style[k] == 0 || style[k] >= 14) {
        border += "/>";
      } else {
        border += std::string(" style=\"") + XlsBorderStyles[style[k]] + "\"><color" + full_color(side_ext[k], " rgb=\"" + xls_argb(palette, color[k]) + "\"") +
                  "/></" + sides[k] + ">";
      }
    }
  }

  if (has_pattern) {
    uint32_t p = u32(pos);
    pos += 4;
    uint32_t fls = (p >> 10) & 0x3F, fg = (p >> 16) & 0x7F, bg = (p >> 23) & 0x7F;
    std::string inner;
    if (!((flags >> 17) & 1)) inner += "<fgColor" + full_color(0x04, " rgb=\"" + xls_argb(palette, fg) + "\"") + "/>";
    if (!((flags >> 18) & 1)) inner += "<bgColor" + full_color(0x05, " rgb=\"" + xls_argb(palette, bg) + "\"") + "/>";
    std::string type;
    if (!((flags >> 16) & 1) && fls < 19) type = std::string(" patternType=\"") + XlsFillPatterns[fls] + "\"";
    if (!type.empty() || !inner.empty()) fill = "<fill><patternFill" + type + ">" + inner + "</patternFill></fill>";
  }

  if (has_protect) pos += 2;

  dxf = "<dxf>";
  if (!font.empty()) dxf += "<font>" + font + "</font>";
  dxf += num + fill;
  if (!border.empty()) dxf += "<border>" + border + "</border>";
  dxf += "</dxf>";

  return pos <= rec.size();
}

// CF record ([MS-XLS] 2.4.42): kind of the rule, comparison, size of two formulas,
// the formatting, and the two formulas.
struct XlsCfRule {
  uint8_t type = 0;        // 1 value of the cell, 2 formula
  uint8_t comparison = 0;  // for type 1
  std::vector<uint8_t> rgce1, rgce2;
  std::string dxf;
};

inline bool xls_cf(const std::vector<uint8_t>& rec, const std::vector<uint32_t>& palette,
                   const std::map<uint16_t, std::string>& formats, int numfmt_id, XlsCfRule& cf,
                   const std::map<uint16_t, XlsCfColor>* ext = nullptr) {
  if (rec.size() < 6) return false;
  cf.type = rec[0];
  cf.comparison = rec[1];
  size_t cce1 = static_cast<size_t>(rec[2] | (rec[3] << 8));
  size_t cce2 = static_cast<size_t>(rec[4] | (rec[5] << 8));

  size_t pos = 6;
  if (!xls_dxfn(rec, pos, palette, formats, numfmt_id, cf.dxf, ext)) return false;
  if (pos + cce1 + cce2 > rec.size()) return false;

  cf.rgce1.assign(rec.begin() + static_cast<std::ptrdiff_t>(pos), rec.begin() + static_cast<std::ptrdiff_t>(pos + cce1));
  cf.rgce2.assign(rec.begin() + static_cast<std::ptrdiff_t>(pos + cce1), rec.begin() + static_cast<std::ptrdiff_t>(pos + cce1 + cce2));
  return true;
}

// A value of a color scale, data bar or icon set (CFVO): how a threshold is calculated.
struct XlsCfvo {
  uint8_t type = 0;  // 1 number, 2 minimum, 3 maximum, 4 percent, 5 percentile, 7 formula
  std::vector<uint8_t> rgce;
  double num = 0;
  bool equal = true;  // icon sets: does a value equal to the threshold pass it
};

// CF12: the rules of Excel 2007 and later. Rules on a formula or a value have
// a dxf like the CF record. The others have color scales, data bars or icon sets.
struct XlsCf12Rule {
  uint8_t type = 0;  // 1 value, 2 formula, 3 color scale, 4 data bar, 5 filter, 6 icon set
  uint8_t comparison = 0;
  std::vector<uint8_t> rgce1, rgce2;
  std::string dxf;
  bool has_dxf = false;
  bool stop_if_true = false;
  uint16_t priority = 0;
  uint16_t icf_template = 0;

  std::vector<XlsCfvo> cfvos;
  std::vector<XlsCfColor> colors;
  uint8_t icon_set = 0;
  bool show_value = true, reverse = false, right_to_left = false;
  uint8_t min_length = 10, max_length = 90;
};

inline bool xls_cf12(const std::vector<uint8_t>& rec, const std::vector<uint32_t>& palette,
                     const std::map<uint16_t, std::string>& formats, int numfmt_id, XlsCf12Rule& cf) {
  size_t n = rec.size();
  auto u8 = [&](size_t at) -> uint32_t { return at < n ? rec[at] : 0u; };
  auto u16 = [&](size_t at) -> uint32_t { return u8(at) | (u8(at + 1) << 8); };
  auto u32 = [&](size_t at) -> uint32_t { return u16(at) | (u16(at + 2) << 16); };
  auto f64 = [&](size_t at) -> double {
    uint64_t v = 0;
    for (size_t i = 0; i < 8; ++i) v |= static_cast<uint64_t>(u8(at + i)) << (8 * i);
    double d = 0;
    std::memcpy(&d, &v, sizeof(d));
    return d;
  };

  // the header of 12 bytes, type, comparison and the sizes of the two formulas
  if (n < 12 + 6) return false;
  size_t pos = 12;
  cf.type = static_cast<uint8_t>(u8(pos));
  cf.comparison = static_cast<uint8_t>(u8(pos + 1));
  size_t cce1 = u16(pos + 2), cce2 = u16(pos + 4);
  pos += 6;

  // DXFN12: the size of the formatting, 0 is followed by 2 unused bytes
  uint32_t cb_dxf = u32(pos);
  pos += 4;
  if (cb_dxf == 0) {
    pos += 2;
  } else {
    size_t at = pos;
    if (!xls_dxfn(rec, at, palette, formats, numfmt_id, cf.dxf)) return false;
    // the extension (XFExtNoFRT): 3 reserved words, the number of properties and the properties
    if (at + 8 <= pos + cb_dxf && pos + cb_dxf <= n) {
      std::map<uint16_t, XlsCfColor> colors;
      uint16_t indent = 0;
      xls_extprops(rec, at + 8, u16(at + 6), colors, indent);
      size_t again = pos;
      if (!colors.empty()) xls_dxfn(rec, again, palette, formats, numfmt_id, cf.dxf, &colors);
    }
    cf.has_dxf = true;
    pos += cb_dxf;
  }

  if (pos + cce1 + cce2 > n) return false;
  cf.rgce1.assign(rec.begin() + static_cast<std::ptrdiff_t>(pos), rec.begin() + static_cast<std::ptrdiff_t>(pos + cce1));
  pos += cce1;
  cf.rgce2.assign(rec.begin() + static_cast<std::ptrdiff_t>(pos), rec.begin() + static_cast<std::ptrdiff_t>(pos + cce2));
  pos += cce2;

  // the formula that tells if the rule is active
  size_t cce_active = u16(pos);
  pos += 2 + cce_active;

  uint32_t flags = u8(pos);
  cf.stop_if_true = (flags & 0x02) != 0;
  cf.priority = static_cast<uint16_t>(u16(pos + 1));
  cf.icf_template = static_cast<uint16_t>(u16(pos + 3));
  size_t cb_template = u8(pos + 5);
  pos += 6 + cb_template;
  if (pos > n) return false;

  auto read_cfvo = [&](XlsCfvo& v) {
    v.type = static_cast<uint8_t>(u8(pos));
    size_t cce = u16(pos + 1);
    pos += 3;
    if (pos + cce > n) return false;
    v.rgce.assign(rec.begin() + static_cast<std::ptrdiff_t>(pos), rec.begin() + static_cast<std::ptrdiff_t>(pos + cce));
    pos += cce;
    if (cce == 0 && v.type != 2 && v.type != 3) {
      v.num = f64(pos);
      pos += 8;
    }
    return pos <= n;
  };
  auto read_color = [&](XlsCfColor& c) {
    c.type = u32(pos);
    c.value = u32(pos + 4);
    c.tint = f64(pos + 8);
    pos += 16;
    return pos <= n;
  };

  if (cf.type == 3) {
    // color scale: unused (2), reserved (1), number of control points (2 x 1),
    // flags (1), then the thresholds (each with a position 0 to 1), then the colors
    size_t count = u8(pos + 3);
    pos += 6;
    for (size_t i = 0; i < count; ++i) {
      XlsCfvo v;
      if (!read_cfvo(v)) return false;
      pos += 8;
      cf.cfvos.push_back(v);
    }
    for (size_t i = 0; i < count; ++i) {
      XlsCfColor c;
      pos += 8;
      if (!read_color(c)) return false;
      cf.colors.push_back(c);
    }
  } else if (cf.type == 4) {
    // data bar: unused (2), reserved (1), flags (1), shortest and longest bar in
    // percent, the color, then the two thresholds
    uint32_t f = u8(pos + 3);
    cf.right_to_left = (f & 0x01) != 0;
    cf.show_value = (f & 0x02) == 0;
    cf.min_length = static_cast<uint8_t>(u8(pos + 4));
    cf.max_length = static_cast<uint8_t>(u8(pos + 5));
    pos += 6;
    XlsCfColor c;
    if (!read_color(c)) return false;
    cf.colors.push_back(c);
    for (size_t i = 0; i < 2; ++i) {
      XlsCfvo v;
      if (!read_cfvo(v)) return false;
      cf.cfvos.push_back(v);
    }
  } else if (cf.type == 6) {
    // icon set: unused (2), reserved (1), number of icons, the set, flags (1),
    // then the thresholds with a flag if the threshold itself passes it
    size_t count = u8(pos + 3);
    cf.icon_set = static_cast<uint8_t>(u8(pos + 4));
    uint32_t f = u8(pos + 5);
    cf.show_value = (f & 0x01) == 0;
    cf.reverse = (f & 0x04) != 0;
    pos += 6;
    for (size_t i = 0; i < count; ++i) {
      XlsCfvo v;
      if (!read_cfvo(v)) return false;
      v.equal = u8(pos) != 0;
      pos += 5;
      cf.cfvos.push_back(v);
    }
  }
  return true;
}

// ---------------------------------------------------------------------------
// hyperlinks
// ---------------------------------------------------------------------------

// What a HLink record holds. A link either points to somewhere outside of the
// workbook (url or file, `target`) or to a place inside of it (`location`). A url
// can have both: http://example.com/page#anchor.
struct XlsHyperlink {
  uint16_t row1 = 0, row2 = 0, col1 = 0, col2 = 0;
  std::string target;
  std::string location;
  std::string display;
};

// HLink record (MS-XLS 2.4.140): the range, the GUID of the hyperlink object, a
// version, flags that tell which of the optional parts follow, and the parts in
// a fixed order: display name, frame name, moniker (the target), location.
// Strings are UTF-16 with a leading character count (including the 0 at the end).
inline XlsHyperlink xls_hlink(const std::vector<uint8_t>& rec) {
  XlsHyperlink h;
  size_t pos = 0;

  auto need = [&](size_t n) {
    if (pos + n > rec.size()) Rcpp::stop("xls: hyperlink record is cut off");
  };

  // u16 and u32 are little endian in the file, whatever the machine is
  auto get16 = [&]() {
    need(2);
    uint16_t v = static_cast<uint16_t>(rec[pos] | (rec[pos + 1] << 8));
    pos += 2;
    return v;
  };
  auto get32 = [&]() {
    need(4);
    uint32_t v = static_cast<uint32_t>(rec[pos]) | (static_cast<uint32_t>(rec[pos + 1]) << 8) |
                 (static_cast<uint32_t>(rec[pos + 2]) << 16) | (static_cast<uint32_t>(rec[pos + 3]) << 24);
    pos += 4;
    return v;
  };

  // a count of characters, then the characters. The 0 at the end is removed.
  auto get_counted_string = [&]() {
    uint32_t cch = get32();
    need(static_cast<size_t>(cch) * 2);
    std::u16string s;
    for (uint32_t i = 0; i < cch; ++i) s.push_back(static_cast<char16_t>(get16()));
    while (!s.empty() && s.back() == u'\0') s.pop_back();
    return to_utf8(s);
  };

  h.row1 = get16();
  h.row2 = get16();
  h.col1 = get16();
  h.col2 = get16();

  need(16);  // GUID of the hyperlink object
  pos += 16;
  get32();   // stream version
  uint32_t flags = get32();

  bool has_moniker = (flags & 0x01) != 0;
  bool has_location = (flags & 0x08) != 0;
  bool has_display = (flags & 0x10) != 0;
  bool has_guid = (flags & 0x20) != 0;
  bool has_time = (flags & 0x40) != 0;
  bool has_frame = (flags & 0x80) != 0;
  bool moniker_is_string = (flags & 0x100) != 0;

  if (has_display) h.display = get_counted_string();
  if (has_frame) get_counted_string();

  if (has_moniker) {
    if (moniker_is_string) {
      h.target = get_counted_string();
    } else {
      // the first 4 bytes of the GUID tell the kind of moniker
      need(16);
      uint32_t kind = static_cast<uint32_t>(rec[pos]) | (static_cast<uint32_t>(rec[pos + 1]) << 8) |
                      (static_cast<uint32_t>(rec[pos + 2]) << 16) | (static_cast<uint32_t>(rec[pos + 3]) << 24);
      pos += 16;

      if (kind == 0x79EAC9E0) {
        // url moniker: size in bytes, the url as UTF-16 with 0 at the end, and
        // for some urls 24 more bytes (GUID, version, flags)
        uint32_t len = get32();
        need(len);
        size_t end = pos + len;
        std::u16string s;
        while (pos + 2 <= end) {
          char16_t c = static_cast<char16_t>(get16());
          if (c == u'\0') break;
          s.push_back(c);
        }
        pos = end;
        h.target = to_utf8(s);
      } else if (kind == 0x00000303) {
        // file moniker: number of ".." in front, the path as ansi, 24 bytes we do
        // not use, and the path as UTF-16 which is preferred if it exists
        uint16_t up = get16();
        uint32_t ansi_len = get32();
        need(ansi_len);
        std::string ansi(reinterpret_cast<const char*>(&rec[pos]), ansi_len);
        pos += ansi_len;
        while (!ansi.empty() && ansi.back() == '\0') ansi.pop_back();
        need(24);
        pos += 24;
        uint32_t unicode_size = get32();
        std::string path = ansi;
        if (unicode_size > 0) {
          uint32_t unicode_bytes = get32();
          get16();  // key value
          need(unicode_bytes);
          std::u16string s;
          for (uint32_t i = 0; i + 2 <= unicode_bytes; i += 2) s.push_back(static_cast<char16_t>(get16()));
          while (!s.empty() && s.back() == u'\0') s.pop_back();
          path = to_utf8(s);
        }
        std::string prefix;
        for (uint16_t i = 0; i < up; ++i) prefix += "../";
        h.target = prefix + path;
        for (size_t i = 0; i < h.target.size(); ++i)
          if (h.target[i] == '\\') h.target[i] = '/';
      } else {
        // a kind of moniker we do not know: the link is kept without a target
        return h;
      }
    }
  }

  if (has_location) h.location = get_counted_string();

  // the GUID and the creation time follow, they are not needed
  (void)has_guid;
  (void)has_time;

  return h;
}

// ---------------------------------------------------------------------------
// tables
// ---------------------------------------------------------------------------

struct XlsTableColumn {
  std::string name;
  uint32_t total_function = 0;  // 0 none, 1 average, 2 count, 3 count numbers, 4 max, 5 min, 6 sum, 7 stdDev, 8 var
  std::string total_label;
};

// A table of a sheet (since Excel 2003 "list"). Only tables of the plain kind
// (source type 0, the data is in the cells) are read.
struct XlsListTable {
  uint32_t id = 0;
  std::string name;
  uint16_t row1 = 0, row2 = 0, col1 = 0, col2 = 0;
  bool header = true;
  bool totals = false;
  bool auto_filter = false;
  std::vector<XlsTableColumn> columns;

  // List12 record, table style
  std::string style;
  bool first_column = false, last_column = false, row_stripes = false, column_stripes = false;
};

// Feature11 and Feature12 record (MS-XLS 2.4.114, 2.4.115) with a TableFeatureType
// (2.5.269). Returns false if the table can not be read: a different kind of table
// or a part of the record that is not supported.
inline bool xls_table(const std::vector<uint8_t>& rec, XlsListTable& t) {
  size_t pos = 0;
  bool ok = true;

  // a read past the end of the record makes the table unusable
  auto need = [&](size_t n) {
    if (pos + n > rec.size()) ok = false;
    return ok;
  };
  auto get16 = [&]() -> uint16_t {
    if (!need(2)) return 0;
    uint16_t v = static_cast<uint16_t>(rec[pos] | (rec[pos + 1] << 8));
    pos += 2;
    return v;
  };
  auto get32 = [&]() -> uint32_t {
    if (!need(4)) return 0;
    uint32_t v = static_cast<uint32_t>(rec[pos]) | (static_cast<uint32_t>(rec[pos + 1]) << 8) |
                 (static_cast<uint32_t>(rec[pos + 2]) << 16) | (static_cast<uint32_t>(rec[pos + 3]) << 24);
    pos += 4;
    return v;
  };
  // XLUnicodeString: characters (16 bit), flags (bit 0: characters are 16 bit)
  auto get_string = [&]() -> std::string {
    size_t cch = get16();
    if (!need(1)) return "";
    uint8_t flags = rec[pos++];
    std::u16string s;
    for (size_t i = 0; i < cch && ok; ++i) {
      if (flags & 0x01) {
        s.push_back(static_cast<char16_t>(get16()));
      } else {
        if (!need(1)) break;
        s.push_back(static_cast<char16_t>(rec[pos++]));
      }
    }
    return to_utf8(s);
  };

  // frtRefHeaderU: record type, flags and a range. The record type tells Feature11 (0x0872)
  // from Feature12 (0x0878), the rest is the same.
  pos += 12;
  uint16_t isf = get16();   // type of shared feature, 5 is a list
  pos += 1 + 4;             // reserved
  uint16_t cref = get16();  // number of ranges
  uint32_t cb_feat = get32();
  pos += 2;                 // reserved
  if (!ok || isf != 5 || cref == 0) return false;

  // ranges of the table: rows (first, last), columns (first, last)
  for (uint16_t i = 0; i < cref; ++i) {
    uint16_t r1 = get16(), r2 = get16(), c1 = get16(), c2 = get16();
    if (i == 0) {
      t.row1 = r1;
      t.row2 = r2;
      t.col1 = c1;
      t.col2 = c2;
    }
  }
  (void)cb_feat;  // 0 means: up to the end of the record

  // TableFeatureType: the fixed part has 64 bytes
  uint32_t lt = get32();  // source: 0 range, 1 SharePoint list, 2 XML, 3 external
  t.id = get32();
  uint32_t crw_header = get32();
  uint32_t crw_totals = get32();
  get32();  // next id of a column
  get32();  // size of the fixed part
  get32();  // build number, unused
  uint32_t flags = get32();
  pos += 4 * 3 + 4 + 16;  // cache position, size, characters, edit mode, hash

  if (!ok || lt != 0) return false;

  t.header = crw_header != 0;
  t.totals = crw_totals != 0;
  t.auto_filter = (flags & 0x0002) != 0;
  bool single_cell = (flags & 0x0200) != 0;
  bool has_csp_name = (flags & 0x4000) != 0;
  bool has_entry_id = (flags & 0x100000) != 0;
  if (single_cell) return false;

  t.name = get_string();
  uint16_t ncol = get16();
  if (has_csp_name) get_string();
  if (has_entry_id) get_string();
  if (!ok || ncol == 0 || ncol > 256) return false;

  // one Feat11FieldDataItem per column (2.5.103)
  for (uint16_t i = 0; i < ncol; ++i) {
    XlsTableColumn c;
    get32();  // id of the column
    get32();  // data type for SharePoint
    get32();  // data type for XML
    c.total_function = get32();
    uint32_t cb_fmt_agg = get32();
    get32();  // style of the total row
    uint32_t cflags = get32();
    uint32_t cb_fmt_insert = get32();
    get32();  // style of the insert row

    get_string();            // name in the data source
    c.name = get_string();   // caption, the text of the header cell

    if (!ok) return false;
    pos += cb_fmt_agg;
    pos += cb_fmt_insert;

    if (t.auto_filter) {
      uint32_t cb_filter = get32();
      pos += 2 + cb_filter;
    }

    // a mapping to XML or a total row with a formula is not read
    bool has_xmap = (cflags & 0x0004) != 0;
    bool has_total_fmla = (cflags & 0x0080) != 0;
    bool save_style_name = (cflags & 0x0200) != 0;
    bool has_total_str = (cflags & 0x0400) != 0;
    if (has_xmap || has_total_fmla || (cflags & 0x0008) != 0) return false;

    if (has_total_str) c.total_label = get_string();

    // the formatting of the column header is stored if the table has no header row
    if (!t.header) {
      uint32_t cb_hdr = get32();
      pos += cb_hdr;
      if (save_style_name) get_string();
    }

    if (!ok) return false;
    t.columns.push_back(c);
  }

  return pos <= rec.size();
}

#endif
