#include <algorithm>
#include <cmath>
#include <cstring>
#include <limits>
#include <iomanip>
#include <map>
#include <set>

#include "openxlsx2.h"
#include "xlsb_defines.h"
#include "xlsb_flags.h"
#include "xlsb_funs.h"
#include "xls_cfb.h"
#include "xls_defines.h"
#include "xls_funs.h"
#include "xls_drawing.h"
#include "xls_pivot.h"

// The xls reader works like the xlsb reader: the workbook is split into the
// parts of an xlsx file (workbook.xml, styles.xml, sharedStrings.xml, sheets),
// that are read by the usual xlsx code. The difference is that an xls file has no
// parts: everything is in one record stream (the globals, followed by the sheets).
// Therefore every function here reads the stream from the beginning.

namespace {

struct XlsSheet {
  std::string name;
  uint32_t offset = 0;  // position of the sheet in the stream
  uint8_t state = 0;    // 0 visible, 1 hidden, 2 very hidden
  uint8_t type = 0;     // 0 worksheet, 2 chart sheet
};

struct XlsDefinedName {
  std::string name;
  std::string formula;
  uint16_t sheet = 0;  // 0 for the workbook, else the number of the sheet
  bool hidden = false;
};

struct XlsFontInfo {
  uint16_t dyHeight = 200, grbit = 0, icv = 0x7FFF, bls = 400, sss = 0;
  uint8_t uls = 0, bFamily = 0, bCharSet = 0;
  std::string name = "Calibri";
  std::string color_attr;  // a color from XFExt, replaces icv

  // the properties as xlsx writes them: <font> with <name> in the styles, <rPr>
  // with <rFont> in a run of rich text
  std::string xml(const std::vector<uint32_t>& palette, const char* outer, const char* name_tag) const {
    std::ostringstream out;
    out << "<" << outer << ">";
    if (bls >= 700) out << "<b/>";
    if (grbit & 0x0002) out << "<i/>";
    if (grbit & 0x0008) out << "<strike/>";
    if (uls == 0x01 || uls == 0x21) out << "<u/>";
    if (uls == 0x02 || uls == 0x22) out << "<u val=\"double\"/>";
    if (sss == 1) out << "<vertAlign val=\"superscript\"/>";
    if (sss == 2) out << "<vertAlign val=\"subscript\"/>";
    out << "<sz val=\"" << xls_num(dyHeight / 20.0) << "\"/>";
    if (!color_attr.empty()) out << "<color" << color_attr << "/>";
    else if (icv != 0x7FFF && icv != 64) out << "<color rgb=\"" << xls_argb(palette, icv) << "\"/>";
    if (!name.empty()) out << "<" << name_tag << " val=\"" << escape_xml(name) << "\"/>";
    if (bFamily) out << "<family val=\"" << static_cast<int>(bFamily) << "\"/>";
    if (bCharSet) out << "<charset val=\"" << static_cast<int>(bCharSet) << "\"/>";
    out << "</" << outer << ">";
    return out.str();
  }
};

struct XlsExternBook {
  std::string path;
  std::vector<std::string> sheets;
};

struct XlsGlobals {
  std::vector<XlsPivotSource> pivot_sources;
  bool date1904 = false;
  std::string code_name;
  bool lock_structure = false, lock_windows = false;
  uint16_t password = 0;
  // window of the workbook (Window1)
  bool have_window = false;
  int16_t win_x = 0, win_y = 0;
  uint16_t win_w = 0, win_h = 0, win_flags = 0, win_active = 0, win_first = 0, win_ratio = 600;
  // calculation (CalcMode, CalcCount, CalcRefMode, CalcIter, CalcDelta, CalcPrecision, CalcSaveRecalc)
  int16_t calc_mode = 1;
  uint16_t calc_count = 100;
  bool calc_a1 = true, calc_iter = false, calc_precision = true, calc_save = true;
  double calc_delta = 0.001;
  bool have_calc = false;
  std::vector<XlsSheet> sheets;
  std::vector<XlsExternBook> externals;
  std::vector<XlsDefinedName> names;
  XlsFormulaContext ctx;
};

// Sheet1! or 'Sheet name'!. A name is quoted if it has more than letters, digits,
// dots and underscores, or starts with a digit. For a range of sheets the
// quotes go around both names: 'Sheet 1:Sheet 3'!
std::string xls_sheet_prefix(const std::string& first, const std::string& last, const std::string& book = "") {
  bool plain = true;
  const std::string* names[2] = {&first, &last};
  for (size_t n = 0; n < 2; ++n) {
    const std::string& name = *names[n];
    if (name.empty() || std::isdigit(static_cast<unsigned char>(name[0]))) plain = false;
    for (size_t i = 0; i < name.size(); ++i) {
      unsigned char c = static_cast<unsigned char>(name[i]);
      if (!(std::isalnum(c) || c == '_' || c == '.')) plain = false;
    }
  }

  std::string sheets = book + first;
  if (last != first) sheets += ":" + last;
  if (plain) return sheets + "!";

  std::string quoted = "'";
  for (size_t i = 0; i < sheets.size(); ++i) {
    if (sheets[i] == '\'') quoted += "'";
    quoted += sheets[i];
  }
  return quoted + "'!";
}

// The path of an external workbook. The first character of an encoded path is
// 0x01, then follow codes: 0x01 a drive letter (0x40: a network share), 0x02 the
// drive of this workbook, 0x03 a separator, 0x04 the parent directory, 0x05 a
// protocol with the length of the rest, 0x06 to 0x08 directories of Excel.
std::string xls_virt_path(const std::u16string& raw) {
  std::string path;
  size_t i = 0, n = raw.size();
  if (n > 0 && raw[0] == 0x01) {
    ++i;
    while (i < n && raw[i] < 0x20 && raw[i] != 0x03) {
      char16_t code = raw[i++];
      if (code == 0x01 && i < n) {
        char16_t v = raw[i++];
        if (v == 0x40) path += "//";
        else path += to_utf8(std::u16string(1, v)) + ":";
      } else if (code == 0x02) {
        path += "/";
      } else if (code == 0x04) {
        path += "../";
      } else if (code == 0x05) {
        ++i;
        break;
      }
    }
  }
  if (i + 1 < n && raw[i + 1] == ':') path.clear();
  std::string rest;
  for (; i < n; ++i) {
    if (raw[i] == 0x03) rest += "/";
    else if (raw[i] >= 0x20) rest += to_utf8(std::u16string(1, raw[i]));
  }
  size_t open = rest.find('[');
  size_t close = rest.find(']');
  if (open != std::string::npos && close != std::string::npos && close > open)
    rest = rest.substr(0, open) + rest.substr(open + 1, close - open - 1);
  return path + rest;
}

// Reads the part of the stream before the first sheet: the names of the sheets,
// where they are, and what is needed to read formulas (the references to sheets
// and the defined names).
XlsGlobals xls_globals(std::istream& bin, bool swapit, bool debug) {
  XlsGlobals g;

  // SupBook: 0 this workbook, 1 add-in functions, 2 another workbook
  std::vector<uint8_t> supbooks;
  std::vector<size_t> supbook_external;
  std::vector<std::vector<std::string>> supbook_sheets;
  std::vector<std::vector<std::string> > supbook_names;
  // ExternSheet: for each XTI the supbook, the first and the last sheet
  std::vector<std::vector<uint16_t> > xtis;

  struct RawName {
    XlsDefinedName name;
    XlsRgce rgce;
  };
  std::vector<RawName> raw_names;

  bin.clear();
  bin.seekg(0, std::ios_base::beg);

  int depth = 0;
  R_xlen_t idx = 0;
  uint16_t rid = 0, rsize = 0;
  while (XLS_RECORD(rid, rsize, bin, swapit)) {
    checkInterrupt(idx);
    std::streampos next = bin.tellg() + std::streamoff(rsize);
    if (debug) Rcpp::Rcout << "globals: " << std::hex << rid << std::dec << ": " << rsize << std::endl;

    switch (rid) {
      case XlsBOF: {
        ++depth;
        break;
      }

      case XlsEOF: {
        --depth;
        break;
      }

      case XlsFilePass: {
        Rcpp::stop("xls: encrypted files are not supported");
        break;
      }

      case XlsProtect: {
        uint16_t flag = 0;
        flag = readbin(flag, bin, swapit);
        g.lock_structure = flag != 0;
        break;
      }

      case XlsWindow1: {
        g.have_window = true;
        g.win_x = readbin(g.win_x, bin, swapit);
        g.win_y = readbin(g.win_y, bin, swapit);
        g.win_w = readbin(g.win_w, bin, swapit);
        g.win_h = readbin(g.win_h, bin, swapit);
        g.win_flags = readbin(g.win_flags, bin, swapit);
        g.win_active = readbin(g.win_active, bin, swapit);
        g.win_first = readbin(g.win_first, bin, swapit);
        bin.seekg(2, std::ios_base::cur);
        g.win_ratio = readbin(g.win_ratio, bin, swapit);
        break;
      }

      case XlsCalcMode: {
        g.have_calc = true;
        g.calc_mode = readbin(g.calc_mode, bin, swapit);
        break;
      }

      case XlsCalcCount: {
        g.have_calc = true;
        g.calc_count = readbin(g.calc_count, bin, swapit);
        break;
      }

      case XlsCalcRefMode: {
        uint16_t v = 1;
        v = readbin(v, bin, swapit);
        g.calc_a1 = v != 0;
        break;
      }

      case XlsCalcIter: {
        uint16_t v = 0;
        v = readbin(v, bin, swapit);
        g.calc_iter = v != 0;
        break;
      }

      case XlsCalcDelta: {
        g.calc_delta = readbin(g.calc_delta, bin, swapit);
        break;
      }

      case XlsCalcPrecision: {
        uint16_t v = 1;
        v = readbin(v, bin, swapit);
        g.calc_precision = v != 0;
        break;
      }

      case XlsCalcSaveRecalc: {
        uint16_t v = 1;
        v = readbin(v, bin, swapit);
        g.calc_save = v != 0;
        break;
      }

      case XlsWinProtect: {
        uint16_t flag = 0;
        flag = readbin(flag, bin, swapit);
        g.lock_windows = flag != 0;
        break;
      }

      case XlsPassword: {
        g.password = readbin(g.password, bin, swapit);
        break;
      }

      case XlsCodeName: {
        if (rsize > 0) g.code_name = XLUnicodeString(bin, swapit, true);
        break;
      }

      case XlsSXStreamID: {
        std::vector<uint8_t> rec(rsize);
        bin.read(reinterpret_cast<char*>(rec.data()), rsize);
        XlsPivotSource source;
        source.id = static_cast<uint16_t>(xls_pv_u16(rec, 0));
        g.pivot_sources.push_back(source);
        break;
      }

      case XlsSXVS: {
        std::vector<uint8_t> rec(rsize);
        bin.read(reinterpret_cast<char*>(rec.data()), rsize);
        if (!g.pivot_sources.empty()) g.pivot_sources.back().type = static_cast<uint16_t>(xls_pv_u16(rec, 0));
        break;
      }

      case XlsDConRef: {
        std::vector<uint8_t> rec(rsize);
        bin.read(reinterpret_cast<char*>(rec.data()), rsize);
        if (!g.pivot_sources.empty()) xls_pv_dconref(rec, g.pivot_sources.back());
        break;
      }

      case XlsDConName: {
        std::vector<uint8_t> rec(rsize);
        bin.read(reinterpret_cast<char*>(rec.data()), rsize);
        if (!g.pivot_sources.empty()) xls_pv_dconname(rec, g.pivot_sources.back());
        break;
      }

      case XlsDate1904: {
        uint16_t flag = 0;
        flag = readbin(flag, bin, swapit);
        g.date1904 = flag == 1;
        break;
      }

      case XlsBoundSheet8: {
        XlsSheet sheet;
        sheet.offset = readbin(sheet.offset, bin, swapit);
        sheet.state = readbin(sheet.state, bin, swapit);
        sheet.type = readbin(sheet.type, bin, swapit);
        sheet.name = XLUnicodeString(bin, swapit, false);
        g.sheets.push_back(sheet);
        break;
      }

      case XlsSupBook: {
        uint16_t ctab = 0, cch = 0;
        ctab = readbin(ctab, bin, swapit);
        cch = readbin(cch, bin, swapit);
        supbooks.push_back(cch == 0x0401 ? 0 : (cch == 0x3A01 ? 1 : 2));
        supbook_external.push_back(0);
        supbook_sheets.push_back(std::vector<std::string>());
        supbook_names.push_back(std::vector<std::string>());
        if (cch != 0x0401 && cch != 0x3A01) {
          XlsExternBook book;
          std::streampos rec_end = next;
          uint8_t flag = 0;
          flag = readbin(flag, bin, swapit);
          book.path = xls_virt_path(xls_chars(bin, swapit, cch, (flag & 0x01) != 0, rec_end));
          for (uint16_t i = 0; i < ctab; ++i) book.sheets.push_back(XLUnicodeString(bin, swapit, true, rec_end));
          size_t found = g.externals.size();
          for (size_t i = 0; i < g.externals.size(); ++i) {
            if (g.externals[i].path == book.path) {
              found = i;
              break;
            }
          }
          if (found == g.externals.size()) {
            g.externals.push_back(book);
          } else {
            for (size_t i = 0; i < book.sheets.size(); ++i) {
              if (std::find(g.externals[found].sheets.begin(), g.externals[found].sheets.end(), book.sheets[i]) == g.externals[found].sheets.end())
                g.externals[found].sheets.push_back(book.sheets[i]);
            }
          }
          supbook_external.back() = found + 1;
          supbook_sheets.back() = book.sheets;
        }
        break;
      }

      case XlsExternName: {
        // a name in an external workbook or an add-in function: options (2),
        // index of the sheet and unused (4), then the name
        if (supbook_names.empty()) break;
        bin.seekg(6, std::ios::cur);
        supbook_names.back().push_back(XLUnicodeString(bin, swapit, false));
        break;
      }

      case XlsExternSheet: {
        uint16_t cxti = 0;
        cxti = readbin(cxti, bin, swapit);
        for (uint16_t i = 0; i < cxti; ++i) {
          std::vector<uint16_t> xti(3, 0);
          for (size_t j = 0; j < 3; ++j) xti[j] = readbin(xti[j], bin, swapit);
          xtis.push_back(xti);
        }
        break;
      }

      case XlsLbl: {
        // flags (2), chKey (1), cch (1), cce (2), unused (2), itab (2), four
        // lengths of strings for macros (4), then the name and the formula
        uint16_t flags = 0, cce = 0, unused = 0, itab = 0;
        uint8_t chKey = 0, cch = 0;
        flags = readbin(flags, bin, swapit);
        chKey = readbin(chKey, bin, swapit);
        cch = readbin(cch, bin, swapit);
        cce = readbin(cce, bin, swapit);
        unused = readbin(unused, bin, swapit);
        itab = readbin(itab, bin, swapit);
        bin.seekg(4, std::ios::cur);

        uint8_t wide = 0;
        wide = readbin(wide, bin, swapit);
        std::streampos no_end = xls_no_end;
        std::u16string chars = xls_chars(bin, swapit, cch, (wide & 0x01) != 0, no_end);

        RawName raw;
        raw.name.hidden = (flags & 0x0001) != 0;
        raw.name.sheet = itab;
        bool builtin = (flags & 0x0020) != 0;
        if (builtin && !chars.empty() && chars[0] < 14) {
          raw.name.name = std::string("_xlnm.") + XlsBuiltInNames[chars[0]];
        } else {
          raw.name.name = to_utf8(chars);
        }

        // the tokens, then what belongs to them (constants of arrays)
        raw.rgce.cce = cce;
        if (next < bin.tellg()) Rcpp::stop("xls: record overrun");
        raw.rgce.data.resize(static_cast<size_t>(next - bin.tellg()));
        if (!raw.rgce.data.empty()) bin.read(reinterpret_cast<char*>(&raw.rgce.data[0]), static_cast<std::streamsize>(raw.rgce.data.size()));
        raw_names.push_back(raw);
        break;
      }

      default:
        break;
    }

    if (depth <= 0) break;
    bin.clear();
    bin.seekg(next, std::ios_base::beg);
  }

  // text in front of a 3D reference. Only references to this workbook are supported
  for (size_t i = 0; i < xtis.size(); ++i) {
    std::string prefix;
    uint16_t book = xtis[i][0], first = xtis[i][1], last = xtis[i][2];
    bool internal = book < supbooks.size() && supbooks[book] == 0;
    if (internal && first < g.sheets.size() && last < g.sheets.size()) {
      prefix = xls_sheet_prefix(g.sheets[first].name, g.sheets[last].name);
    } else if (book < supbooks.size() && supbooks[book] == 2) {
      const std::vector<std::string>& sn = supbook_sheets[book];
      if (first < sn.size() && last < sn.size())
        prefix = xls_sheet_prefix(sn[first], sn[last],
                                  "[" + std::to_string(supbook_external[book]) + "]");
    }
    g.ctx.xti.push_back(prefix);

    std::vector<std::string> xnames;
    if (book < supbooks.size()) {
      xnames = supbook_names[book];
      if (supbooks[book] == 2)
        for (size_t j = 0; j < xnames.size(); ++j) xnames[j] = "[" + std::to_string(supbook_external[book]) + "]!" + xnames[j];
    }
    g.ctx.xnames.push_back(xnames);
  }

  // names are decoded last, a name can use other names
  for (size_t i = 0; i < raw_names.size(); ++i) g.ctx.names.push_back(raw_names[i].name.name);
  for (size_t i = 0; i < raw_names.size(); ++i) {
    XlsDefinedName name = raw_names[i].name;
    name.formula = xls_rgce(raw_names[i].rgce, g.ctx, swapit, debug);
    g.names.push_back(name);
  }

  return g;
}

}  // namespace

// The stream called Workbook inside of the xls file, saved to outPath. This is
// what the other functions read.
// [[Rcpp::export]]
int32_t xls_extract(std::string filePath, std::string outPath, bool debug) {
  std::ifstream xls(filePath, std::ios::in | std::ios::binary | std::ios::ate);
  if (!xls) return -1;

  std::vector<uint8_t> file(static_cast<size_t>(xls.tellg()));
  xls.seekg(0, std::ios_base::beg);
  xls.read(reinterpret_cast<char*>(file.data()), static_cast<std::streamsize>(file.size()));

  std::vector<uint8_t> workbook = cfb_stream(file, "Workbook");
  if (workbook.empty()) {
    // Excel 5 and 95 files call it Book. They have a different record layout
    if (!cfb_stream(file, "Book").empty()) Rcpp::stop("xls: only Excel 97-2003 files are supported");
    Rcpp::stop("xls: file has no Workbook stream");
  }

  // the stream starts with a BOF record, version 0x0600 is BIFF8
  if (workbook.size() < 8 || cfb_u16(workbook, 0) != XlsBOF || cfb_u16(workbook, 4) != 0x0600)
    Rcpp::stop("xls: only Excel 97-2003 files are supported");

  if (!cfb_stream(file, "_VBA_PROJECT").empty() || !cfb_stream(file, "PROJECT").empty())
    Rcpp::warning("The xls file contains a VBA project. Macros are not imported.");

  if (debug) Rcpp::Rcout << "Workbook stream: " << workbook.size() << " bytes" << std::endl;

  std::ofstream out(outPath, std::ios::out | std::ios::binary);
  out.write(reinterpret_cast<const char*>(workbook.data()), static_cast<std::streamsize>(workbook.size()));
  return static_cast<int32_t>(workbook.size());
}

// Document properties: title, author and more are not part of the workbook, they are
// in two streams in OLE property set format. Written to docProps/core.xml and
// docProps/app.xml. Returns 1 if core.xml was written, plus 2 for app.xml.
namespace {

struct XlsProperty {
  uint32_t type = 0;
  std::string text;
  uint64_t time = 0;
  uint32_t number = 0;
};

void xls_utf8(std::string& s, uint32_t c) {
  if (c < 0x80) {
    s += static_cast<char>(c);
  } else if (c < 0x800) {
    s += static_cast<char>(0xC0 | (c >> 6));
    s += static_cast<char>(0x80 | (c & 0x3F));
  } else if (c < 0x10000) {
    s += static_cast<char>(0xE0 | (c >> 12));
    s += static_cast<char>(0x80 | ((c >> 6) & 0x3F));
    s += static_cast<char>(0x80 | (c & 0x3F));
  } else {
    s += static_cast<char>(0xF0 | (c >> 18));
    s += static_cast<char>(0x80 | ((c >> 12) & 0x3F));
    s += static_cast<char>(0x80 | ((c >> 6) & 0x3F));
    s += static_cast<char>(0x80 | (c & 0x3F));
  }
}

// text in the code page of the property set. UTF-8 is kept, the Windows and Mac
// western code pages are converted, everything else is read as Latin 1.
std::string xls_codepage_text(const uint8_t* p, size_t n, uint32_t codepage) {
  static const uint16_t mac[128] = {
      0x00C4, 0x00C5, 0x00C7, 0x00C9, 0x00D1, 0x00D6, 0x00DC, 0x00E1, 0x00E0, 0x00E2, 0x00E4, 0x00E3, 0x00E5, 0x00E7, 0x00E9, 0x00E8,
      0x00EA, 0x00EB, 0x00ED, 0x00EC, 0x00EE, 0x00EF, 0x00F1, 0x00F3, 0x00F2, 0x00F4, 0x00F6, 0x00F5, 0x00FA, 0x00F9, 0x00FB, 0x00FC,
      0x2020, 0x00B0, 0x00A2, 0x00A3, 0x00A7, 0x2022, 0x00B6, 0x00DF, 0x00AE, 0x00A9, 0x2122, 0x00B4, 0x00A8, 0x2260, 0x00C6, 0x00D8,
      0x221E, 0x00B1, 0x2264, 0x2265, 0x00A5, 0x00B5, 0x2202, 0x2211, 0x220F, 0x03C0, 0x222B, 0x00AA, 0x00BA, 0x03A9, 0x00E6, 0x00F8,
      0x00BF, 0x00A1, 0x00AC, 0x221A, 0x0192, 0x2248, 0x2206, 0x00AB, 0x00BB, 0x2026, 0x00A0, 0x00C0, 0x00C3, 0x00D5, 0x0152, 0x0153,
      0x2013, 0x2014, 0x201C, 0x201D, 0x2018, 0x2019, 0x00F7, 0x25CA, 0x00FF, 0x0178, 0x2044, 0x20AC, 0x2039, 0x203A, 0xFB01, 0xFB02,
      0x2021, 0x00B7, 0x201A, 0x201E, 0x2030, 0x00C2, 0x00CA, 0x00C1, 0x00CB, 0x00C8, 0x00CD, 0x00CE, 0x00CF, 0x00CC, 0x00D3, 0x00D4,
      0xF8FF, 0x00D2, 0x00DA, 0x00DB, 0x00D9, 0x0131, 0x02C6, 0x02DC, 0x00AF, 0x02D8, 0x02D9, 0x02DA, 0x00B8, 0x02DD, 0x02DB, 0x02C7};
  static const uint16_t win[32] = {0x20AC, 0x0081, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160,
                                   0x2039, 0x0152, 0x008D, 0x017D, 0x008F, 0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022,
                                   0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x009D, 0x017E, 0x0178};
  std::string out;
  for (size_t i = 0; i < n && p[i] != 0; ++i) {
    uint32_t c = p[i];
    if (codepage == 65001) {
      out += static_cast<char>(c);
      continue;
    }
    if (c >= 0x80) {
      if (codepage == 10000) c = mac[c - 0x80];
      else if (codepage == 1252 && c < 0xA0) c = win[c - 0x80];
    }
    xls_utf8(out, c);
  }
  return out;
}

std::string xls_filetime(uint64_t ft) {
  if (ft == 0) return "";
  int64_t secs = static_cast<int64_t>(ft / 10000000ULL) - 11644473600LL;
  int64_t days = secs >= 0 ? secs / 86400 : -((-secs + 86399) / 86400);
  int64_t rest = secs - days * 86400;
  // civil date from the days since 1970-01-01
  days += 719468;
  int64_t era = (days >= 0 ? days : days - 146096) / 146097;
  int64_t doe = days - era * 146097;
  int64_t yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
  int64_t y = yoe + era * 400;
  int64_t doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
  int64_t mp = (5 * doy + 2) / 153;
  int64_t d = doy - (153 * mp + 2) / 5 + 1;
  int64_t m = mp < 10 ? mp + 3 : mp - 9;
  if (m <= 2) ++y;
  char buf[64];
  std::snprintf(buf, sizeof(buf), "%04d-%02d-%02dT%02d:%02d:%02dZ", static_cast<int>(y), static_cast<int>(m), static_cast<int>(d),
                static_cast<int>(rest / 3600), static_cast<int>((rest % 3600) / 60), static_cast<int>(rest % 60));
  return buf;
}

// the first property set of the stream, by property id
std::map<uint32_t, XlsProperty> xls_property_set(const std::vector<uint8_t>& d) {
  std::map<uint32_t, XlsProperty> props;
  if (d.size() < 48) return props;
  uint32_t nsets = static_cast<uint32_t>(cfb_u32(d, 24));
  if (nsets == 0) return props;
  size_t sec = cfb_u32(d, 44);
  if (sec + 8 > d.size()) return props;
  uint32_t n = static_cast<uint32_t>(cfb_u32(d, sec + 4));
  // the code page is the property 1, it has to be known before the strings
  uint32_t codepage = 1252;
  for (int pass = 0; pass < 2; ++pass) {
    for (uint32_t i = 0; i < n && sec + 8 + i * 8 + 8 <= d.size(); ++i) {
      uint32_t id = static_cast<uint32_t>(cfb_u32(d, sec + 8 + i * 8));
      size_t at = sec + cfb_u32(d, sec + 12 + i * 8);
      if (at + 8 > d.size()) continue;
      uint32_t type = static_cast<uint32_t>(cfb_u32(d, at)) & 0xFFFFu;
      if (pass == 0) {
        if (id == 1 && type == 2) codepage = static_cast<uint32_t>(cfb_u16(d, at + 4));
        continue;
      }
      XlsProperty prop;
      prop.type = type;
      if (type == 0x1E) {
        size_t len = cfb_u32(d, at + 4);
        if (at + 8 + len <= d.size()) prop.text = xls_codepage_text(d.data() + at + 8, len, codepage);
      } else if (type == 0x1F) {
        size_t len = cfb_u32(d, at + 4);
        for (size_t k = 0; k < len && at + 8 + 2 * k + 2 <= d.size(); ++k) {
          uint32_t c = static_cast<uint32_t>(cfb_u16(d, at + 8 + 2 * k));
          if (c == 0) break;
          xls_utf8(prop.text, c);
        }
      } else if (type == 0x40) {
        prop.time = static_cast<uint64_t>(cfb_u32(d, at + 4)) | (static_cast<uint64_t>(cfb_u32(d, at + 8)) << 32);
      } else if (type == 2 || type == 3) {
        prop.number = static_cast<uint32_t>(cfb_u32(d, at + 4));
      }
      props[id] = prop;
    }
  }
  return props;
}

}  // namespace

// [[Rcpp::export]]
int32_t docprops_xls(std::string filePath, std::string corePath, std::string appPath) {
  std::ifstream xls(filePath, std::ios::in | std::ios::binary | std::ios::ate);
  if (!xls) return 0;
  std::vector<uint8_t> file(static_cast<size_t>(xls.tellg()));
  xls.seekg(0, std::ios_base::beg);
  xls.read(reinterpret_cast<char*>(file.data()), static_cast<std::streamsize>(file.size()));

  int32_t result = 0;

  std::vector<uint8_t> summary = cfb_stream(file, "\005SummaryInformation");
  if (!summary.empty()) {
    std::map<uint32_t, XlsProperty> p = xls_property_set(summary);
    auto text = [&](uint32_t id) { return p.count(id) ? escape_xml(p[id].text) : std::string(); };
    std::ofstream out(corePath);
    out << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" << std::endl;
    out << "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" "
        << "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" "
        << "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">";
    if (!text(2).empty()) out << "<dc:title>" << text(2) << "</dc:title>";
    if (!text(3).empty()) out << "<dc:subject>" << text(3) << "</dc:subject>";
    if (!text(4).empty()) out << "<dc:creator>" << text(4) << "</dc:creator>";
    if (!text(5).empty()) out << "<cp:keywords>" << text(5) << "</cp:keywords>";
    if (!text(6).empty()) out << "<dc:description>" << text(6) << "</dc:description>";
    if (!text(8).empty()) out << "<cp:lastModifiedBy>" << text(8) << "</cp:lastModifiedBy>";
    if (p.count(12) && p[12].time) out << "<dcterms:created xsi:type=\"dcterms:W3CDTF\">" << xls_filetime(p[12].time) << "</dcterms:created>";
    if (p.count(13) && p[13].time) out << "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">" << xls_filetime(p[13].time) << "</dcterms:modified>";
    out << "</cp:coreProperties>" << std::endl;
    result += 1;
  }

  std::vector<uint8_t> doc_summary = cfb_stream(file, "\005DocumentSummaryInformation");
  if (!doc_summary.empty()) {
    std::map<uint32_t, XlsProperty> p = xls_property_set(doc_summary);
    if (p.count(14) || p.count(15)) {
      std::ofstream out(appPath);
      out << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" << std::endl;
      out << "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">";
      if (p.count(15)) out << "<Company>" << escape_xml(p[15].text) << "</Company>";
      if (p.count(14)) out << "<Manager>" << escape_xml(p[14].text) << "</Manager>";
      out << "</Properties>" << std::endl;
      result += 2;
    }
  }
  return result;
}

// workbook.xml: sheets and defined names. Returns the number of sheets.
// [[Rcpp::export]]
int32_t workbook_xls(std::string filePath, std::string outPath, bool debug) {
  std::ofstream out(outPath);
  std::ifstream bin(filePath, std::ios::in | std::ios::binary | std::ios::ate);

  bool swapit = is_big_endian();

  if (bin) {
    bin.seekg(0, std::ios_base::beg);
    XlsGlobals g = xls_globals(bin, swapit, debug);

    out << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" << std::endl;
    out << "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ";
    out << "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" << std::endl;
    out << "<workbookPr" << (g.date1904 ? " date1904=\"1\"" : "");
    if (!g.code_name.empty()) out << " codeName=\"" << escape_xml(g.code_name) << "\"";
    out << "/>" << std::endl;

    if (g.lock_structure || g.lock_windows) {
      out << "<workbookProtection";
      if (g.password) {
        char hex[8];
        std::snprintf(hex, sizeof(hex), "%04X", static_cast<unsigned>(g.password));
        out << " workbookPassword=\"" << hex << "\"";
      }
      if (g.lock_structure) out << " lockStructure=\"1\"";
      if (g.lock_windows) out << " lockWindows=\"1\"";
      out << "/>" << std::endl;
    }

    if (g.have_window) {
      out << "<bookViews><workbookView xWindow=\"" << g.win_x << "\" yWindow=\"" << g.win_y << "\" windowWidth=\"" << g.win_w
          << "\" windowHeight=\"" << g.win_h << "\" tabRatio=\"" << g.win_ratio << "\" firstSheet=\"" << g.win_first
          << "\" activeTab=\"" << g.win_active << "\"";
      if (g.win_flags & 0x0001) out << " visibility=\"hidden\"";
      if (g.win_flags & 0x0004) out << " visibility=\"veryHidden\"";
      if (g.win_flags & 0x0002) out << " minimized=\"1\"";
      if (!(g.win_flags & 0x0008)) out << " showHorizontalScroll=\"0\"";
      if (!(g.win_flags & 0x0010)) out << " showVerticalScroll=\"0\"";
      if (!(g.win_flags & 0x0020)) out << " showSheetTabs=\"0\"";
      out << "/></bookViews>" << std::endl;
    }

    out << "<sheets>" << std::endl;
    for (size_t i = 0; i < g.sheets.size(); ++i) {
      out << "<sheet name=\"" << escape_xml(g.sheets[i].name) << "\" sheetId=\"" << i + 1 << "\"";
      if (g.sheets[i].state == 1) out << " state=\"hidden\"";
      if (g.sheets[i].state == 2) out << " state=\"veryHidden\"";
      out << " r:id=\"rId" << i + 1 << "\"/>" << std::endl;
    }
    out << "</sheets>" << std::endl;

    if (!g.externals.empty()) {
      out << "<externalReferences>" << std::endl;
      for (size_t i = 0; i < g.externals.size(); ++i)
        out << "<externalReference r:id=\"rId" << g.sheets.size() + 3 + i << "\"/>" << std::endl;
      out << "</externalReferences>" << std::endl;
    }

    // names we can not read are dropped, they are mostly leftovers of macros
    size_t n_names = 0;
    for (size_t i = 0; i < g.names.size(); ++i)
      if (!g.names[i].formula.empty() && g.names[i].name.compare(0, 6, "_xlfn.") != 0) ++n_names;

    if (n_names > 0) {
      out << "<definedNames>" << std::endl;
      for (size_t i = 0; i < g.names.size(); ++i) {
        const XlsDefinedName& name = g.names[i];
        // Excel adds a hidden name for every function that is new in a later version
        if (name.formula.empty() || name.name.compare(0, 6, "_xlfn.") == 0) continue;
        bool duplicate = false;
        for (size_t j = 0; j < i; ++j)
          if (g.names[j].name == name.name && g.names[j].sheet == name.sheet && !g.names[j].formula.empty()) duplicate = true;
        if (duplicate) continue;
        out << "<definedName name=\"" << escape_xml(name.name) << "\"";
        if (name.sheet > 0) out << " localSheetId=\"" << name.sheet - 1 << "\"";
        if (name.hidden) out << " hidden=\"1\"";
        out << ">" << escape_xml(name.formula) << "</definedName>" << std::endl;
      }
      out << "</definedNames>" << std::endl;
    }

    if (g.have_calc) {
      out << "<calcPr";
      if (g.calc_mode == 0) out << " calcMode=\"manual\"";
      if (g.calc_mode == 2) out << " calcMode=\"autoNoTable\"";
      if (!g.calc_a1) out << " refMode=\"R1C1\"";
      if (g.calc_iter) out << " iterate=\"1\" iterateCount=\"" << g.calc_count << "\" iterateDelta=\"" << xls_num(g.calc_delta) << "\"";
      if (!g.calc_precision) out << " fullPrecision=\"0\"";
      if (!g.calc_save) out << " calcOnSave=\"0\"";
      out << "/>" << std::endl;
    }

    out << "</workbook>" << std::endl;

    out.close();
    bin.close();
    return static_cast<int32_t>(g.sheets.size());
  } else {
    return -1;
  }
}

// xl/externalLinks/externalLinkN.xml and the relationship with the path of each
// external workbook. Returns the number of external workbooks.
// [[Rcpp::export]]
int32_t externallinks_xls(std::string filePath, std::string outDir, bool debug) {
  std::ifstream bin(filePath, std::ios::in | std::ios::binary | std::ios::ate);
  bool swapit = is_big_endian();
  if (!bin) return -1;
  bin.seekg(0, std::ios_base::beg);
  XlsGlobals g = xls_globals(bin, swapit, debug);

  for (size_t i = 0; i < g.externals.size(); ++i) {
    std::string base = outDir + "/externalLink" + std::to_string(i + 1) + ".xml";
    std::ofstream out(base);
    out << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" << std::endl;
    out << "<externalLink xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" << std::endl;
    out << "<externalBook xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId1\">";
    out << "<sheetNames>";
    for (size_t j = 0; j < g.externals[i].sheets.size(); ++j)
      out << "<sheetName val=\"" << escape_xml(g.externals[i].sheets[j]) << "\"/>";
    out << "</sheetNames></externalBook></externalLink>" << std::endl;
    out.close();

    std::ofstream rels(outDir + "/_rels/externalLink" + std::to_string(i + 1) + ".xml.rels");
    rels << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" << std::endl;
    rels << "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
    rels << "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath\" Target=\""
         << escape_xml(g.externals[i].path) << "\" TargetMode=\"External\"/>";
    rels << "</Relationships>" << std::endl;
  }
  return static_cast<int32_t>(g.externals.size());
}

// The theme of the workbook is saved in the Theme record as a zip file. The
// rest of it is in Continue records. Returns the size of the zip file that was
// written, 0 if there is no theme.
// [[Rcpp::export]]
int32_t theme_xls(std::string filePath, std::string outPath, bool debug) {
  std::ifstream bin(filePath, std::ios::in | std::ios::binary | std::ios::ate);
  bool swapit = is_big_endian();
  if (!bin) return -1;
  bin.seekg(0, std::ios_base::beg);

  int depth = 0;
  uint16_t rid = 0, rsize = 0;
  std::string bytes;
  bool in_theme = false;
  while (XLS_RECORD(rid, rsize, bin, swapit)) {
    std::streampos next = bin.tellg() + std::streamoff(rsize);
    if (rid == XlsBOF) ++depth;
    if (rid == XlsEOF) --depth;
    if (depth <= 0) break;

    if (rid == XlsTheme) {
      // frt header (12 bytes) and the version of the theme (4 bytes)
      in_theme = true;
      if (rsize > 16) {
        bin.seekg(16, std::ios::cur);
        bytes.resize(static_cast<size_t>(rsize) - 16);
        bin.read(&bytes[0], static_cast<std::streamsize>(bytes.size()));
      }
    } else if (in_theme && rid == XlsContinue) {
      std::string more(rsize, '\0');
      bin.read(&more[0], rsize);
      bytes += more;
    } else {
      in_theme = false;
    }
    if (debug && in_theme) Rcpp::Rcout << "theme: " << bytes.size() << " bytes" << std::endl;

    bin.clear();
    bin.seekg(next, std::ios_base::beg);
  }

  if (bytes.size() < 4 || bytes.compare(0, 2, "PK") != 0) return 0;
  std::ofstream out(outPath, std::ios::binary);
  out.write(bytes.data(), static_cast<std::streamsize>(bytes.size()));
  return static_cast<int32_t>(bytes.size());
}

// The pictures of the workbook are in the drawing group of the globals, a
// part of the Escher data. They are saved as image1, image2, ... in the
// directory, the position is the number the shapes use. Returns the names of the
// files, empty if a picture is not stored in the file. Compressed metafiles get
// the extension zlib added, they are unpacked in R. The size they have when
// unpacked is the attribute usize.
// [[Rcpp::export]]
Rcpp::CharacterVector media_xls(std::string filePath, std::string outDir, bool debug) {
  std::ifstream bin(filePath, std::ios::in | std::ios::binary | std::ios::ate);
  bool swapit = is_big_endian();
  if (!bin) return Rcpp::CharacterVector();
  bin.seekg(0, std::ios_base::beg);

  int depth = 0;
  uint16_t rid = 0, rsize = 0;
  std::string group;
  bool in_group = false;
  while (XLS_RECORD(rid, rsize, bin, swapit)) {
    std::streampos next = bin.tellg() + std::streamoff(rsize);
    if (rid == XlsBOF) ++depth;
    if (rid == XlsEOF) --depth;
    if (depth <= 0) break;

    if (rid == XlsMsoDrawingGroup || (in_group && rid == XlsContinue)) {
      in_group = true;
      std::string part(rsize, '\0');
      bin.read(&part[0], rsize);
      group += part;
    } else {
      in_group = false;
    }
    bin.clear();
    bin.seekg(next, std::ios_base::beg);
  }

  std::vector<XlsBlip> blips = xls_blips(group);
  Rcpp::CharacterVector names(blips.size());
  Rcpp::NumericVector usize(blips.size());
  for (size_t i = 0; i < blips.size(); ++i) {
    usize[i] = static_cast<double>(blips[i].size);
    if (blips[i].ext.empty()) {
      names[i] = "";
      continue;
    }
    std::string data = blips[i].data, ext = blips[i].ext;
    if (ext == "dib") {
      data = xls_dib_to_bmp(data);
      ext = "bmp";
    }
    if (data.empty()) {
      names[i] = "";
      continue;
    }
    std::string name = "image" + std::to_string(i + 1) + "." + ext + (blips[i].deflated ? ".zlib" : "");
    std::ofstream out(outDir + "/" + name, std::ios::binary);
    out.write(data.data(), static_cast<std::streamsize>(data.size()));
    names[i] = name;
    if (debug) Rcpp::Rcout << "media: " << name << " (" << data.size() << " bytes)" << std::endl;
  }
  names.attr("usize") = usize;
  return names;
}

// sharedStrings.xml
// [[Rcpp::export]]
int32_t sharedstrings_xls(std::string filePath, std::string outPath, bool debug) {
  std::ofstream out(outPath);
  std::ifstream bin(filePath, std::ios::in | std::ios::binary | std::ios::ate);

  bool swapit = is_big_endian();

  if (bin) {
    bin.seekg(0, std::ios_base::beg);

    out << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" << std::endl;
    out << "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" << std::endl;

    std::vector<XlsFontInfo> fonts;
    std::vector<uint32_t> palette(XlsDefaultPalette, XlsDefaultPalette + 56);

    int depth = 0;
    R_xlen_t idx = 0;
    uint16_t rid = 0, rsize = 0;
    while (XLS_RECORD(rid, rsize, bin, swapit)) {
      checkInterrupt(idx);
      std::streampos next = bin.tellg() + std::streamoff(rsize);

      if (rid == XlsBOF) ++depth;
      if (rid == XlsEOF) --depth;
      if (depth <= 0) break;

      if (rid == XlsFont) {
        XlsFontInfo f;
        f.dyHeight = readbin(f.dyHeight, bin, swapit);
        f.grbit = readbin(f.grbit, bin, swapit);
        f.icv = readbin(f.icv, bin, swapit);
        f.bls = readbin(f.bls, bin, swapit);
        f.sss = readbin(f.sss, bin, swapit);
        f.uls = readbin(f.uls, bin, swapit);
        f.bFamily = readbin(f.bFamily, bin, swapit);
        f.bCharSet = readbin(f.bCharSet, bin, swapit);
        bin.seekg(1, std::ios::cur);
        f.name = XLUnicodeString(bin, swapit, false);
        fonts.push_back(f);
      }

      if (rid == XlsPalette) {
        uint16_t ccv = 0;
        ccv = readbin(ccv, bin, swapit);
        for (size_t i = 0; i < ccv && i < palette.size(); ++i) {
          uint8_t r = 0, g = 0, b = 0, reserved = 0;
          r = readbin(r, bin, swapit);
          g = readbin(g, bin, swapit);
          b = readbin(b, bin, swapit);
          reserved = readbin(reserved, bin, swapit);
          palette[i] = (static_cast<uint32_t>(r) << 16) | (static_cast<uint32_t>(g) << 8) | b;
        }
      }

      if (rid == XlsSST) {
        uint32_t cstTotal = 0, cstUnique = 0;
        cstTotal = readbin(cstTotal, bin, swapit);
        cstUnique = readbin(cstUnique, bin, swapit);
        if (debug) Rcpp::Rcout << "SST: " << cstTotal << " strings, " << cstUnique << " unique" << std::endl;

        // the strings go on in Continue records
        std::streampos rec_end = next;
        for (uint32_t i = 0; i < cstUnique; ++i) {
          if (bin.tellg() >= rec_end) {
            uint16_t crid = 0, crsize = 0;
            if (!XLS_RECORD(crid, crsize, bin, swapit) || crid != XlsContinue)
              Rcpp::stop("xls: shared strings are cut off");
            rec_end = bin.tellg() + std::streamoff(crsize);
          }
          std::vector<std::pair<uint16_t, uint16_t> > runs;
          std::u16string chars = XLUnicodeRichString(bin, swapit, rec_end, runs);
          if (runs.empty()) {
            out << "<si><t xml:space=\"preserve\">" << escape_xml(to_utf8(chars)) << "</t></si>" << std::endl;
          } else {
            out << "<si>";
            size_t from = 0;
            for (size_t r = 0; r <= runs.size(); ++r) {
              size_t to = r < runs.size() ? std::min<size_t>(runs[r].first, chars.size()) : chars.size();
              if (to > from) {
                out << "<r>";
                if (r > 0) {
                  // the number 4 is not used for fonts
                  size_t fontId = runs[r - 1].second;
                  if (fontId > 4) --fontId;
                  if (fontId < fonts.size()) out << fonts[fontId].xml(palette, "rPr", "rFont");
                }
                out << "<t xml:space=\"preserve\">" << escape_xml(to_utf8(chars.substr(from, to - from))) << "</t></r>";
                from = to;
              }
            }
            out << "</si>" << std::endl;
          }
        }
        break;
      }

      bin.seekg(next, std::ios_base::beg);
    }

    out << "</sst>" << std::endl;

    out.close();
    bin.close();
    return 1;
  } else {
    return -1;
  }
}

// styles.xml. Fonts, number formats and cell formats (XF) are collected, since
// the palette they use can come after them, and written at the end.
// [[Rcpp::export]]
int32_t styles_xls(std::string filePath, std::string outPath, bool debug) {
  std::ofstream out(outPath);
  std::ifstream bin(filePath, std::ios::in | std::ios::binary | std::ios::ate);

  bool swapit = is_big_endian();

  if (bin) {
    bin.seekg(0, std::ios_base::beg);

    struct Xf {
      uint16_t ifnt = 0, ifmt = 0, flags = 1;
      uint8_t align = 0, rot = 0, indent = 0;
      uint16_t border = 0, border_icv = 0, fill = 0;
      uint32_t border_icv2 = 0;
    };

    std::vector<XlsFontInfo> fonts;
    std::vector<Xf> xfs;
    // exact colors and the indent of the XF records, by number of the XF
    std::map<uint16_t, std::map<uint16_t, XlsCfColor> > xf_colors;
    std::map<uint16_t, uint16_t> xf_indent;
    std::map<uint16_t, std::string> formats;
    struct NamedStyle {
      uint16_t ixfe = 0;
      bool builtin = false;
      int istyBuiltIn = -1, iLevel = 0;
      std::string name;
    };
    std::vector<NamedStyle> named_styles;
    std::vector<uint32_t> palette(XlsDefaultPalette, XlsDefaultPalette + 56);

    int depth = 0;
    R_xlen_t idx = 0;
    uint16_t rid = 0, rsize = 0;
    while (XLS_RECORD(rid, rsize, bin, swapit)) {
      checkInterrupt(idx);
      std::streampos next = bin.tellg() + std::streamoff(rsize);
      if (debug) Rcpp::Rcout << "styles: " << std::hex << rid << std::dec << ": " << rsize << std::endl;

      if (rid == XlsBOF) ++depth;
      if (rid == XlsEOF) --depth;
      if (depth <= 0) break;

      switch (rid) {
        case XlsFont: {
          XlsFontInfo f;
          f.dyHeight = readbin(f.dyHeight, bin, swapit);
          f.grbit = readbin(f.grbit, bin, swapit);
          f.icv = readbin(f.icv, bin, swapit);
          f.bls = readbin(f.bls, bin, swapit);
          f.sss = readbin(f.sss, bin, swapit);
          f.uls = readbin(f.uls, bin, swapit);
          f.bFamily = readbin(f.bFamily, bin, swapit);
          f.bCharSet = readbin(f.bCharSet, bin, swapit);
          bin.seekg(1, std::ios::cur);
          f.name = XLUnicodeString(bin, swapit, false);
          fonts.push_back(f);
          break;
        }

        case XlsFormat: {
          uint16_t ifmt = 0;
          ifmt = readbin(ifmt, bin, swapit);
          formats[ifmt] = XLUnicodeString(bin, swapit, true);
          break;
        }

        case XlsXF: {
          // font, number format, flags (locked, hidden, style xf), alignment,
          // rotation, indent, used attributes, border styles, border colors,
          // fill pattern and the colors of the fill
          Xf x;
          uint8_t used = 0;
          x.ifnt = readbin(x.ifnt, bin, swapit);
          x.ifmt = readbin(x.ifmt, bin, swapit);
          x.flags = readbin(x.flags, bin, swapit);
          x.align = readbin(x.align, bin, swapit);
          x.rot = readbin(x.rot, bin, swapit);
          x.indent = readbin(x.indent, bin, swapit);
          used = readbin(used, bin, swapit);
          x.border = readbin(x.border, bin, swapit);
          x.border_icv = readbin(x.border_icv, bin, swapit);
          x.border_icv2 = readbin(x.border_icv2, bin, swapit);
          x.fill = readbin(x.fill, bin, swapit);
          xfs.push_back(x);
          break;
        }

        case XlsStyle: {
          NamedStyle st;
          uint16_t word = 0;
          word = readbin(word, bin, swapit);
          st.ixfe = word & 0x0FFF;
          st.builtin = (word & 0x8000) != 0;
          if (st.builtin) {
            uint8_t id = 0, level = 0;
            id = readbin(id, bin, swapit);
            level = readbin(level, bin, swapit);
            st.istyBuiltIn = id;
            st.iLevel = level;
          } else {
            st.name = XLUnicodeString(bin, swapit, true);
          }
          named_styles.push_back(st);
          break;
        }

        case XlsStyleExt: {
          std::vector<uint8_t> rec(rsize);
          bin.read(reinterpret_cast<char*>(rec.data()), rsize);
          if (rsize < 18) break;
          bool ext_builtin = (rec[12] & 0x01) != 0;
          int ext_id = rec[14];
          if (!ext_builtin || named_styles.empty()) break;
          size_t cch = static_cast<size_t>(rec[16] | (rec[17] << 8));
          if (rec.size() < 18 + 2 * cch) break;
          std::u16string chars;
          for (size_t c = 0; c < cch; ++c) chars += static_cast<char16_t>(rec[18 + 2 * c] | (rec[19 + 2 * c] << 8));
          std::string ext_name = to_utf8(chars);
          for (size_t k = 0; k < named_styles.size(); ++k) {
            if (!named_styles[k].builtin && named_styles[k].name == ext_name) {
              named_styles[k].istyBuiltIn = ext_id;
              break;
            }
          }
          break;
        }

        case XlsXFExt: {
          std::vector<uint8_t> rec(rsize);
          bin.read(reinterpret_cast<char*>(rec.data()), rsize);
          if (rsize < 20) break;
          uint16_t ixfe = static_cast<uint16_t>(rec[14] | (rec[15] << 8));
          uint16_t cexts = static_cast<uint16_t>(rec[18] | (rec[19] << 8));
          uint16_t indent = 0;
          xls_extprops(rec, 20, cexts, xf_colors[ixfe], indent);
          if (indent > 0) xf_indent[ixfe] = indent;
          break;
        }

        case XlsPalette: {
          uint16_t ccv = 0;
          ccv = readbin(ccv, bin, swapit);
          for (size_t i = 0; i < ccv && i < palette.size(); ++i) {
            uint8_t r = 0, g = 0, b = 0, reserved = 0;
            r = readbin(r, bin, swapit);
            g = readbin(g, bin, swapit);
            b = readbin(b, bin, swapit);
            reserved = readbin(reserved, bin, swapit);
            palette[i] = (static_cast<uint32_t>(r) << 16) | (static_cast<uint32_t>(g) << 8) | b;
          }
          break;
        }

        default:
          break;
      }

      bin.seekg(next, std::ios_base::beg);
    }

    if (fonts.empty()) fonts.push_back(XlsFontInfo());
    if (xfs.empty()) xfs.push_back(Xf());

    // a text color in XFExt is a font of its own, since fonts and cell formats are separate in xlsx
    std::vector<size_t> xf_font(xfs.size(), 0);
    {
      std::map<std::pair<size_t, std::string>, size_t> derived;
      for (size_t i = 0; i < xfs.size(); ++i) {
        size_t fontId = xfs[i].ifnt;
        if (fontId > 4) --fontId;
        if (fontId >= fonts.size()) fontId = 0;
        xf_font[i] = fontId;

        std::map<uint16_t, std::map<uint16_t, XlsCfColor> >::const_iterator ext = xf_colors.find(static_cast<uint16_t>(i));
        if (ext == xf_colors.end()) continue;
        std::map<uint16_t, XlsCfColor>::const_iterator text = ext->second.find(0x0D);
        if (text == ext->second.end()) continue;
        std::string attr = xls_color_attr(text->second, palette);
        if (attr.empty()) continue;

        std::pair<size_t, std::string> key(fontId, attr);
        if (derived.find(key) == derived.end()) {
          XlsFontInfo f = fonts[fontId];
          f.color_attr = attr;
          fonts.push_back(f);
          derived[key] = fonts.size() - 1;
        }
        xf_font[i] = derived[key];
      }
    }

    out << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" << std::endl;
    out << "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" << std::endl;

    size_t n_formats = 0;
    for (std::map<uint16_t, std::string>::const_iterator it = formats.begin(); it != formats.end(); ++it)
      if (it->first >= 164) ++n_formats;
    if (n_formats > 0) {
      out << "<numFmts count=\"" << n_formats << "\">" << std::endl;
      for (std::map<uint16_t, std::string>::const_iterator it = formats.begin(); it != formats.end(); ++it) {
        if (it->first < 164) continue;
        out << "<numFmt numFmtId=\"" << it->first << "\" formatCode=\"" << escape_xml(it->second) << "\"/>" << std::endl;
      }
      out << "</numFmts>" << std::endl;
    }

    out << "<fonts count=\"" << fonts.size() << "\">" << std::endl;
    for (size_t i = 0; i < fonts.size(); ++i) {
      out << fonts[i].xml(palette, "font", "name") << std::endl;
    }
    out << "</fonts>" << std::endl;

    // fills and borders are not listed in an xls file, every XF has its own
    // description. Here the same ones are stored once.
    std::vector<std::string> fills;
    fills.push_back("<fill><patternFill patternType=\"none\"/></fill>");
    fills.push_back("<fill><patternFill patternType=\"gray125\"/></fill>");
    std::vector<std::string> borders;
    std::vector<size_t> fill_id(xfs.size(), 0), border_id(xfs.size(), 0);

    static const char* const sides[5] = {"left", "right", "top", "bottom", "diagonal"};
    for (size_t i = 0; i < xfs.size(); ++i) {
      const Xf& x = xfs[i];

      uint32_t pattern = x.border_icv2 >> 26;
      std::string fill = fills[0];
      if (pattern > 0 && pattern < 19) {
        std::string fg = " rgb=\"" + xls_argb(palette, x.fill & 0x7Fu) + "\"";
        std::string bg = " rgb=\"" + xls_argb(palette, (x.fill >> 7) & 0x7Fu) + "\"";
        std::map<uint16_t, std::map<uint16_t, XlsCfColor> >::const_iterator ext = xf_colors.find(static_cast<uint16_t>(i));
        if (ext != xf_colors.end()) {
          std::map<uint16_t, XlsCfColor>::const_iterator c = ext->second.find(0x04);
          if (c != ext->second.end() && !xls_color_attr(c->second, palette).empty()) fg = xls_color_attr(c->second, palette);
          c = ext->second.find(0x05);
          if (c != ext->second.end() && !xls_color_attr(c->second, palette).empty()) bg = xls_color_attr(c->second, palette);
        }
        fill = std::string("<fill><patternFill patternType=\"") + XlsFillPatterns[pattern] + "\"><fgColor" + fg +
               "/><bgColor" + bg + "/></patternFill></fill>";
      }
      fill_id[i] = static_cast<size_t>(std::find(fills.begin(), fills.end(), fill) - fills.begin());
      if (fill_id[i] == fills.size()) fills.push_back(fill);

      // style of left, right, top, bottom: 4 bits each. Colors of left and
      // right, and of top, bottom and the diagonal: 7 bits each
      uint32_t b1 = x.border, b2 = x.border_icv, b3 = x.border_icv2;
      uint32_t style[5] = {b1 & 0xFu, (b1 >> 4) & 0xFu, (b1 >> 8) & 0xFu, (b1 >> 12) & 0xFu, (b3 >> 21) & 0xFu};
      uint32_t color[5] = {b2 & 0x7Fu, (b2 >> 7) & 0x7Fu, b3 & 0x7Fu, (b3 >> 7) & 0x7Fu, (b3 >> 14) & 0x7Fu};
      bool diag_down = (b2 & 0x4000u) != 0;
      bool diag_up = (b2 & 0x8000u) != 0;
      if (!diag_down && !diag_up) style[4] = 0;

      std::string border = "<border";
      if (diag_up) border += " diagonalUp=\"1\"";
      if (diag_down) border += " diagonalDown=\"1\"";
      border += ">";
      for (size_t s = 0; s < 5; ++s) {
        if (style[s] > 0 && style[s] < 14) {
          border += std::string("<") + sides[s] + " style=\"" + XlsBorderStyles[style[s]] + "\">";
          // types of XFExt: top 7, bottom 8, left 9, right 10, diagonal 11
          static const uint16_t ext_type[5] = {0x09, 0x0A, 0x07, 0x08, 0x0B};
          std::string exact;
          std::map<uint16_t, std::map<uint16_t, XlsCfColor> >::const_iterator ext = xf_colors.find(static_cast<uint16_t>(i));
          if (ext != xf_colors.end()) {
            std::map<uint16_t, XlsCfColor>::const_iterator c = ext->second.find(ext_type[s]);
            if (c != ext->second.end()) exact = xls_color_attr(c->second, palette);
          }
          if (!exact.empty()) border += "<color" + exact + "/>";
          else if (color[s] == 64) border += "<color indexed=\"64\"/>";
          else border += "<color rgb=\"" + xls_argb(palette, color[s]) + "\"/>";
          border += std::string("</") + sides[s] + ">";
        } else {
          border += std::string("<") + sides[s] + "/>";
        }
      }
      border += "</border>";
      border_id[i] = static_cast<size_t>(std::find(borders.begin(), borders.end(), border) - borders.begin());
      if (border_id[i] == borders.size()) borders.push_back(border);
    }

    out << "<fills count=\"" << fills.size() << "\">" << std::endl;
    for (size_t i = 0; i < fills.size(); ++i) out << fills[i] << std::endl;
    out << "</fills>" << std::endl;

    out << "<borders count=\"" << borders.size() << "\">" << std::endl;
    for (size_t i = 0; i < borders.size(); ++i) out << borders[i] << std::endl;
    out << "</borders>" << std::endl;

    std::vector<size_t> style_xf_index(xfs.size(), 0);
    std::vector<size_t> style_xfs;
    for (size_t i = 0; i < xfs.size(); ++i) {
      if (xfs[i].flags & 0x0004) {
        style_xf_index[i] = style_xfs.size();
        style_xfs.push_back(i);
      }
    }
    if (style_xfs.empty()) style_xfs.push_back(0);

    static const char* const horizontal[8] = {"", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"};
    static const char* const vertical[5] = {"top", "center", "", "justify", "distributed"};

    auto write_xf = [&](size_t i, bool is_cell_xf) {
      const Xf& x = xfs[i];

      uint32_t numFmtId = x.ifmt;
      if (numFmtId >= 164 && formats.find(x.ifmt) == formats.end()) numFmtId = 0;

      size_t fontId = xf_font[i];

      std::string align;
      unsigned alc = x.align & 0x07u, alcv = (x.align >> 4) & 0x07u, indent = x.indent & 0x0Fu;
      if (xf_indent.find(static_cast<uint16_t>(i)) != xf_indent.end()) indent = xf_indent[static_cast<uint16_t>(i)];
      if (alc > 0) align += std::string(" horizontal=\"") + horizontal[alc] + "\"";
      if (alcv < 5 && alcv != 2) align += std::string(" vertical=\"") + vertical[alcv] + "\"";
      if (x.align & 0x08) align += " wrapText=\"1\"";
      if (x.rot != 0) align += " textRotation=\"" + std::to_string(x.rot) + "\"";
      if (indent > 0) align += " indent=\"" + std::to_string(indent) + "\"";
      if (x.indent & 0x10) align += " shrinkToFit=\"1\"";

      std::string prot;
      bool locked = (x.flags & 0x0001) != 0, hidden = (x.flags & 0x0002) != 0;
      if (!locked || hidden)
        prot = std::string("<protection locked=\"") + (locked ? "1" : "0") + "\" hidden=\"" + (hidden ? "1" : "0") + "\"/>";

      out << "<xf numFmtId=\"" << numFmtId << "\" fontId=\"" << fontId << "\" fillId=\"" << fill_id[i]
          << "\" borderId=\"" << border_id[i] << "\"";
      if (is_cell_xf) {
        size_t parent = x.flags >> 4;
        size_t xfId = (parent < xfs.size() && (xfs[parent].flags & 0x0004)) ? style_xf_index[parent] : 0;
        out << " xfId=\"" << xfId << "\"";
      }
      out << " applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\""
          << (prot.empty() ? "" : " applyProtection=\"1\"");
      if (align.empty() && prot.empty()) {
        out << "/>" << std::endl;
      } else {
        out << ">";
        if (!align.empty()) out << "<alignment" << align << "/>";
        out << prot << "</xf>" << std::endl;
      }
    };

    out << "<cellStyleXfs count=\"" << style_xfs.size() << "\">" << std::endl;
    for (size_t k = 0; k < style_xfs.size(); ++k) write_xf(style_xfs[k], false);
    out << "</cellStyleXfs>" << std::endl;

    out << "<cellXfs count=\"" << xfs.size() << "\">" << std::endl;
    for (size_t i = 0; i < xfs.size(); ++i) write_xf(i, true);
    out << "</cellXfs>" << std::endl;

    static const char* const builtin_names[10] = {"Normal", "RowLevel_", "ColLevel_", "Comma", "Currency", "Percent",
                                                  "Comma [0]", "Currency [0]", "Hyperlink", "Followed Hyperlink"};

    std::string cell_styles;
    size_t n_cell_styles = 0;
    std::set<std::string> style_seen;
    for (size_t k = 0; k < named_styles.size(); ++k) {
      const NamedStyle& st = named_styles[k];
      if (st.ixfe >= xfs.size() || !(xfs[st.ixfe].flags & 0x0004)) continue;
      std::string name = st.name;
      if (st.builtin) {
        if (st.istyBuiltIn < 0 || st.istyBuiltIn > 9) continue;
        name = builtin_names[st.istyBuiltIn];
        if (st.istyBuiltIn == 1 || st.istyBuiltIn == 2) name += std::to_string(st.iLevel + 1);
      }
      if (name.empty() || !style_seen.insert(name).second) continue;

      std::string xml_name = escape_xml(name);
      cell_styles += "<cellStyle name=\"" + xml_name + "\" xfId=\"" + std::to_string(style_xf_index[st.ixfe]) + "\"";
      if (st.istyBuiltIn >= 0) cell_styles += " builtinId=\"" + std::to_string(st.istyBuiltIn) + "\"";
      if (st.builtin && (st.istyBuiltIn == 1 || st.istyBuiltIn == 2)) cell_styles += " iLevel=\"" + std::to_string(st.iLevel) + "\"";
      cell_styles += "/>";
      ++n_cell_styles;
    }
    if (n_cell_styles == 0) {
      cell_styles = "<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>";
      n_cell_styles = 1;
    }
    out << "<cellStyles count=\"" << n_cell_styles << "\">" << cell_styles << "</cellStyles>" << std::endl;

    // The formatting of conditional formats is in the CF records on the sheets.
    // They are numbered in the order of the stream, the sheets use the same numbers.
    std::vector<std::string> dxfs;
    std::vector<std::vector<uint8_t>> cf_recs;
    size_t cf_base = 0;
    bin.clear();
    bin.seekg(0, std::ios_base::beg);
    while (XLS_RECORD(rid, rsize, bin, swapit)) {
      std::streampos next = bin.tellg() + std::streamoff(rsize);
      if (rid == XlsCondFmt) cf_base = dxfs.size();
      if (rid == XlsCF) {
        std::vector<uint8_t> rec(rsize);
        bin.read(reinterpret_cast<char*>(rec.data()), rsize);
        XlsCfRule cf;
        if (xls_cf(rec, palette, formats, 164 + static_cast<int>(dxfs.size()), cf)) dxfs.push_back(cf.dxf);
        else dxfs.push_back("<dxf/>");
        cf_recs.push_back(rec);
      }
      if (rid == XlsCFEx && rsize > 30) {
        // full colors of a CF rule: fIsCF12 (4 bytes), nID, icf, cp, template, priority, flags, fHasDXF, DXFN12
        std::vector<uint8_t> rec(rsize);
        bin.read(reinterpret_cast<char*>(rec.data()), rsize);
        auto u16 = [&](size_t at) { return static_cast<size_t>(rec[at] | (rec[at + 1] << 8)); };
        size_t idx = cf_base + u16(18);
        size_t cb_dxf = u16(26) | (u16(28) << 16);
        if (u16(12) == 0 && u16(14) == 0 && rec[25] == 1 && idx < cf_recs.size() && !cf_recs[idx].empty() && 30 + cb_dxf <= rec.size()) {
          size_t at = 30;
          std::string unused;
          if (xls_dxfn(rec, at, palette, formats, 164, unused) && at + 8 <= 30 + cb_dxf) {
            std::map<uint16_t, XlsCfColor> colors;
            uint16_t indent = 0;
            xls_extprops(rec, at + 8, u16(at + 6), colors, indent);
            XlsCfRule cf;
            if (!colors.empty() && xls_cf(cf_recs[idx], palette, formats, 164 + static_cast<int>(idx), cf, &colors)) dxfs[idx] = cf.dxf;
          }
        }
      }
      if (rid == XlsCF12) {
        std::vector<uint8_t> rec(rsize);
        bin.read(reinterpret_cast<char*>(rec.data()), rsize);
        XlsCf12Rule cf;
        // only rules on a value or a formula have a dxf
        if (rec.size() > 12 && (rec[12] == 1 || rec[12] == 2)) {
          if (xls_cf12(rec, palette, formats, 164 + static_cast<int>(dxfs.size()), cf) && cf.has_dxf) dxfs.push_back(cf.dxf);
          else dxfs.push_back("<dxf/>");
          cf_recs.emplace_back();
        }
      }
      bin.clear();
      bin.seekg(next, std::ios_base::beg);
    }
    if (!dxfs.empty()) {
      out << "<dxfs count=\"" << dxfs.size() << "\">" << std::endl;
      for (size_t i = 0; i < dxfs.size(); ++i) out << dxfs[i] << std::endl;
      out << "</dxfs>" << std::endl;
    }

    out << "</styleSheet>" << std::endl;

    out.close();
    bin.close();
    return 1;
  } else {
    return -1;
  }
}

// type of each sheet: 0 worksheet, 2 chart sheet
// [[Rcpp::export]]
Rcpp::IntegerVector sheettypes_xls(std::string filePath) {
  std::ifstream bin(filePath, std::ios::in | std::ios::binary | std::ios::ate);
  bool swapit = is_big_endian();
  Rcpp::IntegerVector types;
  if (bin) {
    bin.seekg(0, std::ios_base::beg);
    XlsGlobals g = xls_globals(bin, swapit, false);
    for (size_t i = 0; i < g.sheets.size(); ++i) types.push_back(g.sheets[i].type == 0 ? 0 : 2);
  }
  return types;
}

// sheetN.xml for the sheet with the number sheet (counting from 1). Cells are
// written as they are read, with the XML around them: the rows come from the
// ROW records, that come before the cells of the row.
// [[Rcpp::export]]
Rcpp::IntegerVector worksheet_xls(std::string filePath, int32_t sheet, std::string outPath, std::string relsPath, std::string commentsPath,
                                  std::string vmlPath, int32_t comment_file, std::string tablesDir, int32_t table_file,
                                  Rcpp::CharacterVector media, std::string drawingPath, std::string drawingRelsPath,
                                  int32_t drawing_file, std::string chartsDir, int32_t chart_file, std::string xlDir, std::string xlsPath, int32_t pivot_file,
                                  Rcpp::IntegerVector known_caches, bool data_only, bool debug) {
  std::ofstream out(outPath);
  std::ifstream bin(filePath, std::ios::in | std::ios::binary | std::ios::ate);

  bool swapit = is_big_endian();

  if (bin) {
    bin.seekg(0, std::ios_base::beg);
    XlsGlobals g = xls_globals(bin, swapit, debug);

    if (sheet < 1 || static_cast<size_t>(sheet) > g.sheets.size()) Rcpp::stop("xls: sheet does not exist");

    out << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" << std::endl;

    // a chart sheet: the substream of the sheet is the chart. It is drawn by a
    // drawing that has the chart as its only object.
    if (g.sheets[static_cast<size_t>(sheet) - 1].type != 0) {
      bin.clear();
      bin.seekg(static_cast<std::streamoff>(g.sheets[static_cast<size_t>(sheet) - 1].offset), std::ios_base::beg);
      std::vector<XlsChartRecord> recs;
      int cdepth = 0;
      uint16_t crid = 0, crsize = 0;
      while (XLS_RECORD(crid, crsize, bin, swapit)) {
        std::streampos cnext = bin.tellg() + std::streamoff(crsize);
        if (crid == XlsBOF) {
          ++cdepth;
        } else if (crid == XlsEOF) {
          --cdepth;
          if (cdepth <= 0) break;
        } else if (cdepth == 1) {
          XlsChartRecord cr;
          cr.id = crid;
          cr.d.resize(crsize);
          bin.read(reinterpret_cast<char*>(cr.d.data()), crsize);
          recs.push_back(cr);
        }
        bin.seekg(cnext, std::ios_base::beg);
      }
      XlsFormulaContext chart_ctx = g.ctx;
      std::string cx = data_only ? std::string() : xls_chart_xml(recs, chart_ctx, swapit, debug);

      out << "<chartsheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
          << "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" << std::endl;
      out << "<sheetViews><sheetView workbookViewId=\"0\" zoomToFit=\"1\"/></sheetViews>" << std::endl;
      if (cx.empty()) {
        out << "</chartsheet>" << std::endl;
        return Rcpp::IntegerVector::create(0, 0, 0, 0, 0);
      }
      out << "<drawing r:id=\"rId1\"/>" << std::endl << "</chartsheet>" << std::endl;

      std::string cname = "chart" + std::to_string(chart_file) + ".xml";
      std::ofstream cf(chartsDir + "/" + cname);
      cf << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" << std::endl << cx << std::endl;
      cf.close();

      std::ofstream dr(drawingPath);
      dr << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" << std::endl;
      dr << "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
         << "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
         << "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
         << "<xdr:absoluteAnchor><xdr:pos x=\"0\" y=\"0\"/><xdr:ext cx=\"9525000\" cy=\"6350000\"/>"
         << "<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"2\" name=\"Chart 1\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>"
         << "<xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm>"
         << "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">"
         << "<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"rId1\"/>"
         << "</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>" << std::endl;
      dr.close();

      std::ofstream drel(drawingRelsPath);
      drel << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" << std::endl;
      drel << "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" << std::endl;
      drel << "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" "
           << "Target=\"../charts/" << cname << "\"/>" << std::endl << "</Relationships>" << std::endl;
      drel.close();

      std::ofstream rel(relsPath);
      rel << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" << std::endl;
      rel << "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" << std::endl;
      rel << "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" "
          << "Target=\"../drawings/drawing" << drawing_file << ".xml\"/>" << std::endl << "</Relationships>" << std::endl;
      return Rcpp::IntegerVector::create(0, 0, 1, 1, 0);
    }

    out << "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
        << "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" << std::endl;

    struct Row {
      bool custom = false, hidden = false;
      double height = 0;
      uint8_t outline = 0;
    };
    // everything about a cell that is written to the file
    struct Cell {
      uint32_t row = 0, col = 0, xf = 0;
      std::string type;       // the t attribute
      std::string formula;    // text of the formula, empty if none
      std::string formula_attr;
      std::string body;       // value, after the formula
    };

    std::map<uint32_t, Row> rows;
    std::map<uint16_t, double> col_px;  // widths of columns in pixels

    // pictures and charts: the shapes of the last drawing record wait for the object record
    struct DrawingItem {
      bool chart = false;
      std::string chart_xml;
      bool chart_done = false;
      bool autoshape = false;
      uint32_t blip = 0;
      XlsShape shape;
    };
    std::vector<XlsShape> pending_shapes;
    std::vector<DrawingItem> items;
    std::vector<XlsChartRecord> chart_recs;
    std::vector<XlsPivotView> pv_views;
    bool in_chart = false;
    bool after_obj = false;  // CONTINUE records that follow an object may hold the shape of the next one
    std::vector<std::string> merges;
    std::vector<XlsHyperlink> hlinks;

    // view, print and page settings
    std::string tab_color;
    bool fit_to_page = false, have_window = false, have_pane = false;
    uint16_t window_flags = 0, zoom = 0;
    uint16_t pane_x = 0, pane_y = 0, pane_row = 0, pane_col = 0;
    double margins[6] = {0.7, 0.7, 0.75, 0.75, 0.3, 0.3};
    bool have_margin = false;
    uint16_t paper = 0, scale = 100, page_start = 0, fit_w = 1, fit_h = 1, setup_flags = 0, copies = 1;
    bool have_setup = false;
    std::string code_name;
    std::string dimension_ref;
    std::string ignored_errors;
    int sort_sfp = -1;
    std::string sort_ref, sort_conditions;
    uint32_t sort_pending = 0;
    bool outline_below = true, outline_right = true, apply_outline_styles = false;
    bool have_row_height = false, row_height_custom = false, row_height_zero = false;
    double default_row_height = 15;
    int default_col_chars = -1;
    uint16_t outline_row = 0, outline_col = 0;
    std::string header, footer;
    std::string even_header, even_footer, first_header, first_footer;
    bool hf_diff_even = false, hf_diff_first = false, hf_no_scale = false, hf_no_align = false;
    struct Selection {
      uint8_t pane = 3;
      uint16_t row = 0, col = 0;
      std::string sqref;
    };
    std::vector<Selection> selections;
    bool h_center = false, v_center = false, print_grid = false, print_headers = false;
    std::map<uint16_t, std::string> filter_columns;
    bool have_autofilter = false, protect_sheet = false, protect_objects = false, protect_scenarios = false;
    uint16_t password = 0;
    std::vector<uint16_t> row_breaks, col_breaks;
    std::vector<XlsListTable> tables;
    std::vector<std::string> validations;

    // conditional formats: CondFmt has the ranges and how many CF records follow
    struct CondFmt {
      std::string sqref;
      uint32_t row = 0, col = 0;  // top left cell, formulas are relative to it
      size_t pending = 0;         // CF records still to come
      std::string xml;
      std::string x14;
    };
    std::vector<CondFmt> cond_fmts;
    uint32_t x14_id = 0;
    int32_t cf_priority = 0;
    int32_t dxf_id = 0;

    // Comments are spread over three records. Note has the cell, the author and
    // the number of the object. Obj is the object (a note has type 0x19) with
    // that number, and the TxO that follows it has the text.
    struct Note {
      uint16_t row = 0, col = 0, id = 0;
      std::string author;
    };
    std::vector<Note> notes;
    std::map<uint16_t, std::u16string> note_text;
    std::map<uint16_t, std::vector<std::pair<uint16_t, uint16_t> > > note_runs;  // start of the run and font
    std::map<uint16_t, XlsShape> note_shapes;
    int32_t last_note_object = -1;
    std::map<std::string, std::string> tooltips;  // by range
    std::stringstream cols;
    // The formulas of shared and array formulas are stored once. All other cells
    // of the formula refer to the first cell with a PtgExp token.
    std::map<std::pair<uint32_t, uint32_t>, XlsRgce> shared_formulas;
    std::map<std::pair<uint32_t, uint32_t>, std::pair<std::string, XlsRgce> > array_formulas;

    bool sheet_data_open = false;
    bool row_open = false;
    int64_t current_row = -1;

    // a formula cell waits for the records that can follow it: the text result
    // (String) and the formula it shares with other cells (ShrFmla, Array)
    bool has_pending = false;
    Cell pending;
    XlsRgce pending_rgce;

    int n_unsupported = 0;
    int first_unsupported = -1;

    auto row_attributes = [](const Row& r) {
      std::string attr;
      if (r.custom) attr += " ht=\"" + xls_num(r.height) + "\" customHeight=\"1\"";
      if (r.hidden) attr += " hidden=\"1\"";
      if (r.outline) attr += " outlineLevel=\"" + std::to_string(r.outline) + "\"";
      return attr;
    };

    // Rows without cells are only written when they have a height, are hidden or
    // grouped. Writes those between the current row and the row `until`.
    auto write_empty_rows = [&](uint32_t until) {
      std::map<uint32_t, Row>::const_iterator it = current_row < 0 ? rows.begin() : rows.upper_bound(static_cast<uint32_t>(current_row));
      for (; it != rows.end() && it->first < until; ++it) {
        std::string attr = row_attributes(it->second);
        if (!attr.empty()) out << "<row r=\"" << it->first + 1 << "\"" << attr << "/>" << std::endl;
      }
    };

    auto open_sheet_data = [&]() {
      if (sheet_data_open) return;
      bool outline_pr = !outline_below || !outline_right || apply_outline_styles;
      if (fit_to_page || !tab_color.empty() || !code_name.empty() || outline_pr) {
        out << "<sheetPr";
        if (!code_name.empty()) out << " codeName=\"" << escape_xml(code_name) << "\"";
        out << ">";
        if (!tab_color.empty()) out << "<tabColor" << tab_color << "/>";
        if (outline_pr) {
          out << "<outlinePr";
          if (apply_outline_styles) out << " applyStyles=\"1\"";
          if (!outline_below) out << " summaryBelow=\"0\"";
          if (!outline_right) out << " summaryRight=\"0\"";
          out << "/>";
        }
        if (fit_to_page) out << "<pageSetUpPr fitToPage=\"1\"/>";
        out << "</sheetPr>" << std::endl;
      }
      if (!dimension_ref.empty()) out << "<dimension ref=\"" << dimension_ref << "\"/>" << std::endl;
      if (have_window) {
        out << "<sheetViews><sheetView";
        if (!(window_flags & 0x0002)) out << " showGridLines=\"0\"";
        if (!(window_flags & 0x0004)) out << " showRowColHeaders=\"0\"";
        if (!(window_flags & 0x0010)) out << " showZeros=\"0\"";
        if (window_flags & 0x0200) out << " tabSelected=\"1\"";
        if (zoom > 0 && zoom != 100) out << " zoomScale=\"" << zoom << "\" zoomScaleNormal=\"" << zoom << "\"";
        out << " workbookViewId=\"0\"";
        bool has_pane = have_pane && (window_flags & 0x0008) && (pane_x > 0 || pane_y > 0);
        if (!has_pane && selections.empty()) {
          out << "/></sheetViews>" << std::endl;
        } else {
          out << ">";
          if (has_pane) {
            out << "<pane";
            if (pane_x > 0) out << " xSplit=\"" << pane_x << "\"";
            if (pane_y > 0) out << " ySplit=\"" << pane_y << "\"";
            out << " topLeftCell=\"" << int_to_col(pane_col + 1u) << pane_row + 1u << "\" activePane=\""
                << (pane_x > 0 && pane_y > 0 ? "bottomRight" : (pane_y > 0 ? "bottomLeft" : "topRight"))
                << "\" state=\"frozen\"/>";
          }
          static const char* const pane_names[4] = {"bottomRight", "topRight", "bottomLeft", "topLeft"};
          for (size_t k = 0; k < selections.size(); ++k) {
            const Selection& sel = selections[k];
            if (!has_pane && sel.pane != 3) continue;
            out << "<selection";
            if (sel.pane < 3) out << " pane=\"" << pane_names[sel.pane] << "\"";
            out << " activeCell=\"" << int_to_col(sel.col + 1u) << sel.row + 1u << "\"";
            if (!sel.sqref.empty()) out << " sqref=\"" << sel.sqref << "\"";
            out << "/>";
          }
          out << "</sheetView></sheetViews>" << std::endl;
        }
      }
      if (have_row_height || outline_row > 1 || outline_col > 1 || (default_col_chars >= 0 && default_col_chars != 8)) {
        out << "<sheetFormatPr";
        if (default_col_chars >= 0 && default_col_chars != 8) out << " baseColWidth=\"" << default_col_chars << "\"";
        out << " defaultRowHeight=\"" << xls_num(default_row_height) << "\"";
        if (row_height_custom) out << " customHeight=\"1\"";
        if (row_height_zero) out << " zeroHeight=\"1\"";
        if (outline_row > 1) out << " outlineLevelRow=\"" << outline_row - 1 << "\"";
        if (outline_col > 1) out << " outlineLevelCol=\"" << outline_col - 1 << "\"";
        out << "/>" << std::endl;
      }
      if (!cols.str().empty()) out << "<cols>" << std::endl << cols.str() << "</cols>" << std::endl;
      out << "<sheetData>" << std::endl;
      sheet_data_open = true;
    };

    // Writes a cell. A cell in a new row closes the row before it.
    auto write_cell = [&](const Cell& c) {
      open_sheet_data();

      if (static_cast<int64_t>(c.row) != current_row) {
        if (row_open) out << "</row>" << std::endl;
        write_empty_rows(c.row);
        std::map<uint32_t, Row>::const_iterator it = rows.find(c.row);
        out << "<row r=\"" << c.row + 1 << "\"" << (it == rows.end() ? "" : row_attributes(it->second)) << ">" << std::endl;
        row_open = true;
        current_row = c.row;
      }

      out << "<c r=\"" << int_to_col(c.col + 1) << c.row + 1 << "\"";
      if (c.xf) out << " s=\"" << c.xf << "\"";
      if (!c.type.empty()) out << " t=\"" << c.type << "\"";
      if (c.formula.empty() && c.body.empty()) {
        out << "/>" << std::endl;
      } else {
        out << ">";
        if (!c.formula.empty()) out << "<f" << c.formula_attr << ">" << escape_xml(c.formula) << "</f>";
        out << c.body << "</c>" << std::endl;
      }
    };

    // The formula of the pending cell is written when it is complete.
    auto finish_pending = [&]() {
      if (!has_pending) return;
      has_pending = false;

      XlsFormulaContext& ctx = g.ctx;
      ctx.row = pending.row;
      ctx.col = pending.col;
      ctx.unsupported = -1;

      // the first token PtgExp says that the formula is stored in a ShrFmla or Array record
      XlsRgce* rgce = &pending_rgce;
      if (pending_rgce.cce >= 5 && pending_rgce.data[0] == PtgExp) {
        std::pair<uint32_t, uint32_t> key(
            static_cast<uint32_t>(pending_rgce.data[1]) | (static_cast<uint32_t>(pending_rgce.data[2]) << 8),
            static_cast<uint32_t>(pending_rgce.data[3]) | (static_cast<uint32_t>(pending_rgce.data[4]) << 8));
        rgce = nullptr;
        if (shared_formulas.count(key)) {
          rgce = &shared_formulas[key];
        } else if (array_formulas.count(key) && key.first == pending.row && key.second == pending.col) {
          // only the first cell of an array formula has the formula
          rgce = &array_formulas[key].second;
          pending.formula_attr = " t=\"array\" ref=\"" + array_formulas[key].first + "\"";
        }
      }

      if (rgce != nullptr) {
        pending.formula = xls_rgce(*rgce, ctx, swapit, debug);
        if (pending.formula.empty() && ctx.unsupported >= 0) {
          if (n_unsupported == 0) first_unsupported = ctx.unsupported;
          ++n_unsupported;
          pending.formula_attr.clear();
        }
      }
      write_cell(pending);
    };

    // the formatting of the rules is in the styles, numbered over all sheets
    {
      bin.clear();
      bin.seekg(0, std::ios_base::beg);
      uint16_t crid = 0, crsize = 0;
      std::streampos sheet_start = static_cast<std::streamoff>(g.sheets[static_cast<size_t>(sheet) - 1].offset);
      while (bin.tellg() < sheet_start && XLS_RECORD(crid, crsize, bin, swapit)) {
        std::streampos cnext = bin.tellg() + std::streamoff(crsize);
        if (crid == XlsCF) ++dxf_id;
        if (crid == XlsCF12 && crsize > 12) {
          uint8_t ct = 0;
          bin.seekg(12, std::ios::cur);
          ct = readbin(ct, bin, swapit);
          if (ct == 1 || ct == 2) ++dxf_id;
        }
        bin.seekg(cnext, std::ios_base::beg);
      }
    }

    // the settings of the view come after the cells, but are written before them
    {
      bin.clear();
      bin.seekg(static_cast<std::streamoff>(g.sheets[static_cast<size_t>(sheet) - 1].offset), std::ios_base::beg);
      int pdepth = 0;
      uint16_t prid = 0, prsize = 0;
      while (XLS_RECORD(prid, prsize, bin, swapit)) {
        std::streampos pnext = bin.tellg() + std::streamoff(prsize);
        if (prid == XlsBOF) ++pdepth;
        if (prid == XlsEOF && --pdepth <= 0) break;
        if (pdepth == 1) {
          switch (prid) {
            case XlsWsBool: {
              uint16_t flags = 0;
              flags = readbin(flags, bin, swapit);
              fit_to_page = (flags & 0x0100) != 0;
              outline_below = (flags & 0x0040) != 0;
              outline_right = (flags & 0x0080) != 0;
              apply_outline_styles = (flags & 0x0020) != 0;
              break;
            }

            case XlsCodeName: {
              if (prsize > 0) code_name = XLUnicodeString(bin, swapit, true);
              break;
            }

            case XlsDimensions: {
              uint32_t row_first = 0, row_last = 0;
              uint16_t col_first = 0, col_last = 0;
              row_first = readbin(row_first, bin, swapit);
              row_last = readbin(row_last, bin, swapit);
              col_first = readbin(col_first, bin, swapit);
              col_last = readbin(col_last, bin, swapit);
              dimension_ref = "A1";
              if (row_last > row_first && col_last > col_first) {
                dimension_ref = int_to_col(static_cast<uint32_t>(col_first) + 1) + std::to_string(row_first + 1);
                if (row_last - row_first > 1 || col_last - col_first > 1)
                  dimension_ref += ":" + int_to_col(static_cast<uint32_t>(col_last)) + std::to_string(row_last);
              }
              break;
            }

            case XlsDefColWidth: {
              uint16_t cch = 0;
              cch = readbin(cch, bin, swapit);
              default_col_chars = cch;
              break;
            }

            case XlsDefaultRowHeight: {
              uint16_t flags = 0, twips = 0;
              flags = readbin(flags, bin, swapit);
              twips = readbin(twips, bin, swapit);
              have_row_height = true;
              row_height_custom = (flags & 0x0001) != 0;
              row_height_zero = (flags & 0x0002) != 0;
              default_row_height = twips / 20.0;
              break;
            }

            case XlsGuts: {
              bin.seekg(4, std::ios_base::cur);
              outline_row = readbin(outline_row, bin, swapit);
              outline_col = readbin(outline_col, bin, swapit);
              break;
            }

            case XlsSheetExt: {
              // frt header (12 bytes), size (4), the tab color as palette index
              // (7 bits), and 20 optional bytes with the color as CFColor
              std::vector<uint8_t> rec(prsize);
              bin.read(reinterpret_cast<char*>(rec.data()), prsize);
              if (rec.size() < 20) break;
              std::vector<uint32_t> pal(XlsDefaultPalette, XlsDefaultPalette + 56);
              unsigned icv = rec[16] & 0x7Fu;
              if (rec.size() >= 40) {
                XlsCfColor c;
                c.type = static_cast<uint32_t>(rec[24] | (rec[25] << 8) | (rec[26] << 16) | (static_cast<uint32_t>(rec[27]) << 24));
                c.value = static_cast<uint32_t>(rec[28] | (rec[29] << 8) | (rec[30] << 16) | (static_cast<uint32_t>(rec[31]) << 24));
                if (c.type == 2 || c.type == 3) tab_color = xls_color_attr(c, pal);
              }
              if (tab_color.empty() && icv >= 8 && icv <= 0x3F) tab_color = " rgb=\"" + xls_argb(pal, icv) + "\"";
              break;
            }

            case XlsWindow2: {
              window_flags = readbin(window_flags, bin, swapit);
              have_window = true;
              break;
            }

            case XlsScl: {
              uint16_t num = 0, denom = 1;
              num = readbin(num, bin, swapit);
              denom = readbin(denom, bin, swapit);
              if (denom > 0) zoom = static_cast<uint16_t>(100u * num / denom);
              break;
            }

            case XlsSelection: {
              // pane, active cell (row, column), index of the active range, the ranges
              Selection sel;
              uint16_t iref = 0, cref = 0;
              sel.pane = readbin(sel.pane, bin, swapit);
              sel.row = readbin(sel.row, bin, swapit);
              sel.col = readbin(sel.col, bin, swapit);
              iref = readbin(iref, bin, swapit);
              cref = readbin(cref, bin, swapit);
              (void)iref;
              for (uint16_t k = 0; k < cref; ++k) {
                uint16_t row1 = 0, row2 = 0;
                uint8_t col1 = 0, col2 = 0;
                row1 = readbin(row1, bin, swapit);
                row2 = readbin(row2, bin, swapit);
                col1 = readbin(col1, bin, swapit);
                col2 = readbin(col2, bin, swapit);
                std::string a = int_to_col(col1 + 1u) + std::to_string(row1 + 1u);
                std::string b = int_to_col(col2 + 1u) + std::to_string(row2 + 1u);
                if (!sel.sqref.empty()) sel.sqref += " ";
                sel.sqref += a == b ? a : a + ":" + b;
              }
              selections.push_back(sel);
              break;
            }

            case XlsHeaderFooter: {
              // header of a FRT record (12), view (16), flags, 4 lengths and the strings, each with a flag for the encoding
              bin.seekg(28, std::ios_base::cur);
              uint16_t flags = 0, len[4] = {0, 0, 0, 0};
              flags = readbin(flags, bin, swapit);
              for (int k = 0; k < 4; ++k) len[k] = readbin(len[k], bin, swapit);
              std::string* target[4] = {&even_header, &even_footer, &first_header, &first_footer};
              std::streampos rec_end = pnext;
              for (int k = 0; k < 4; ++k) {
                if (len[k] == 0) continue;
                uint8_t enc = 0;
                enc = readbin(enc, bin, swapit);
                *target[k] = to_utf8(xls_chars(bin, swapit, len[k], (enc & 0x01) != 0, rec_end));
              }
              hf_diff_even = (flags & 0x0001) != 0;
              hf_diff_first = (flags & 0x0002) != 0;
              hf_no_scale = (flags & 0x0004) == 0;
              hf_no_align = (flags & 0x0008) == 0;
              break;
            }

            case XlsPane: {
              pane_x = readbin(pane_x, bin, swapit);
              pane_y = readbin(pane_y, bin, swapit);
              pane_row = readbin(pane_row, bin, swapit);
              pane_col = readbin(pane_col, bin, swapit);
              have_pane = true;
              break;
            }

            default:
              break;
          }
        }
        bin.seekg(pnext, std::ios_base::beg);
      }
    }

    // the fonts of the workbook for the formatted text of comments
    std::vector<XlsFontInfo> note_fonts;
    std::vector<uint32_t> note_palette(XlsDefaultPalette, XlsDefaultPalette + 56);
    if (!data_only) {
      bin.clear();
      bin.seekg(0, std::ios_base::beg);
      int gdepth = 0;
      uint16_t grid = 0, grsize = 0;
      while (XLS_RECORD(grid, grsize, bin, swapit)) {
        std::streampos gnext = bin.tellg() + std::streamoff(grsize);
        if (grid == XlsBOF) ++gdepth;
        if (grid == XlsEOF) --gdepth;
        if (gdepth <= 0) break;
        if (grid == XlsFont) {
          XlsFontInfo f;
          f.dyHeight = readbin(f.dyHeight, bin, swapit);
          f.grbit = readbin(f.grbit, bin, swapit);
          f.icv = readbin(f.icv, bin, swapit);
          f.bls = readbin(f.bls, bin, swapit);
          f.sss = readbin(f.sss, bin, swapit);
          f.uls = readbin(f.uls, bin, swapit);
          f.bFamily = readbin(f.bFamily, bin, swapit);
          f.bCharSet = readbin(f.bCharSet, bin, swapit);
          bin.seekg(1, std::ios::cur);
          f.name = XLUnicodeString(bin, swapit, false);
          note_fonts.push_back(f);
        }
        if (grid == XlsPalette) {
          uint16_t ccv = 0;
          ccv = readbin(ccv, bin, swapit);
          for (size_t k = 0; k < ccv && k < note_palette.size(); ++k) {
            uint8_t r = 0, gg = 0, b = 0, reserved = 0;
            r = readbin(r, bin, swapit);
            gg = readbin(gg, bin, swapit);
            b = readbin(b, bin, swapit);
            reserved = readbin(reserved, bin, swapit);
            note_palette[k] = (static_cast<uint32_t>(r) << 16) | (static_cast<uint32_t>(gg) << 8) | b;
          }
        }
        bin.seekg(gnext, std::ios_base::beg);
      }
    }

    bin.clear();
    bin.seekg(static_cast<std::streamoff>(g.sheets[static_cast<size_t>(sheet) - 1].offset), std::ios_base::beg);

    int depth = 0;
    R_xlen_t idx = 0;
    uint16_t rid = 0, rsize = 0;
    while (XLS_RECORD(rid, rsize, bin, swapit)) {
      checkInterrupt(idx);
      std::streampos next = bin.tellg() + std::streamoff(rsize);
      if (debug) Rcpp::Rcout << "sheet: " << std::hex << rid << std::dec << ": " << rsize << std::endl;

      if (rid != XlsContinue && rid != XlsObj) after_obj = false;

      if (rid != XlsString && rid != XlsShrFmla && rid != XlsArray && rid != XlsTable && rid != XlsContinue)
        finish_pending();

      // charts and other objects have a substream of their own
      if (rid == XlsBOF) {
        uint16_t bof_version = 0, bof_type = 0;
        bof_version = readbin(bof_version, bin, swapit);
        bof_type = readbin(bof_type, bin, swapit);
        if (!data_only && depth == 1 && bof_type == 0x0020) {
          in_chart = true;
          chart_recs.clear();
        }
        ++depth;
        bin.seekg(next, std::ios_base::beg);
        continue;
      }
      if (rid == XlsEOF) {
        --depth;
        if (in_chart && depth == 1) {
          in_chart = false;
          XlsFormulaContext chart_ctx = g.ctx;
          std::string cx = xls_chart_xml(chart_recs, chart_ctx, swapit, debug);
          for (size_t k = items.size(); k > 0; --k) {
            if (items[k - 1].chart && !items[k - 1].chart_done) {
              items[k - 1].chart_xml = cx;
              items[k - 1].chart_done = true;
              break;
            }
          }
        }
        if (depth <= 0) break;
        bin.seekg(next, std::ios_base::beg);
        continue;
      }
      if (depth != 1) {
        if (in_chart && depth == 2) {
          XlsChartRecord cr;
          cr.id = rid;
          cr.d.resize(rsize);
          bin.read(reinterpret_cast<char*>(cr.d.data()), rsize);
          chart_recs.push_back(cr);
        }
        bin.seekg(next, std::ios_base::beg);
        continue;
      }

      if (data_only) {
        switch (rid) {
          case XlsMsoDrawing:
          case XlsObj:
          case XlsContinue:
          case XlsTxO:
          case XlsNote:
          case XlsList12:
          case XlsFeature11:
          case XlsFeature12:
          case XlsFeat:
          case XlsCondFmt:
          case XlsCondFmt12:
          case XlsCF:
          case XlsCF12:
          case XlsDv:
          case XlsHLink:
          case XlsHLinkTooltip:
          case XlsSxView:
          case XlsSxvd:
          case XlsSXVI:
          case XlsSxIvd:
          case XlsSXPI:
          case XlsSXDI:
          case XlsSXAddl:
            bin.seekg(next, std::ios_base::beg);
            continue;
          default:
            break;
        }
      }

      std::vector<Cell> cells;
      switch (rid) {
        case XlsLabelSst: {
          // row, column, format, index of the string
          Cell c;
          uint16_t row = 0, col = 0, ixfe = 0;
          uint32_t isst = 0;
          row = readbin(row, bin, swapit);
          col = readbin(col, bin, swapit);
          ixfe = readbin(ixfe, bin, swapit);
          isst = readbin(isst, bin, swapit);
          c.row = row;
          c.col = col;
          c.xf = ixfe;
          c.type = "s";
          c.body = "<v>" + std::to_string(isst) + "</v>";
          cells.push_back(c);
          break;
        }

        case XlsLabel: {
          Cell c;
          uint16_t row = 0, col = 0, ixfe = 0;
          row = readbin(row, bin, swapit);
          col = readbin(col, bin, swapit);
          ixfe = readbin(ixfe, bin, swapit);
          std::string val = XLUnicodeString(bin, swapit, true);
          c.row = row;
          c.col = col;
          c.xf = ixfe;
          c.type = "inlineStr";
          c.body = "<is><t xml:space=\"preserve\">" + escape_xml(val) + "</t></is>";
          cells.push_back(c);
          break;
        }

        case XlsNumber: {
          Cell c;
          uint16_t row = 0, col = 0, ixfe = 0;
          double num = 0;
          row = readbin(row, bin, swapit);
          col = readbin(col, bin, swapit);
          ixfe = readbin(ixfe, bin, swapit);
          num = readbin(num, bin, swapit);
          c.row = row;
          c.col = col;
          c.xf = ixfe;
          if (std::isfinite(num)) {
            c.body = "<v>" + xls_num(num) + "</v>";
          } else {
            c.type = "e";
            c.body = "<v>#NUM!</v>";
          }
          cells.push_back(c);
          break;
        }

        case XlsRK: {
          Cell c;
          uint16_t row = 0, col = 0, ixfe = 0;
          int32_t rk = 0;
          row = readbin(row, bin, swapit);
          col = readbin(col, bin, swapit);
          ixfe = readbin(ixfe, bin, swapit);
          rk = readbin(rk, bin, swapit);
          c.row = row;
          c.col = col;
          c.xf = ixfe;
          c.body = "<v>" + xls_num(RkNumber(rk)) + "</v>";
          cells.push_back(c);
          break;
        }

        case XlsMulRk: {
          // row, first column, then a format and a number for every column, the last column
          uint16_t row = 0, col_first = 0;
          row = readbin(row, bin, swapit);
          col_first = readbin(col_first, bin, swapit);
          for (size_t i = 0; i < (rsize - 6u) / 6u; ++i) {
            Cell c;
            uint16_t ixfe = 0;
            int32_t rk = 0;
            ixfe = readbin(ixfe, bin, swapit);
            rk = readbin(rk, bin, swapit);
            c.row = row;
            c.col = col_first + static_cast<uint32_t>(i);
            c.xf = ixfe;
            c.body = "<v>" + xls_num(RkNumber(rk)) + "</v>";
            cells.push_back(c);
          }
          break;
        }

        case XlsBoolErr: {
          Cell c;
          uint16_t row = 0, col = 0, ixfe = 0;
          uint8_t value = 0, is_error = 0;
          row = readbin(row, bin, swapit);
          col = readbin(col, bin, swapit);
          ixfe = readbin(ixfe, bin, swapit);
          value = readbin(value, bin, swapit);
          is_error = readbin(is_error, bin, swapit);
          c.row = row;
          c.col = col;
          c.xf = ixfe;
          if (is_error) {
            c.type = "e";
            c.body = "<v>" + xls_error(value) + "</v>";
          } else {
            c.type = "b";
            c.body = std::string("<v>") + (value ? "1" : "0") + "</v>";
          }
          cells.push_back(c);
          break;
        }

        case XlsBlank: {
          // empty cells are only there for their format
          Cell c;
          uint16_t row = 0, col = 0, ixfe = 0;
          row = readbin(row, bin, swapit);
          col = readbin(col, bin, swapit);
          ixfe = readbin(ixfe, bin, swapit);
          c.row = row;
          c.col = col;
          c.xf = ixfe;
          if (ixfe != 0) cells.push_back(c);
          break;
        }

        case XlsMulBlank: {
          uint16_t row = 0, col_first = 0;
          row = readbin(row, bin, swapit);
          col_first = readbin(col_first, bin, swapit);
          for (size_t i = 0; i < (rsize - 6u) / 2u; ++i) {
            Cell c;
            uint16_t ixfe = 0;
            ixfe = readbin(ixfe, bin, swapit);
            c.row = row;
            c.col = col_first + static_cast<uint32_t>(i);
            c.xf = ixfe;
            if (ixfe != 0) cells.push_back(c);
          }
          break;
        }

        case XlsFormula: {
          // row, column, format, the last result (8 bytes), flags, 4 unused
          // bytes, size of the formula, the formula. If the last 2 bytes of the
          // result are 0xFFFF it is not a number: 0 text (in the next record), 1
          // boolean, 2 error, 3 empty text.
          uint16_t row = 0, col = 0, ixfe = 0, flags = 0, cce = 0;
          uint8_t result[8] = {0, 0, 0, 0, 0, 0, 0, 0};
          uint32_t chn = 0;
          row = readbin(row, bin, swapit);
          col = readbin(col, bin, swapit);
          ixfe = readbin(ixfe, bin, swapit);
          for (size_t i = 0; i < 8; ++i) result[i] = readbin(result[i], bin, swapit);
          flags = readbin(flags, bin, swapit);
          chn = readbin(chn, bin, swapit);
          cce = readbin(cce, bin, swapit);

          pending = Cell();
          pending.row = row;
          pending.col = col;
          pending.xf = ixfe;
          pending_rgce.cce = cce;
          if (next < bin.tellg()) Rcpp::stop("xls: record overrun");
          pending_rgce.data.assign(static_cast<size_t>(next - bin.tellg()), 0);
          if (!pending_rgce.data.empty()) bin.read(reinterpret_cast<char*>(&pending_rgce.data[0]), static_cast<std::streamsize>(pending_rgce.data.size()));

          if (result[6] == 0xFF && result[7] == 0xFF) {
            if (result[0] == 1) {
              pending.type = "b";
              pending.body = std::string("<v>") + (result[2] ? "1" : "0") + "</v>";
            } else if (result[0] == 2) {
              pending.type = "e";
              pending.body = "<v>" + xls_error(result[2]) + "</v>";
            } else {
              pending.type = "str";
            }
          } else {
            double num = 0;
            std::memcpy(&num, result, sizeof(num));
            if (swapit) num = swap_endian(num);
            pending.body = "<v>" + xls_num(num) + "</v>";
          }
          has_pending = true;
          break;
        }

        case XlsString: {
          if (has_pending && pending.type == "str") {
            std::string val = XLUnicodeString(bin, swapit, true);
            pending.body = "<v>" + escape_xml(val) + "</v>";
          }
          break;
        }

        case XlsShrFmla: {
          // the cells it is used in (first row, last row, first column, last
          // column), 2 bytes of counters, the size of the formula, the formula
          uint16_t row1 = 0, row2 = 0, cce = 0;
          uint8_t col1 = 0, col2 = 0, unused = 0, cuse = 0;
          row1 = readbin(row1, bin, swapit);
          row2 = readbin(row2, bin, swapit);
          col1 = readbin(col1, bin, swapit);
          col2 = readbin(col2, bin, swapit);
          unused = readbin(unused, bin, swapit);
          cuse = readbin(cuse, bin, swapit);
          cce = readbin(cce, bin, swapit);
          XlsRgce rgce;
          rgce.cce = cce;
          if (next < bin.tellg()) Rcpp::stop("xls: record overrun");
          rgce.data.assign(static_cast<size_t>(next - bin.tellg()), 0);
          if (!rgce.data.empty()) bin.read(reinterpret_cast<char*>(&rgce.data[0]), static_cast<std::streamsize>(rgce.data.size()));
          shared_formulas[has_pending ? std::make_pair(pending.row, pending.col)
                                      : std::make_pair(static_cast<uint32_t>(row1), static_cast<uint32_t>(col1))] = rgce;
          break;
        }

        case XlsArray: {
          // like ShrFmla, with flags and 4 unused bytes before the formula
          uint16_t row1 = 0, row2 = 0, flags = 0, cce = 0;
          uint8_t col1 = 0, col2 = 0;
          uint32_t chn = 0;
          row1 = readbin(row1, bin, swapit);
          row2 = readbin(row2, bin, swapit);
          col1 = readbin(col1, bin, swapit);
          col2 = readbin(col2, bin, swapit);
          flags = readbin(flags, bin, swapit);
          chn = readbin(chn, bin, swapit);
          cce = readbin(cce, bin, swapit);
          XlsRgce rgce;
          rgce.cce = cce;
          if (next < bin.tellg()) Rcpp::stop("xls: record overrun");
          rgce.data.assign(static_cast<size_t>(next - bin.tellg()), 0);
          if (!rgce.data.empty()) bin.read(reinterpret_cast<char*>(&rgce.data[0]), static_cast<std::streamsize>(rgce.data.size()));
          std::string ref = int_to_col(col1 + 1) + std::to_string(row1 + 1);
          if (row1 != row2 || col1 != col2) ref += ":" + int_to_col(col2 + 1) + std::to_string(row2 + 1);
          array_formulas[has_pending ? std::make_pair(pending.row, pending.col)
                                     : std::make_pair(static_cast<uint32_t>(row1), static_cast<uint32_t>(col1))] =
              std::make_pair(ref, rgce);
          break;
        }

        case XlsRow: {
          // row, first and last column (not needed), height (the highest bit says
          // that it is the default), 4 unused bytes, flags
          uint16_t row = 0, col_first = 0, col_last = 0, height = 0, unused1 = 0, unused2 = 0;
          uint32_t flags = 0;
          row = readbin(row, bin, swapit);
          col_first = readbin(col_first, bin, swapit);
          col_last = readbin(col_last, bin, swapit);
          height = readbin(height, bin, swapit);
          unused1 = readbin(unused1, bin, swapit);
          unused2 = readbin(unused2, bin, swapit);
          flags = readbin(flags, bin, swapit);
          Row r;
          r.height = (height & 0x7FFF) / 20.0;
          r.outline = static_cast<uint8_t>(flags & 0x07u);
          r.hidden = (flags & 0x20u) != 0;
          r.custom = (flags & 0x40u) != 0;
          rows[row] = r;
          break;
        }

        case XlsColInfo: {
          // first and last column, width (1/256 of the width of a character),
          // format, flags (hidden, outline level)
          uint16_t col_first = 0, col_last = 0, width = 0, ixfe = 0, flags = 0;
          col_first = readbin(col_first, bin, swapit);
          col_last = readbin(col_last, bin, swapit);
          width = readbin(width, bin, swapit);
          ixfe = readbin(ixfe, bin, swapit);
          flags = readbin(flags, bin, swapit);
          for (uint32_t c = col_first; c <= col_last && c < 256; ++c) col_px[static_cast<uint16_t>(c)] = std::floor(width / 256.0 * 7.0 + 0.5) + 5.0;
          cols << "<col min=\"" << col_first + 1 << "\" max=\"" << (col_last >= 255 ? 16384 : col_last + 1) << "\" width=\"" << xls_num(width / 256.0)
               << "\" customWidth=\"1\"";
          if (ixfe) cols << " style=\"" << ixfe << "\"";
          if (flags & 0x0001) cols << " hidden=\"1\"";
          if ((flags >> 8) & 0x07) cols << " outlineLevel=\"" << ((flags >> 8) & 0x07) << "\"";
          cols << "/>" << std::endl;
          break;
        }

        case XlsLeftMargin:
        case XlsRightMargin:
        case XlsTopMargin:
        case XlsBottomMargin: {
          double v = 0;
          v = readbin(v, bin, swapit);
          margins[rid - XlsLeftMargin] = v;
          have_margin = true;
          break;
        }

        case XlsSetup: {
          paper = readbin(paper, bin, swapit);
          scale = readbin(scale, bin, swapit);
          page_start = readbin(page_start, bin, swapit);
          fit_w = readbin(fit_w, bin, swapit);
          fit_h = readbin(fit_h, bin, swapit);
          setup_flags = readbin(setup_flags, bin, swapit);
          uint16_t res = 0, vres = 0;
          res = readbin(res, bin, swapit);
          vres = readbin(vres, bin, swapit);
          margins[4] = readbin(margins[4], bin, swapit);
          margins[5] = readbin(margins[5], bin, swapit);
          copies = readbin(copies, bin, swapit);
          have_setup = true;
          have_margin = true;
          break;
        }

        case XlsHeader:
          if (rsize > 0) header = XLUnicodeString(bin, swapit, true);
          break;

        case XlsFooter:
          if (rsize > 0) footer = XLUnicodeString(bin, swapit, true);
          break;

        case XlsHCenter:
        case XlsVCenter:
        case XlsPrintGrid:
        case XlsPrintRowCol: {
          uint16_t v = 0;
          v = readbin(v, bin, swapit);
          if (rid == XlsHCenter) h_center = v != 0;
          if (rid == XlsVCenter) v_center = v != 0;
          if (rid == XlsPrintGrid) print_grid = v != 0;
          if (rid == XlsPrintRowCol) print_headers = v != 0;
          break;
        }

        case XlsAutoFilterInfo:
          have_autofilter = true;
          break;

        case XlsAutoFilter: {
          // column, options, two operands of 10 bytes: type, comparison and 8 bytes of value.
          // The text of a string operand follows after both of them.
          std::vector<uint8_t> rec(rsize);
          bin.read(reinterpret_cast<char*>(rec.data()), rsize);
          if (rsize < 24) break;
          auto u16 = [&](size_t at) -> uint32_t { return static_cast<uint32_t>(rec[at] | (rec[at + 1] << 8)); };
          uint16_t col = static_cast<uint16_t>(u16(0));
          uint32_t opt = u16(2);
          bool join_or = (opt & 0x0003) != 0;
          bool top_n = (opt & 0x0010) != 0;

          if (top_n) {
            filter_columns[col] = std::string("<filterColumn colId=\"") + std::to_string(col) + "\"><top10" +
                                  ((opt & 0x0020) ? " top=\"1\"" : " top=\"0\"") + ((opt & 0x0040) ? " percent=\"1\"" : "") +
                                  " val=\"" + std::to_string(opt >> 7) + "\"/></filterColumn>";
            break;
          }

          static const char* const comparison[7] = {"", "lessThan", "equal", "lessThanOrEqual", "greaterThan", "notEqual",
                                                    "greaterThanOrEqual"};
          size_t text_at = 24;
          struct Operand {
            bool valid = false, blank = false, nonblank = false;
            std::string comparison, value;
            bool is_equal = false;
          } ops[2];
          for (size_t k = 0; k < 2; ++k) {
            size_t at = 4 + k * 10;
            uint8_t vt = rec[at], sign = rec[at + 1];
            Operand& o = ops[k];
            if (sign >= 1 && sign <= 6) o.comparison = comparison[sign];
            o.is_equal = sign == 2;
            if (vt == 0x02) {
              uint32_t rk = u16(at + 2) | (u16(at + 4) << 16);
              double d = 0;
              if (rk & 2) {
                d = static_cast<double>(static_cast<int32_t>(rk) >> 2);
              } else {
                uint64_t bits = static_cast<uint64_t>(rk & 0xFFFFFFFCu) << 32;
                std::memcpy(&d, &bits, sizeof(d));
              }
              if (rk & 1) d /= 100.0;
              o.value = xls_num(d);
              o.valid = true;
            } else if (vt == 0x04) {
              uint64_t bits = 0;
              for (size_t b = 0; b < 8; ++b) bits |= static_cast<uint64_t>(rec[at + 2 + b]) << (8 * b);
              double d = 0;
              std::memcpy(&d, &bits, sizeof(d));
              o.value = xls_num(d);
              o.valid = true;
            } else if (vt == 0x06) {
              size_t cch = rec[at + 6];
              if (text_at >= rec.size()) continue;
              bool wide = (rec[text_at] & 1) != 0;
              ++text_at;
              std::u16string txt;
              for (size_t b = 0; b < cch; ++b) {
                if (wide && text_at + 1 < rec.size()) {
                  txt.push_back(static_cast<char16_t>(u16(text_at)));
                  text_at += 2;
                } else if (!wide && text_at < rec.size()) {
                  txt.push_back(static_cast<char16_t>(rec[text_at++]));
                }
              }
              o.value = to_utf8(txt);
              o.valid = true;
            } else if (vt == 0x08) {
              o.value = rec[at + 2] ? "TRUE" : "FALSE";
              o.valid = true;
            } else if (vt == 0x0C) {
              o.blank = true;
              o.valid = true;
            } else if (vt == 0x0E) {
              o.nonblank = true;
              o.valid = true;
            }
          }

          std::string x = "<filterColumn colId=\"" + std::to_string(col) + "\">";
          bool both = ops[0].valid && ops[1].valid;
          if (ops[0].valid && !both && ops[0].blank) {
            x += "<filters blank=\"1\"/>";
          } else if (ops[0].valid && ((!both && ops[0].is_equal) || (both && join_or && ops[0].is_equal && ops[1].is_equal)) &&
                     !ops[0].blank && !ops[0].nonblank) {
            x += "<filters>";
            for (size_t k = 0; k < (both ? 2u : 1u); ++k) x += "<filter val=\"" + escape_xml(ops[k].value) + "\"/>";
            x += "</filters>";
          } else {
            x += std::string("<customFilters") + (both && !join_or ? " and=\"1\"" : "") + ">";
            for (size_t k = 0; k < 2; ++k) {
              if (!ops[k].valid) continue;
              if (ops[k].blank) x += "<customFilter operator=\"equal\" val=\"\"/>";
              else if (ops[k].nonblank) x += "<customFilter operator=\"notEqual\" val=\" \"/>";
              else {
                x += "<customFilter";
                if (!ops[k].comparison.empty() && ops[k].comparison != "equal") x += " operator=\"" + ops[k].comparison + "\"";
                x += " val=\"" + escape_xml(ops[k].value) + "\"/>";
              }
            }
            x += "</customFilters>";
          }
          x += "</filterColumn>";
          filter_columns[col] = x;
          break;
        }

        case XlsProtect:
        case XlsObjProtect:
        case XlsScenarioProtect:
        case XlsPassword: {
          uint16_t v = 0;
          v = readbin(v, bin, swapit);
          if (rid == XlsProtect) protect_sheet = v != 0;
          if (rid == XlsObjProtect) protect_objects = v != 0;
          if (rid == XlsScenarioProtect) protect_scenarios = v != 0;
          if (rid == XlsPassword) password = v;
          break;
        }

        case XlsHorizontalPageBreaks:
        case XlsVerticalPageBreaks: {
          uint16_t count = 0;
          count = readbin(count, bin, swapit);
          for (uint16_t i = 0; i < count; ++i) {
            uint16_t brk = 0, from = 0, to = 0;
            brk = readbin(brk, bin, swapit);
            from = readbin(from, bin, swapit);
            to = readbin(to, bin, swapit);
            (rid == XlsHorizontalPageBreaks ? row_breaks : col_breaks).push_back(brk);
          }
          break;
        }

        case XlsMergeCells: {
          uint16_t count = 0;
          count = readbin(count, bin, swapit);
          for (uint16_t i = 0; i < count; ++i) {
            uint16_t row1 = 0, row2 = 0, col1 = 0, col2 = 0;
            row1 = readbin(row1, bin, swapit);
            row2 = readbin(row2, bin, swapit);
            col1 = readbin(col1, bin, swapit);
            col2 = readbin(col2, bin, swapit);
            merges.push_back(int_to_col(col1 + 1) + std::to_string(row1 + 1) + ":" + int_to_col(col2 + 1) +
                             std::to_string(row2 + 1));
          }
          break;
        }

        case XlsSortData: {
          std::vector<uint8_t> rec(rsize);
          bin.read(reinterpret_cast<char*>(rec.data()), rsize);
          if (rsize < 38) break;
          uint16_t flags = static_cast<uint16_t>(rec[12] | (rec[13] << 8));
          uint32_t v[4];
          for (size_t k = 0; k < 4; ++k)
            v[k] = static_cast<uint32_t>(rec[14 + 4 * k]) | (static_cast<uint32_t>(rec[15 + 4 * k]) << 8) |
                   (static_cast<uint32_t>(rec[16 + 4 * k]) << 16) | (static_cast<uint32_t>(rec[17 + 4 * k]) << 24);
          sort_pending = static_cast<uint32_t>(rec[30]) | (static_cast<uint32_t>(rec[31]) << 8) |
                         (static_cast<uint32_t>(rec[32]) << 16) | (static_cast<uint32_t>(rec[33]) << 24);
          sort_sfp = (flags >> 3) & 0x07;
          sort_conditions.clear();

          // the range of an AutoFilter includes its header row, the sort state does not
          uint32_t first_row = v[0] + (sort_sfp == 2 ? 1 : 0);
          sort_ref = int_to_col(v[2] + 1) + std::to_string(first_row + 1) + ":" + int_to_col(v[3] + 1) + std::to_string(v[1] + 1);
          break;
        }

        case XlsContinueFrt12: {
          if (sort_pending == 0) break;
          std::vector<uint8_t> rec(rsize);
          bin.read(reinterpret_cast<char*>(rec.data()), rsize);
          if (rsize < 30) break;
          --sort_pending;
          uint16_t flags = static_cast<uint16_t>(rec[12] | (rec[13] << 8));
          if (((flags >> 1) & 0x0F) != 0) break;
          uint32_t v[4];
          for (size_t k = 0; k < 4; ++k)
            v[k] = static_cast<uint32_t>(rec[14 + 4 * k]) | (static_cast<uint32_t>(rec[15 + 4 * k]) << 8) |
                   (static_cast<uint32_t>(rec[16 + 4 * k]) << 16) | (static_cast<uint32_t>(rec[17 + 4 * k]) << 24);
          sort_conditions += "<sortCondition";
          if (flags & 0x01) sort_conditions += " descending=\"1\"";
          sort_conditions += " ref=\"" + int_to_col(v[2] + 1) + std::to_string(v[0] + 1) + ":" + int_to_col(v[3] + 1) +
                             std::to_string(v[1] + 1) + "\"/>";
          break;
        }

        case XlsFeat: {
          std::vector<uint8_t> rec(rsize);
          bin.read(reinterpret_cast<char*>(rec.data()), rsize);
          if (rsize < 27) break;
          uint16_t isf = static_cast<uint16_t>(rec[12] | (rec[13] << 8));
          size_t cref = static_cast<size_t>(rec[19] | (rec[20] << 8));
          if (isf != 0x0003 || rsize < 27 + cref * 8 + 1) break;

          std::string sqref;
          for (size_t i = 0; i < cref; ++i) {
            const uint8_t* r = &rec[27 + i * 8];
            unsigned row1 = static_cast<unsigned>(r[0] | (r[1] << 8)), row2 = static_cast<unsigned>(r[2] | (r[3] << 8));
            uint32_t col1 = static_cast<uint32_t>(r[4] | (r[5] << 8)), col2 = static_cast<uint32_t>(r[6] | (r[7] << 8));
            if (!sqref.empty()) sqref += " ";
            sqref += int_to_col(col1 + 1) + std::to_string(row1 + 1);
            if (row1 != row2 || col1 != col2) sqref += ":" + int_to_col(col2 + 1) + std::to_string(row2 + 1);
          }

          static const char* const ignored[8] = {"evalError", "emptyCellReference", "numberStoredAsText", "formulaRange",
                                                 "formula", "twoDigitTextYear", "unlockedFormula", "listDataValidation"};
          uint8_t flags = rec[27 + cref * 8];
          std::string attrs;
          for (size_t bit = 0; bit < 8; ++bit)
            if (flags & (1u << bit)) attrs += std::string(" ") + ignored[bit] + "=\"1\"";
          if (!sqref.empty() && !attrs.empty()) ignored_errors += "<ignoredError sqref=\"" + sqref + "\"" + attrs + "/>";
          break;
        }

        case XlsFeature11:
        case XlsFeature12: {
          std::vector<uint8_t> rec(rsize);
          bin.read(reinterpret_cast<char*>(rec.data()), rsize);
          XlsListTable t;
          if (xls_table(rec, t)) tables.push_back(t);
          break;
        }

        case XlsList12: {
          // frt header (12 bytes), kind of the data, table id. Kind 1 is the style.
          std::vector<uint8_t> rec(rsize);
          bin.read(reinterpret_cast<char*>(rec.data()), rsize);
          if (rsize < 18) break;
          uint16_t lsd = static_cast<uint16_t>(rec[12] | (rec[13] << 8));
          uint32_t id = static_cast<uint32_t>(rec[14]) | (static_cast<uint32_t>(rec[15]) << 8) |
                        (static_cast<uint32_t>(rec[16]) << 16) | (static_cast<uint32_t>(rec[17]) << 24);
          if (lsd != 1 || rsize < 21) break;

          // options (2 bytes), name of the style (XLUnicodeString)
          uint16_t opt = static_cast<uint16_t>(rec[18] | (rec[19] << 8));
          size_t cch = static_cast<size_t>(rec[20] | (rec[21] << 8));
          bool wide = rsize > 22 && (rec[22] & 0x01);
          std::u16string name;
          for (size_t i = 0; i < cch; ++i) {
            size_t at = 23 + i * (wide ? 2 : 1);
            if (at + (wide ? 1 : 0) >= rsize) break;
            name.push_back(wide ? static_cast<char16_t>(rec[at] | (rec[at + 1] << 8)) : static_cast<char16_t>(rec[at]));
          }
          for (size_t i = 0; i < tables.size(); ++i) {
            if (tables[i].id != id) continue;
            tables[i].style = to_utf8(name);
            tables[i].first_column = (opt & 0x01) != 0;
            tables[i].last_column = (opt & 0x02) != 0;
            tables[i].row_stripes = (opt & 0x04) != 0;
            tables[i].column_stripes = (opt & 0x08) != 0;
          }
          break;
        }

        case XlsDv: {
          // flags, four texts (input title, error title, input message, error
          // message), two formulas and the ranges the rule is used for
          uint32_t flags = 0;
          flags = readbin(flags, bin, swapit);
          std::string prompt_title = XLUnicodeString(bin, swapit, true);
          std::string error_title = XLUnicodeString(bin, swapit, true);
          std::string prompt = XLUnicodeString(bin, swapit, true);
          std::string error = XLUnicodeString(bin, swapit, true);

          // an unused text is a single 0 character
          auto trim = [](std::string x) {
            while (!x.empty() && x.back() == '\0') x.pop_back();
            return x;
          };

          std::string formula[2];
          for (int k = 0; k < 2; ++k) {
            uint16_t cce = 0, unused = 0;
            cce = readbin(cce, bin, swapit);
            unused = readbin(unused, bin, swapit);
            std::vector<uint8_t> tokens(cce, 0);
            if (cce > 0) bin.read(reinterpret_cast<char*>(tokens.data()), cce);
            XlsRgce rgce(tokens);
            if (cce > 0) {
              XlsFormulaContext& dctx = g.ctx;
              dctx.row = 0;
              dctx.col = 0;
              dctx.unsupported = -1;
              formula[k] = xls_rgce(rgce, dctx, swapit, debug);
            }
          }

          uint16_t cref = 0;
          cref = readbin(cref, bin, swapit);
          std::string sqref;
          for (uint16_t i = 0; i < cref; ++i) {
            uint16_t row1 = 0, row2 = 0, col1 = 0, col2 = 0;
            row1 = readbin(row1, bin, swapit);
            row2 = readbin(row2, bin, swapit);
            col1 = readbin(col1, bin, swapit);
            col2 = readbin(col2, bin, swapit);
            std::string a = int_to_col(col1 + 1) + std::to_string(row1 + 1);
            std::string b = int_to_col(col2 + 1) + std::to_string(row2 + 1);
            if (!sqref.empty()) sqref += " ";
            sqref += a == b ? a : a + ":" + b;
          }

          static const char* types[] = {"none", "whole", "decimal", "list", "date", "time", "textLength", "custom"};
          static const char* operators[] = {"between", "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual"};
          uint32_t type = flags & 0x0F;
          uint32_t style = (flags >> 4) & 0x07;
          uint32_t oper = (flags >> 20) & 0x0F;
          if (type == 0 || type > 7 || oper > 7 || sqref.empty()) break;

          // an explicit list is a text with the entries separated by a 0 character
          if (type == 3 && (flags & 0x80)) {
            for (size_t i = 0; i < formula[0].size(); ++i)
              if (formula[0][i] == '\0') formula[0][i] = ',';
          }

          std::string dv = std::string("<dataValidation type=\"") + types[type] + "\"";
          if (style == 1) dv += " errorStyle=\"warning\"";
          if (style == 2) dv += " errorStyle=\"information\"";
          if (oper != 0) dv += std::string(" operator=\"") + operators[oper] + "\"";
          if (flags & 0x100) dv += " allowBlank=\"1\"";
          if (flags & 0x200) dv += " showDropDown=\"1\"";
          if (flags & 0x40000) dv += " showInputMessage=\"1\"";
          if (flags & 0x80000) dv += " showErrorMessage=\"1\"";
          if (!trim(error_title).empty()) dv += " errorTitle=\"" + escape_xml(trim(error_title)) + "\"";
          if (!trim(error).empty()) dv += " error=\"" + escape_xml(trim(error)) + "\"";
          if (!trim(prompt_title).empty()) dv += " promptTitle=\"" + escape_xml(trim(prompt_title)) + "\"";
          if (!trim(prompt).empty()) dv += " prompt=\"" + escape_xml(trim(prompt)) + "\"";
          dv += " sqref=\"" + sqref + "\">";
          dv += "<formula1>" + escape_xml(formula[0]) + "</formula1>";
          if (!formula[1].empty()) dv += "<formula2>" + escape_xml(formula[1]) + "</formula2>";
          dv += "</dataValidation>";
          validations.push_back(dv);
          break;
        }

        case XlsMsoDrawing: {
          // Escher data, goes on in Continue records
          std::string dg(rsize, '\0');
          bin.read(&dg[0], rsize);
          while (true) {
            bin.clear();
            bin.seekg(next, std::ios_base::beg);
            uint16_t crid = 0, crsize = 0;
            if (!XLS_RECORD(crid, crsize, bin, swapit) || crid != XlsContinue) break;
            std::string more(crsize, '\0');
            bin.read(&more[0], crsize);
            dg += more;
            next = bin.tellg();
          }
          std::vector<XlsShape> shapes = xls_shapes(dg);
          for (size_t i = 0; i < shapes.size(); ++i)
            if (!shapes[i].group) pending_shapes.push_back(shapes[i]);
          break;
        }

        case XlsObj: {
          // the first sub record is ftCmo: 0x15, size, type of the object, number
          std::vector<uint8_t> rec(rsize);
          bin.read(reinterpret_cast<char*>(rec.data()), rsize);
          last_note_object = -1;
          if (rsize >= 8 && rec[0] == 0x15 && rec[1] == 0x00) {
            uint16_t type = static_cast<uint16_t>(rec[4] | (rec[5] << 8));
            uint16_t id = static_cast<uint16_t>(rec[6] | (rec[7] << 8));
            if (type == 0x19) last_note_object = id;

            // every object takes the next shape
            if (!pending_shapes.empty()) {
              XlsShape shape = pending_shapes.front();
              pending_shapes.erase(pending_shapes.begin());
              if (type == 0x19 && shape.has_anchor) note_shapes[id] = shape;
              bool is_shape = type == 0x01 || type == 0x02 || type == 0x03 || type == 0x04 || type == 0x09 || type == 0x1E;
              if (shape.has_anchor && ((type == 0x08 && shape.blip > 0) || type == 0x05 || is_shape)) {
                DrawingItem item;
                item.chart = type == 0x05;
                item.autoshape = is_shape;
                item.blip = shape.blip;
                item.shape = shape;
                items.push_back(item);
              }
            }
          }
          after_obj = true;
          break;
        }

        case XlsContinue: {
          // Escher data of the shape of an object can follow the object as CONTINUE
          // records, when the drawing record before was full
          if (!after_obj || rsize < 8) break;
          std::string dg(rsize, '\0');
          bin.read(&dg[0], rsize);
          if (static_cast<uint8_t>(dg[0]) != 0x0F || static_cast<uint8_t>(dg[2]) != 0x04 || static_cast<uint8_t>(dg[3]) != 0xF0) break;
          while (true) {
            bin.clear();
            bin.seekg(next, std::ios_base::beg);
            uint16_t crid = 0, crsize = 0;
            if (!XLS_RECORD(crid, crsize, bin, swapit) || crid != XlsContinue) break;
            std::string more(crsize, '\0');
            bin.read(&more[0], crsize);
            dg += more;
            next = bin.tellg();
          }
          std::vector<XlsShape> shapes = xls_shapes(dg);
          for (size_t i = 0; i < shapes.size(); ++i)
            if (!shapes[i].group) pending_shapes.push_back(shapes[i]);
          break;
        }

        case XlsTxO: {
          // options, rotation, 6 bytes we do not use, length of the text, length
          // of the formatting. The text is in the CONTINUE record that follows, then
          // the formatting: a run of 8 bytes for every change of the font.
          uint16_t cch = 0, cb_runs = 0;
          bin.seekg(10, std::ios_base::cur);
          cch = readbin(cch, bin, swapit);
          cb_runs = readbin(cb_runs, bin, swapit);

          if (last_note_object >= 0 && cch > 0) {
            bin.seekg(next, std::ios_base::beg);
            uint16_t crid = 0, crsize = 0;
            if (!XLS_RECORD(crid, crsize, bin, swapit) || crid != XlsContinue) Rcpp::stop("xls: text of a comment is missing");
            std::streampos text_end = bin.tellg() + std::streamoff(crsize);
            uint8_t flags = 0;
            flags = readbin(flags, bin, swapit);
            uint16_t note_id = static_cast<uint16_t>(last_note_object);
            note_text[note_id] = xls_chars(bin, swapit, cch, (flags & 0x01) != 0, text_end);

            bin.seekg(text_end, std::ios_base::beg);
            if (cb_runs >= 8 && XLS_RECORD(crid, crsize, bin, swapit) && crid == XlsContinue) {
              for (uint32_t k = 0; k + 8 <= crsize; k += 8) {
                uint16_t ich = 0, ifnt = 0;
                uint32_t reserved = 0;
                ich = readbin(ich, bin, swapit);
                ifnt = readbin(ifnt, bin, swapit);
                reserved = readbin(reserved, bin, swapit);
                note_runs[note_id].push_back(std::make_pair(ich, ifnt));
              }
            }
          }
          last_note_object = -1;
          break;
        }

        case XlsNote: {
          // cell, options, number of the object, author
          Note n;
          n.row = readbin(n.row, bin, swapit);
          n.col = readbin(n.col, bin, swapit);
          bin.seekg(2, std::ios_base::cur);
          n.id = readbin(n.id, bin, swapit);
          n.author = XLUnicodeString(bin, swapit, true);
          // a cell has one comment
          bool replaced = false;
          for (size_t i = 0; i < notes.size(); ++i) {
            if (notes[i].row == n.row && notes[i].col == n.col) {
              notes[i] = n;
              replaced = true;
              break;
            }
          }
          if (!replaced) notes.push_back(n);
          break;
        }

        case XlsCondFmt:
        case XlsCondFmt12: {
          // number of rules, options, the range of all cells, then the ranges.
          // CondFmt12 has a header of 12 bytes in front.
          CondFmt c;
          uint16_t ccf = 0, cref = 0;
          if (rid == XlsCondFmt12) bin.seekg(12, std::ios::cur);
          ccf = readbin(ccf, bin, swapit);
          bin.seekg(10, std::ios_base::cur);
          cref = readbin(cref, bin, swapit);
          for (uint16_t i = 0; i < cref; ++i) {
            uint16_t row1 = 0, row2 = 0, col1 = 0, col2 = 0;
            row1 = readbin(row1, bin, swapit);
            row2 = readbin(row2, bin, swapit);
            col1 = readbin(col1, bin, swapit);
            col2 = readbin(col2, bin, swapit);
            if (i == 0) {
              c.row = row1;
              c.col = col1;
            }
            std::string a = int_to_col(col1 + 1) + std::to_string(row1 + 1);
            std::string b = int_to_col(col2 + 1) + std::to_string(row2 + 1);
            if (!c.sqref.empty()) c.sqref += " ";
            c.sqref += a == b ? a : a + ":" + b;
          }
          c.pending = ccf;
          cond_fmts.push_back(c);
          break;
        }

        case XlsCF: {
          std::vector<uint8_t> rec(rsize);
          bin.read(reinterpret_cast<char*>(rec.data()), rsize);
          int32_t this_dxf = dxf_id++;
          if (cond_fmts.empty() || cond_fmts.back().pending == 0) break;
          CondFmt& c = cond_fmts.back();
          --c.pending;

          std::vector<uint32_t> default_palette(XlsDefaultPalette, XlsDefaultPalette + 56);
          std::map<uint16_t, std::string> no_formats;
          XlsCfRule cf;
          if (!xls_cf(rec, default_palette, no_formats, 164, cf)) break;

          // the formulas are relative to the first cell of the range
          XlsFormulaContext& cctx = g.ctx;
          cctx.row = c.row;
          cctx.col = c.col;
          cctx.unsupported = -1;
          std::string f1 = xls_rgce(XlsRgce(cf.rgce1), cctx, swapit, debug);
          std::string f2 = cf.rgce2.empty() ? "" : xls_rgce(XlsRgce(cf.rgce2), cctx, swapit, debug);
          if (f1.empty()) break;

          static const char* const operators[9] = {"", "between", "notBetween", "equal", "notEqual", "greaterThan", "lessThan",
                                                   "greaterThanOrEqual", "lessThanOrEqual"};
          std::string rule = "<cfRule";
          if (cf.type == 1 && cf.comparison >= 1 && cf.comparison <= 8) {
            rule += std::string(" type=\"cellIs\" operator=\"") + operators[cf.comparison] + "\"";
          } else if (cf.type == 2) {
            rule += " type=\"expression\"";
          } else {
            break;
          }
          rule += " dxfId=\"" + std::to_string(this_dxf) + "\" priority=\"" + std::to_string(++cf_priority) + "\">";
          rule += "<formula>" + escape_xml(f1) + "</formula>";
          if (!f2.empty()) rule += "<formula>" + escape_xml(f2) + "</formula>";
          rule += "</cfRule>";
          c.xml += rule;
          break;
        }

        case XlsCF12: {
          std::vector<uint8_t> rec(rsize);
          bin.read(reinterpret_cast<char*>(rec.data()), rsize);
          if (cond_fmts.empty() || cond_fmts.back().pending == 0) break;
          CondFmt& c = cond_fmts.back();
          --c.pending;

          std::vector<uint32_t> default_palette(XlsDefaultPalette, XlsDefaultPalette + 56);
          std::map<uint16_t, std::string> no_formats;
          XlsCf12Rule cf;
          if (!xls_cf12(rec, default_palette, no_formats, 164, cf)) {
            if (rec.size() > 12 && (rec[12] == 1 || rec[12] == 2)) ++dxf_id;
            break;
          }
          int32_t this_dxf = (cf.type == 1 || cf.type == 2) ? dxf_id++ : -1;

          XlsFormulaContext& cctx = g.ctx;
          cctx.row = c.row;
          cctx.col = c.col;
          cctx.unsupported = -1;

          // color of a rule: a palette entry, RGB (r, g, b, a) or a theme color
          auto color = [&](const XlsCfColor& col, const char* tag) {
            std::string attr = xls_color_attr(col, default_palette);
            return std::string("<") + tag + (attr.empty() ? " auto=\"1\"" : attr) + "/>";
          };
          auto cfvo = [&](const XlsCfvo& v) {
            static const char* const names[8] = {"", "num", "min", "max", "percent", "percentile", "", "formula"};
            if (v.type < 1 || v.type > 7 || v.type == 6) return std::string("<cfvo type=\"num\" val=\"0\"/>");
            std::string x = std::string("<cfvo type=\"") + names[v.type] + "\"";
            if (!v.rgce.empty()) x += " val=\"" + escape_xml(xls_rgce(XlsRgce(v.rgce), cctx, swapit, debug)) + "\"";
            else if (v.type != 2 && v.type != 3) x += " val=\"" + xls_num(v.num) + "\"";
            if (!v.equal) x += " gte=\"0\"";
            return x + "/>";
          };

          std::string rule;
          if (cf.type == 3 && !cf.cfvos.empty()) {
            rule = "<cfRule type=\"colorScale\" priority=\"" + std::to_string(++cf_priority) + "\"><colorScale>";
            for (size_t i = 0; i < cf.cfvos.size(); ++i) rule += cfvo(cf.cfvos[i]);
            for (size_t i = 0; i < cf.colors.size(); ++i) rule += color(cf.colors[i], "color");
            rule += "</colorScale></cfRule>";
          } else if (cf.type == 4 && cf.cfvos.size() == 2 && cf.colors.size() == 1) {
            std::string cfvo_min = cfvo(cf.cfvos[0]), cfvo_max = cfvo(cf.cfvos[1]);
            bool modern = cf.min_length == 0 && cf.max_length == 100 && cfvo_min == "<cfvo type=\"min\"/>" && cfvo_max == "<cfvo type=\"max\"/>";
            rule = "<cfRule type=\"dataBar\" priority=\"" + std::to_string(++cf_priority) + "\"><dataBar";
            if (!modern) rule += " minLength=\"" + std::to_string(cf.min_length) + "\" maxLength=\"" + std::to_string(cf.max_length) + "\"";
            if (!cf.show_value) rule += " showValue=\"0\"";
            rule += ">" + cfvo_min + cfvo_max + color(cf.colors[0], "color") + "</dataBar>";
            if (modern) {
              // Excel 2010 and later draw the bar of an xls with a border and a separate negative bar. The
              // formatting of a data bar in the xls is the same for all of them, so the extension is written
              // with these defaults.
              char guid[48];
              std::snprintf(guid, sizeof(guid), "{%08X-0000-4000-8000-%012X}", static_cast<unsigned>(sheet + 1), static_cast<unsigned>(++x14_id));
              std::string border_attr = xls_color_attr(cf.colors[0], default_palette);
              rule += std::string("<extLst><ext uri=\"{B025F937-C7B1-47D3-B67F-A62EFF666E3E}\" "
                                  "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:id>") + guid + "</x14:id></ext></extLst>";
              c.x14 += std::string("<x14:cfRule type=\"dataBar\" id=\"") + guid + "\"><x14:dataBar minLength=\"0\" maxLength=\"100\"" +
                       (cf.show_value ? "" : " showValue=\"0\"") + " border=\"1\" negativeBarBorderColorSameAsPositive=\"0\">"
                       "<x14:cfvo type=\"autoMin\"/><x14:cfvo type=\"autoMax\"/><x14:borderColor" + border_attr + "/>"
                       "<x14:negativeFillColor rgb=\"FFFF0000\"/><x14:negativeBorderColor rgb=\"FFFF0000\"/>"
                       "<x14:axisColor rgb=\"FF000000\"/></x14:dataBar></x14:cfRule>";
            }
            rule += "</cfRule>";
          } else if (cf.type == 6 && !cf.cfvos.empty()) {
            static const char* const icons[17] = {"3Arrows", "3ArrowsGray", "3Flags", "3TrafficLights1", "3Signs",
                                                  "3TrafficLights2", "3Symbols", "3Symbols2", "4Arrows", "4ArrowsGray",
                                                  "4RedToBlack", "4Rating", "4TrafficLights", "5Arrows", "5ArrowsGray",
                                                  "5Rating", "5Quarters"};
            rule = "<cfRule type=\"iconSet\" priority=\"" + std::to_string(++cf_priority) + "\"><iconSet";
            rule += std::string(" iconSet=\"") + icons[cf.icon_set < 17 ? cf.icon_set : 0] + "\"";
            if (!cf.show_value) rule += " showValue=\"0\"";
            if (cf.reverse) rule += " reverse=\"1\"";
            rule += ">";
            for (size_t i = 0; i < cf.cfvos.size(); ++i) rule += cfvo(cf.cfvos[i]);
            rule += "</iconSet></cfRule>";
          } else if (cf.type == 1 || cf.type == 2) {
            std::string f1 = xls_rgce(XlsRgce(cf.rgce1), cctx, swapit, debug);
            std::string f2 = cf.rgce2.empty() ? "" : xls_rgce(XlsRgce(cf.rgce2), cctx, swapit, debug);
            if (f1.empty()) break;
            static const char* const operators[9] = {"", "between", "notBetween", "equal", "notEqual", "greaterThan",
                                                     "lessThan", "greaterThanOrEqual", "lessThanOrEqual"};
            rule = "<cfRule";
            if (cf.type == 1 && cf.comparison >= 1 && cf.comparison <= 8)
              rule += std::string(" type=\"cellIs\" operator=\"") + operators[cf.comparison] + "\"";
            else if (cf.type == 2) rule += " type=\"expression\"";
            else break;
            rule += " dxfId=\"" + std::to_string(this_dxf) + "\" priority=\"" + std::to_string(++cf_priority) + "\"";
            if (cf.stop_if_true) rule += " stopIfTrue=\"1\"";
            rule += "><formula>" + escape_xml(f1) + "</formula>";
            if (!f2.empty()) rule += "<formula>" + escape_xml(f2) + "</formula>";
            rule += "</cfRule>";
          }
          c.xml += rule;
          break;
        }

        case XlsHLink: {
          std::vector<uint8_t> rec(rsize);
          bin.read(reinterpret_cast<char*>(rec.data()), rsize);
          hlinks.push_back(xls_hlink(rec));
          break;
        }

        case XlsSxView: {
          std::vector<uint8_t> rec(rsize);
          bin.read(reinterpret_cast<char*>(rec.data()), rsize);
          XlsPivotView v;
          xls_pv_view(rec, v);
          if (v.cache < g.pivot_sources.size()) v.source = g.pivot_sources[v.cache];
          pv_views.push_back(v);
          break;
        }

        case XlsSXAddl: {
          std::vector<uint8_t> rec(rsize);
          bin.read(reinterpret_cast<char*>(rec.data()), rsize);
          // table style of the pivot table: class view, type 0x1E; flags, the length and the name
          if (pv_views.empty() || rec.size() < 16 || rec[4] != 0x00 || rec[5] != 0x1E) break;
          size_t cch = static_cast<size_t>(rec[14] | (rec[15] << 8));
          if (16 + 2 * cch > rec.size()) break;
          std::u16string name;
          for (size_t k = 0; k < cch; ++k) name.push_back(static_cast<char16_t>(rec[16 + 2 * k] | (rec[17 + 2 * k] << 8)));
          if (!name.empty()) pv_views.back().style = to_utf8(name);
          pv_views.back().style_flags = static_cast<uint16_t>(rec[12] | (rec[13] << 8));
          break;
        }

        case XlsSxvd: {
          std::vector<uint8_t> rec(rsize);
          bin.read(reinterpret_cast<char*>(rec.data()), rsize);
          if (pv_views.empty()) break;
          XlsPivotField f;
          xls_pv_field(rec, f);
          pv_views.back().fields.push_back(f);
          break;
        }

        case XlsSXVI: {
          std::vector<uint8_t> rec(rsize);
          bin.read(reinterpret_cast<char*>(rec.data()), rsize);
          if (pv_views.empty() || pv_views.back().fields.empty()) break;
          XlsPivotItem item;
          xls_pv_item(rec, item);
          pv_views.back().fields.back().items.push_back(item);
          break;
        }

        case XlsSxIvd: {
          std::vector<uint8_t> rec(rsize);
          bin.read(reinterpret_cast<char*>(rec.data()), rsize);
          if (pv_views.empty()) break;
          std::vector<uint16_t> axis;
          for (size_t i = 0; i + 2 <= rec.size(); i += 2) axis.push_back(static_cast<uint16_t>(xls_pv_u16(rec, i)));
          pv_views.back().axes.push_back(axis);
          break;
        }

        case XlsSXPI: {
          std::vector<uint8_t> rec(rsize);
          bin.read(reinterpret_cast<char*>(rec.data()), rsize);
          if (pv_views.empty()) break;
          for (size_t i = 0; i + 6 <= rec.size(); i += 6) {
            XlsPivotPage page;
            page.field = static_cast<uint16_t>(xls_pv_u16(rec, i));
            page.item = static_cast<uint16_t>(xls_pv_u16(rec, i + 2));
            pv_views.back().pages.push_back(page);
          }
          break;
        }

        case XlsSXDI: {
          std::vector<uint8_t> rec(rsize);
          bin.read(reinterpret_cast<char*>(rec.data()), rsize);
          if (pv_views.empty()) break;
          XlsPivotDataField d;
          xls_pv_data(rec, d);
          pv_views.back().data.push_back(d);
          break;
        }

        case XlsHLinkTooltip: {
          // frt header (4 bytes), the range (rows 16 bit, columns 8 bit), then the text up to the 0 at the end
          std::vector<uint8_t> rec(rsize);
          bin.read(reinterpret_cast<char*>(rec.data()), rsize);
          if (rsize < 10) break;
          auto u16 = [&](size_t at) { return static_cast<uint16_t>(rec[at] | (rec[at + 1] << 8)); };
          std::u16string text;
          for (size_t i = 10; i + 2 <= rsize; i += 2) {
            char16_t c = static_cast<char16_t>(u16(i));
            if (c == u'\0') break;
            text.push_back(c);
          }
          std::string ref = int_to_col(rec[8] + 1) + std::to_string(u16(4) + 1) + ":" + int_to_col(rec[9] + 1) + std::to_string(u16(6) + 1);
          tooltips[ref] = to_utf8(text);
          break;
        }

        default:
          break;
      }

      for (size_t i = 0; i < cells.size(); ++i) write_cell(cells[i]);

      bin.clear();
      bin.seekg(next, std::ios_base::beg);
    }

    finish_pending();
    open_sheet_data();
    if (row_open) out << "</row>" << std::endl;
    write_empty_rows(std::numeric_limits<uint32_t>::max());
    out << "</sheetData>" << std::endl;

    if (protect_sheet) {
      out << "<sheetProtection sheet=\"1\"";
      if (protect_objects) out << " objects=\"1\"";
      if (protect_scenarios) out << " scenarios=\"1\"";
      if (password) {
        char hex[8];
        std::snprintf(hex, sizeof(hex), "%04X", static_cast<unsigned>(password));
        out << " password=\"" << hex << "\"";
      }
      out << "/>" << std::endl;
    }

    std::string sort_state;
    if (!sort_conditions.empty() && (sort_sfp == 0 || sort_sfp == 2))
      sort_state = "<sortState ref=\"" + sort_ref + "\">" + sort_conditions + "</sortState>";

    if (have_autofilter) {
      std::string ref;
      for (size_t i = 0; i < g.names.size(); ++i) {
        if (g.names[i].name != "_xlnm._FilterDatabase" || g.names[i].sheet != sheet) continue;
        ref = g.names[i].formula.substr(g.names[i].formula.rfind('!') + 1);
        ref.erase(std::remove(ref.begin(), ref.end(), '$'), ref.end());
        break;
      }
      std::string filter_state = sort_sfp == 2 ? sort_state : "";
      if (ref.empty() && sort_sfp == 2) {
        // without the name the range is the one of the sort, including the header row
        ref = sort_ref;
        size_t colon = ref.find(':');
        size_t digits = ref.find_first_of("0123456789");
        ref = ref.substr(0, digits) + std::to_string(std::stoul(ref.substr(digits, colon - digits)) - 1) + ref.substr(colon);
      }
      if (!ref.empty()) {
        if (filter_columns.empty() && filter_state.empty()) {
          out << "<autoFilter ref=\"" << ref << "\"/>" << std::endl;
        } else {
          out << "<autoFilter ref=\"" << ref << "\">" << std::endl;
          for (std::map<uint16_t, std::string>::const_iterator it = filter_columns.begin(); it != filter_columns.end(); ++it)
            out << it->second << std::endl;
          out << filter_state << "</autoFilter>" << std::endl;
        }
      }
    }
    if (sort_sfp == 0) out << sort_state << std::endl;

    if (!merges.empty()) {
      out << "<mergeCells count=\"" << merges.size() << "\">" << std::endl;
      for (size_t i = 0; i < merges.size(); ++i) out << "<mergeCell ref=\"" << merges[i] << "\"/>" << std::endl;
      out << "</mergeCells>" << std::endl;
    }

    for (size_t i = 0; i < cond_fmts.size(); ++i) {
      if (cond_fmts[i].xml.empty()) continue;
      out << "<conditionalFormatting sqref=\"" << cond_fmts[i].sqref << "\">" << cond_fmts[i].xml << "</conditionalFormatting>" << std::endl;
    }

    if (!validations.empty()) {
      out << "<dataValidations count=\"" << validations.size() << "\">" << std::endl;
      for (size_t i = 0; i < validations.size(); ++i) out << validations[i] << std::endl;
      out << "</dataValidations>" << std::endl;
    }

    // each table is a part of its own, the sheet has a relationship to it
    std::vector<int32_t> table_files;
    for (size_t i = 0; i < tables.size(); ++i) {
      const XlsListTable& t = tables[i];
      int32_t file = table_file + static_cast<int32_t>(i);
      table_files.push_back(file);

      std::string first = int_to_col(t.col1 + 1) + std::to_string(t.row1 + 1);
      std::string last = int_to_col(t.col2 + 1) + std::to_string(t.row2 + 1);
      // the filter covers the header and the data, but not the total row
      std::string filter_last = int_to_col(t.col2 + 1) + std::to_string(t.row2 + 1 - (t.totals ? 1 : 0));

      std::ofstream tab(tablesDir + "/table" + std::to_string(file) + ".xml");
      tab << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" << std::endl;
      tab << "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"" << file << "\" name=\""
          << escape_xml(t.name) << "\" displayName=\"" << escape_xml(t.name) << "\" ref=\"" << first << ":" << last << "\"";
      if (!t.header) tab << " headerRowCount=\"0\"";
      if (t.totals) tab << " totalsRowCount=\"1\"";
      tab << ">" << std::endl;
      if (t.auto_filter && t.header) tab << "<autoFilter ref=\"" << first << ":" << filter_last << "\"/>" << std::endl;

      static const char* functions[] = {"", "average", "count", "countNums", "max", "min", "sum", "stdDev", "var"};
      tab << "<tableColumns count=\"" << t.columns.size() << "\">" << std::endl;
      for (size_t j = 0; j < t.columns.size(); ++j) {
        const XlsTableColumn& c = t.columns[j];
        tab << "<tableColumn id=\"" << j + 1 << "\" name=\"" << escape_xml(c.name) << "\"";
        if (t.totals && c.total_function >= 1 && c.total_function <= 8)
          tab << " totalsRowFunction=\"" << functions[c.total_function] << "\"";
        if (t.totals && !c.total_label.empty()) tab << " totalsRowLabel=\"" << escape_xml(c.total_label) << "\"";
        tab << "/>" << std::endl;
      }
      tab << "</tableColumns>" << std::endl;
      if (!t.style.empty()) {
        tab << "<tableStyleInfo name=\"" << escape_xml(t.style) << "\" showFirstColumn=\"" << t.first_column
            << "\" showLastColumn=\"" << t.last_column << "\" showRowStripes=\"" << t.row_stripes
            << "\" showColumnStripes=\"" << t.column_stripes << "\"/>" << std::endl;
      }
      tab << "</table>" << std::endl;
    }

    // External targets are relationships of the sheet, the hyperlink points to
    // them with an id. Links inside of the workbook only have a location.
    std::vector<std::string> rels;
    if (!hlinks.empty()) {
      out << "<hyperlinks>" << std::endl;
      for (size_t i = 0; i < hlinks.size(); ++i) {
        const XlsHyperlink& h = hlinks[i];
        std::string first = int_to_col(h.col1 + 1) + std::to_string(h.row1 + 1);
        std::string last = int_to_col(h.col2 + 1) + std::to_string(h.row2 + 1);
        std::string ref = first == last ? first : first + ":" + last;

        out << "<hyperlink ref=\"" << ref << "\"";
        if (!h.target.empty()) {
          rels.push_back(h.target);
          out << " r:id=\"rId" << rels.size() << "\"";
        }
        if (!h.location.empty()) out << " location=\"" << escape_xml(h.location) << "\"";
        std::string tip = tooltips[first == last ? first + ":" + first : ref];
        if (!tip.empty()) out << " tooltip=\"" << escape_xml(tip) << "\"";
        if (!h.display.empty()) out << " display=\"" << escape_xml(h.display) << "\"";
        out << "/>" << std::endl;
      }
      out << "</hyperlinks>" << std::endl;
    }

    if (h_center || v_center || print_grid || print_headers) {
      out << "<printOptions";
      if (h_center) out << " horizontalCentered=\"1\"";
      if (v_center) out << " verticalCentered=\"1\"";
      if (print_headers) out << " headings=\"1\"";
      if (print_grid) out << " gridLines=\"1\"";
      out << "/>" << std::endl;
    }

    if (have_margin) {
      out << "<pageMargins left=\"" << xls_num(margins[0]) << "\" right=\"" << xls_num(margins[1]) << "\" top=\""
          << xls_num(margins[2]) << "\" bottom=\"" << xls_num(margins[3]) << "\" header=\"" << xls_num(margins[4])
          << "\" footer=\"" << xls_num(margins[5]) << "\"/>" << std::endl;
    }

    if (have_setup) {
      bool no_orient = (setup_flags & 0x0040) != 0;
      bool valid = (setup_flags & 0x0004) == 0;
      out << "<pageSetup";
      if (valid) {
        if (paper > 0) out << " paperSize=\"" << paper << "\"";
        if (scale != 100 && !fit_to_page) out << " scale=\"" << scale << "\"";
        if (setup_flags & 0x0080) out << " firstPageNumber=\"" << page_start << "\" useFirstPageNumber=\"1\"";
        if (fit_to_page) out << " fitToWidth=\"" << fit_w << "\" fitToHeight=\"" << fit_h << "\"";
        if (!no_orient) out << " orientation=\"" << ((setup_flags & 0x0002) ? "portrait" : "landscape") << "\"";
        if (setup_flags & 0x0001) out << " pageOrder=\"overThenDown\"";
        if (setup_flags & 0x0008) out << " blackAndWhite=\"1\"";
        if (setup_flags & 0x0010) out << " draft=\"1\"";
        if (copies > 1) out << " copies=\"" << copies << "\"";
      }
      out << "/>" << std::endl;
    }

    if (!header.empty() || !footer.empty() || !even_header.empty() || !even_footer.empty() || !first_header.empty() || !first_footer.empty()) {
      out << "<headerFooter";
      if (hf_diff_even) out << " differentOddEven=\"1\"";
      if (hf_diff_first) out << " differentFirst=\"1\"";
      if (hf_no_scale) out << " scaleWithDoc=\"0\"";
      if (hf_no_align) out << " alignWithMargins=\"0\"";
      out << ">";
      if (!header.empty()) out << "<oddHeader>" << escape_xml(header) << "</oddHeader>";
      if (!footer.empty()) out << "<oddFooter>" << escape_xml(footer) << "</oddFooter>";
      if (!even_header.empty()) out << "<evenHeader>" << escape_xml(even_header) << "</evenHeader>";
      if (!even_footer.empty()) out << "<evenFooter>" << escape_xml(even_footer) << "</evenFooter>";
      if (!first_header.empty()) out << "<firstHeader>" << escape_xml(first_header) << "</firstHeader>";
      if (!first_footer.empty()) out << "<firstFooter>" << escape_xml(first_footer) << "</firstFooter>";
      out << "</headerFooter>" << std::endl;
    }

    if (!row_breaks.empty()) {
      out << "<rowBreaks count=\"" << row_breaks.size() << "\" manualBreakCount=\"" << row_breaks.size() << "\">";
      for (size_t i = 0; i < row_breaks.size(); ++i) out << "<brk id=\"" << row_breaks[i] << "\" max=\"16383\" man=\"1\"/>";
      out << "</rowBreaks>" << std::endl;
    }
    if (!col_breaks.empty()) {
      out << "<colBreaks count=\"" << col_breaks.size() << "\" manualBreakCount=\"" << col_breaks.size() << "\">";
      for (size_t i = 0; i < col_breaks.size(); ++i) out << "<brk id=\"" << col_breaks[i] << "\" max=\"1048575\" man=\"1\"/>";
      out << "</colBreaks>" << std::endl;
    }

    if (!ignored_errors.empty()) out << "<ignoredErrors>" << ignored_errors << "</ignoredErrors>" << std::endl;

    // the ids of the relationships: external links first, then comments (2), the drawing, then tables
    size_t drawing_rel = rels.size() + (notes.empty() ? 0 : 2) + 1;
    size_t first_table_rel = drawing_rel + (items.empty() ? 0 : 1);
    if (!items.empty()) out << "<drawing r:id=\"rId" << drawing_rel << "\"/>" << std::endl;
    if (!notes.empty()) out << "<legacyDrawing r:id=\"rId" << rels.size() + 2 << "\"/>" << std::endl;
    if (!table_files.empty()) {
      out << "<tableParts count=\"" << table_files.size() << "\">" << std::endl;
      for (size_t i = 0; i < table_files.size(); ++i) out << "<tablePart r:id=\"rId" << first_table_rel + i << "\"/>" << std::endl;
      out << "</tableParts>" << std::endl;
    }

    std::string x14_formats;
    for (size_t i = 0; i < cond_fmts.size(); ++i)
      if (!cond_fmts[i].x14.empty() && !cond_fmts[i].xml.empty())
        x14_formats += "<x14:conditionalFormatting>" + cond_fmts[i].x14 + "<xm:sqref>" + cond_fmts[i].sqref + "</xm:sqref></x14:conditionalFormatting>";
    if (!x14_formats.empty())
      out << "<extLst><ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">"
          << "<x14:conditionalFormattings xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">" << x14_formats
          << "</x14:conditionalFormattings></ext></extLst>" << std::endl;

    out << "</worksheet>" << std::endl;

    if (!notes.empty()) {
      std::ofstream com(commentsPath);
      std::vector<std::string> authors;
      for (size_t i = 0; i < notes.size(); ++i) {
        if (std::find(authors.begin(), authors.end(), notes[i].author) == authors.end()) authors.push_back(notes[i].author);
      }

      com << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" << std::endl;
      com << "<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" << std::endl;
      com << "<authors>" << std::endl;
      for (size_t i = 0; i < authors.size(); ++i) com << "<author>" << escape_xml(authors[i]) << "</author>" << std::endl;
      com << "</authors>" << std::endl;
      com << "<commentList>" << std::endl;
      for (size_t i = 0; i < notes.size(); ++i) {
        size_t author_id = static_cast<size_t>(std::find(authors.begin(), authors.end(), notes[i].author) - authors.begin());
        com << "<comment ref=\"" << int_to_col(notes[i].col + 1) << notes[i].row + 1 << "\" authorId=\"" << author_id << "\">"
            << "<text>";
        const std::u16string& chars = note_text[notes[i].id];
        const std::vector<std::pair<uint16_t, uint16_t> >& runs = note_runs[notes[i].id];
        if (runs.size() < 2) {
          com << "<r><rPr><sz val=\"9\"/><color indexed=\"81\"/><rFont val=\"Tahoma\"/><family val=\"2\"/></rPr>"
              << "<t xml:space=\"preserve\">" << escape_xml(to_utf8(chars)) << "</t></r>";
        } else {
          // the last run only marks the end of the text
          for (size_t r = 0; r + 1 < runs.size(); ++r) {
            size_t from = std::min<size_t>(runs[r].first, chars.size());
            size_t to = std::min<size_t>(runs[r + 1].first, chars.size());
            if (r + 2 == runs.size()) to = chars.size();
            if (to <= from) continue;
            // the number 4 is not used for fonts
            size_t font_id = runs[r].second;
            if (font_id > 4) --font_id;
            com << "<r>";
            if (font_id < note_fonts.size()) com << note_fonts[font_id].xml(note_palette, "rPr", "rFont");
            com << "<t xml:space=\"preserve\">" << escape_xml(to_utf8(chars.substr(from, to - from))) << "</t></r>";
          }
        }
        com << "</text></comment>" << std::endl;
      }
      com << "</commentList>" << std::endl;
      com << "</comments>" << std::endl;
    }

    if (!notes.empty()) {
      // the shapes of the comments, hidden until the cell is hovered
      std::ofstream vml(vmlPath);
      vml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" << std::endl;
      vml << "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" "
          << "xmlns:x=\"urn:schemas-microsoft-com:office:excel\">" << std::endl;
      // a block of 1024 shape numbers per idmap entry. Every sheet has 64 blocks for itself
      int32_t first_block = (comment_file - 1) * 64 + 1;
      int32_t n_blocks = static_cast<int32_t>(notes.size() / 1024) + 1;
      vml << "<o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"";
      for (int32_t b = 0; b < n_blocks; ++b) vml << (b > 0 ? "," : "") << first_block + b;
      vml << "\"/></o:shapelayout>" << std::endl;
      vml << "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\">"
          << "<v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype>" << std::endl;
      for (size_t i = 0; i < notes.size(); ++i) {
        uint32_t col = notes[i].col, row = notes[i].row;
        vml << "<v:shape id=\"_x0000_s" << first_block * 1024 + 1 + static_cast<int32_t>(i) << "\" type=\"#_x0000_t202\" "
            << "style=\"position:absolute;margin-left:107.25pt;margin-top:172.5pt;width:147pt;height:96pt;z-index:" << i + 1
            << ";visibility:hidden;mso-wrap-style:tight\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\">"
            << "<v:fill color2=\"#ffffe1\"/><v:shadow color=\"black\" obscured=\"t\"/><v:path o:connecttype=\"none\"/>"
            << "<v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"/></v:textbox>"
            << "<x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>";
        std::map<uint16_t, XlsShape>::const_iterator ns = note_shapes.find(notes[i].id);
        if (ns != note_shapes.end()) {
          // offsets in the file are fractions of the cell, in the shape they are pixels
          const XlsShape& a = ns->second;
          double w1 = col_px.count(a.col1) ? col_px[a.col1] : 64.0, w2 = col_px.count(a.col2) ? col_px[a.col2] : 64.0;
          double h1 = rows.count(a.row1) && rows[a.row1].custom ? rows[a.row1].height * 4.0 / 3.0 : 20.0;
          double h2 = rows.count(a.row2) && rows[a.row2].custom ? rows[a.row2].height * 4.0 / 3.0 : 20.0;
          vml << a.col1 << ", " << std::lround(a.dx1 / 1024.0 * w1) << ", " << a.row1 << ", " << std::lround(a.dy1 / 256.0 * h1) << ", "
              << a.col2 << ", " << std::lround(a.dx2 / 1024.0 * w2) << ", " << a.row2 << ", " << std::lround(a.dy2 / 256.0 * h2);
        } else {
          vml << col + 1 << ", 15, " << (row > 0 ? row - 1 : 0) << ", 10, " << col + 3 << ", 15, " << row + 4 << ", 4";
        }
        vml << "</x:Anchor><x:AutoFill>False</x:AutoFill>"
            << "<x:Row>" << row << "</x:Row><x:Column>" << col << "</x:Column></x:ClientData></v:shape>" << std::endl;
      }
      vml << "</xml>" << std::endl;
    }

    std::vector<std::pair<std::string, std::string> > drawing_rels;  // type, target
    int32_t nchart = 0;
    if (!items.empty()) {
      // the position: cell and distance to the corner as fraction of the size of the cell, in EMU
      auto col_emu = [&](uint16_t c) {
        std::map<uint16_t, double>::const_iterator it = col_px.find(c);
        return (it == col_px.end() ? 64.0 : it->second) * 9525.0;
      };
      auto row_emu = [&](uint32_t r) {
        std::map<uint32_t, Row>::const_iterator it = rows.find(r);
        return (it == rows.end() || it->second.height <= 0 ? 15.0 : it->second.height) * 12700.0;
      };

      std::ofstream dr(drawingPath);
      dr << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" << std::endl;
      dr << "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
         << "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
         << "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" << std::endl;
      size_t shape_no = 1;
      for (size_t i = 0; i < items.size(); ++i) {
        const DrawingItem& item = items[i];
        const XlsShape& sh = item.shape;
        auto emu = [](double v) { return static_cast<long long>(std::floor(v)); };
        if (item.chart) {
          if (item.chart_xml.empty()) continue;
          ++nchart;
          std::string cname = "chart" + std::to_string(chart_file + nchart - 1) + ".xml";
          std::ofstream cf(chartsDir + "/" + cname);
          cf << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" << std::endl << item.chart_xml << std::endl;
          cf.close();
          drawing_rels.push_back(std::make_pair(std::string("chart"), cname));
          ++shape_no;
          dr << "<xdr:twoCellAnchor><xdr:from><xdr:col>" << sh.col1 << "</xdr:col><xdr:colOff>"
             << emu(sh.dx1 / 1024.0 * col_emu(sh.col1)) << "</xdr:colOff><xdr:row>" << sh.row1 << "</xdr:row><xdr:rowOff>"
             << emu(sh.dy1 / 256.0 * row_emu(sh.row1)) << "</xdr:rowOff></xdr:from><xdr:to><xdr:col>" << sh.col2 << "</xdr:col><xdr:colOff>"
             << emu(sh.dx2 / 1024.0 * col_emu(sh.col2)) << "</xdr:colOff><xdr:row>" << sh.row2 << "</xdr:row><xdr:rowOff>"
             << emu(sh.dy2 / 256.0 * row_emu(sh.row2)) << "</xdr:rowOff></xdr:to>"
             << "<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"" << shape_no << "\" name=\"Chart " << nchart
             << "\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm>"
             << "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">"
             << "<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"rId" << drawing_rels.size()
             << "\"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>" << std::endl;
          continue;
        }
        if (item.autoshape) {
          ++shape_no;
          std::string fill = sh.fill.empty() ? "" : "<a:solidFill><a:srgbClr val=\"" + sh.fill + "\"/></a:solidFill>";
          std::string ln;
          if (!sh.line.empty() || sh.line_width > 0) {
            ln = "<a:ln";
            if (sh.line_width > 0) ln += " w=\"" + std::to_string(sh.line_width) + "\"";
            ln += ">";
            if (!sh.line.empty()) ln += "<a:solidFill><a:srgbClr val=\"" + sh.line + "\"/></a:solidFill>";
            ln += "</a:ln>";
          }
          dr << "<xdr:twoCellAnchor><xdr:from><xdr:col>" << sh.col1 << "</xdr:col><xdr:colOff>"
             << emu(sh.dx1 / 1024.0 * col_emu(sh.col1)) << "</xdr:colOff><xdr:row>" << sh.row1 << "</xdr:row><xdr:rowOff>"
             << emu(sh.dy1 / 256.0 * row_emu(sh.row1)) << "</xdr:rowOff></xdr:from><xdr:to><xdr:col>" << sh.col2 << "</xdr:col><xdr:colOff>"
             << emu(sh.dx2 / 1024.0 * col_emu(sh.col2)) << "</xdr:colOff><xdr:row>" << sh.row2 << "</xdr:row><xdr:rowOff>"
             << emu(sh.dy2 / 256.0 * row_emu(sh.row2)) << "</xdr:rowOff></xdr:to>"
             << "<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"" << shape_no << "\" name=\"Shape " << shape_no - 1
             << "\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm>"
             << "<a:prstGeom prst=\"" << xls_preset_geometry(sh.shape_type) << "\"><a:avLst/></a:prstGeom>" << fill << ln
             << "</xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>" << std::endl;
          continue;
        }
        R_xlen_t blip = static_cast<R_xlen_t>(item.blip);
        if (blip > media.size() || Rcpp::as<std::string>(media[blip - 1]).empty()) continue;
        std::string file = Rcpp::as<std::string>(media[blip - 1]);
        drawing_rels.push_back(std::make_pair(std::string("image"), file));
        ++shape_no;

        dr << "<xdr:twoCellAnchor editAs=\"oneCell\"><xdr:from><xdr:col>" << sh.col1 << "</xdr:col><xdr:colOff>"
           << emu(sh.dx1 / 1024.0 * col_emu(sh.col1)) << "</xdr:colOff><xdr:row>" << sh.row1 << "</xdr:row><xdr:rowOff>"
           << emu(sh.dy1 / 256.0 * row_emu(sh.row1)) << "</xdr:rowOff></xdr:from><xdr:to><xdr:col>" << sh.col2 << "</xdr:col><xdr:colOff>"
           << emu(sh.dx2 / 1024.0 * col_emu(sh.col2)) << "</xdr:colOff><xdr:row>" << sh.row2 << "</xdr:row><xdr:rowOff>"
           << emu(sh.dy2 / 256.0 * row_emu(sh.row2)) << "</xdr:rowOff></xdr:to>"
           << "<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"" << shape_no << "\" name=\"Picture " << drawing_rels.size()
           << "\"/><xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr>"
           << "<xdr:blipFill><a:blip r:embed=\"rId" << drawing_rels.size() << "\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>"
           << "<xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr>"
           << "</xdr:pic><xdr:clientData/></xdr:twoCellAnchor>" << std::endl;
      }
      dr << "</xdr:wsDr>" << std::endl;
      dr.close();

      std::ofstream drel(drawingRelsPath);
      drel << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" << std::endl;
      drel << "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" << std::endl;
      for (size_t i = 0; i < drawing_rels.size(); ++i)
        drel << "<Relationship Id=\"rId" << i + 1 << "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
             << drawing_rels[i].first << "\" Target=\"../" << (drawing_rels[i].first == "chart" ? "charts/" : "media/")
             << drawing_rels[i].second << "\"/>" << std::endl;
      drel << "</Relationships>" << std::endl;
    }

    // pivot tables: the cache of a table is written once, the first time it is used
    std::vector<int32_t> pivot_files;
    std::vector<int32_t> cache_ids(known_caches.begin(), known_caches.end());
    size_t n_known = cache_ids.size();
    if (!pv_views.empty()) {
      std::ifstream xf(xlsPath, std::ios::in | std::ios::binary | std::ios::ate);
      std::vector<uint8_t> xls_bytes(static_cast<size_t>(xf.tellg()));
      xf.seekg(0, std::ios_base::beg);
      xf.read(reinterpret_cast<char*>(xls_bytes.data()), static_cast<std::streamsize>(xls_bytes.size()));

      for (size_t i = 0; i < pv_views.size(); ++i) {
        XlsPivotView& v = pv_views[i];
        if (!v.source.supported()) {
          Rcpp::warning("xls: the pivot table %s on sheet %d is not based on a worksheet and is not imported.", v.name.c_str(), sheet);
          continue;
        }

        size_t at = 0;
        while (at < cache_ids.size() && cache_ids[at] != v.source.id) ++at;
        if (at == cache_ids.size()) {
          char stream_name[8];
          std::snprintf(stream_name, sizeof(stream_name), "%04X", v.source.id);
          std::vector<uint8_t> cache = cfb_stream(xls_bytes, stream_name);
          if (cache.empty()) {
            Rcpp::warning("xls: the cache of the pivot table %s on sheet %d is missing and the table is not imported.", v.name.c_str(), sheet);
            continue;
          }
          cache_ids.push_back(v.source.id);
          int32_t number = static_cast<int32_t>(cache_ids.size());

          std::ofstream def(xlDir + "/pivotCache/pivotCacheDefinition" + std::to_string(number) + ".xml");
          def << xls_pivot_cache_xml(cache, v.source);
          std::ofstream rec_rel(xlDir + "/pivotCache/_rels/pivotCacheDefinition" + std::to_string(number) + ".xml.rels");
          rec_rel << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" << std::endl;
          rec_rel << "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                  << "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords\" "
                  << "Target=\"pivotCacheRecords" << number << ".xml\"/></Relationships>" << std::endl;
          std::ofstream records(xlDir + "/pivotCache/pivotCacheRecords" + std::to_string(number) + ".xml");
          records << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" << std::endl;
          records << "<pivotCacheRecords xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\"/>" << std::endl;
        }

        int32_t cache_number = static_cast<int32_t>(at + 1);
        int32_t file = pivot_file + static_cast<int32_t>(pivot_files.size());
        pivot_files.push_back(file);

        std::ofstream pt(xlDir + "/pivotTables/pivotTable" + std::to_string(file) + ".xml");
        pt << xls_pivot_view_xml(v, cache_number);
        std::ofstream ptrel(xlDir + "/pivotTables/_rels/pivotTable" + std::to_string(file) + ".xml.rels");
        ptrel << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" << std::endl;
        ptrel << "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
              << "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition\" "
              << "Target=\"../pivotCache/pivotCacheDefinition" << cache_number << ".xml\"/></Relationships>" << std::endl;
      }
    }

    if (!rels.empty() || !notes.empty() || !table_files.empty() || !items.empty() || !pivot_files.empty()) {
      std::ofstream rel(relsPath);
      rel << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" << std::endl;
      rel << "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" << std::endl;
      for (size_t i = 0; i < rels.size(); ++i) {
        rel << "<Relationship Id=\"rId" << i + 1 << "\" "
            << "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" "
            << "Target=\"" << escape_xml(rels[i]) << "\" TargetMode=\"External\"/>" << std::endl;
      }
      if (!notes.empty()) {
        rel << "<Relationship Id=\"rId" << rels.size() + 1 << "\" "
            << "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" "
            << "Target=\"../comments" << comment_file << ".xml\"/>" << std::endl;
        rel << "<Relationship Id=\"rId" << rels.size() + 2 << "\" "
            << "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" "
            << "Target=\"../drawings/vmlDrawing" << comment_file << ".vml\"/>" << std::endl;
      }
      if (!items.empty()) {
        rel << "<Relationship Id=\"rId" << drawing_rel << "\" "
            << "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" "
            << "Target=\"../drawings/drawing" << drawing_file << ".xml\"/>" << std::endl;
      }
      for (size_t i = 0; i < table_files.size(); ++i) {
        rel << "<Relationship Id=\"rId" << first_table_rel + i << "\" "
            << "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" "
            << "Target=\"../tables/table" << table_files[i] << ".xml\"/>" << std::endl;
      }
      for (size_t i = 0; i < pivot_files.size(); ++i) {
        rel << "<Relationship Id=\"rId" << first_table_rel + table_files.size() + i << "\" "
            << "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable\" "
            << "Target=\"../pivotTables/pivotTable" << pivot_files[i] << ".xml\"/>" << std::endl;
      }
      rel << "</Relationships>" << std::endl;
    }

    if (n_unsupported > 0)
      Rcpp::warning("xls: %d formulas on sheet %d use a feature that is not supported (token 0x%02X). Their values are kept.",
                    n_unsupported, sheet, first_unsupported);

    out.close();
    bin.close();
    Rcpp::IntegerVector result = Rcpp::IntegerVector::create(static_cast<int>(notes.size()), static_cast<int>(tables.size()),
                                                             static_cast<int>(items.empty() ? 0 : 1), static_cast<int>(nchart),
                                                             static_cast<int>(pivot_files.size()));
    for (size_t i = n_known; i < cache_ids.size(); ++i) result.push_back(cache_ids[i]);
    return result;
  } else {
    Rcpp::stop("xls: the file can not be read");
  }
}
