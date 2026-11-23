#ifndef XLS_LEGACY_H
#define XLS_LEGACY_H

#include <cstdint>
#include <cstring>
#include <map>
#include <set>
#include <string>
#include <utility>
#include <vector>

#include "xls_cfb.h"
#include "xls_defines.h"

// Files of BIFF2 to BIFF7 are converted to the stream of a BIFF8 file, which the xls
// reader then reads like any other. Text is converted to UTF-16 with the code page of
// the file. BIFF5 and BIFF7 also have lines, rectangles, ovals, text boxes, bitmaps and
// charts, which become the drawing records of BIFF8, BIFF3 and BIFF4 only have lines,
// rectangles and ovals. Not converted are text boxes, pictures and charts of BIFF3 and
// BIFF4, arcs, polygons, controls and metafile pictures, pivot tables, references to
// other workbooks, array constants and the defined names of BIFF2. A formula that can
// not be converted keeps its last result as a value.

typedef std::string (*XlsTextFn)(const uint8_t*, size_t, uint32_t);

struct XlsBytes {
  std::vector<uint8_t> b;
  void u8(uint32_t v) { b.push_back(static_cast<uint8_t>(v & 0xFFu)); }
  void u16(uint32_t v) {
    u8(v);
    u8(v >> 8);
  }
  void u32(uint32_t v) {
    u16(v);
    u16(v >> 16);
  }
  void raw(const std::vector<uint8_t>& v, size_t from, size_t n) {
    if (from + n > v.size()) Rcpp::stop("xls: file is truncated");
    b.insert(b.end(), v.begin() + static_cast<std::ptrdiff_t>(from), v.begin() + static_cast<std::ptrdiff_t>(from + n));
  }
  void raw(const XlsBytes& v) { b.insert(b.end(), v.b.begin(), v.b.end()); }
};

struct XlsLegacyXf {
  uint16_t font = 0, fmt = 0, parent = 0xFFF;
  bool locked = true, hidden = false, style = false;
  uint8_t halign = 0, wrap = 0, valign = 2, rotation = 0, used = 0xFC;
  uint8_t border_style[4] = {0, 0, 0, 0};  // left, right, top, bottom
  uint8_t border_color[4] = {0, 0, 0, 0};
  uint8_t pattern = 0, fore = 64, back = 65;
};

inline bool xls_legacy_detect(const std::vector<uint8_t>& file) {
  if (file.size() < 8) return false;
  bool cfb = file[0] == 0xD0 && file[1] == 0xCF && file[2] == 0x11 && file[3] == 0xE0;
  if (!cfb) {
    uint32_t id = cfb_u16(file, 0);
    return id == 0x0009 || id == 0x0209 || id == 0x0409 || id == 0x0809;
  }
  std::vector<uint8_t> workbook = cfb_stream(file, "Workbook");
  if (workbook.empty()) return !cfb_stream(file, "Book").empty();
  return workbook.size() >= 6 && cfb_u16(workbook, 0) == 0x0809 && cfb_u16(workbook, 4) == 0x0500;
}

struct XlsLegacy {
  XlsTextFn text_fn;
  int ver = 0;  // 2, 3, 4, or 5 for BIFF5 and BIFF7
  uint32_t codepage = 1252;

  std::vector<uint8_t> g_fonts, g_styles, g_misc, g_names, g_palette;
  std::vector<std::pair<uint16_t, std::u16string> > formats;  // as numbered in the file
  std::vector<XlsLegacyXf> xfs;
  bool has_xf = false;
  std::map<uint32_t, uint16_t> attr_xf;
  uint16_t ixfe = 0;
  uint32_t nformats = 0;
  std::vector<std::vector<uint16_t> > xtis;
  std::vector<uint8_t>* cur = nullptr;
  std::set<uint32_t> dead_exp;

  struct Pending {
    bool active = false, ok = false, extra_failed = false, has_string = false;
    uint16_t row = 0, col = 0, xf = 0;
    uint8_t result[8] = {0, 0, 0, 0, 0, 0, 0, 0};
    XlsBytes formula;
    std::vector<std::pair<uint16_t, XlsBytes> > extra;
    std::u16string text;
  } pend;

  struct Note {
    bool active = false;
    uint16_t row = 0, col = 0;
    size_t total = 0;
    std::string bytes;
  } note;
  uint16_t next_object = 1;

  // pictures wait for the IMDATA records that follow the object
  struct Pic {
    bool active = false;
    std::vector<uint8_t> obj;
    std::string data;
  } pic;
  std::vector<std::string> blips;  // bitmaps as DIB, in the order of the objects
  std::vector<uint32_t> palette = std::vector<uint32_t>(XlsDefaultPalette, XlsDefaultPalette + 56);
  uint32_t next_shape = 0x400;

  explicit XlsLegacy(XlsTextFn fn) : text_fn(fn) {}

  uint32_t xti_index(uint16_t first, uint16_t last) {
    std::vector<uint16_t> xti(3, 0);
    xti[1] = first;
    xti[2] = last;
    size_t ixti = 0;
    while (ixti < xtis.size() && xtis[ixti] != xti) ++ixti;
    if (ixti == xtis.size()) xtis.push_back(xti);
    return static_cast<uint32_t>(ixti);
  }

  std::u16string str8(const std::vector<uint8_t>& d, size_t pos, size_t n) const {
    if (pos + n > d.size()) Rcpp::stop("xls: file is truncated");
    std::string u8 = text_fn(d.data() + pos, n, codepage);
    std::u16string out;
    size_t i = 0;
    while (i < u8.size()) {
      uint32_t c = static_cast<uint8_t>(u8[i]);
      size_t extra = 0;
      if (c >= 0xF0) {
        c &= 0x07;
        extra = 3;
      } else if (c >= 0xE0) {
        c &= 0x0F;
        extra = 2;
      } else if (c >= 0xC0) {
        c &= 0x1F;
        extra = 1;
      }
      ++i;
      for (size_t k = 0; k < extra && i < u8.size(); ++k, ++i) c = (c << 6) | (static_cast<uint8_t>(u8[i]) & 0x3Fu);
      if (c >= 0x10000) {
        c -= 0x10000;
        out += static_cast<char16_t>(0xD800 + (c >> 10));
        out += static_cast<char16_t>(0xDC00 + (c & 0x3FF));
      } else {
        out += static_cast<char16_t>(c);
      }
    }
    return out;
  }

  static void ustr(XlsBytes& o, const std::u16string& s, bool long_cch) {
    if (long_cch) o.u16(static_cast<uint32_t>(s.size()));
    else o.u8(static_cast<uint32_t>(s.size()));
    o.u8(1);
    for (size_t i = 0; i < s.size(); ++i) o.u16(static_cast<uint32_t>(s[i]));
  }

  static void rec(std::vector<uint8_t>& out, uint16_t id, const XlsBytes& data) {
    if (data.b.size() > 8224) Rcpp::stop("xls: record is too long");
    XlsBytes h;
    h.u16(id);
    h.u16(static_cast<uint32_t>(data.b.size()));
    out.insert(out.end(), h.b.begin(), h.b.end());
    out.insert(out.end(), data.b.begin(), data.b.end());
  }

  static uint16_t color3(uint32_t c) { return static_cast<uint16_t>(c < 16 ? c + 8 : c); }

  uint16_t xf_from_attr(const std::vector<uint8_t>& d, size_t p) {
    if (p + 3 > d.size()) Rcpp::stop("xls: file is truncated");
    uint32_t a0 = d[p], a1 = d[p + 1], a2 = d[p + 2];
    if (has_xf) {
      uint16_t x = static_cast<uint16_t>(a0 & 0x3Fu);
      return x == 0x3F ? ixfe : x;
    }
    if (xfs.empty()) {
      for (size_t i = 0; i < 16; ++i) {
        XlsLegacyXf x;
        x.style = i < 15;
        x.parent = x.style ? 0xFFF : 0;
        xfs.push_back(x);
      }
      attr_xf[0x000040u] = 15;
    }
    uint32_t key = a0 | (a1 << 8) | (a2 << 16);
    std::map<uint32_t, uint16_t>::const_iterator it = attr_xf.find(key);
    if (it != attr_xf.end()) return it->second;
    XlsLegacyXf x;
    x.parent = 0;
    x.locked = (a0 & 0x40u) != 0;
    x.hidden = (a0 & 0x80u) != 0;
    x.fmt = static_cast<uint16_t>(a1 & 0x3Fu);
    x.font = static_cast<uint16_t>((a1 & 0xC0u) >> 6);
    x.halign = static_cast<uint8_t>(a2 & 7u);
    for (size_t k = 0; k < 4; ++k) {
      if (a2 & (0x08u << k)) {
        x.border_style[k] = 1;
        x.border_color[k] = 8;
      }
    }
    if (a2 & 0x80u) {
      x.pattern = 4;
      x.fore = 8;
      x.back = 9;
    }
    xfs.push_back(x);
    uint16_t idx = static_cast<uint16_t>(xfs.size() - 1);
    attr_xf[key] = idx;
    return idx;
  }

  bool rgce(const std::vector<uint8_t>& d, size_t from, size_t len, XlsBytes& o, bool in_name = false) {
    size_t p = from, end = from + len;
    if (end > d.size()) return false;
    std::vector<std::pair<size_t, size_t> > mem;  // end of the sub expression, position of its size in o

    while (p < end) {
      while (!mem.empty() && p >= mem.back().first) {
        size_t at = mem.back().second;
        size_t size = o.b.size() - (at + 2);
        o.b[at] = static_cast<uint8_t>(size & 0xFFu);
        o.b[at + 1] = static_cast<uint8_t>((size >> 8) & 0xFFu);
        mem.pop_back();
      }

      uint8_t op = d[p];
      uint8_t base = op < 0x20 ? op : static_cast<uint8_t>(0x20 | (op & 0x1F));
      size_t left = end - p;

      if (base == 0x01) {
        if (ver == 2) {
          if (left < 4) return false;
          o.u8(op);
          o.u16(cfb_u16(d, p + 1));
          o.u16(d[p + 3]);
          p += 4;
        } else {
          if (left < 5) return false;
          o.raw(d, p, 5);
          p += 5;
        }
      } else if (base >= 0x03 && base <= 0x16) {
        o.u8(op);
        p += 1;
      } else if (base == 0x17) {
        if (left < 2) return false;
        size_t n = d[p + 1];
        if (left < 2 + n) return false;
        std::u16string s = str8(d, p + 2, n);
        o.u8(op);
        ustr(o, s, false);
        p += 2 + n;
      } else if (base == 0x19) {
        if (left < 4) return false;
        size_t size = 4;
        if (d[p + 1] == 0x04) size = 2 * static_cast<size_t>(cfb_u16(d, p + 2)) + 6;
        if (left < size) return false;
        o.raw(d, p, size);
        p += size;
      } else if (base >= 0x1C && base <= 0x1F) {
        size_t size = base == 0x1C || base == 0x1D ? 2 : (base == 0x1E ? 3 : 9);
        if (left < size) return false;
        o.raw(d, p, size);
        p += size;
      } else if (base == 0x21) {
        if (ver <= 3) {
          if (left < 2) return false;
          o.u8(op);
          o.u16(d[p + 1]);
          p += 2;
        } else {
          if (left < 3) return false;
          o.raw(d, p, 3);
          p += 3;
        }
      } else if (base == 0x22) {
        if (ver <= 3) {
          if (left < 3) return false;
          o.u8(op);
          o.u8(d[p + 1]);
          o.u16(d[p + 2]);
          p += 3;
        } else {
          if (left < 4) return false;
          o.raw(d, p, 4);
          p += 4;
        }
      } else if (base == 0x23) {
        size_t size = ver == 5 ? 15 : 11;
        if (ver == 2 || left < size) return false;
        o.u8(op);
        o.u16(cfb_u16(d, p + 1));
        o.u16(0);
        p += size;
      } else if (base == 0x24) {
        if (left < 4) return false;
        uint32_t rf = cfb_u16(d, p + 1);
        if (in_name && ver < 5) {
          o.u8((op & 0x60u) | 0x3Au);
          o.u16(xti_index(0, 0));
        } else {
          o.u8(op);
        }
        o.u16(rf & 0x3FFFu);
        o.u16(d[p + 3] | (rf & 0xC000u));
        p += 4;
      } else if (base == 0x25) {
        if (left < 7) return false;
        uint32_t rf = cfb_u16(d, p + 1), rl = cfb_u16(d, p + 3);
        uint32_t r1 = rf & 0x3FFFu, r2 = rl & 0x3FFFu;
        if (r1 == 0 && r2 == 0x3FFF) r2 = 0xFFFF;
        if (in_name && ver < 5) {
          o.u8((op & 0x60u) | 0x3Bu);
          o.u16(xti_index(0, 0));
        } else {
          o.u8(op);
        }
        o.u16(r1);
        o.u16(r2);
        o.u16(d[p + 5] | (rf & 0xC000u));
        o.u16(d[p + 6] | (rl & 0xC000u));
        p += 7;
      } else if (base == 0x29) {
        if (left < 3) return false;
        size_t sub_end = p + 3 + cfb_u16(d, p + 1);
        if (sub_end > end) return false;
        o.u8(op);
        mem.push_back(std::make_pair(sub_end, o.b.size()));
        o.u16(0);
        p += 3;
      } else if (base == 0x2A) {
        if (left < 4) return false;
        o.u8(op);
        o.u32(0);
        p += 4;
      } else if (base == 0x2B) {
        if (left < 7) return false;
        o.u8(op);
        o.u32(0);
        o.u32(0);
        p += 7;
      } else if (base == 0x2C) {
        if (left < 4) return false;
        uint32_t rf = cfb_u16(d, p + 1);
        uint32_t row = rf & 0x3FFFu;
        if ((rf & 0x8000u) && (row & 0x2000u)) row |= 0xC000u;
        o.u8(op);
        o.u16(row);
        o.u16(d[p + 3] | (rf & 0xC000u));
        p += 4;
      } else if (base == 0x2D) {
        if (left < 7) return false;
        uint32_t rf = cfb_u16(d, p + 1), rl = cfb_u16(d, p + 3);
        uint32_t r1 = rf & 0x3FFFu, r2 = rl & 0x3FFFu;
        if ((rf & 0x8000u) && (r1 & 0x2000u)) r1 |= 0xC000u;
        if ((rl & 0x8000u) && (r2 & 0x2000u)) r2 |= 0xC000u;
        o.u8(op);
        o.u16(r1);
        o.u16(r2);
        o.u16(d[p + 5] | (rf & 0xC000u));
        o.u16(d[p + 6] | (rl & 0xC000u));
        p += 7;
      } else if (base == 0x3A || base == 0x3B || base == 0x3C || base == 0x3D) {
        size_t size = (base == 0x3A || base == 0x3C) ? 18 : 21;
        if (ver != 5 || left < size) return false;
        int16_t ixals = static_cast<int16_t>(cfb_u16(d, p + 1));
        int16_t first = static_cast<int16_t>(cfb_u16(d, p + 11));
        int16_t last = static_cast<int16_t>(cfb_u16(d, p + 13));
        if (ixals > 0 || first < 0 || last < first) return false;
        uint32_t ixti = xti_index(static_cast<uint16_t>(first), static_cast<uint16_t>(last));
        o.u8(op);
        o.u16(ixti);
        if (base == 0x3A) {
          uint32_t rf = cfb_u16(d, p + 15);
          o.u16(rf & 0x3FFFu);
          o.u16(d[p + 17] | (rf & 0xC000u));
        } else if (base == 0x3B) {
          uint32_t rf = cfb_u16(d, p + 15), rl = cfb_u16(d, p + 17);
          uint32_t r1 = rf & 0x3FFFu, r2 = rl & 0x3FFFu;
          if (r1 == 0 && r2 == 0x3FFF) r2 = 0xFFFF;
          o.u16(r1);
          o.u16(r2);
          o.u16(d[p + 19] | (rf & 0xC000u));
          o.u16(d[p + 20] | (rl & 0xC000u));
        } else if (base == 0x3C) {
          o.u32(0);
        } else {
          o.u32(0);
          o.u32(0);
        }
        p += size;
      } else {
        return false;
      }
    }

    while (!mem.empty()) {
      size_t at = mem.back().second;
      size_t size = o.b.size() - (at + 2);
      o.b[at] = static_cast<uint8_t>(size & 0xFFu);
      o.b[at + 1] = static_cast<uint8_t>((size >> 8) & 0xFFu);
      mem.pop_back();
    }
    return true;
  }

  void emit_note() {
    if (!note.active) return;
    note.active = false;
    std::vector<uint8_t>& out = *cur;
    std::vector<uint8_t> raw(note.bytes.begin(), note.bytes.end());
    std::u16string text = str8(raw, 0, raw.size());
    if (text.size() > 4000) text.resize(4000);
    uint16_t id = next_object++;

    XlsBytes o;
    o.u16(0x0015);
    o.u16(0x0012);
    o.u16(0x0019);
    o.u16(id);
    o.u16(0x4011);
    for (size_t i = 0; i < 12; ++i) o.u8(0);
    o.u32(0);
    rec(out, 0x005D, o);

    o = XlsBytes();
    o.u16(0x0212);
    o.u16(0);
    for (size_t i = 0; i < 6; ++i) o.u8(0);
    o.u16(static_cast<uint32_t>(text.size()));
    o.u16(0);
    o.u32(0);
    rec(out, 0x01B6, o);

    if (!text.empty()) {
      o = XlsBytes();
      o.u8(1);
      for (size_t i = 0; i < text.size(); ++i) o.u16(static_cast<uint32_t>(text[i]));
      rec(out, 0x003C, o);
    }

    o = XlsBytes();
    o.u16(note.row);
    o.u16(note.col);
    o.u16(0);
    o.u16(id);
    o.u16(0);
    o.u8(0);
    rec(out, 0x001C, o);
  }

  static void esc(XlsBytes& o, uint32_t ver_inst, uint32_t type, const XlsBytes& body) {
    o.u16(ver_inst);
    o.u16(type);
    o.u32(static_cast<uint32_t>(body.b.size()));
    o.raw(body);
  }

  // The drawing objects of BIFF5 are converted to the Escher shapes of BIFF8: lines,
  // rectangles, ovals, text boxes and bitmaps. Groups are flattened, arcs, polygons and
  // controls are dropped.
  void shape5(const std::vector<uint8_t>& d, uint32_t blip) {
    uint32_t ot = cfb_u16(d, 4);
    uint32_t id = cfb_u16(d, 6);
    if (!(d[9] & 0x02u)) return;
    bool line_object = ot == 1;
    if (ot != 1 && ot != 2 && ot != 3 && ot != 5 && ot != 6 && ot != 8) return;
    if (ver != 5 && ot != 1 && ot != 2 && ot != 3) return;
    if (ot == 8 && blip == 0) return;
    size_t base = ver == 5 ? 34 : 30;
    size_t cb_text = 0, cb_runs = 0;
    if (ot == 6) {
      if (d.size() < 71) return;
      cb_text = cfb_u16(d, 44);
      cb_runs = cfb_u16(d, 48);
      if (cb_text + (cb_text & 1u) + cb_runs + 71 > d.size()) return;
    }

    auto color = [&](uint32_t icv) {
      uint32_t rgb = 0;
      if (icv < 8) rgb = XlsDefaultPalette[icv];
      else if (icv < 64) rgb = palette[icv - 8];
      else if (icv == 65) rgb = 0xFFFFFF;
      return ((rgb >> 16) & 0xFFu) | (rgb & 0xFF00u) | ((rgb & 0xFFu) << 16);
    };

    std::vector<std::pair<uint32_t, uint32_t> > props;
    uint32_t flags = 0x0A00;
    uint32_t type = 1;
    size_t border = base + 4;
    if (line_object) {
      type = 20;
      border = base;
    } else {
      if (d.size() < base + 10) return;
      bool auto_fill = (d[base + 3] & 1u) != 0;
      uint32_t fls = d[base + 2];
      if (auto_fill) {
        props.push_back(std::make_pair(0x0181u, 0x00FFFFFFu));
        props.push_back(std::make_pair(0x01BFu, 0x00100010u));
      } else if (fls == 0) {
        props.push_back(std::make_pair(0x01BFu, 0x00100000u));
      } else {
        props.push_back(std::make_pair(0x0181u, color(d[base + 1])));
        props.push_back(std::make_pair(0x01BFu, 0x00100010u));
      }
      if (ot == 2 && (cfb_u16(d, base + 8) & 1u)) type = 2;
      if (ot == 3) type = 3;
      if (ot == 5) type = 201;
      if (ot == 6) type = 202;
      if (ot == 8) {
        type = 75;
        props.push_back(std::make_pair(0x4104u, blip));
      }
    }
    if (d.size() < border + (line_object ? 7 : 4)) return;
    uint32_t icv = d[border], lns = d[border + 1], lnw = d[border + 2];
    bool auto_line = (d[border + 3] & 1u) != 0;
    if (auto_line) {
      props.push_back(std::make_pair(0x01C0u, 0u));
      props.push_back(std::make_pair(0x01CBu, 12700u));
      props.push_back(std::make_pair(0x01FFu, 0x00080008u));
    } else if (lns == 5) {
      props.push_back(std::make_pair(0x01FFu, 0x00080000u));
    } else {
      static const uint32_t widths[4] = {3175, 12700, 25400, 38100};
      props.push_back(std::make_pair(0x01C0u, color(icv)));
      props.push_back(std::make_pair(0x01CBu, widths[lnw & 3u]));
      if (lns >= 1 && lns <= 4) props.push_back(std::make_pair(0x01CEu, lns));
      props.push_back(std::make_pair(0x01FFu, 0x00080008u));
    }
    if (line_object) {
      uint32_t sest = d[border + 4] & 0x0Fu;
      if (sest != 0) props.push_back(std::make_pair(0x01D1u, (sest == 2 || sest == 4) ? 1u : 5u));
      uint32_t iqu = d[border + 6];
      if (iqu == 1 || iqu == 2) flags |= 0x40;
      if (iqu == 2 || iqu == 3) flags |= 0x80;
    }
    std::sort(props.begin(), props.end());

    XlsBytes sp;
    sp.u32(next_shape++);
    sp.u32(flags);
    XlsBytes opt;
    for (size_t i = 0; i < props.size(); ++i) {
      opt.u16(props[i].first);
      opt.u32(props[i].second);
    }
    XlsBytes anchor;
    anchor.u16(0);
    anchor.u16(cfb_u16(d, 10));
    anchor.u16(cfb_u16(d, 12));
    anchor.u16(cfb_u16(d, 14));
    anchor.u16(cfb_u16(d, 16) / 4);
    anchor.u16(cfb_u16(d, 18));
    anchor.u16(cfb_u16(d, 20));
    anchor.u16(cfb_u16(d, 22));
    anchor.u16(cfb_u16(d, 24) / 4);
    XlsBytes inner;
    esc(inner, 0x2 | (type << 4), 0xF00A, sp);
    esc(inner, 0x3 | (static_cast<uint32_t>(props.size()) << 4), 0xF00B, opt);
    esc(inner, 0x0, 0xF010, anchor);
    XlsBytes outer;
    esc(outer, 0xF, 0xF004, inner);
    std::vector<uint8_t>& out = *cur;
    rec(out, 0x00EC, outer);

    XlsBytes o;
    o.u16(0x0015);
    o.u16(0x0012);
    o.u16(ot == 8 ? 0x0008 : ot);
    o.u16(id);
    o.u16(0x4011);
    for (size_t i = 0; i < 12; ++i) o.u8(0);
    o.u32(0);
    rec(out, 0x005D, o);

    if (ot != 6) return;
    size_t text_at = d.size() - cb_runs - (cb_text & 1u) - cb_text;
    std::u16string text = str8(d, text_at, cb_text);
    o = XlsBytes();
    o.u16((d[54] & 0x7Eu) | ((d[55] & 0x02u) ? 0x200u : 0u));
    o.u16(cfb_u16(d, 56));
    for (size_t i = 0; i < 6; ++i) o.u8(0);
    o.u16(static_cast<uint32_t>(text.size()));
    o.u16(static_cast<uint32_t>(cb_runs));
    o.u32(0);
    rec(out, 0x01B6, o);
    if (text.empty()) return;
    o = XlsBytes();
    o.u8(1);
    for (size_t i = 0; i < text.size(); ++i) o.u16(static_cast<uint32_t>(text[i]));
    rec(out, 0x003C, o);
    if (cb_runs == 0) return;
    o = XlsBytes();
    for (size_t i = 0; i + 8 <= cb_runs; i += 8) {
      size_t first = cfb_u16(d, d.size() - cb_runs + i);
      o.u16(static_cast<uint32_t>(std::min(first, text.size())));
      o.u16(cfb_u16(d, d.size() - cb_runs + i + 2));
      o.u32(0);
    }
    rec(out, 0x003C, o);
  }

  // The bitmap of an IMDATA record starts with a BITMAPCOREHEADER, the palette has 3
  // bytes per color. DIBs need a BITMAPINFOHEADER and 4 bytes per color.
  void emit_pic() {
    if (!pic.active) return;
    pic.active = false;
    const std::string& p = pic.data;
    if (p.size() < 8 || static_cast<uint8_t>(p[0]) != 9 || p[1] != 0) return;
    std::vector<uint8_t> raw(p.begin() + 8, p.end());
    if (raw.size() < 12 || cfb_u32(raw, 0) != 12) return;
    uint32_t width = cfb_u16(raw, 4), height = cfb_u16(raw, 6), bits = cfb_u16(raw, 10);
    if (width == 0 || height == 0) return;
    if (bits != 1 && bits != 4 && bits != 8 && bits != 24) return;
    size_t colors = bits <= 8 ? (static_cast<size_t>(1) << bits) : 0;
    size_t row = ((static_cast<size_t>(width) * bits + 31) / 32) * 4;
    size_t pixels_at = 12 + colors * 3;
    if (pixels_at + row * height > raw.size()) return;
    XlsBytes dib;
    dib.u32(40);
    dib.u32(width);
    dib.u32(height);
    dib.u16(1);
    dib.u16(bits);
    for (size_t i = 0; i < 6; ++i) dib.u32(0);
    for (size_t i = 0; i < colors; ++i) {
      dib.u8(raw[12 + 3 * i]);
      dib.u8(raw[13 + 3 * i]);
      dib.u8(raw[14 + 3 * i]);
      dib.u8(0);
    }
    dib.b.insert(dib.b.end(), raw.begin() + static_cast<std::ptrdiff_t>(pixels_at),
                 raw.begin() + static_cast<std::ptrdiff_t>(pixels_at + row * height));
    blips.push_back(std::string(dib.b.begin(), dib.b.end()));
    shape5(pic.obj, static_cast<uint32_t>(blips.size()));
  }

  // The records of a chart are the same in BIFF5 and BIFF8, except for the strings,
  // the formulas of the references to the data and a few records that got more fields.
  void chart5(uint16_t id, const std::vector<uint8_t>& d, std::vector<uint8_t>& out) {
    XlsBytes o;
    if (id == 0x100D && d.size() >= 3) {
      size_t cch = std::min<size_t>(d[2], d.size() - 3);
      o.u16(cfb_u16(d, 0));
      ustr(o, str8(d, 3, cch), false);
    } else if (id == 0x1051 && d.size() >= 8) {
      size_t cce = cfb_u16(d, 6);
      XlsBytes rg;
      bool ok = cce > 0 && rgce(d, 8, cce, rg);
      o.raw(d, 0, 6);
      if (!ok) {
        o.b[1] = 0;
        rg = XlsBytes();
      }
      o.u16(static_cast<uint32_t>(rg.b.size()));
      o.raw(rg);
    } else if (id == 0x1009 && d.size() >= 12) {
      o.raw(d, 0, 12);
      o.u16(0);
      o.u16(0);
      o.u32(100);
    } else if (id == 0x1025 && d.size() >= 26) {
      static const uint8_t rotation[4] = {0, 0xFF, 90, 180};
      o.raw(d, 0, 26);
      o.b[25] = static_cast<uint8_t>(o.b[25] & ~7u);
      o.u16(0x4D);
      o.u16(rotation[d[25] & 3u]);
    } else if (id == 0x0204 && d.size() >= 8) {
      size_t cch = std::min<size_t>(cfb_u16(d, 6), d.size() - 8);
      o.raw(d, 0, 6);
      ustr(o, str8(d, 8, cch), true);
    } else if ((id >= 0x1000 && id < 0x2000) || id == 0x0203 || id == 0x0201) {
      o.raw(d, 0, d.size());
    } else {
      return;
    }
    rec(out, id, o);
  }

  void flush() {
    if (!pend.active) return;
    pend.active = false;
    std::vector<uint8_t>& out = *cur;
    bool is_text = pend.result[6] == 0xFF && pend.result[7] == 0xFF && (pend.result[0] == 0 || pend.result[0] == 3);

    if (pend.ok && !pend.extra_failed) {
      rec(out, 0x0006, pend.formula);
      for (size_t i = 0; i < pend.extra.size(); ++i) rec(out, pend.extra[i].first, pend.extra[i].second);
      if (is_text) {
        XlsBytes s;
        ustr(s, pend.text, true);
        rec(out, 0x0207, s);
      }
      return;
    }

    XlsBytes v;
    v.u16(pend.row);
    v.u16(pend.col);
    v.u16(pend.xf);
    if (pend.result[6] != 0xFF || pend.result[7] != 0xFF) {
      for (size_t i = 0; i < 8; ++i) v.u8(pend.result[i]);
      rec(out, 0x0203, v);
    } else if (is_text) {
      ustr(v, pend.text, true);
      rec(out, 0x0204, v);
    } else {
      v.u8(pend.result[2]);
      v.u8(pend.result[0] == 2 ? 1 : 0);
      rec(out, 0x0205, v);
    }
  }

  void decode_xf(const std::vector<uint8_t>& d) {
    XlsLegacyXf x;
    if (ver == 2) {
      if (d.size() < 4) Rcpp::stop("xls: file is truncated");
      uint32_t f = d[2], a = d[3];
      x.font = d[0];
      x.fmt = static_cast<uint16_t>(f & 0x3Fu);
      x.locked = (f & 0x40u) != 0;
      x.hidden = (f & 0x80u) != 0;
      x.halign = static_cast<uint8_t>(a & 7u);
      for (size_t k = 0; k < 4; ++k) {
        if (a & (0x08u << k)) {
          x.border_style[k] = 1;
          x.border_color[k] = 8;
        }
      }
      if (a & 0x80u) {
        x.pattern = 4;
        x.fore = 8;
        x.back = 9;
      }
      x.style = xfs.size() < 15;
      x.parent = x.style ? 0xFFF : 0;
    } else if (ver == 3 || ver == 4) {
      if (d.size() < 12) Rcpp::stop("xls: file is truncated");
      uint32_t tp, ao;
      x.font = d[0];
      x.fmt = d[1];
      if (ver == 3) {
        tp = d[2];
        ao = cfb_u16(d, 4);
        x.used = static_cast<uint8_t>(d[3] & 0xFCu);
        x.parent = static_cast<uint16_t>(ao >> 4);
        x.halign = static_cast<uint8_t>(ao & 7u);
        x.wrap = static_cast<uint8_t>((ao >> 3) & 1u);
      } else {
        tp = cfb_u16(d, 2);
        ao = d[4];
        x.used = static_cast<uint8_t>(d[5] & 0xFCu);
        x.parent = static_cast<uint16_t>(tp >> 4);
        x.halign = static_cast<uint8_t>(ao & 7u);
        x.wrap = static_cast<uint8_t>((ao >> 3) & 1u);
        x.valign = static_cast<uint8_t>((ao >> 4) & 3u);
        static const uint8_t rot[4] = {0, 255, 90, 180};
        x.rotation = rot[(ao >> 6) & 3u];
      }
      x.locked = (tp & 1u) != 0;
      x.hidden = (tp & 2u) != 0;
      x.style = (tp & 4u) != 0;
      uint32_t bkg = cfb_u16(d, 6), brd = cfb_u32(d, 8);
      x.pattern = static_cast<uint8_t>(bkg & 0x3Fu);
      x.fore = static_cast<uint8_t>(color3((bkg >> 6) & 0x1Fu));
      x.back = static_cast<uint8_t>(color3((bkg >> 11) & 0x1Fu));
      x.border_style[2] = static_cast<uint8_t>(brd & 7u);
      x.border_color[2] = static_cast<uint8_t>(color3((brd >> 3) & 0x1Fu));
      x.border_style[0] = static_cast<uint8_t>((brd >> 8) & 7u);
      x.border_color[0] = static_cast<uint8_t>(color3((brd >> 11) & 0x1Fu));
      x.border_style[3] = static_cast<uint8_t>((brd >> 16) & 7u);
      x.border_color[3] = static_cast<uint8_t>(color3((brd >> 19) & 0x1Fu));
      x.border_style[1] = static_cast<uint8_t>((brd >> 24) & 7u);
      x.border_color[1] = static_cast<uint8_t>(color3((brd >> 27) & 0x1Fu));
    } else {
      if (d.size() < 16) Rcpp::stop("xls: file is truncated");
      uint32_t tp = cfb_u16(d, 4), al = d[6], ou = d[7], b1 = cfb_u32(d, 8), b2 = cfb_u32(d, 12);
      x.font = static_cast<uint16_t>(cfb_u16(d, 0));
      x.fmt = static_cast<uint16_t>(cfb_u16(d, 2));
      x.locked = (tp & 1u) != 0;
      x.hidden = (tp & 2u) != 0;
      x.style = (tp & 4u) != 0;
      x.parent = static_cast<uint16_t>(tp >> 4);
      x.halign = static_cast<uint8_t>(al & 7u);
      x.wrap = static_cast<uint8_t>((al >> 3) & 1u);
      x.valign = static_cast<uint8_t>((al >> 4) & 7u);
      static const uint8_t rot[4] = {0, 255, 90, 180};
      x.rotation = rot[ou & 3u];
      x.used = static_cast<uint8_t>(ou & 0xFCu);
      x.fore = static_cast<uint8_t>(b1 & 0x7Fu);
      x.back = static_cast<uint8_t>((b1 >> 7) & 0x7Fu);
      x.pattern = static_cast<uint8_t>((b1 >> 16) & 0x3Fu);
      x.border_style[3] = static_cast<uint8_t>((b1 >> 22) & 7u);
      x.border_color[3] = static_cast<uint8_t>((b1 >> 25) & 0x7Fu);
      x.border_style[2] = static_cast<uint8_t>(b2 & 7u);
      x.border_style[0] = static_cast<uint8_t>((b2 >> 3) & 7u);
      x.border_style[1] = static_cast<uint8_t>((b2 >> 6) & 7u);
      x.border_color[2] = static_cast<uint8_t>((b2 >> 9) & 0x7Fu);
      x.border_color[0] = static_cast<uint8_t>((b2 >> 16) & 0x7Fu);
      x.border_color[1] = static_cast<uint8_t>((b2 >> 23) & 0x7Fu);
    }
    xfs.push_back(x);
    has_xf = true;
  }

  // BIFF2 has its own record numbers and layouts. They are changed to the ones of
  // BIFF3, with the format of a cell as an XF number. Returns 0 to drop the record.
  uint16_t upgrade_biff2(uint16_t id, std::vector<uint8_t>& d) {
    XlsBytes t;
    switch (id) {
      case 0x0000:
        t.raw(d, 0, 8);
        d = t.b;
        return 0x0200;
      case 0x0001: {
        t.raw(d, 0, 4);
        t.u16(xf_from_attr(d, 4));
        d = t.b;
        return 0x0201;
      }
      case 0x0002: {
        t.raw(d, 0, 4);
        t.u16(xf_from_attr(d, 4));
        double v = static_cast<double>(cfb_u16(d, 7));
        uint64_t bits = 0;
        std::memcpy(&bits, &v, sizeof(bits));
        t.u32(static_cast<uint32_t>(bits & 0xFFFFFFFFu));
        t.u32(static_cast<uint32_t>(bits >> 32));
        d = t.b;
        return 0x0203;
      }
      case 0x0003: {
        t.raw(d, 0, 4);
        t.u16(xf_from_attr(d, 4));
        t.raw(d, 7, 8);
        d = t.b;
        return 0x0203;
      }
      case 0x0004: {
        if (d.size() < 8) Rcpp::stop("xls: file is truncated");
        size_t n = d[7];
        t.raw(d, 0, 4);
        t.u16(xf_from_attr(d, 4));
        t.u16(static_cast<uint32_t>(n));
        t.raw(d, 8, n);
        d = t.b;
        return 0x0204;
      }
      case 0x0005: {
        t.raw(d, 0, 4);
        t.u16(xf_from_attr(d, 4));
        t.raw(d, 7, 2);
        d = t.b;
        return 0x0205;
      }
      case 0x0006: {
        if (d.size() < 17) Rcpp::stop("xls: file is truncated");
        t.raw(d, 0, 4);
        t.u16(xf_from_attr(d, 4));
        t.raw(d, 7, 8);
        t.u16(d[15]);
        t.u32(0);
        t.u16(d[16]);
        t.raw(d, 17, d.size() - 17);
        d = t.b;
        return 0x0206;
      }
      case 0x0007: {
        if (d.empty()) Rcpp::stop("xls: file is truncated");
        size_t n = d[0];
        t.u16(static_cast<uint32_t>(n));
        t.raw(d, 1, n);
        d = t.b;
        return 0x0207;
      }
      case 0x0008: {
        t.raw(d, 0, 6);
        t.raw(d, 6, 2);
        t.u32(0);
        t.u32(0);
        d = t.b;
        return 0x0208;
      }
      case 0x0021:
        return 0x0221;
      case 0x0024: {
        if (d.size() < 4) Rcpp::stop("xls: file is truncated");
        t.u16(d[0]);
        t.u16(d[1]);
        t.raw(d, 2, 2);
        t.u16(15);
        t.u16(0);
        t.u16(0);
        d = t.b;
        return 0x007D;
      }
      case 0x003E: {
        if (d.size() < 13) Rcpp::stop("xls: file is truncated");
        uint32_t flags = (d[0] ? 1u : 0u) | (d[1] ? 2u : 0u) | (d[2] ? 4u : 0u) | (d[3] ? 8u : 0u) | (d[4] ? 0x10u : 0u) | (d[9] ? 0x20u : 0u);
        t.u16(flags);
        t.raw(d, 5, 4);
        t.raw(d, 10, 3);
        t.u8(0);
        d = t.b;
        return 0x023E;
      }
      case 0x0025:
      case 0x0036:
      case 0x0037:
        return 0;
      default:
        return id;
    }
  }

  void record(uint16_t id, std::vector<uint8_t> d, bool glob) {
    uint16_t cid = id;
    if (ver == 2) cid = upgrade_biff2(id, d);
    if (cid == 0) return;

    if (cid != 0x007F && cid != 0x003C) emit_pic();
    if (cid != 0x001C) emit_note();
    if (cid != 0x0221 && cid != 0x04BC && cid != 0x0207) flush();
    std::vector<uint8_t>& out = *cur;
    XlsBytes o;

    switch (cid) {
      case 0x0042: {
        uint32_t cp = cfb_u16(d, 0);
        codepage = cp == 32768 ? 10000 : (cp == 32769 ? 1252 : cp);
        break;
      }

      case 0x0022:
      case 0x000C:
      case 0x000D:
      case 0x000E:
      case 0x000F:
      case 0x0010:
      case 0x0011:
      case 0x005F:
      case 0x0019:
        if (ver < 5 || glob) {
          o.raw(d, 0, d.size());
          rec(g_misc, cid, o);
        }
        break;

      case 0x0012:
      case 0x0013:
      case 0x0063:
      case 0x00DD:
        o.raw(d, 0, d.size());
        rec(ver == 5 && glob ? g_misc : out, cid, o);
        break;

      case 0x003D: {
        if (!(ver < 5 || glob)) break;
        if (d.size() >= 18) {
          o.raw(d, 0, 18);
        } else {
          o.raw(d, 0, 8);
          o.u16(d.size() > 8 && d[8] ? 1 : 0x38);
          o.u16(0);
          o.u16(0);
          o.u16(1);
          o.u16(600);
        }
        rec(g_misc, 0x003D, o);
        break;
      }

      case 0x0031:
      case 0x0231: {
        size_t name_at = ver == 2 ? 4 : (ver == 5 ? 14 : 6);
        if (d.size() < name_at + 1) Rcpp::stop("xls: file is truncated");
        uint32_t grbit = cfb_u16(d, 2);
        o.raw(d, 0, 4);
        if (ver == 2) {
          o.u16(0x7FFF);
        } else if (ver == 5) {
          o.raw(d, 4, 10);
        } else {
          uint32_t icv = cfb_u16(d, 4);
          o.u16(icv < 16 ? icv + 8 : icv);
        }
        if (ver < 5) {
          o.u16(grbit & 1u ? 700 : 400);
          o.u16(0);
          o.u8(grbit & 4u ? 1 : 0);
          o.u8(0);
          o.u8(0);
          o.u8(0);
        }
        size_t n = d[name_at];
        ustr(o, str8(d, name_at + 1, n), false);
        rec(g_fonts, 0x0031, o);
        break;
      }

      case 0x001E:
      case 0x041E: {
        size_t at = 0;
        uint16_t key = static_cast<uint16_t>(nformats);
        if (cid == 0x041E) {
          if (d.size() < 3) Rcpp::stop("xls: file is truncated");
          if (ver == 5) key = static_cast<uint16_t>(cfb_u16(d, 0));
          at = 2;
        }
        if (d.size() < at + 1) Rcpp::stop("xls: file is truncated");
        size_t n = d[at];
        formats.push_back(std::make_pair(key, str8(d, at + 1, n)));
        ++nformats;
        break;
      }

      case 0x0043:
      case 0x0243:
      case 0x0443:
      case 0x00E0:
        decode_xf(d);
        break;

      case 0x0044:
        ixfe = static_cast<uint16_t>(cfb_u16(d, 0));
        break;

      case 0x0293: {
        if (ver < 3 || d.size() < 3) break;
        uint32_t x = cfb_u16(d, 0);
        o.u16(x);
        if (x & 0x8000u) {
          o.u8(d[2]);
          o.u8(d.size() > 3 ? d[3] : 0);
        } else {
          size_t n = d[2];
          ustr(o, str8(d, 3, n), true);
        }
        rec(g_styles, 0x0293, o);
        break;
      }

      case 0x0092:
        if (ver >= 3) {
          o.raw(d, 0, d.size());
          rec(g_palette, 0x0092, o);
          size_t ccv = cfb_u16(d, 0);
          for (size_t i = 0; i < ccv && i < palette.size() && 6 + 4 * i <= d.size(); ++i)
            palette[i] = (static_cast<uint32_t>(d[2 + 4 * i]) << 16) | (static_cast<uint32_t>(d[3 + 4 * i]) << 8) | d[4 + 4 * i];
        }
        break;

      case 0x005D:
        if (ver >= 3 && ver <= 5 && (!glob || ver < 5) && d.size() >= (ver == 5 ? 34u : 30u)) {
          if (ver == 5 && cfb_u16(d, 4) == 8) {
            pic.active = true;
            pic.obj = d;
            pic.data.clear();
          } else {
            shape5(d, 0);
          }
        }
        break;

      case 0x007F:
        if (pic.active) pic.data.assign(d.begin(), d.end());
        break;

      case 0x003C:
        if (pic.active) pic.data.append(d.begin(), d.end());
        break;

      case 0x0018:
      case 0x0218: {
        if (!glob) break;
        uint32_t grbit = 0;
        size_t cch = 0, cce = 0, name_at = 0, formula_at = 0;
        uint8_t key = 0;
        uint16_t itab = 0;
        if (ver == 5) {
          if (d.size() < 14) Rcpp::stop("xls: file is truncated");
          grbit = cfb_u16(d, 0);
          key = d[2];
          cch = d[3];
          cce = cfb_u16(d, 4);
          itab = static_cast<uint16_t>(cfb_u16(d, 8));
          name_at = 14;
          formula_at = 14 + cch + d[10] + d[11] + d[12] + d[13];
        } else if (ver == 2) {
          break;
        } else {
          if (d.size() < 6) Rcpp::stop("xls: file is truncated");
          grbit = cfb_u16(d, 0);
          key = d[2];
          cch = d[3];
          cce = cfb_u16(d, 4);
          name_at = 6;
          formula_at = 6 + cch;
        }
        if (ver < 5 && (grbit & 0x20u)) itab = 1;

        XlsBytes rg;
        if (!rgce(d, formula_at, cce, rg, true)) {
          rg.b.clear();
          rg.u8(0x1C);
          rg.u8(0x17);
        }
        bool builtin = (grbit & 0x20u) != 0;
        o.u16(grbit);
        o.u8(key);
        o.u8(builtin ? 1u : static_cast<uint32_t>(cch));
        o.u16(static_cast<uint32_t>(rg.b.size()));
        o.u16(0);
        o.u16(itab);
        o.u32(0);
        if (builtin) {
          if (name_at >= d.size()) Rcpp::stop("xls: file is truncated");
          o.u8(0);
          o.u8(d[name_at]);
        } else {
          std::u16string name = str8(d, name_at, cch);
          o.u8(1);
          for (size_t i = 0; i < name.size(); ++i) o.u16(static_cast<uint32_t>(name[i]));
        }
        o.raw(rg);
        rec(g_names, 0x0018, o);
        break;
      }

      case 0x0200: {
        if (d.size() < 8) Rcpp::stop("xls: file is truncated");
        o.u32(cfb_u16(d, 0));
        o.u32(cfb_u16(d, 2));
        o.u16(cfb_u16(d, 4));
        o.u16(cfb_u16(d, 6));
        o.u16(0);
        rec(out, 0x0200, o);
        break;
      }

      case 0x0201:
      case 0x0203:
      case 0x0205:
      case 0x027E:
      case 0x00BD:
      case 0x00BE:
      case 0x0055:
      case 0x0225:
      case 0x0099:
      case 0x001D:
      case 0x0026:
      case 0x0027:
      case 0x0028:
      case 0x0029:
      case 0x0083:
      case 0x0084:
      case 0x002A:
      case 0x002B:
      case 0x0080:
      case 0x0081:
      case 0x00A0:
        o.raw(d, 0, d.size());
        rec(out, cid, o);
        break;

      case 0x0204: {
        if (d.size() < 8) Rcpp::stop("xls: file is truncated");
        size_t n = cfb_u16(d, 6);
        o.raw(d, 0, 6);
        ustr(o, str8(d, 8, n), true);
        rec(out, 0x0204, o);
        break;
      }

      case 0x0208: {
        o.raw(d, 0, d.size());
        while (o.b.size() < 16) o.u8(0);
        rec(out, 0x0208, o);
        break;
      }

      case 0x007D: {
        o.raw(d, 0, d.size());
        while (o.b.size() < 12) o.u8(0);
        rec(out, 0x007D, o);
        break;
      }

      case 0x023E: {
        if (d.size() < 10) Rcpp::stop("xls: file is truncated");
        o.raw(d, 0, 10);
        o.u32(0);
        o.u32(0);
        o.u16(0);
        rec(out, 0x023E, o);
        break;
      }

      case 0x0041: {
        o.raw(d, 0, d.size());
        while (o.b.size() < 10) o.u8(0);
        rec(out, 0x0041, o);
        break;
      }

      case 0x0014:
      case 0x0015: {
        if (!d.empty()) {
          size_t n = d[0];
          ustr(o, str8(d, 1, n), true);
        }
        rec(out, cid, o);
        break;
      }

      case 0x00A1:
        if (d.size() == 34) {
          o.raw(d, 0, d.size());
          rec(out, 0x00A1, o);
        }
        break;

      case 0x001A:
      case 0x001B: {
        size_t n = cfb_u16(d, 0);
        o.u16(static_cast<uint32_t>(n));
        for (size_t i = 0; i < n; ++i) {
          o.u16(cfb_u16(d, 2 + 2 * i));
          o.u16(0);
          o.u16(cid == 0x001B ? 0xFF : 0xFFFF);
        }
        rec(out, cid, o);
        break;
      }

      case 0x0206:
      case 0x0406:
      case 0x0006: {
        // BIFF3 and BIFF4 have no chain field, the length of the formula follows the flags
        size_t cce_at = (ver == 3 || ver == 4) ? 16 : 20;
        if (d.size() < cce_at + 2) Rcpp::stop("xls: file is truncated");
        pend = Pending();
        pend.active = true;
        pend.row = static_cast<uint16_t>(cfb_u16(d, 0));
        pend.col = static_cast<uint16_t>(cfb_u16(d, 2));
        pend.xf = static_cast<uint16_t>(cfb_u16(d, 4));
        for (size_t i = 0; i < 8; ++i) pend.result[i] = d[6 + i];
        XlsBytes rg;
        pend.ok = rgce(d, cce_at + 2, cfb_u16(d, cce_at), rg);
        if (pend.ok && rg.b.size() == 5 && rg.b[0] == 0x01) {
          uint32_t key = (static_cast<uint32_t>(rg.b[1] | (rg.b[2] << 8)) << 16) | static_cast<uint32_t>(rg.b[3] | (rg.b[4] << 8));
          if (dead_exp.count(key)) pend.ok = false;
        }
        pend.formula.raw(d, 0, 16);
        pend.formula.u32(0);
        pend.formula.u16(static_cast<uint32_t>(rg.b.size()));
        pend.formula.raw(rg);
        break;
      }

      case 0x0207: {
        if (!pend.active || d.size() < 2) break;
        size_t n = cfb_u16(d, 0);
        pend.text = str8(d, 2, n);
        pend.has_string = true;
        break;
      }

      case 0x0221:
      case 0x04BC: {
        size_t head = ver == 2 ? 8 : (cid == 0x0221 ? 14 : 10);
        if (!pend.active) break;
        if (d.size() < head) Rcpp::stop("xls: file is truncated");
        uint32_t key = (cfb_u16(d, 0) << 16) | d[4];
        size_t cce = ver == 2 ? d[7] : cfb_u16(d, head - 2);
        XlsBytes rg;
        if (!rgce(d, head, cce, rg)) {
          dead_exp.insert(key);
          pend.extra_failed = true;
          break;
        }
        if (ver == 2) {
          o.raw(d, 0, 6);
          o.u16(d[6]);
          o.u32(0);
        } else {
          o.raw(d, 0, head - 2);
        }
        o.u16(static_cast<uint32_t>(rg.b.size()));
        o.raw(rg);
        pend.extra.push_back(std::make_pair(cid, o));
        break;
      }

      case 0x001C: {
        if (d.size() < 6) Rcpp::stop("xls: file is truncated");
        size_t cch = cfb_u16(d, 4);
        if (cfb_u16(d, 0) == 0xFFFF) {
          if (note.active) note.bytes.append(reinterpret_cast<const char*>(d.data()) + 6, std::min(cch, d.size() - 6));
        } else {
          emit_note();
          note = Note();
          note.active = true;
          note.row = static_cast<uint16_t>(cfb_u16(d, 0));
          note.col = static_cast<uint16_t>(cfb_u16(d, 2));
          note.total = cch;
          note.bytes.assign(reinterpret_cast<const char*>(d.data()) + 6, std::min(cch, d.size() - 6));
        }
        break;
      }

      default:
        break;
    }
  }

  std::vector<uint8_t> run(const std::vector<uint8_t>& file) {
    std::vector<uint8_t> stream;
    bool cfb = file.size() >= 8 && file[0] == 0xD0 && file[1] == 0xCF && file[2] == 0x11 && file[3] == 0xE0;
    if (cfb) {
      stream = cfb_stream(file, "Book");
      if (stream.empty()) stream = cfb_stream(file, "Workbook");
    } else {
      stream = file;
    }
    if (stream.size() < 8) Rcpp::stop("xls: file is truncated");

    uint32_t first = cfb_u16(stream, 0);
    if (first == 0x0009) ver = 2;
    else if (first == 0x0209) ver = 3;
    else if (first == 0x0409) ver = 4;
    else ver = 5;
    if (ver == 4 && cfb_u16(stream, 6) == 0x0100) Rcpp::stop("xls: Excel 4 workspace files are not supported");

    struct Sheet {
      uint32_t offset = 0;
      uint8_t state = 0, type = 0;
      std::u16string name;
      std::vector<uint8_t> body;
      bool empty = false;
    };
    std::vector<Sheet> sheets;

    if (ver < 5) {
      Sheet s;
      s.name = u"Sheet1";
      s.empty = cfb_u16(stream, 4) == 0x0020;
      sheets.push_back(s);
      cur = &sheets[0].body;
      size_t pos = 0;
      int depth = 0;
      while (pos + 4 <= stream.size()) {
        uint16_t id = static_cast<uint16_t>(cfb_u16(stream, pos));
        size_t len = cfb_u16(stream, pos + 2);
        if (pos + 4 + len > stream.size()) break;
        std::vector<uint8_t> d(stream.begin() + static_cast<std::ptrdiff_t>(pos + 4), stream.begin() + static_cast<std::ptrdiff_t>(pos + 4 + len));
        pos += 4 + len;
        if (id == 0x0009 || id == 0x0209 || id == 0x0409 || id == 0x0809) {
          ++depth;
          continue;
        }
        if (id == 0x000A) {
          if (--depth <= 0) break;
          continue;
        }
        if (depth != 1) continue;
        if (id == 0x002F) Rcpp::stop("xls: encrypted files are not supported");
        if (!sheets[0].empty) record(id, d, true);
      }
      flush();
      emit_note();
    } else {
      size_t pos = 0;
      int depth = 0;
      while (pos + 4 <= stream.size()) {
        uint16_t id = static_cast<uint16_t>(cfb_u16(stream, pos));
        size_t len = cfb_u16(stream, pos + 2);
        if (pos + 4 + len > stream.size()) break;
        std::vector<uint8_t> d(stream.begin() + static_cast<std::ptrdiff_t>(pos + 4), stream.begin() + static_cast<std::ptrdiff_t>(pos + 4 + len));
        pos += 4 + len;
        if (id == 0x0809) {
          ++depth;
          continue;
        }
        if (id == 0x000A) {
          if (--depth <= 0) break;
          continue;
        }
        if (id == 0x002F) Rcpp::stop("xls: encrypted files are not supported");
        if (id == 0x0085) {
          if (d.size() < 8) Rcpp::stop("xls: file is truncated");
          Sheet s;
          s.offset = static_cast<uint32_t>(cfb_u32(d, 0));
          s.state = static_cast<uint8_t>(d[4] & 3u);
          s.type = d[5];
          s.name = str8(d, 7, d[6]);
          sheets.push_back(s);
          continue;
        }
        cur = &g_misc;
        record(id, d, true);
      }

      for (size_t i = 0; i < sheets.size(); ++i) {
        cur = &sheets[i].body;
        if (sheets[i].type > 2) {
          sheets[i].empty = true;
          continue;
        }
        bool chart_sheet = sheets[i].type == 2;
        bool in_chart = false;
        size_t spos = sheets[i].offset;
        int sdepth = 0;
        while (spos + 4 <= stream.size()) {
          uint16_t id = static_cast<uint16_t>(cfb_u16(stream, spos));
          size_t len = cfb_u16(stream, spos + 2);
          if (spos + 4 + len > stream.size()) break;
          std::vector<uint8_t> d(stream.begin() + static_cast<std::ptrdiff_t>(spos + 4),
                                 stream.begin() + static_cast<std::ptrdiff_t>(spos + 4 + len));
          spos += 4 + len;
          if (id == 0x0809) {
            ++sdepth;
            if (sdepth == 2 && !chart_sheet && cfb_u16(d, 2) == 0x0020) {
              flush();
              emit_note();
              emit_pic();
              XlsBytes bof;
              bof.u16(0x0600);
              bof.u16(0x0020);
              bof.u16(0x0DBB);
              bof.u16(0x07CC);
              bof.u32(0);
              bof.u32(6);
              rec(*cur, 0x0809, bof);
              in_chart = true;
            }
            continue;
          }
          if (id == 0x000A) {
            if (in_chart && sdepth == 2) {
              rec(*cur, 0x000A, XlsBytes());
              in_chart = false;
            }
            if (--sdepth <= 0) break;
            continue;
          }
          if (chart_sheet || in_chart) {
            if (sdepth == (chart_sheet ? 1 : 2)) chart5(id, d, *cur);
            continue;
          }
          if (sdepth != 1) continue;
          record(id, d, false);
        }
        flush();
        emit_note();
        emit_pic();
        next_object = 1;
      }
    }

    while (xfs.size() < 16) {
      XlsLegacyXf x;
      x.style = xfs.size() < 15;
      x.parent = x.style ? 0xFFF : 0;
      xfs.push_back(x);
    }

    std::map<uint16_t, uint16_t> fmt_map;
    XlsBytes fmt_records;
    for (size_t i = 0; i < formats.size(); ++i) {
      uint16_t key = formats[i].first;
      uint16_t out_key = key;
      if (ver < 5) out_key = formats[i].second == u"General" ? 0 : static_cast<uint16_t>(164 + key);
      fmt_map[key] = out_key;
      if (out_key == 0) continue;
      XlsBytes f;
      f.u16(out_key);
      ustr(f, formats[i].second, true);
      rec(fmt_records.b, 0x041E, f);
    }

    std::vector<uint8_t> out;
    XlsBytes b;

    b.u16(0x0600);
    b.u16(0x0005);
    b.u16(0x0DBB);
    b.u16(0x07CC);
    b.u32(0);
    b.u32(6);
    rec(out, 0x0809, b);

    b = XlsBytes();
    b.u16(1200);
    rec(out, 0x0042, b);

    out.insert(out.end(), g_misc.begin(), g_misc.end());
    out.insert(out.end(), g_fonts.begin(), g_fonts.end());
    out.insert(out.end(), fmt_records.b.begin(), fmt_records.b.end());

    for (size_t i = 0; i < xfs.size(); ++i) {
      const XlsLegacyXf& x = xfs[i];
      std::map<uint16_t, uint16_t>::const_iterator it = fmt_map.find(x.fmt);
      b = XlsBytes();
      b.u16(x.font);
      b.u16(it == fmt_map.end() ? x.fmt : it->second);
      b.u16((x.locked ? 1u : 0u) | (x.hidden ? 2u : 0u) | (x.style ? 4u : 0u) | (static_cast<uint32_t>(x.parent) << 4));
      b.u8(static_cast<uint32_t>(x.halign) | (static_cast<uint32_t>(x.wrap) << 3) | (static_cast<uint32_t>(x.valign) << 4));
      b.u8(x.rotation);
      b.u8(0);
      b.u8(x.used);
      b.u32(static_cast<uint32_t>(x.border_style[0]) | (static_cast<uint32_t>(x.border_style[1]) << 4) |
            (static_cast<uint32_t>(x.border_style[2]) << 8) | (static_cast<uint32_t>(x.border_style[3]) << 12) |
            (static_cast<uint32_t>(x.border_color[0]) << 16) | (static_cast<uint32_t>(x.border_color[1]) << 23));
      b.u32(static_cast<uint32_t>(x.border_color[2]) | (static_cast<uint32_t>(x.border_color[3]) << 7) | (static_cast<uint32_t>(x.pattern) << 26));
      b.u16(static_cast<uint32_t>(x.fore) | (static_cast<uint32_t>(x.back) << 7));
      rec(out, 0x00E0, b);
    }

    out.insert(out.end(), g_styles.begin(), g_styles.end());
    out.insert(out.end(), g_palette.begin(), g_palette.end());

    if (!blips.empty()) {
      XlsBytes store;
      for (size_t i = 0; i < blips.size(); ++i) {
        XlsBytes uid;
        uid.u32(static_cast<uint32_t>(i + 1));
        for (size_t k = 0; k < 12; ++k) uid.u8(0);
        XlsBytes blip;
        blip.raw(uid);
        blip.u8(0xFF);
        blip.b.insert(blip.b.end(), blips[i].begin(), blips[i].end());
        XlsBytes blip_record;
        esc(blip_record, 0x7A8u << 4, 0xF01F, blip);
        XlsBytes fbse;
        fbse.u8(7);
        fbse.u8(7);
        fbse.raw(uid);
        fbse.u16(0xFF);
        fbse.u32(static_cast<uint32_t>(blip_record.b.size()));
        fbse.u32(1);
        fbse.u32(0);
        fbse.u32(0);
        fbse.raw(blip_record);
        esc(store, 0x2 | (7u << 4), 0xF007, fbse);
      }
      XlsBytes bstore;
      esc(bstore, 0xF | (static_cast<uint32_t>(blips.size()) << 4), 0xF001, store);
      XlsBytes group;
      esc(group, 0xF, 0xF000, bstore);
      for (size_t at = 0; at < group.b.size(); at += 8224) {
        XlsBytes part;
        part.b.assign(group.b.begin() + static_cast<std::ptrdiff_t>(at),
                      group.b.begin() + static_cast<std::ptrdiff_t>(std::min(at + 8224, group.b.size())));
        rec(out, at == 0 ? 0x00EB : 0x003C, part);
      }
    }

    std::vector<size_t> offset_at;
    for (size_t i = 0; i < sheets.size(); ++i) {
      b = XlsBytes();
      offset_at.push_back(out.size() + 4);
      b.u32(0);
      b.u8(sheets[i].state);
      b.u8(sheets[i].type == 2 ? 2 : 0);
      ustr(b, sheets[i].name, false);
      rec(out, 0x0085, b);
    }

    b = XlsBytes();
    b.u16(static_cast<uint32_t>(sheets.size()));
    b.u16(0x0401);
    rec(out, 0x01AE, b);

    if (!xtis.empty()) {
      b = XlsBytes();
      b.u16(static_cast<uint32_t>(xtis.size()));
      for (size_t i = 0; i < xtis.size(); ++i) {
        b.u16(xtis[i][0]);
        b.u16(xtis[i][1]);
        b.u16(xtis[i][2]);
      }
      rec(out, 0x0017, b);
    }

    out.insert(out.end(), g_names.begin(), g_names.end());
    rec(out, 0x000A, XlsBytes());

    for (size_t i = 0; i < sheets.size(); ++i) {
      uint32_t at = static_cast<uint32_t>(out.size());
      for (size_t k = 0; k < 4; ++k) out[offset_at[i] + k] = static_cast<uint8_t>((at >> (8 * k)) & 0xFFu);
      b = XlsBytes();
      b.u16(0x0600);
      b.u16(sheets[i].type == 2 ? 0x0020 : 0x0010);
      b.u16(0x0DBB);
      b.u16(0x07CC);
      b.u32(0);
      b.u32(6);
      rec(out, 0x0809, b);
      out.insert(out.end(), sheets[i].body.begin(), sheets[i].body.end());
      rec(out, 0x000A, XlsBytes());
    }

    return out;
  }
};

#endif
