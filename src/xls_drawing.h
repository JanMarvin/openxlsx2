#ifndef XLS_DRAWING_H
#define XLS_DRAWING_H

#include <algorithm>
#include <cstdint>
#include <cstring>
#include <functional>
#include <map>
#include <sstream>
#include <string>
#include <utility>
#include <vector>

#include "xls_funs.h"

// ---------------------------------------------------------------------------
// pictures and shapes (Escher)
// ---------------------------------------------------------------------------

// Pictures and shapes of an xls file are stored as Escher records (the drawing
// format of Office 97 to 2003): a header of 8 bytes (version and instance, type,
// length), followed by the content. A version of 0xF is a container of records.

struct XlsEscherRecord {
  uint8_t ver = 0;
  uint16_t inst = 0, type = 0;
  size_t body = 0, len = 0;
};

inline uint32_t xls_esc_u16(const std::string& d, size_t at) {
  if (at + 2 > d.size()) return 0;
  return static_cast<uint32_t>(static_cast<uint8_t>(d[at]) | (static_cast<uint8_t>(d[at + 1]) << 8));
}

inline uint32_t xls_esc_u32(const std::string& d, size_t at) {
  return xls_esc_u16(d, at) | (xls_esc_u16(d, at + 2) << 16);
}

inline void xls_escher_walk(const std::string& d, size_t pos, size_t end, const std::function<void(const XlsEscherRecord&)>& visit) {
  if (end > d.size()) end = d.size();
  while (pos + 8 <= end) {
    XlsEscherRecord r;
    uint32_t vi = xls_esc_u16(d, pos);
    r.ver = static_cast<uint8_t>(vi & 0xFu);
    r.inst = static_cast<uint16_t>(vi >> 4);
    r.type = static_cast<uint16_t>(xls_esc_u16(d, pos + 2));
    r.body = pos + 8;
    r.len = xls_esc_u32(d, pos + 4);
    if (r.body + r.len > end) r.len = end - r.body;
    visit(r);
    if (r.ver == 0xF) xls_escher_walk(d, r.body, r.body + r.len, visit);
    pos = r.body + r.len;
  }
}

struct XlsBlip {
  std::string ext;   // empty if the picture is not in the file
  std::string data;
  bool deflated = false;  // metafiles are compressed
  uint32_t size = 0;      // size of a metafile when it is unpacked
};

// The pictures of the workbook (BLIP store), in the order the shapes count them
inline std::vector<XlsBlip> xls_blips(const std::string& group) {
  std::vector<XlsBlip> blips;
  xls_escher_walk(group, 0, group.size(), [&](const XlsEscherRecord& r) {
    // FBSE: the entry of the store. Type, type on Mac, uid (16), tag (2), size (4),
    // count (4), position (4), 3 unused, length of the name; then the name and the picture
    if (r.type != 0xF007) return;
    XlsBlip b;
    if (r.len < 36) {
      blips.push_back(b);
      return;
    }
    size_t cb_name = static_cast<uint8_t>(group[r.body + 33]);
    size_t at = r.body + 36 + cb_name;
    size_t end = r.body + r.len;
    if (at + 8 > end) {
      blips.push_back(b);
      return;
    }
    uint32_t vi = xls_esc_u16(group, at);
    uint16_t inst = static_cast<uint16_t>(vi >> 4);
    uint16_t type = static_cast<uint16_t>(xls_esc_u16(group, at + 2));
    size_t len = xls_esc_u32(group, at + 4);
    size_t pos = at + 8;
    if (pos + len > end) len = end - pos;
    size_t stop = pos + len;

    // one uid, or two if the lowest bit of the instance is set
    pos += (inst & 1) ? 32 : 16;
    switch (type) {
      case 0xF01E: b.ext = "png"; pos += 1; break;
      case 0xF01D: b.ext = "jpeg"; pos += 1; break;
      case 0xF01F: b.ext = "dib"; pos += 1; break;
      case 0xF029: b.ext = "tiff"; pos += 1; break;
      case 0xF01A:
      case 0xF01B: {
        // metafiles: size, bounds (16), size (8), size of the data, compression, filter
        b.ext = type == 0xF01A ? "emf" : "wmf";
        b.size = pos + 4 <= stop ? xls_esc_u32(group, pos) : 0;
        b.deflated = pos + 33 < stop && static_cast<uint8_t>(group[pos + 32]) == 0;
        pos += 34;
        break;
      }
      default: break;
    }
    if (!b.ext.empty() && pos <= stop) b.data = group.substr(pos, stop - pos);
    if (b.data.empty()) b.ext.clear();
    blips.push_back(b);
  });
  return blips;
}

// A bitmap in a file without the header: the file header with the position of the pixels is added
inline std::string xls_dib_to_bmp(const std::string& dib) {
  if (dib.size() < 40) return std::string();
  uint32_t header = xls_esc_u32(dib, 0);
  uint32_t bits = xls_esc_u16(dib, 14);
  uint32_t compression = xls_esc_u32(dib, 16);
  uint32_t used = xls_esc_u32(dib, 32);
  uint32_t colors = bits <= 8 ? (used ? used : (1u << bits)) : used;
  uint32_t offset = 14 + header + colors * 4 + ((compression == 3 && header == 40) ? 12 : 0);
  uint32_t size = static_cast<uint32_t>(dib.size()) + 14;
  std::string out = "BM";
  for (int i = 0; i < 4; ++i) out.push_back(static_cast<char>((size >> (8 * i)) & 0xFF));
  out.append(4, '\0');
  for (int i = 0; i < 4; ++i) out.push_back(static_cast<char>((offset >> (8 * i)) & 0xFF));
  return out + dib;
}

// A shape of a sheet, together with the position
struct XlsShape {
  bool group = false;
  uint16_t shape_type = 0;
  uint32_t spid = 0;
  uint32_t blip = 0;  // number of the picture, counting from 1
  std::string fill, line;  // colors as RRGGBB, empty if not set
  uint32_t line_width = 0;  // EMU, 0 if not set
  bool has_anchor = false;
  // anchor: column, offset in 1/1024 of the cell, row, offset in 1/256 of the cell
  uint16_t col1 = 0, dx1 = 0, row1 = 0, dy1 = 0, col2 = 0, dx2 = 0, row2 = 0, dy2 = 0;
};

inline std::vector<XlsShape> xls_shapes(const std::string& dg) {
  std::vector<XlsShape> shapes;
  xls_escher_walk(dg, 0, dg.size(), [&](const XlsEscherRecord& r) {
    if (r.type == 0xF004) {
      shapes.push_back(XlsShape());
      return;
    }
    if (shapes.empty()) return;
    XlsShape& s = shapes.back();
    if (r.type == 0xF00A && r.len >= 8) {
      s.shape_type = r.inst;
      s.spid = xls_esc_u32(dg, r.body);
      s.group = (xls_esc_u32(dg, r.body + 4) & 0x01u) != 0;
    } else if (r.type == 0xF00B) {
      // properties: id (14 bits, then a flag for a picture and one for data after the table) and a value
      for (size_t i = 0; i < r.inst && (i + 1) * 6 <= r.len; ++i) {
        uint32_t id = xls_esc_u16(dg, r.body + i * 6);
        uint32_t val = xls_esc_u32(dg, r.body + i * 6 + 2);
        uint32_t pid = id & 0x3FFFu;
        if (pid == 0x0104 && !(id & 0x8000u)) s.blip = val;
        // colors: red, green, blue and flags. Scheme and system colors are indexes.
        auto color = [&](uint32_t v) {
          if (v & 0x18000000u) return std::string();
          static const char* hex = "0123456789ABCDEF";
          std::string c;
          for (int k = 0; k < 3; ++k) {
            uint32_t b = (v >> (8 * k)) & 0xFFu;
            c += hex[b >> 4];
            c += hex[b & 0xF];
          }
          return c;
        };
        if (pid == 0x0181) s.fill = color(val);
        if (pid == 0x01C0) s.line = color(val);
        if (pid == 0x01CB) s.line_width = val;
      }
    } else if (r.type == 0xF010 && r.len >= 18) {
      s.has_anchor = true;
      s.col1 = static_cast<uint16_t>(xls_esc_u16(dg, r.body + 2));
      s.dx1 = static_cast<uint16_t>(xls_esc_u16(dg, r.body + 4));
      s.row1 = static_cast<uint16_t>(xls_esc_u16(dg, r.body + 6));
      s.dy1 = static_cast<uint16_t>(xls_esc_u16(dg, r.body + 8));
      s.col2 = static_cast<uint16_t>(xls_esc_u16(dg, r.body + 10));
      s.dx2 = static_cast<uint16_t>(xls_esc_u16(dg, r.body + 12));
      s.row2 = static_cast<uint16_t>(xls_esc_u16(dg, r.body + 14));
      s.dy2 = static_cast<uint16_t>(xls_esc_u16(dg, r.body + 16));
    }
  });
  return shapes;
}

// preset geometry of DrawingML for the shape types of Office 97 (msosptXXX)
inline std::string xls_preset_geometry(uint16_t type) {
  switch (type) {
    case 1: return "rect";
    case 2: return "roundRect";
    case 3: return "ellipse";
    case 4: return "diamond";
    case 5: return "triangle";
    case 6: return "rtTriangle";
    case 7: return "parallelogram";
    case 8: return "trapezoid";
    case 9: return "hexagon";
    case 10: return "octagon";
    case 11: return "plus";
    case 12: return "star5";
    case 13: return "rightArrow";
    case 15: return "homePlate";
    case 16: return "cube";
    case 20: return "line";
    case 21: return "plaque";
    case 22: return "can";
    case 23: return "donut";
    case 56: return "pentagon";
    case 55: return "chevron";
    case 63: return "star8";
    case 67: return "star16";
    case 66: return "star24";
    case 68: return "star32";
    case 71: return "wedgeRectCallout";
    case 72: return "wedgeRoundRectCallout";
    case 74: return "rect";
    case 84: return "leftArrow";
    case 85: return "upArrow";
    case 86: return "downArrow";
    case 87: return "leftRightArrow";
    case 88: return "upDownArrow";
    case 89: return "bentArrow";
    case 91: return "curvedRightArrow";
    case 92: return "stripedRightArrow";
    case 93: return "notchedRightArrow";
    case 94: return "bevel";
    case 95: return "foldedCorner";
    case 96: return "smileyFace";
    case 97: return "noSmoking";
    case 98: return "flowChartProcess";
    case 119: return "flowChartTerminator";
    case 120: return "flowChartDecision";
    case 121: return "flowChartInputOutput";
    case 183: return "sun";
    case 184: return "moon";
    case 185: return "bracketPair";
    case 186: return "bracePair";
    case 187: return "star4";
    case 188: return "doubleWave";
    case 189: return "blockArc";
    case 190: return "heart";
    case 191: return "lightningBolt";
    case 192: return "star6";
    default: return "rect";
  }
}

// ---------------------------------------------------------------------------
// charts
// ---------------------------------------------------------------------------

// A chart of an xls file is a substream of its own (BOF of type 0x20). The
// records between BOF and EOF are read into this structure and written as
// DrawingML. Only what is needed to draw the chart again is imported: the type
// of the chart, series with their references and cached values, titles, the
// legend, axes, and the colors, lines and markers of series.

struct XlsChartRecord {
  uint16_t id = 0;
  std::vector<uint8_t> d;
};

namespace xls_chart_detail {

inline uint32_t u8(const std::vector<uint8_t>& d, size_t at) { return at < d.size() ? d[at] : 0u; }
inline uint32_t u16(const std::vector<uint8_t>& d, size_t at) { return u8(d, at) | (u8(d, at + 1) << 8); }
inline uint32_t u32(const std::vector<uint8_t>& d, size_t at) { return u16(d, at) | (u16(d, at + 2) << 16); }
inline int32_t i16(const std::vector<uint8_t>& d, size_t at) { return static_cast<int16_t>(u16(d, at)); }
inline double f64(const std::vector<uint8_t>& d, size_t at) {
  uint64_t v = static_cast<uint64_t>(u32(d, at)) | (static_cast<uint64_t>(u32(d, at + 4)) << 32);
  double x;
  std::memcpy(&x, &v, sizeof(x));
  return x;
}

inline std::string esc(const std::string& s) {
  std::string o;
  for (size_t i = 0; i < s.size(); ++i) {
    switch (s[i]) {
      case '&': o += "&amp;"; break;
      case '<': o += "&lt;"; break;
      case '>': o += "&gt;"; break;
      case '"': o += "&quot;"; break;
      default: o += s[i];
    }
  }
  return o;
}

inline std::string rgb(const std::vector<uint8_t>& d, size_t at) {
  static const char* hex = "0123456789ABCDEF";
  std::string s;
  for (size_t i = 0; i < 3; ++i) {
    uint8_t b = static_cast<uint8_t>(u8(d, at + i));
    s += hex[b >> 4];
    s += hex[b & 0xF];
  }
  return s;
}

// ShortXLUnicodeString: length (1 byte), flags (1 byte), characters
inline std::string short_string(const std::vector<uint8_t>& d, size_t at) {
  size_t cch = u8(d, at);
  bool wide = (u8(d, at + 1) & 1) != 0;
  std::string out;
  size_t p = at + 2;
  for (size_t i = 0; i < cch && p < d.size(); ++i) {
    uint32_t c;
    if (wide) {
      c = u16(d, p);
      p += 2;
    } else {
      c = u8(d, p);
      p += 1;
    }
    if (c < 0x80) {
      out += static_cast<char>(c);
    } else if (c < 0x800) {
      out += static_cast<char>(0xC0 | (c >> 6));
      out += static_cast<char>(0x80 | (c & 0x3F));
    } else if (c >= 0xD800 && c < 0xDC00 && wide && p + 1 < d.size()) {
      uint32_t lo = u16(d, p);
      p += 2;
      ++i;
      c = 0x10000 + ((c - 0xD800) << 10) + (lo - 0xDC00);
      out += static_cast<char>(0xF0 | (c >> 18));
      out += static_cast<char>(0x80 | ((c >> 12) & 0x3F));
      out += static_cast<char>(0x80 | ((c >> 6) & 0x3F));
      out += static_cast<char>(0x80 | (c & 0x3F));
    } else {
      out += static_cast<char>(0xE0 | (c >> 12));
      out += static_cast<char>(0x80 | ((c >> 6) & 0x3F));
      out += static_cast<char>(0x80 | (c & 0x3F));
    }
  }
  return out;
}

// Excel 2007 and later write the DrawingML properties of chart elements into
// the file as XML (ShapePropsStream, TextPropsStream, RichTextStream). The
// XML follows a header of fixed size, and starts with the element itself.
inline std::string frt_xml(const std::vector<uint8_t>& d, const std::string& from, const std::string& to) {
  std::string x;
  size_t start = std::string::npos;
  for (size_t i = 12; i + 1 < d.size(); ++i) {
    if (d[i] == '<') {
      start = i;
      break;
    }
  }
  if (start == std::string::npos) return "";
  x.assign(d.begin() + static_cast<std::ptrdiff_t>(start), d.end());
  while (!x.empty() && (x.back() == '\0' || x.back() == '\n' || x.back() == '\r')) x.pop_back();
  if (x.compare(0, 5, "<?xml") == 0) {
    size_t e = x.find("?>");
    if (e == std::string::npos) return "";
    x = x.substr(e + 2);
    while (!x.empty() && (x[0] == '\n' || x[0] == '\r')) x.erase(0, 1);
  }
  if (x.compare(0, from.size() + 1, "<" + from) != 0) return "";
  size_t close = x.rfind("</" + from + ">");
  if (close == std::string::npos) return "";
  x = "<" + to + x.substr(from.size() + 1, close - from.size() - 1) + "</" + to + ">";
  return x;
}

struct Line {
  bool set = false;
  bool automatic = true;
  bool auto_color = true;
  std::string color;
  uint16_t style = 0;
  int32_t weight = 0;
};

struct Area {
  bool set = false;
  bool automatic = true;
  uint16_t pattern = 1;
  std::string color;
};

struct Marker {
  bool set = false;
  bool automatic = true;
  uint16_t type = 0;
  std::string fore, back;
  uint32_t size = 100;
};

struct Fmt {
  std::string sppr_xml, marker_sppr_xml;  // properties as written by Excel 2007 and later
  Line line;
  Area area;
  Marker marker;
  bool has_label = false;
  uint16_t label = 0;
  bool smooth = false;
  uint16_t explosion = 0;
};

struct Cached {
  bool is_string = false;
  std::string text;
};

struct Series {
  uint16_t type_x = 1, type_y = 1;
  uint16_t rt[4] = {0, 0, 0, 0};
  std::string ref[4];       // title, values, categories, bubble sizes
  std::string title_text;
  Fmt fmt;
  std::map<uint16_t, Fmt> points;
  int group = 0;
  std::map<uint16_t, std::map<uint32_t, Cached> > cache;  // 1 values, 2 categories, 3 bubble sizes
  bool no_data[4] = {false, false, false, false};
};

enum GroupType { GroupOther, GroupBar, GroupLine, GroupPie, GroupArea, GroupScatter, GroupRadar, GroupSurface };

struct Group {
  GroupType type = GroupOther;
  int icrt = 0;
  int axis_parent = 0;
  bool vary = false;
  bool horizontal = false, stacked = false, percent = false;
  int overlap = 0, gap = 150;
  uint16_t angle = 0, donut = 0;
  bool bubbles = false;
  int bubble_scale = 100;
  bool bubble_width = false, show_negative_bubbles = false;
  bool filled = false;
  bool three_d = false;
  int gap_depth = 150;
  std::string shape = "box";
  std::vector<size_t> series;
};

struct Axis {
  bool present = false;
  bool has_range = false;
  double min = 0, max = 0, major = 0, minor = 0, cross = 0;
  uint16_t range_flags = 0x1F;
  bool reverse = false, log = false, between = true;
  std::string partner_crosses;
  uint8_t major_tick = 3, minor_tick = 0, label_pos = 3;
  bool major_grid = false, minor_grid = false;
  bool has_line = false;
  Line line;
  std::string title;
  std::string title_sppr_xml, title_rich_xml;
  std::string sppr_xml, grid_xml, minor_grid_xml, txpr_xml;
};

struct AxisParent {
  uint16_t index = 0;
  Axis axes[3];
  std::vector<size_t> groups;
};

struct Chart {
  std::vector<Series> series;
  std::vector<Group> groups;
  std::vector<AxisParent> parents;
  std::string title;
  std::string title_sppr_xml, title_rich_xml;
  std::string sppr_xml, txpr_xml, plot_sppr_xml;
  bool has_3d = false;
  int rot_x = 15, rot_y = 20, height_percent = 100, depth_percent = 100;
  bool right_angle_axes = true, perspective = false;
  bool has_legend = false;
  uint8_t legend_pos = 3;
  bool legend_manual = false;
  double legend_rect[4] = {0, 0, 0, 0};
  std::string legend_sppr_xml, legend_txpr_xml;
};

inline std::string line_xml(const Line& l) {
  if (!l.set || l.automatic) return "";
  if (l.style == 5) return "<a:ln><a:noFill/></a:ln>";
  std::string x = "<a:ln";
  if (l.weight == -1) x += " w=\"3175\"";
  else x += " w=\"" + std::to_string(12700 * (l.weight + 1)) + "\"";
  x += ">";
  if (!l.auto_color) x += "<a:solidFill><a:srgbClr val=\"" + l.color + "\"/></a:solidFill>";
  static const char* dash[] = {"solid", "dash", "sysDot", "dashDot", "lgDashDotDot", "solid", "solid", "solid", "solid"};
  if (l.style > 0 && l.style < 5) x += std::string("<a:prstDash val=\"") + dash[l.style] + "\"/>";
  x += "</a:ln>";
  return x;
}

inline std::string fill_xml(const Area& a) {
  if (!a.set || a.automatic) return "";
  if (a.pattern == 0) return "<a:noFill/>";
  return "<a:solidFill><a:srgbClr val=\"" + a.color + "\"/></a:solidFill>";
}

inline std::string sppr(const Fmt& f, bool fill) {
  if (!f.sppr_xml.empty()) return f.sppr_xml;
  std::string x = (fill ? fill_xml(f.area) : std::string()) + line_xml(f.line);
  if (x.empty()) return "";
  return "<c:spPr>" + x + "</c:spPr>";
}

inline std::string marker_xml(const Marker& m, const std::string& sppr_xml) {
  if (!m.set) return "";
  static const char* sym[] = {"none", "square", "diamond", "triangle", "x", "star", "dot", "dash", "circle", "plus"};
  if (m.automatic && sppr_xml.empty()) return "";
  std::string x = "<c:marker><c:symbol val=\"" + std::string(sym[m.type > 9 ? 0 : m.type]) + "\"/>";
  if (m.type != 0) {
    uint32_t pt = m.size / 20;
    if (pt < 2) pt = 2;
    if (pt > 72) pt = 72;
    x += "<c:size val=\"" + std::to_string(pt) + "\"/>";
    if (!sppr_xml.empty())
      x += sppr_xml;
    else
      x += "<c:spPr><a:solidFill><a:srgbClr val=\"" + m.fore + "\"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val=\"" + m.back +
           "\"/></a:solidFill></a:ln></c:spPr>";
  }
  x += "</c:marker>";
  return x;
}

inline std::string cache_xml(const std::map<uint32_t, Cached>& pts, bool strings) {
  std::string x;
  uint32_t n = 0;
  for (std::map<uint32_t, Cached>::const_iterator it = pts.begin(); it != pts.end(); ++it)
    if (it->first + 1 > n) n = it->first + 1;
  x += std::string(strings ? "<c:strCache>" : "<c:numCache><c:formatCode>General</c:formatCode>");
  x += "<c:ptCount val=\"" + std::to_string(n) + "\"/>";
  for (std::map<uint32_t, Cached>::const_iterator it = pts.begin(); it != pts.end(); ++it)
    x += "<c:pt idx=\"" + std::to_string(it->first) + "\"><c:v>" + esc(it->second.text) + "</c:v></c:pt>";
  x += strings ? "</c:strCache>" : "</c:numCache>";
  return x;
}

// data of a series: a reference with cached values, or only the cached values
inline std::string data_xml(const Series& s, int which, const char* tag, bool force_strings) {
  uint16_t sdt = which == 2 ? s.type_x : s.type_y;
  std::map<uint16_t, std::map<uint32_t, Cached> >::const_iterator c = s.cache.find(static_cast<uint16_t>(which));
  bool have_cache = c != s.cache.end() && !c->second.empty();
  bool strings = force_strings || (which == 2 && sdt == 3);
  if (have_cache && !strings)
    for (std::map<uint32_t, Cached>::const_iterator it = c->second.begin(); it != c->second.end(); ++it)
      if (it->second.is_string) strings = true;
  const std::string& f = s.ref[which];
  bool by_ref = s.rt[which] == 2 && !f.empty();
  if (!by_ref && !have_cache) return "";

  std::string x = std::string("<c:") + tag + ">";
  if (by_ref) {
    x += strings ? "<c:strRef>" : "<c:numRef>";
    x += "<c:f>" + esc(f) + "</c:f>";
    if (have_cache) x += cache_xml(c->second, strings);
    x += strings ? "</c:strRef>" : "</c:numRef>";
  } else {
    x += strings ? "<c:strLit>" : "<c:numLit><c:formatCode>General</c:formatCode>";
    uint32_t n = 0;
    for (std::map<uint32_t, Cached>::const_iterator it = c->second.begin(); it != c->second.end(); ++it)
      if (it->first + 1 > n) n = it->first + 1;
    x += "<c:ptCount val=\"" + std::to_string(n) + "\"/>";
    for (std::map<uint32_t, Cached>::const_iterator it = c->second.begin(); it != c->second.end(); ++it)
      x += "<c:pt idx=\"" + std::to_string(it->first) + "\"><c:v>" + esc(it->second.text) + "</c:v></c:pt>";
    x += strings ? "</c:strLit>" : "</c:numLit>";
  }
  x += std::string("</c:") + tag + ">";
  return x;
}

inline std::string labels_xml(const Fmt& f) {
  if (!f.has_label) return "";
  uint16_t l = f.label;
  bool val = (l & 0x0001) != 0, pct = (l & 0x0002) != 0 || (l & 0x0004) != 0, cat = (l & 0x0010) != 0 || (l & 0x0004) != 0;
  bool ser = (l & 0x0040) != 0, bub = (l & 0x0020) != 0;
  if (!val && !pct && !cat && !ser && !bub) return "";
  std::string x = "<c:dLbls>";
  x += "<c:showLegendKey val=\"0\"/>";
  x += std::string("<c:showVal val=\"") + (val ? "1" : "0") + "\"/>";
  x += std::string("<c:showCatName val=\"") + (cat ? "1" : "0") + "\"/>";
  x += std::string("<c:showSerName val=\"") + (ser ? "1" : "0") + "\"/>";
  x += std::string("<c:showPercent val=\"") + (pct ? "1" : "0") + "\"/>";
  x += std::string("<c:showBubbleSize val=\"") + (bub ? "1" : "0") + "\"/>";
  x += "</c:dLbls>";
  return x;
}

inline size_t utf8_length(const std::string& s) {
  size_t n = 0;
  for (size_t i = 0; i < s.size(); ++i)
    if ((static_cast<uint8_t>(s[i]) & 0xC0u) != 0x80u) ++n;
  return n;
}

// The runs of the rich text hold placeholders of the length of the text of the run.
// The text is in a record of the chart, the runs get their part of it.
inline std::string fill_runs(const std::string& rich, const std::string& text) {
  std::string out;
  size_t pos = 0, used = 0;
  while (true) {
    size_t open = rich.find("<a:t>", pos);
    if (open == std::string::npos) break;
    size_t close = rich.find("</a:t>", open);
    if (close == std::string::npos) break;
    size_t len = utf8_length(rich.substr(open + 5, close - open - 5));
    out += rich.substr(pos, open + 5 - pos);
    size_t start = used, k = 0;
    while (used < text.size() && k < len) {
      ++used;
      while (used < text.size() && (static_cast<uint8_t>(text[used]) & 0xC0u) == 0x80u) ++used;
      ++k;
    }
    out += esc(text.substr(start, used - start));
    pos = close;
  }
  out += rich.substr(pos);
  // text beyond the runs goes into the last one
  if (used < text.size()) {
    size_t last = out.rfind("</a:t>");
    if (last != std::string::npos) out.insert(last, esc(text.substr(used)));
  }
  return out;
}

inline std::string title_xml(const std::string& text, const std::string& rich, const std::string& sppr) {
  std::string x = "<c:title><c:tx>";
  if (!rich.empty() && !text.empty()) x += fill_runs(rich, text);
  else if (!rich.empty()) x += rich;
  else x += "<c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" + esc(text) + "</a:t></a:r></a:p></c:rich>";
  x += "</c:tx><c:overlay val=\"0\"/>" + sppr + "</c:title>";
  return x;
}

inline std::string series_xml(const Chart& ch, size_t si, const Group& g) {
  const Series& s = ch.series[si];
  std::string x = "<c:ser><c:idx val=\"" + std::to_string(si) + "\"/><c:order val=\"" + std::to_string(si) + "\"/>";

  if (s.rt[0] == 2 && !s.ref[0].empty()) {
    x += "<c:tx><c:strRef><c:f>" + esc(s.ref[0]) + "</c:f>";
    if (!s.title_text.empty()) x += "<c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>" + esc(s.title_text) + "</c:v></c:pt></c:strCache>";
    x += "</c:strRef></c:tx>";
  } else if (!s.title_text.empty()) {
    x += "<c:tx><c:v>" + esc(s.title_text) + "</c:v></c:tx>";
  }

  bool bubble = g.type == GroupScatter && g.bubbles;
  bool lines = g.type == GroupLine || (g.type == GroupScatter && !bubble) || (g.type == GroupRadar && !g.filled);
  x += sppr(s.fmt, !lines);
  if (lines) x += marker_xml(s.fmt.marker, s.fmt.marker_sppr_xml);
  else if (g.type == GroupBar || bubble) x += "<c:invertIfNegative val=\"0\"/>";
  else if (g.type == GroupPie && s.fmt.explosion > 0) x += "<c:explosion val=\"" + std::to_string(s.fmt.explosion) + "\"/>";

  for (std::map<uint16_t, Fmt>::const_iterator p = s.points.begin(); p != s.points.end(); ++p) {
    std::string pp = sppr(p->second, !lines);
    if (pp.empty()) continue;
    x += "<c:dPt><c:idx val=\"" + std::to_string(p->first) + "\"/>";
    if (g.type == GroupBar) x += "<c:invertIfNegative val=\"0\"/>";
    if (lines) x += marker_xml(p->second.marker, p->second.marker_sppr_xml);
    if (g.type == GroupPie && p->second.explosion > 0) x += "<c:explosion val=\"" + std::to_string(p->second.explosion) + "\"/>";
    x += pp + "</c:dPt>";
  }
  x += labels_xml(s.fmt);

  if (bubble) {
    x += data_xml(s, 2, "xVal", false);
    x += data_xml(s, 1, "yVal", false);
    x += data_xml(s, 3, "bubbleSize", false);
    x += "<c:bubble3D val=\"0\"/>";
  } else if (g.type == GroupScatter) {
    x += data_xml(s, 2, "xVal", false);
    x += data_xml(s, 1, "yVal", false);
    x += std::string("<c:smooth val=\"") + (s.fmt.smooth ? "1" : "0") + "\"/>";
  } else {
    x += data_xml(s, 2, "cat", false);
    x += data_xml(s, 1, "val", false);
    if (g.type == GroupLine) x += std::string("<c:smooth val=\"") + (s.fmt.smooth ? "1" : "0") + "\"/>";
  }
  x += "</c:ser>";
  return x;
}

inline std::string axis_xml(const Axis& a, int id, int cross_id, bool value_axis, const char* pos, bool secondary,
                            const std::string& crosses) {
  std::string x = value_axis ? "<c:valAx>" : "<c:catAx>";
  x += "<c:axId val=\"" + std::to_string(id) + "\"/><c:scaling>";
  if (a.log) x += "<c:logBase val=\"10\"/>";
  x += std::string("<c:orientation val=\"") + (a.reverse ? "maxMin" : "minMax") + "\"/>";
  if (a.has_range && !(a.range_flags & 0x02)) x += "<c:max val=\"" + xls_num(a.max) + "\"/>";
  if (a.has_range && !(a.range_flags & 0x01)) x += "<c:min val=\"" + xls_num(a.min) + "\"/>";
  x += "</c:scaling>";
  x += std::string("<c:delete val=\"") + ((secondary && !value_axis) ? "1" : "0") + "\"/>";
  x += std::string("<c:axPos val=\"") + pos + "\"/>";
  if (a.major_grid) x += a.grid_xml.empty() ? "<c:majorGridlines/>" : "<c:majorGridlines>" + a.grid_xml + "</c:majorGridlines>";
  if (a.minor_grid) x += a.minor_grid_xml.empty() ? "<c:minorGridlines/>" : "<c:minorGridlines>" + a.minor_grid_xml + "</c:minorGridlines>";
  if (!a.title.empty() || !a.title_rich_xml.empty()) x += title_xml(a.title, a.title_rich_xml, a.title_sppr_xml);
  static const char* tick[] = {"none", "in", "out", "cross"};
  x += std::string("<c:majorTickMark val=\"") + tick[a.major_tick & 3] + "\"/>";
  x += std::string("<c:minorTickMark val=\"") + tick[a.minor_tick & 3] + "\"/>";
  static const char* lbl[] = {"none", "low", "high", "nextTo"};
  x += std::string("<c:tickLblPos val=\"") + lbl[a.label_pos & 3] + "\"/>";
  if (!a.sppr_xml.empty()) {
    x += a.sppr_xml;
  } else if (a.has_line) {
    std::string ln = line_xml(a.line);
    if (!ln.empty()) x += "<c:spPr>" + ln + "</c:spPr>";
  }
  x += a.txpr_xml;
  x += "<c:crossAx val=\"" + std::to_string(cross_id) + "\"/>";
  x += crosses.empty() ? "<c:crosses val=\"autoZero\"/>" : crosses;
  if (value_axis) {
    x += std::string("<c:crossBetween val=\"") + (a.between ? "between" : "midCat") + "\"/>";
    if (a.has_range && !(a.range_flags & 0x04)) x += "<c:majorUnit val=\"" + xls_num(a.major) + "\"/>";
    if (a.has_range && !(a.range_flags & 0x08)) x += "<c:minorUnit val=\"" + xls_num(a.minor) + "\"/>";
  } else {
    x += "<c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/><c:noMultiLvlLbl val=\"0\"/>";
  }
  x += value_axis ? "</c:valAx>" : "</c:catAx>";
  return x;
}

}  // namespace xls_chart_detail

// Reads the records of a chart substream (without BOF and EOF) and returns the
// chart part (chartN.xml). An empty string if nothing can be drawn.
inline std::string xls_chart_xml(const std::vector<XlsChartRecord>& recs, XlsFormulaContext& ctx, bool swapit, bool debug) {
  using namespace xls_chart_detail;
  Chart ch;

  enum Ctx { CNone, CChart, CSeries, CDataFormat, CFrame, CPlotFrame, CText, CAxis, CAxisParent, CChartFormat, CLegend, COther };
  std::vector<Ctx> stack;
  Ctx last = CNone;

  int cur_series = -1;
  int cur_group = -1;
  int cur_parent = -1;
  int cur_axis = -1;
  int axis_line = -1;
  bool df_valid = false;
  size_t df_series = 0;
  uint16_t df_point = 0xFFFF;
  std::string text;
  std::string text_sppr, text_rich;
  std::string* title_rich_target = nullptr;
  bool saw_plot_area = false;
  uint16_t cache_type = 0;
  bool in_data = false;

  for (size_t r = 0; r < recs.size(); ++r) {
    const XlsChartRecord& rec = recs[r];
    const std::vector<uint8_t>& d = rec.d;

    if (rec.id == XlsBegin) {
      stack.push_back(last);
      continue;
    }
    if (rec.id == XlsEnd) {
      Ctx closed = stack.empty() ? CNone : stack.back();
      if (!stack.empty()) stack.pop_back();
      if (closed == CDataFormat) df_valid = false;
      if (closed == CAxis) {
        cur_axis = -1;
        axis_line = -1;
      }
      if (closed == CSeries) cur_series = -1;
      if (closed == CAxisParent) {
        cur_parent = -1;
        cur_axis = -1;
      }
      if (closed == CChartFormat) cur_group = -1;
      last = CNone;
      continue;
    }
    Ctx top = stack.empty() ? CNone : stack.back();

    last = COther;
    switch (rec.id) {
      case XlsPlotArea:
        saw_plot_area = true;
        break;

      case XlsFrame:
        last = saw_plot_area ? CPlotFrame : CFrame;
        saw_plot_area = false;
        break;

      case XlsShapePropsStream: {
        std::string x = frt_xml(d, "a:spPr", "c:spPr");
        if (x.empty()) break;
        uint16_t what = static_cast<uint16_t>(u16(d, 12));
        Ctx parent = stack.size() >= 2 ? stack[stack.size() - 2] : CNone;
        if (top == CDataFormat && df_valid) {
          Fmt& f = df_point == 0xFFFF ? ch.series[df_series].fmt : ch.series[df_series].points[df_point];
          if (what == 0) f.sppr_xml = x;
          else if (what == 1) f.marker_sppr_xml = x;
        } else if (top == CAxis && cur_parent >= 0 && cur_axis >= 0) {
          Axis& a = ch.parents[static_cast<size_t>(cur_parent)].axes[cur_axis];
          if (what == 0) a.sppr_xml = x;
          else if (what == 1) a.grid_xml = x;
          else if (what == 2) a.minor_grid_xml = x;
        } else if (top == CPlotFrame) {
          ch.plot_sppr_xml = x;
        } else if (top == CFrame && parent == CChart) {
          ch.sppr_xml = x;
        } else if (top == CFrame && parent == CLegend) {
          ch.legend_sppr_xml = x;
        } else if (top == CFrame && parent == CText) {
          text_sppr = x;
        }
        break;
      }

      case XlsTextPropsStream: {
        std::string x = frt_xml(d, "a:txPr", "c:txPr");
        if (x.empty()) break;
        if (top == CChart) ch.txpr_xml = x;
        else if (top == CLegend) ch.legend_txpr_xml = x;
        else if (top == CAxis && cur_parent >= 0 && cur_axis >= 0) ch.parents[static_cast<size_t>(cur_parent)].axes[cur_axis].txpr_xml = x;
        break;
      }

      case XlsRichTextStream: {
        std::string x = frt_xml(d, "a:richFrt", "c:rich");
        if (x.empty() || top != CText) break;
        if (title_rich_target) *title_rich_target = x;
        else text_rich = x;
        break;
      }

      case XlsChart:
        last = CChart;
        break;

      case XlsSeries: {
        Series s;
        s.type_x = static_cast<uint16_t>(u16(d, 0));
        s.type_y = static_cast<uint16_t>(u16(d, 2));
        ch.series.push_back(s);
        cur_series = static_cast<int>(ch.series.size()) - 1;
        last = CSeries;
        break;
      }

      case XlsBRAI: {
        uint16_t id = static_cast<uint16_t>(u8(d, 0));
        uint16_t rt = static_cast<uint16_t>(u8(d, 1));
        uint16_t cce = static_cast<uint16_t>(u16(d, 6));
        if (cur_series >= 0 && top == CSeries && id < 4) {
          Series& s = ch.series[static_cast<size_t>(cur_series)];
          s.rt[id] = rt;
          if (rt == 2 && cce > 0 && d.size() >= 8u + cce) {
            XlsRgce rgce(std::vector<uint8_t>(d.begin() + 8, d.begin() + 8 + cce));
            ctx.unsupported = -1;
            ctx.row = 0;
            ctx.col = 0;
            s.ref[id] = xls_rgce(rgce, ctx, swapit, debug);
          }
        }
        break;
      }

      case XlsSeriesText: {
        std::string t = short_string(d, 2);
        if (top == CSeries && cur_series >= 0) ch.series[static_cast<size_t>(cur_series)].title_text = t;
        else if (top == CText) text = t;
        break;
      }

      case XlsSerToCrt:
        if (cur_series >= 0) ch.series[static_cast<size_t>(cur_series)].group = static_cast<int>(u16(d, 0));
        break;

      case XlsDataFormat: {
        uint16_t xi = static_cast<uint16_t>(u16(d, 0));
        uint16_t yi = static_cast<uint16_t>(u16(d, 2));
        uint16_t iss = static_cast<uint16_t>(u16(d, 4));
        (void)iss;
        df_valid = yi < ch.series.size();
        df_series = yi;
        df_point = xi;
        last = CDataFormat;
        break;
      }

      case XlsLineFormat:
      case XlsAreaFormat:
      case XlsMarkerFormat:
      case XlsPieFormat:
      case XlsSerFmt:
      case XlsAttachedLabel: {
        if (top == CDataFormat && df_valid) {
          Series& s = ch.series[df_series];
          Fmt& f = df_point == 0xFFFF ? s.fmt : s.points[df_point];
          if (rec.id == XlsLineFormat) {
            f.line.set = true;
            f.line.color = rgb(d, 0);
            f.line.style = static_cast<uint16_t>(u16(d, 4));
            f.line.weight = i16(d, 6);
            uint16_t fl = static_cast<uint16_t>(u16(d, 8));
            f.line.automatic = (fl & 0x01) != 0;
            f.line.auto_color = (fl & 0x08) != 0;
          } else if (rec.id == XlsAreaFormat) {
            f.area.set = true;
            f.area.color = rgb(d, 0);
            f.area.pattern = static_cast<uint16_t>(u16(d, 8));
            f.area.automatic = (u16(d, 10) & 0x01) != 0;
          } else if (rec.id == XlsMarkerFormat) {
            f.marker.set = true;
            f.marker.fore = rgb(d, 0);
            f.marker.back = rgb(d, 4);
            f.marker.type = static_cast<uint16_t>(u16(d, 8));
            f.marker.automatic = (u16(d, 10) & 0x01) != 0;
            f.marker.size = u32(d, 16);
          } else if (rec.id == XlsPieFormat) {
            f.explosion = static_cast<uint16_t>(u16(d, 0));
          } else if (rec.id == XlsSerFmt) {
            f.smooth = (u16(d, 0) & 0x01) != 0;
          } else {
            f.has_label = true;
            f.label = static_cast<uint16_t>(u16(d, 0));
          }
        } else if (rec.id == XlsLineFormat && top == CAxis && cur_parent >= 0 && cur_axis >= 0) {
          Axis& a = ch.parents[static_cast<size_t>(cur_parent)].axes[cur_axis];
          if (axis_line == 0) {
            a.has_line = true;
            a.line.set = true;
            a.line.color = rgb(d, 0);
            a.line.style = static_cast<uint16_t>(u16(d, 4));
            a.line.weight = i16(d, 6);
            uint16_t fl = static_cast<uint16_t>(u16(d, 8));
            a.line.automatic = (fl & 0x01) != 0;
            a.line.auto_color = (fl & 0x08) != 0;
          }
        }
        break;
      }

      case XlsChartFormat: {
        Group g;
        g.vary = (u16(d, 16) & 0x01) != 0;
        g.icrt = static_cast<int>(u16(d, 18));
        g.axis_parent = cur_parent;
        ch.groups.push_back(g);
        cur_group = static_cast<int>(ch.groups.size()) - 1;
        if (cur_parent >= 0) ch.parents[static_cast<size_t>(cur_parent)].groups.push_back(static_cast<size_t>(cur_group));
        last = CChartFormat;
        break;
      }

      case XlsBar:
        if (cur_group >= 0) {
          Group& g = ch.groups[static_cast<size_t>(cur_group)];
          uint16_t fl = static_cast<uint16_t>(u16(d, 4));
          g.type = GroupBar;
          g.overlap = -i16(d, 0);
          g.gap = static_cast<int>(u16(d, 2));
          g.horizontal = (fl & 0x01) != 0;
          g.stacked = (fl & 0x02) != 0;
          g.percent = (fl & 0x04) != 0;
        }
        break;

      case XlsChart3d:
        if (cur_group >= 0) {
          ch.groups[static_cast<size_t>(cur_group)].three_d = true;
          ch.groups[static_cast<size_t>(cur_group)].gap_depth = static_cast<int>(u16(d, 10));
          uint16_t fl = static_cast<uint16_t>(u16(d, 12));
          ch.has_3d = true;
          ch.rot_y = static_cast<int>(u16(d, 0));
          ch.rot_x = i16(d, 2);
          ch.height_percent = static_cast<int>(u16(d, 6));
          ch.depth_percent = static_cast<int>(u16(d, 8));
          ch.perspective = (fl & 0x01) != 0;
          ch.right_angle_axes = (fl & 0x02) != 0;
        }
        break;

      case XlsChart3DBarShape:
        if (cur_group >= 0 && d.size() >= 2) {
          uint32_t riser = u8(d, 0), taper = u8(d, 1);
          static const char* const shapes[2][3] = {{"box", "pyramidToMax", "pyramid"}, {"cylinder", "coneToMax", "cone"}};
          if (riser < 2 && taper < 3) ch.groups[static_cast<size_t>(cur_group)].shape = shapes[riser][taper];
        }
        break;

      case XlsLine:
      case XlsArea:
        if (cur_group >= 0) {
          Group& g = ch.groups[static_cast<size_t>(cur_group)];
          uint16_t fl = static_cast<uint16_t>(u16(d, 0));
          g.type = rec.id == XlsLine ? GroupLine : GroupArea;
          g.stacked = (fl & 0x01) != 0;
          g.percent = (fl & 0x02) != 0;
        }
        break;

      case XlsPie:
        if (cur_group >= 0) {
          Group& g = ch.groups[static_cast<size_t>(cur_group)];
          g.type = GroupPie;
          g.angle = static_cast<uint16_t>(u16(d, 0));
          g.donut = static_cast<uint16_t>(u16(d, 2));
        }
        break;

      case XlsRadar:
      case XlsRadarArea:
        if (cur_group >= 0) {
          Group& g = ch.groups[static_cast<size_t>(cur_group)];
          g.type = GroupRadar;
          g.filled = rec.id == XlsRadarArea;
        }
        break;

      case XlsSurf:
        if (cur_group >= 0) {
          Group& g = ch.groups[static_cast<size_t>(cur_group)];
          g.type = GroupSurface;
          g.filled = (u16(d, 0) & 0x01) != 0;
        }
        break;

      case XlsScatter:
        if (cur_group >= 0) {
          Group& g = ch.groups[static_cast<size_t>(cur_group)];
          g.type = GroupScatter;
          g.bubbles = (u16(d, 4) & 0x01) != 0;
          g.show_negative_bubbles = (u16(d, 4) & 0x02) != 0;
          g.bubble_scale = static_cast<int>(u16(d, 0));
          g.bubble_width = u16(d, 2) == 2;
        }
        break;

      case XlsLegend:
        ch.has_legend = true;
        ch.legend_pos = static_cast<uint8_t>(u8(d, 16));
        if ((u16(d, 18) & 0x01) == 0) {
          // not placed automatically: position and size in 1/4000 of the chart area
          ch.legend_manual = true;
          for (size_t k = 0; k < 4; ++k) ch.legend_rect[k] = static_cast<double>(static_cast<int32_t>(u32(d, k * 4))) / 4000.0;
        }
        last = CLegend;
        break;

      case XlsAxisParent: {
        AxisParent p;
        p.index = static_cast<uint16_t>(u16(d, 0));
        ch.parents.push_back(p);
        cur_parent = static_cast<int>(ch.parents.size()) - 1;
        last = CAxisParent;
        break;
      }

      case XlsAxis:
        if (cur_parent >= 0) {
          uint16_t t = static_cast<uint16_t>(u16(d, 0));
          if (t < 3) {
            cur_axis = t;
            ch.parents[static_cast<size_t>(cur_parent)].axes[t].present = true;
          }
        }
        last = CAxis;
        break;

      case XlsAxisLine:
        axis_line = static_cast<int>(u16(d, 0));
        if (cur_parent >= 0 && cur_axis >= 0) {
          Axis& a = ch.parents[static_cast<size_t>(cur_parent)].axes[cur_axis];
          if (axis_line == 1) a.major_grid = true;
          if (axis_line == 2) a.minor_grid = true;
        }
        break;

      case XlsValueRange:
        if (cur_parent >= 0 && cur_axis >= 0) {
          Axis& a = ch.parents[static_cast<size_t>(cur_parent)].axes[cur_axis];
          a.has_range = true;
          a.min = f64(d, 0);
          a.max = f64(d, 8);
          a.major = f64(d, 16);
          a.minor = f64(d, 24);
          a.cross = f64(d, 32);
          a.range_flags = static_cast<uint16_t>(u16(d, 40));
          a.log = (a.range_flags & 0x20) != 0;
          a.reverse = (a.range_flags & 0x40) != 0;
          if (!(a.range_flags & 0x10)) a.partner_crosses = "<c:crossesAt val=\"" + xls_num(a.cross) + "\"/>";
          else a.partner_crosses = std::string("<c:crosses val=\"") + ((a.range_flags & 0x80) ? "max" : "autoZero") + "\"/>";
        }
        break;

      case XlsCatSerRange:
        if (cur_parent >= 0 && cur_axis >= 0) {
          Axis& a = ch.parents[static_cast<size_t>(cur_parent)].axes[cur_axis];
          uint16_t fl = static_cast<uint16_t>(u16(d, 6));
          a.between = (fl & 0x01) != 0;
          a.reverse = (fl & 0x04) != 0;
          uint32_t cat_cross = u16(d, 0);
          if (fl & 0x02) a.partner_crosses = "<c:crosses val=\"max\"/>";
          else if (cat_cross > 1) a.partner_crosses = "<c:crossesAt val=\"" + std::to_string(cat_cross) + "\"/>";
          else a.partner_crosses = "<c:crosses val=\"autoZero\"/>";
        }
        break;

      case XlsTick:
        if (cur_parent >= 0 && cur_axis >= 0) {
          Axis& a = ch.parents[static_cast<size_t>(cur_parent)].axes[cur_axis];
          a.major_tick = static_cast<uint8_t>(u8(d, 0));
          a.minor_tick = static_cast<uint8_t>(u8(d, 1));
          a.label_pos = static_cast<uint8_t>(u8(d, 2));
        }
        break;

      case XlsText:
        text.clear();
        text_sppr.clear();
        text_rich.clear();
        title_rich_target = nullptr;
        last = CText;
        break;

      case XlsObjectLink: {
        uint16_t what = static_cast<uint16_t>(u16(d, 0));
        if (what == 1) {
          ch.title = text;
          ch.title_sppr_xml = text_sppr;
          ch.title_rich_xml = text_rich;
          title_rich_target = &ch.title_rich_xml;
        } else if ((what == 2 || what == 3) && !ch.parents.empty()) {
          size_t p = cur_parent >= 0 ? static_cast<size_t>(cur_parent) : 0;
          Axis& a = ch.parents[p].axes[what == 3 ? 0 : 1];
          a.title = text;
          a.title_sppr_xml = text_sppr;
          a.title_rich_xml = text_rich;
          title_rich_target = &a.title_rich_xml;
        }
        break;
      }

      case XlsSIIndex:
        cache_type = static_cast<uint16_t>(u16(d, 0));
        in_data = true;
        break;

      case XlsNumber:
      case XlsLabel:
      case XlsBlank:
        if (in_data) {
          uint32_t pt = u16(d, 0);
          size_t si = u16(d, 2);
          if (si < ch.series.size()) {
            Cached c;
            if (rec.id == XlsNumber) {
              c.text = xls_num(f64(d, 6));
            } else if (rec.id == XlsLabel) {
              c.is_string = true;
              uint32_t cch = u16(d, 6);
              std::vector<uint8_t> sd(d.begin() + std::min<size_t>(d.size(), 8), d.end());
              // XLUnicodeString: the flags follow the 2 bytes of the length
              std::vector<uint8_t> tmp;
              tmp.push_back(static_cast<uint8_t>(cch & 0xFF));
              tmp.insert(tmp.end(), sd.begin(), sd.end());
              c.text = short_string(tmp, 0);
            } else {
              break;
            }
            ch.series[si].cache[cache_type][pt] = c;
          }
        }
        break;

      default:
        break;
    }
  }

  // no ChartFormat: nothing to draw
  if (ch.groups.empty()) return "";
  for (size_t i = 0; i < ch.groups.size(); ++i)
    if (ch.groups[i].type == GroupOther) return "";

  for (size_t i = 0; i < ch.series.size(); ++i) {
    int icrt = ch.series[i].group;
    for (size_t g = 0; g < ch.groups.size(); ++g) {
      if (ch.groups[g].icrt == icrt) {
        ch.groups[g].series.push_back(i);
        break;
      }
    }
  }

  std::string x;
  x += "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" "
       "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
       "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">";
  x += "<c:roundedCorners val=\"0\"/><c:chart>";
  bool has_title = !ch.title.empty() || !ch.title_rich_xml.empty();
  if (has_title) x += title_xml(ch.title, ch.title_rich_xml, ch.title_sppr_xml);
  x += std::string("<c:autoTitleDeleted val=\"") + (has_title ? "0" : "1") + "\"/>";
  if (ch.has_3d) {
    x += "<c:view3D><c:rotX val=\"" + std::to_string(ch.rot_x) + "\"/>";
    if (!ch.right_angle_axes) x += "<c:hPercent val=\"" + std::to_string(ch.height_percent) + "\"/>";
    x += "<c:rotY val=\"" + std::to_string(ch.rot_y) + "\"/><c:depthPercent val=\"" + std::to_string(ch.depth_percent) + "\"/>" +
         "<c:rAngAx val=\"" + (ch.right_angle_axes ? "1" : "0") + "\"/>" + (ch.perspective ? "<c:perspective val=\"30\"/>" : "") + "</c:view3D>";
  }
  x += "<c:plotArea><c:layout/>";

  std::string axes;
  for (size_t p = 0; p < ch.parents.size(); ++p) {
    const AxisParent& ap = ch.parents[p];
    if (ap.groups.empty()) continue;
    bool secondary = p > 0;
    int cat_id = 50010 + static_cast<int>(p) * 2;
    int val_id = cat_id + 1;
    bool pie_only = true;
    bool scatter = false, horizontal = false, ser_axis_needed = false;
    for (size_t k = 0; k < ap.groups.size(); ++k) {
      const Group& g = ch.groups[ap.groups[k]];
      if (g.type != GroupPie) pie_only = false;
      if (g.type == GroupScatter) scatter = true;
      if (g.type == GroupBar && g.horizontal) horizontal = true;
    }

    for (size_t k = 0; k < ap.groups.size(); ++k) {
      const Group& g = ch.groups[ap.groups[k]];
      std::string sers;
      for (size_t s = 0; s < g.series.size(); ++s) sers += series_xml(ch, g.series[s], g);
      std::string ids = pie_only ? "" : "<c:axId val=\"" + std::to_string(cat_id) + "\"/><c:axId val=\"" + std::to_string(val_id) + "\"/>";
      std::string grouping = g.percent ? "percentStacked" : (g.stacked ? "stacked" : "standard");
      std::string dlbls = "<c:dLbls><c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/><c:showCatName val=\"0\"/><c:showSerName val=\"0\"/>"
                          "<c:showPercent val=\"0\"/><c:showBubbleSize val=\"0\"/></c:dLbls>";
      std::string gap_depth = "<c:gapDepth val=\"" + std::to_string(g.gap_depth) + "\"/>";
      bool serax = g.three_d && !g.stacked && !g.percent && !(g.type == GroupBar && ch.right_angle_axes);
      if (g.three_d && (g.type == GroupLine || g.type == GroupSurface || serax)) {
        ids += "<c:axId val=\"" + std::to_string(50100 + static_cast<int>(p)) + "\"/>";
        ser_axis_needed = true;
      } else if (g.three_d && g.type != GroupPie) {
        ids += "<c:axId val=\"0\"/>";
      }
      switch (g.type) {
        case GroupBar:
          x += std::string(g.three_d ? "<c:bar3DChart>" : "<c:barChart>") + "<c:barDir val=\"" + (g.horizontal ? "bar" : "col") + "\"/><c:grouping val=\"" +
               (g.percent ? "percentStacked" : (g.stacked ? "stacked" : (g.three_d && serax ? "standard" : "clustered"))) + "\"/><c:varyColors val=\"" +
               (g.vary ? "1" : "0") + "\"/>" + sers + dlbls + "<c:gapWidth val=\"" + std::to_string(g.gap) + "\"/>";
          if (g.three_d) {
            x += gap_depth + "<c:shape val=\"" + g.shape + "\"/>";
          } else {
            if (g.stacked || g.percent) x += "<c:overlap val=\"100\"/>";
            else if (g.overlap != 0) x += "<c:overlap val=\"" + std::to_string(g.overlap) + "\"/>";
          }
          x += ids + (g.three_d ? "</c:bar3DChart>" : "</c:barChart>");
          break;
        case GroupLine:
          if (g.three_d)
            x += "<c:line3DChart><c:grouping val=\"" + grouping + "\"/><c:varyColors val=\"0\"/>" + sers + dlbls + gap_depth + ids + "</c:line3DChart>";
          else
            x += "<c:lineChart><c:grouping val=\"" + grouping + "\"/><c:varyColors val=\"0\"/>" + sers + dlbls + "<c:marker val=\"1\"/>" + ids + "</c:lineChart>";
          break;
        case GroupArea:
          if (g.three_d)
            x += "<c:area3DChart><c:grouping val=\"" + grouping + "\"/><c:varyColors val=\"0\"/>" + sers + dlbls + gap_depth + ids + "</c:area3DChart>";
          else
            x += "<c:areaChart><c:grouping val=\"" + grouping + "\"/><c:varyColors val=\"0\"/>" + sers + dlbls + ids + "</c:areaChart>";
          break;
        case GroupPie:
          if (g.three_d) {
            x += "<c:pie3DChart><c:varyColors val=\"1\"/>" + sers + dlbls + "</c:pie3DChart>";
          } else if (g.donut > 0) {
            x += std::string("<c:doughnutChart><c:varyColors val=\"1\"/>") + sers + dlbls + "<c:firstSliceAng val=\"" + std::to_string(g.angle) +
                 "\"/><c:holeSize val=\"" + std::to_string(g.donut) + "\"/></c:doughnutChart>";
          } else {
            x += std::string("<c:pieChart><c:varyColors val=\"1\"/>") + sers + dlbls + "<c:firstSliceAng val=\"" + std::to_string(g.angle) + "\"/></c:pieChart>";
          }
          break;
        case GroupRadar:
          x += std::string("<c:radarChart><c:radarStyle val=\"") + (g.filled ? "filled" : "marker") + "\"/><c:varyColors val=\"0\"/>" + sers + dlbls + ids +
               "</c:radarChart>";
          break;
        case GroupSurface:
          x += std::string("<c:surface3DChart><c:wireframe val=\"") + (g.filled ? "0" : "1") + "\"/>" + sers + ids + "</c:surface3DChart>";
          break;
        case GroupScatter:
          if (g.bubbles) {
            x += "<c:bubbleChart><c:varyColors val=\"0\"/>" + sers + dlbls + "<c:bubbleScale val=\"" + std::to_string(g.bubble_scale) +
                 "\"/><c:showNegBubbles val=\"" + (g.show_negative_bubbles ? "1" : "0") + "\"/><c:sizeRepresents val=\"" +
                 (g.bubble_width ? "w" : "area") + "\"/>" + ids + "</c:bubbleChart>";
            break;
          }
          x += "<c:scatterChart><c:scatterStyle val=\"lineMarker\"/><c:varyColors val=\"0\"/>" + sers + dlbls + ids + "</c:scatterChart>";
          break;
        default:
          break;
      }
    }

    if (!pie_only) {
      Axis a0 = ap.axes[0], a1 = ap.axes[1];
      if (scatter) {
        a0.between = false;
        a1.between = false;
      }
      const char* p0 = horizontal ? "l" : "b";
      const char* p1 = horizontal ? "b" : "l";
      std::string cat_crosses = a1.partner_crosses, val_crosses = a0.partner_crosses;
      if (secondary) {
        p1 = horizontal ? "t" : "r";
        val_crosses = "<c:crosses val=\"max\"/>";
      }
      axes += axis_xml(a0, cat_id, val_id, scatter, p0, secondary, cat_crosses);
      axes += axis_xml(a1, val_id, cat_id, true, p1, secondary, val_crosses);
      if (ser_axis_needed) {
        axes += "<c:serAx><c:axId val=\"" + std::to_string(50100 + static_cast<int>(p)) + "\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling>"
                "<c:delete val=\"0\"/><c:axPos val=\"b\"/><c:majorTickMark val=\"out\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/>"
                "<c:crossAx val=\"" + std::to_string(val_id) + "\"/><c:crosses val=\"autoZero\"/></c:serAx>";
      }
    }
  }
  x += axes;
  x += ch.plot_sppr_xml;
  x += "</c:plotArea>";

  if (ch.has_legend) {
    static const char* pos[] = {"b", "tr", "t", "r", "l", "r", "r", "r"};
    std::string layout;
    if (ch.legend_manual)
      layout = "<c:layout><c:manualLayout><c:xMode val=\"edge\"/><c:yMode val=\"edge\"/><c:x val=\"" + xls_num(ch.legend_rect[0]) + "\"/><c:y val=\"" +
               xls_num(ch.legend_rect[1]) + "\"/><c:w val=\"" + xls_num(ch.legend_rect[2]) + "\"/><c:h val=\"" + xls_num(ch.legend_rect[3]) +
               "\"/></c:manualLayout></c:layout>";
    x += std::string("<c:legend><c:legendPos val=\"") + pos[ch.legend_pos & 7] + "\"/>" + layout + "<c:overlay val=\"0\"/>" + ch.legend_sppr_xml + ch.legend_txpr_xml + "</c:legend>";
  }
  x += "<c:plotVisOnly val=\"1\"/><c:dispBlanksAs val=\"gap\"/></c:chart>" + ch.sppr_xml + ch.txpr_xml + "</c:chartSpace>";
  return x;
}

#endif
