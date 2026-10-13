#ifndef XLS_PIVOT_H
#define XLS_PIVOT_H

#include <cmath>
#include <cstring>
#include <string>
#include <vector>

#include "xls_defines.h"
#include "xls_funs.h"

// Pivot tables. A table is a group of records in the stream of its sheet
// (SXVIEW, SXVD, SXVI, ...). The data it is based on, the pivot cache, is a
// stream of its own in the storage _SX_DB_CUR. Both are written the way
// openxlsx2 writes its pivot tables: the cache has no records and is refreshed
// when the file is opened, so only the layout and the items of the fields are
// needed.

struct XlsPivotSource {
  uint16_t id = 0;
  uint16_t type = 0;   // 1 worksheet, 2 external, 4 consolidation, 0x10 scenario
  std::string sheet;
  std::string ref;
  std::string name;    // named range
  bool supported() const { return type == 1 && (!ref.empty() || !name.empty()); }
};

struct XlsPivotItem {
  uint16_t type = 0;
  uint16_t flags = 0;
  uint16_t cache = 0xFFFF;
  std::string name;
};

struct XlsPivotField {
  uint16_t axis = 0;
  std::string name;
  std::vector<XlsPivotItem> items;
  bool default_subtotal = true, multi = false;
  uint32_t numfmt = 0;
  uint8_t sort = 0;
};

struct XlsPivotDataField {
  uint16_t field = 0, function = 0, display = 0, base_field = 0, base_item = 0, numfmt = 0;
  std::string name;
};

struct XlsPivotPage {
  uint16_t field = 0, item = 0;
};

struct XlsPivotLine {
  uint16_t reuse = 0;
  uint8_t type = 0;
  uint32_t data = 0;
  std::vector<uint32_t> x;
};

struct XlsPivotFormat {
  uint32_t dxf = 0;
  int32_t field = -1;
  uint8_t type = 0, flags = 0, axis = 0, position = 0;
};

struct XlsPivotView {
  XlsPivotSource source;
  uint16_t cache = 0;
  uint16_t row1 = 0, row2 = 0, col1 = 0, col2 = 0, head_row = 0, data_row = 0, data_col = 0, grbit = 0;
  uint16_t n_row = 0, n_col = 0;
  std::string name, data_name;
  std::vector<XlsPivotField> fields;
  std::vector<std::vector<uint16_t>> axes;  // SXIVD: row fields, then column fields
  std::vector<XlsPivotPage> pages;
  std::vector<XlsPivotDataField> data;
  std::vector<XlsPivotLine> lines[2];
  bool row_grand = true, col_grand = true, outline = false;
  unsigned version = 8;
  std::vector<XlsPivotFormat> formats;
  std::string style = "PivotStyleLight16";
  uint16_t style_flags = 0x32;
};

inline uint32_t xls_pv_u16(const std::vector<uint8_t>& r, size_t at) {
  return at + 2 <= r.size() ? static_cast<uint32_t>(r[at] | (r[at + 1] << 8)) : 0u;
}

inline uint32_t xls_pv_u32(const std::vector<uint8_t>& r, size_t at) {
  return xls_pv_u16(r, at) | (xls_pv_u16(r, at + 2) << 16);
}

// cch characters with the flags byte in front, 0xFFFF as cch means no string
inline std::string xls_pv_string(const std::vector<uint8_t>& r, size_t& pos, size_t cch) {
  if (cch == 0xFFFF || pos >= r.size()) return std::string();
  bool wide = (r[pos] & 0x01) != 0;
  ++pos;
  std::u16string s;
  for (size_t i = 0; i < cch && pos < r.size(); ++i) {
    if (wide) {
      s.push_back(static_cast<char16_t>(xls_pv_u16(r, pos)));
      pos += 2;
    } else {
      s.push_back(static_cast<char16_t>(r[pos]));
      ++pos;
    }
  }
  return to_utf8(s);
}

// DCONREF: the area on a sheet. The name of the sheet starts with 0x02 if it is
// in this file.
inline void xls_pv_dconref(const std::vector<uint8_t>& r, XlsPivotSource& src) {
  size_t pos = 6;
  size_t cch = xls_pv_u16(r, pos);
  pos += 2;
  std::string file = xls_pv_string(r, pos, cch);
  if (file.empty() || file[0] != '\x02') {
    src.type = 2;
    return;
  }
  src.sheet = file.substr(1);
  src.ref = int_to_col(r.size() > 5 ? static_cast<size_t>(r[4]) + 1 : 1) + std::to_string(xls_pv_u16(r, 0) + 1) + ":" +
            int_to_col(r.size() > 5 ? static_cast<size_t>(r[5]) + 1 : 1) + std::to_string(xls_pv_u16(r, 2) + 1);
}

// DCONNAME: a named range
inline void xls_pv_dconname(const std::vector<uint8_t>& r, XlsPivotSource& src) {
  size_t pos = 2;
  size_t cch = r.empty() ? 0 : r[0];
  src.name = xls_pv_string(r, pos, cch);
}

inline void xls_pv_view(const std::vector<uint8_t>& r, XlsPivotView& v) {
  v.row1 = static_cast<uint16_t>(xls_pv_u16(r, 0));
  v.row2 = static_cast<uint16_t>(xls_pv_u16(r, 2));
  v.col1 = static_cast<uint16_t>(xls_pv_u16(r, 4));
  v.col2 = static_cast<uint16_t>(xls_pv_u16(r, 6));
  v.head_row = static_cast<uint16_t>(xls_pv_u16(r, 8));
  v.data_row = static_cast<uint16_t>(xls_pv_u16(r, 10));
  v.data_col = static_cast<uint16_t>(xls_pv_u16(r, 12));
  v.cache = static_cast<uint16_t>(xls_pv_u16(r, 14));
  v.n_row = static_cast<uint16_t>(xls_pv_u16(r, 24));
  v.n_col = static_cast<uint16_t>(xls_pv_u16(r, 26));
  v.grbit = static_cast<uint16_t>(xls_pv_u16(r, 36));
  size_t cch_table = xls_pv_u16(r, 40), cch_data = xls_pv_u16(r, 42);
  size_t pos = 44;
  v.name = xls_pv_string(r, pos, cch_table);
  v.data_name = xls_pv_string(r, pos, cch_data);
}

inline void xls_pv_field(const std::vector<uint8_t>& r, XlsPivotField& f) {
  f.axis = static_cast<uint16_t>(xls_pv_u16(r, 0));
  size_t cch = xls_pv_u16(r, 8);
  size_t pos = 10;
  f.name = xls_pv_string(r, pos, cch);
}

inline void xls_pv_item(const std::vector<uint8_t>& r, XlsPivotItem& i) {
  i.type = static_cast<uint16_t>(xls_pv_u16(r, 0));
  i.flags = static_cast<uint16_t>(xls_pv_u16(r, 2));
  i.cache = static_cast<uint16_t>(xls_pv_u16(r, 4));
  size_t cch = xls_pv_u16(r, 6);
  size_t pos = 8;
  i.name = xls_pv_string(r, pos, cch);
}

inline void xls_pv_data(const std::vector<uint8_t>& r, XlsPivotDataField& d) {
  d.field = static_cast<uint16_t>(xls_pv_u16(r, 0));
  d.function = static_cast<uint16_t>(xls_pv_u16(r, 2));
  d.display = static_cast<uint16_t>(xls_pv_u16(r, 4));
  d.base_field = static_cast<uint16_t>(xls_pv_u16(r, 6));
  d.base_item = static_cast<uint16_t>(xls_pv_u16(r, 8));
  d.numfmt = static_cast<uint16_t>(xls_pv_u16(r, 10));
  size_t cch = xls_pv_u16(r, 12);
  size_t pos = 14;
  d.name = xls_pv_string(r, pos, cch);
}

struct XlsPivotCacheField {
  std::string name;
  std::string items;
  std::string attrs;
  size_t expected = 0, count = 0, n_str = 0, n_num = 0, n_int = 0, n_blank = 0, n_other = 0, n_date = 0;
  double min = 0, max = 0;
  std::string min_date, max_date;
  uint32_t flags = 0;
  bool shared = false;
};

inline std::string xls_pivot_cache_write(const std::vector<XlsPivotCacheField>& fields, uint32_t n_records, const XlsPivotSource& src,
                                         uint32_t slicer_id = 0, bool refresh = true) {
  std::string out = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
  out += "<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
         "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId1\" refreshedBy=\"openxlsx2\" " +
         std::string(refresh ? "invalid=\"1\" refreshOnLoad=\"1\" " : "") + "refreshedDate=\"1\" createdVersion=\"8\" refreshedVersion=\"8\" "
         "minRefreshableVersion=\"3\" recordCount=\"" + std::to_string(n_records) + "\">";
  out += "<cacheSource type=\"worksheet\"><worksheetSource ";
  if (!src.name.empty()) {
    out += "name=\"" + escape_xml(src.name) + "\"/>";
  } else {
    out += "ref=\"" + src.ref + "\" sheet=\"" + escape_xml(src.sheet) + "\"/>";
  }
  out += "</cacheSource><cacheFields count=\"" + std::to_string(fields.size()) + "\">";
  for (size_t i = 0; i < fields.size(); ++i) {
    const XlsPivotCacheField& f = fields[i];
    out += "<cacheField name=\"" + escape_xml(f.name) + "\" numFmtId=\"0\"><sharedItems";
    if (!f.attrs.empty()) {
      out += f.attrs;
    } else if (f.count > 0 && f.n_date + f.n_blank == f.count && f.n_date > 0) {
      out += std::string(f.n_blank == 0 ? " containsSemiMixedTypes=\"0\"" : "") + " containsNonDate=\"0\" containsDate=\"1\" containsString=\"0\" minDate=\"" +
             f.min_date + "\" maxDate=\"" + f.max_date + "\"";
    } else if (f.count > 0 && f.n_num == f.count) {
      out += " containsSemiMixedTypes=\"0\" containsString=\"0\" containsNumber=\"1\"";
      if (f.n_int == f.n_num) out += " containsInteger=\"1\"";
      out += " minValue=\"" + xls_num(f.min) + "\" maxValue=\"" + xls_num(f.max) + "\"";
    } else if (f.n_num > 0 && f.n_str > 0) {
      out += " containsMixedTypes=\"1\" containsNumber=\"1\"";
    } else if (f.n_num > 0) {
      out += " containsSemiMixedTypes=\"0\" containsString=\"0\" containsNumber=\"1\"";
    }
    if (f.n_blank > 0) out += " containsBlank=\"1\"";
    if (f.count == 0) {
      out += "/>";
    } else {
      out += " count=\"" + std::to_string(f.count) + "\">" + f.items + "</sharedItems>";
    }
    out += "</cacheField>";
  }
  out += "</cacheFields>";
  if (slicer_id != 0) {
    out += "<extLst><ext uri=\"{725AE2AE-9491-48be-B2B4-4EB974FC3084}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">"
           "<x14:pivotCacheDefinition pivotCacheId=\"" + std::to_string(slicer_id) + "\"/></ext></extLst>";
  }
  out += "</pivotCacheDefinition>\n";
  return out;
}

// The definition of the cache, from the records of its stream: SXDB, then per
// field an SXFDB followed by the items (SXSTRING, SXNUM, ...) of the field. The
// records of the cache follow after the last field and are not read.
inline std::string xls_pivot_cache_xml(const std::vector<uint8_t>& stream, const XlsPivotSource& src) {
  std::vector<XlsPivotCacheField> fields;
  uint32_t n_records = 0;

  size_t pos = 0;
  while (pos + 4 <= stream.size()) {
    uint16_t id = static_cast<uint16_t>(xls_pv_u16(stream, pos));
    size_t size = xls_pv_u16(stream, pos + 2);
    if (pos + 4 + size > stream.size()) break;
    std::vector<uint8_t> r(stream.begin() + static_cast<std::ptrdiff_t>(pos + 4), stream.begin() + static_cast<std::ptrdiff_t>(pos + 4 + size));
    pos += 4 + size;

    if (id == XlsSXDB) {
      n_records = xls_pv_u32(r, 0);
    } else if (id == XlsSXFDB) {
      XlsPivotCacheField f;
      f.expected = xls_pv_u16(r, 12);
      size_t cch = xls_pv_u16(r, 14);
      size_t at = 16;
      f.name = xls_pv_string(r, at, cch);
      fields.push_back(f);
    } else if (fields.empty() || fields.back().count >= fields.back().expected) {
      continue;
    } else if (id == XlsSXString) {
      XlsPivotCacheField& f = fields.back();
      size_t at = 2;
      f.items += "<s v=\"" + escape_xml(xls_pv_string(r, at, xls_pv_u16(r, 0))) + "\"/>";
      ++f.count;
      ++f.n_str;
    } else if (id == XlsSXNum) {
      XlsPivotCacheField& f = fields.back();
      uint64_t bits = static_cast<uint64_t>(xls_pv_u32(r, 0)) | (static_cast<uint64_t>(xls_pv_u32(r, 4)) << 32);
      double d = 0;
      std::memcpy(&d, &bits, sizeof(d));
      if (f.n_num == 0 || d < f.min) f.min = d;
      if (f.n_num == 0 || d > f.max) f.max = d;
      if (d == std::floor(d)) ++f.n_int;
      f.items += "<n v=\"" + xls_num(d) + "\"/>";
      ++f.count;
      ++f.n_num;
    } else if (id == XlsSxBool) {
      XlsPivotCacheField& f = fields.back();
      f.items += std::string("<b v=\"") + (xls_pv_u16(r, 0) ? "1" : "0") + "\"/>";
      ++f.count;
      ++f.n_other;
    } else if (id == XlsSxErr) {
      XlsPivotCacheField& f = fields.back();
      f.items += "<e v=\"" + xls_error(static_cast<uint8_t>(xls_pv_u16(r, 0))) + "\"/>";
      ++f.count;
      ++f.n_other;
    } else if (id == 0x00CF) {
      XlsPivotCacheField& f = fields.back();
      f.items += "<m/>";
      ++f.count;
      ++f.n_blank;
    }
  }

  return xls_pivot_cache_write(fields, n_records, src);
}

// The table itself
inline std::string xls_pivot_view_xml(const XlsPivotView& v, int32_t cache_number) {
  static const char* const subtotal[11] = {"sum", "count", "average", "max", "min", "product", "countNums", "stdDev", "stdDevp", "var", "varp"};
  static const char* const item_type[14] = {"", "default", "sum", "countA", "count", "avg", "max", "min", "product", "count", "stdDev", "stdDevP", "var", "varP"};
  static const char* const show_as[9] = {"normal", "difference", "percent", "percentDiff", "runTotal", "percentOfRow", "percentOfCol", "percentOfTotal", "index"};

  std::string ref = int_to_col(v.col1 + 1u) + std::to_string(v.row1 + 1u) + ":" + int_to_col(v.col2 + 1u) + std::to_string(v.row2 + 1u);

  bool data_on_rows = false;
  if (!v.axes.empty()) {
    for (size_t i = 0; i < v.axes[0].size(); ++i)
      if (v.axes[0][i] == 0xFFFE) data_on_rows = true;
  }
  size_t n_page = v.pages.size();

  std::string out = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
  out += "<pivotTableDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"" + escape_xml(v.name) +
         "\" cacheId=\"" + std::to_string(cache_number) + "\" applyNumberFormats=\"0\" applyBorderFormats=\"0\" applyFontFormats=\"0\" "
         "applyPatternFormats=\"0\" applyAlignmentFormats=\"0\" applyWidthHeightFormats=\"1\"";
  if (!v.data_name.empty()) out += " dataCaption=\"" + escape_xml(v.data_name) + "\"";
  if (data_on_rows) out += " dataOnRows=\"1\"";
  if (!v.row_grand) out += " rowGrandTotals=\"0\"";
  if (!v.col_grand) out += " colGrandTotals=\"0\"";
  out += " updatedVersion=\"" + std::to_string(v.version) + "\" minRefreshableVersion=\"3\" useAutoFormatting=\"1\" itemPrintTitles=\"1\" "
         "createdVersion=\"" + std::to_string(v.version) + "\" indent=\"0\"";
  if (v.outline) out += " outline=\"1\"";
  out += " outlineData=\"1\" multipleFieldFilters=\"0\">";

  out += "<location ref=\"" + ref + "\" firstHeaderRow=\"" + std::to_string(v.head_row >= v.row1 ? v.head_row - v.row1 : 0) +
         "\" firstDataRow=\"" + std::to_string(v.data_row >= v.row1 ? v.data_row - v.row1 : 0) +
         "\" firstDataCol=\"" + std::to_string(v.data_col >= v.col1 ? v.data_col - v.col1 : 0) + "\"";
  if (n_page > 0) out += " rowPageCount=\"" + std::to_string(n_page) + "\" colPageCount=\"1\"";
  out += "/>";

  out += "<pivotFields count=\"" + std::to_string(v.fields.size()) + "\">";
  for (size_t i = 0; i < v.fields.size(); ++i) {
    const XlsPivotField& f = v.fields[i];
    out += "<pivotField";
    if (!f.name.empty()) out += " name=\"" + escape_xml(f.name) + "\"";
    if (f.axis & 0x01) out += " axis=\"axisRow\"";
    else if (f.axis & 0x02) out += " axis=\"axisCol\"";
    else if (f.axis & 0x04) out += " axis=\"axisPage\"";
    if (f.axis & 0x08) out += " dataField=\"1\"";
    out += " showAll=\"0\"";
    bool has_hidden = false;
    for (size_t k = 0; k < f.items.size(); ++k)
      if (f.items[k].flags & 0x01) has_hidden = true;
    if ((has_hidden || f.multi) && (f.axis & 0x07) != 0) out += " multipleItemSelectionAllowed=\"1\"";
    if (f.numfmt != 0) out += " numFmtId=\"" + std::to_string(f.numfmt) + "\"";
    if (f.sort == 1) out += " sortType=\"ascending\"";
    if (f.sort == 2) out += " sortType=\"descending\"";
    if (!f.default_subtotal) out += " defaultSubtotal=\"0\"";
    if (f.items.empty() || (f.axis & 0x07) == 0) {
      out += "/>";
      continue;
    }
    out += "><items count=\"" + std::to_string(f.items.size()) + "\">";
    for (size_t k = 0; k < f.items.size(); ++k) {
      const XlsPivotItem& it = f.items[k];
      out += "<item";
      if (!it.name.empty()) out += " n=\"" + escape_xml(it.name) + "\"";
      if (it.flags & 0x01) out += " h=\"1\"";
      if (it.flags & 0x02) out += " sd=\"0\"";
      if (it.flags & 0x08) out += " m=\"1\"";
      if (it.type == 0) {
        out += " x=\"" + std::to_string(it.cache) + "\"";
      } else if (it.type < 14) {
        out += std::string(" t=\"") + item_type[it.type] + "\"";
      }
      out += "/>";
    }
    out += "</items></pivotField>";
  }
  out += "</pivotFields>";

  for (size_t a = 0; a < 2; ++a) {
    bool has_fields = a < v.axes.size() && !v.axes[a].empty();
    if (has_fields) {
      const char* tag = a == 0 ? "rowFields" : "colFields";
      out += std::string("<") + tag + " count=\"" + std::to_string(v.axes[a].size()) + "\">";
      for (size_t k = 0; k < v.axes[a].size(); ++k) {
        int x = v.axes[a][k] == 0xFFFE ? -2 : static_cast<int>(v.axes[a][k]);
        out += "<field x=\"" + std::to_string(x) + "\"/>";
      }
      out += std::string("</") + tag + ">";
    }

    // the lines of the table: one for each shown item of the first field and the grand
    // total, for the data fields one for each data field. Excel writes them again when
    // the table is refreshed.
    std::string lines;
    size_t n_lines = 0;
    if (!v.lines[a].empty()) {
      static const char* line_type[] = {"data", "default", "sum", "countA", "avg", "max", "min", "product",
                                        "count", "stdDev", "stdDevP", "var", "varP", "grand", "blank"};
      for (const XlsPivotLine& l : v.lines[a]) {
        lines += "<i";
        if (l.reuse > 0) lines += " r=\"" + std::to_string(l.reuse) + "\"";
        if (l.type > 0 && l.type < 15) lines += std::string(" t=\"") + line_type[l.type] + "\"";
        if (l.data > 0) lines += " i=\"" + std::to_string(l.data) + "\"";
        lines += ">";
        for (uint32_t x : l.x) {
          if (x == 0x001000FEu) continue;
          lines += "<x";
          if (x > 0) lines += " v=\"" + std::to_string(x) + "\"";
          lines += "/>";
        }
        lines += "</i>";
      }
      n_lines = v.lines[a].size();
    } else if (!has_fields) {
      if (a == 1) {
        lines = "<i/>";
        n_lines = 1;
      }
    } else if (v.axes[a][0] == 0xFFFE) {
      for (size_t k = 0; k < v.data.size(); ++k) {
        lines += "<i";
        if (k > 0) lines += " i=\"" + std::to_string(k) + "\"";
        lines += "><x";
        if (k > 0) lines += " v=\"" + std::to_string(k) + "\"";
        lines += "/></i>";
        ++n_lines;
      }
    } else if (v.axes[a][0] < v.fields.size()) {
      const XlsPivotField& f = v.fields[v.axes[a][0]];
      for (size_t k = 0; k < f.items.size(); ++k) {
        if (f.items[k].type != 0 || (f.items[k].flags & 0x01)) continue;
        lines += "<i><x";
        if (k > 0) lines += " v=\"" + std::to_string(k) + "\"";
        lines += "/></i>";
        ++n_lines;
      }
      lines += "<i t=\"grand\"><x/></i>";
      ++n_lines;
    }
    if (n_lines > 0) {
      const char* tag = a == 0 ? "rowItems" : "colItems";
      out += std::string("<") + tag + " count=\"" + std::to_string(n_lines) + "\">" + lines + "</" + tag + ">";
    }
  }

  if (n_page > 0) {
    out += "<pageFields count=\"" + std::to_string(n_page) + "\">";
    for (size_t k = 0; k < n_page; ++k) {
      out += "<pageField fld=\"" + std::to_string(v.pages[k].field) + "\"";
      if (v.pages[k].item != 0x7FFD) out += " item=\"" + std::to_string(v.pages[k].item) + "\"";
      out += " hier=\"-1\"/>";
    }
    out += "</pageFields>";
  }

  if (!v.data.empty()) {
    out += "<dataFields count=\"" + std::to_string(v.data.size()) + "\">";
    for (size_t k = 0; k < v.data.size(); ++k) {
      const XlsPivotDataField& d = v.data[k];
      out += "<dataField name=\"" + escape_xml(d.name) + "\" fld=\"" + std::to_string(d.field) + "\"";
      if (d.function > 0 && d.function < 11) out += std::string(" subtotal=\"") + subtotal[d.function] + "\"";
      if (d.display > 0 && d.display < 9) {
        out += std::string(" showDataAs=\"") + show_as[d.display] + "\"";
        out += " baseField=\"" + std::to_string(d.base_field) + "\" baseItem=\"" + std::to_string(d.base_item) + "\"";
      } else {
        out += " baseField=\"0\" baseItem=\"0\"";
      }
      if (d.numfmt != 0) out += " numFmtId=\"" + std::to_string(d.numfmt) + "\"";
      out += "/>";
    }
    out += "</dataFields>";
  }

  if (!v.formats.empty()) {
    static const char* const area_type[7] = {"none", "normal", "data", "all", "origin", "button", "topRight"};
    out += "<formats count=\"" + std::to_string(v.formats.size()) + "\">";
    for (const XlsPivotFormat& f : v.formats) {
      out += "<format dxfId=\"" + std::to_string(f.dxf) + "\"><pivotArea";
      if (f.field >= 0) out += " field=\"" + std::to_string(f.field) + "\"";
      if (f.type < 7) out += std::string(" type=\"") + area_type[f.type] + "\"";
      if (!(f.flags & 0x01)) out += " dataOnly=\"0\"";
      if (f.flags & 0x02) out += " labelOnly=\"1\"";
      if (f.flags & 0x04) out += " grandRow=\"1\"";
      if (f.flags & 0x08) out += " grandCol=\"1\"";
      if (f.flags & 0x10) out += " cacheIndex=\"1\"";
      if (!(f.flags & 0x20)) out += " outline=\"0\"";
      if (f.axis == 1) out += " axis=\"axisRow\"";
      else if (f.axis == 2) out += " axis=\"axisCol\"";
      else if (f.axis == 4) out += " axis=\"axisPage\"";
      else if (f.axis == 8) out += " axis=\"axisData\"";
      out += " fieldPosition=\"" + std::to_string(f.position) + "\"";
      out += "/></format>";
    }
    out += "</formats>";
  }

  out += "<pivotTableStyleInfo name=\"" + escape_xml(v.style) + "\"";
  if (v.style_flags & 0x01) out += " showFirstColumn=\"1\"";
  out += std::string(" showRowHeaders=\"") + ((v.style_flags & 0x10) ? "1" : "0") + "\"";
  out += std::string(" showColHeaders=\"") + ((v.style_flags & 0x20) ? "1" : "0") + "\"";
  out += std::string(" showRowStripes=\"") + ((v.style_flags & 0x04) ? "1" : "0") + "\"";
  out += std::string(" showColStripes=\"") + ((v.style_flags & 0x08) ? "1" : "0") + "\"";
  out += std::string(" showLastColumn=\"") + ((v.style_flags & 0x02) ? "1" : "0") + "\"/>";
  out += "</pivotTableDefinition>\n";
  return out;
}

#endif
