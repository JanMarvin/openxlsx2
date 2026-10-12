#include <fstream>
#include <iterator>

#include "openxlsx2.h"
#include "xlsb_defines.h"
#include "xlsb_flags.h"
#include "xlsb_funs.h"
#include "xls_defines.h"
#include "xls_funs.h"
#include "xls_pivot.h"

// Pivot tables and pivot caches of xlsb files. The parts are converted the way
// the xls parts are: the cache gets the layout and the items of its fields, the
// records of the cache are not imported and the cache is refreshed when the
// file is opened.

namespace {

struct BinRecord {
  int32_t id = 0;
  std::vector<uint8_t> data;
};

std::vector<BinRecord> pivot_records(const std::string& path) {
  std::vector<BinRecord> out;
  std::ifstream in(path, std::ios::in | std::ios::binary);
  if (!in) return out;
  std::vector<uint8_t> bytes((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());

  size_t pos = 0, n = bytes.size();
  while (pos < n) {
    BinRecord rec;
    uint32_t id = bytes[pos++];
    if ((id & 0x80) != 0) {
      if (pos >= n) break;
      id = (id & 0x7Fu) | ((bytes[pos++] & 0x7Fu) << 7);
    }
    uint32_t size = 0;
    bool done = false;
    for (unsigned k = 0; k < 4 && pos < n; ++k) {
      uint8_t b = bytes[pos++];
      size |= static_cast<uint32_t>(b & 0x7Fu) << (7 * k);
      if ((b & 0x80) == 0) {
        done = true;
        break;
      }
    }
    if (!done || size > n - pos) break;
    rec.id = static_cast<int32_t>(id);
    rec.data.assign(bytes.begin() + static_cast<std::ptrdiff_t>(pos), bytes.begin() + static_cast<std::ptrdiff_t>(pos + size));
    pos += size;
    out.push_back(rec);
  }
  return out;
}

// XLWideString and XLNullableWideString: the number of characters and the characters
std::string pivot_string(const std::vector<uint8_t>& r, size_t& pos) {
  if (pos + 4 > r.size()) {
    pos = r.size();
    return std::string();
  }
  uint32_t len = xls_pv_u32(r, pos);
  pos += 4;
  if (len == 0xFFFFFFFFu) return std::string();
  if (len > (r.size() - pos) / 2) {
    pos = r.size();
    return std::string();
  }
  std::u16string s;
  for (uint32_t i = 0; i < len; ++i) {
    s.push_back(static_cast<char16_t>(xls_pv_u16(r, pos)));
    pos += 2;
  }
  return to_utf8(s);
}

}  // namespace

// [[Rcpp::export]]
int32_t pivotcache_bin(std::string filePath, std::string outPath, bool debug) {
  std::vector<BinRecord> recs = pivot_records(filePath);
  if (recs.empty()) return -1;

  XlsPivotSource src;
  uint32_t src_type = 0, n_records = 0, slicer_id = 0;
  std::vector<XlsPivotCacheField> fields;
  bool in_items = false;

  for (size_t i = 0; i < recs.size(); ++i) {
    const std::vector<uint8_t>& d = recs[i].data;
    if (debug) Rcpp::Rcout << "pivot cache record " << recs[i].id << ": " << d.size() << std::endl;

    switch (recs[i].id) {
      case BrtBeginPivotCacheDef:
        n_records = xls_pv_u32(d, 17);
        break;

      case BrtBeginPCDSource:
        src_type = xls_pv_u32(d, 0);
        break;

      case BrtBeginPCD14:
        slicer_id = xls_pv_u32(d, 5);
        break;

      case BrtBeginPCDSRange: {
        if (d.size() < 3) break;
        bool is_name = (d[0] & 0x01) != 0, has_rel = (d[2] & 0x01) != 0, has_sheet = (d[2] & 0x02) != 0;
        size_t pos = 3;
        if (has_sheet) src.sheet = pivot_string(d, pos);
        if (has_rel) {
          src_type = 1;
          pivot_string(d, pos);
        }
        if (is_name) {
          src.name = pivot_string(d, pos);
        } else {
          src.ref = int_to_col(xls_pv_u32(d, pos + 8) + 1u) + std::to_string(xls_pv_u32(d, pos) + 1u) + ":" +
                    int_to_col(xls_pv_u32(d, pos + 12) + 1u) + std::to_string(xls_pv_u32(d, pos + 4) + 1u);
        }
        break;
      }

      case BrtBeginPCDField: {
        XlsPivotCacheField f;
        size_t pos = 20;
        f.name = pivot_string(d, pos);
        fields.push_back(f);
        break;
      }

      case BrtEndPCDField:
        in_items = false;
        break;

      case BrtBeginPCDFAtbl: {
        in_items = true;
        if (fields.empty()) break;
        XlsPivotCacheField& f = fields.back();
        uint32_t fl = xls_pv_u16(d, 0);
        // only the combinations that are known are written for fields without items
        if ((fl & ~0x0543u) == 0 && (fl & 0x0040u) != 0 && d.size() >= 22) {
          uint64_t b1 = static_cast<uint64_t>(xls_pv_u32(d, 6)) | (static_cast<uint64_t>(xls_pv_u32(d, 10)) << 32);
          uint64_t b2 = static_cast<uint64_t>(xls_pv_u32(d, 14)) | (static_cast<uint64_t>(xls_pv_u32(d, 18)) << 32);
          double lo = 0, hi = 0;
          std::memcpy(&lo, &b1, sizeof(lo));
          std::memcpy(&hi, &b2, sizeof(hi));
          f.attrs = std::string(" containsSemiMixedTypes=\"0\" containsString=\"0\" containsNumber=\"1\" minValue=\"") + xls_num(lo) +
                    "\" maxValue=\"" + xls_num(hi) + "\"";
        }
        break;
      }

      case BrtEndPCDFAtbl:
        in_items = false;
        break;

      case BrtPCDIString:
      case BrtPCDINumber:
      case BrtPCDIBoolean:
      case BrtPCDIError:
      case BrtPCDIMissing:
      case BrtPCDIDatetime: {
        if (!in_items || fields.empty()) break;
        XlsPivotCacheField& f = fields.back();
        f.attrs.clear();
        if (recs[i].id == BrtPCDIString) {
          size_t pos = 0;
          f.items += "<s v=\"" + escape_xml(pivot_string(d, pos)) + "\"/>";
          ++f.n_str;
        } else if (recs[i].id == BrtPCDINumber) {
          uint64_t bits = static_cast<uint64_t>(xls_pv_u32(d, 0)) | (static_cast<uint64_t>(xls_pv_u32(d, 4)) << 32);
          double x = 0;
          std::memcpy(&x, &bits, sizeof(x));
          if (f.n_num == 0 || x < f.min) f.min = x;
          if (f.n_num == 0 || x > f.max) f.max = x;
          if (x == std::floor(x)) ++f.n_int;
          f.items += "<n v=\"" + xls_num(x) + "\"/>";
          ++f.n_num;
        } else if (recs[i].id == BrtPCDIBoolean) {
          f.items += std::string("<b v=\"") + ((!d.empty() && d[0] != 0) ? "1" : "0") + "\"/>";
          ++f.n_other;
        } else if (recs[i].id == BrtPCDIError) {
          f.items += "<e v=\"" + xls_error(d.empty() ? static_cast<uint8_t>(0) : d[0]) + "\"/>";
          ++f.n_other;
        } else {
          f.items += "<m/>";
          ++f.n_blank;
        }
        ++f.count;
        break;
      }

      default:
        break;
    }
  }

  if (src_type != 0 || (src.name.empty() && src.ref.empty()) || fields.empty()) return -1;

  std::ofstream out(outPath);
  out << xls_pivot_cache_write(fields, n_records, src, slicer_id);
  return static_cast<int32_t>(fields.size());
}

// Returns the id of the cache of the table, -1 if the table is not imported
// [[Rcpp::export]]
int32_t pivottable_bin(std::string filePath, std::string outPath, bool debug) {
  std::vector<BinRecord> recs = pivot_records(filePath);
  if (recs.empty()) return -1;

  XlsPivotView v;
  int64_t cache_id = -1;
  v.axes.resize(2);

  for (size_t i = 0; i < recs.size(); ++i) {
    const std::vector<uint8_t>& d = recs[i].data;
    if (debug) Rcpp::Rcout << "pivot table record " << recs[i].id << ": " << d.size() << std::endl;

    switch (recs[i].id) {
      case BrtBeginSXView: {
        cache_id = xls_pv_u32(d, 28);
        size_t pos = 32;
        v.name = pivot_string(d, pos);
        v.data_name = pivot_string(d, pos);
        break;
      }

      case BrtBeginSXLocation:
        v.row1 = static_cast<uint16_t>(xls_pv_u32(d, 0));
        v.row2 = static_cast<uint16_t>(xls_pv_u32(d, 4));
        v.col1 = static_cast<uint16_t>(xls_pv_u32(d, 8));
        v.col2 = static_cast<uint16_t>(xls_pv_u32(d, 12));
        v.head_row = static_cast<uint16_t>(xls_pv_u32(d, 16));
        v.data_row = static_cast<uint16_t>(xls_pv_u32(d, 20));
        v.data_col = static_cast<uint16_t>(xls_pv_u32(d, 24));
        break;

      case BrtBeginSXVD: {
        XlsPivotField f;
        f.axis = static_cast<uint16_t>(xls_pv_u32(d, 0) & 0x0F);
        size_t pos = 16;
        f.name = pivot_string(d, pos);
        v.fields.push_back(f);
        break;
      }

      case BrtBeginSXVI: {
        if (v.fields.empty() || d.size() < 7) break;
        XlsPivotItem it;
        it.type = d[0];
        it.flags = static_cast<uint16_t>(xls_pv_u16(d, 1));
        uint32_t c = xls_pv_u32(d, 3);
        it.cache = c > 0xFFFEu ? static_cast<uint16_t>(0xFFFF) : static_cast<uint16_t>(c);
        if (d.size() > 7) {
          size_t pos = 7;
          it.name = pivot_string(d, pos);
        }
        v.fields.back().items.push_back(it);
        break;
      }

      case BrtBeginISXVDRws:
      case BrtBeginISXVDCols: {
        std::vector<uint16_t>& axis = v.axes[recs[i].id == BrtBeginISXVDRws ? 0 : 1];
        uint32_t n = xls_pv_u32(d, 0);
        for (uint32_t k = 0; k < n && 4 + 4 * static_cast<size_t>(k) + 4 <= d.size(); ++k) {
          uint32_t x = xls_pv_u32(d, 4 + 4 * static_cast<size_t>(k));
          axis.push_back(x == 0xFFFFFFFEu ? static_cast<uint16_t>(0xFFFE) : static_cast<uint16_t>(x));
        }
        break;
      }

      case BrtBeginSXPI: {
        XlsPivotPage p;
        p.field = static_cast<uint16_t>(xls_pv_u32(d, 0));
        uint32_t item = xls_pv_u32(d, 4);
        p.item = item <= 0x00100000u ? static_cast<uint16_t>(item) : static_cast<uint16_t>(0x7FFD);
        v.pages.push_back(p);
        break;
      }

      case BrtBeginSXDI: {
        XlsPivotDataField df;
        df.field = static_cast<uint16_t>(xls_pv_u32(d, 0));
        df.function = static_cast<uint16_t>(xls_pv_u32(d, 4));
        df.display = static_cast<uint16_t>(xls_pv_u32(d, 8));
        df.base_field = static_cast<uint16_t>(xls_pv_u32(d, 12));
        df.base_item = static_cast<uint16_t>(xls_pv_u32(d, 16));
        df.numfmt = static_cast<uint16_t>(xls_pv_u32(d, 20));
        if (d.size() > 25) {
          size_t pos = 25;
          df.name = pivot_string(d, pos);
        }
        v.data.push_back(df);
        break;
      }

      default:
        break;
    }
  }

  if (cache_id < 0 || v.fields.empty()) return -1;

  std::ofstream out(outPath);
  out << xls_pivot_view_xml(v, static_cast<int32_t>(cache_id));
  return static_cast<int32_t>(cache_id);
}

// [[Rcpp::export]]
int32_t slicer_bin(std::string filePath, std::string outPath, bool debug) {
  std::vector<BinRecord> recs = pivot_records(filePath);
  if (recs.empty()) return -1;

  std::string slicers;
  int32_t n = 0;
  for (size_t i = 0; i < recs.size(); ++i) {
    const std::vector<uint8_t>& d = recs[i].data;
    if (debug) Rcpp::Rcout << "slicer record " << recs[i].id << ": " << d.size() << std::endl;
    if (recs[i].id != BrtBeginSlicer || d.size() < 17) continue;

    uint8_t flags = d[0];
    uint32_t start = xls_pv_u32(d, 1), columns = xls_pv_u32(d, 5), level = xls_pv_u32(d, 9), height = xls_pv_u32(d, 13);
    size_t pos = 17;
    std::string name = pivot_string(d, pos);
    std::string cache = pivot_string(d, pos);
    std::string caption, style;
    if ((flags & 0x02) != 0) caption = pivot_string(d, pos);
    if ((flags & 0x04) != 0) style = pivot_string(d, pos);

    slicers += "<slicer name=\"" + escape_xml(name) + "\" cache=\"" + escape_xml(cache) + "\"";
    if ((flags & 0x02) != 0) slicers += " caption=\"" + escape_xml(caption) + "\"";
    if ((flags & 0x01) == 0) slicers += " showCaption=\"0\"";
    if (start != 0) slicers += " startItem=\"" + std::to_string(start) + "\"";
    if (columns != 1) slicers += " columnCount=\"" + std::to_string(columns) + "\"";
    if (level != 0) slicers += " level=\"" + std::to_string(level) + "\"";
    if ((flags & 0x04) != 0) slicers += " style=\"" + escape_xml(style) + "\"";
    if ((flags & 0x08) != 0) slicers += " lockedPosition=\"1\"";
    slicers += " rowHeight=\"" + std::to_string(height) + "\"/>";
    ++n;
  }
  if (n == 0) return -1;

  std::ofstream out(outPath);
  out << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      << "<slicers xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" "
         "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x\" "
         "xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      << slicers << "</slicers>\n";
  return n;
}

// [[Rcpp::export]]
int32_t slicercache_bin(std::string filePath, std::string outPath, bool debug) {
  std::vector<BinRecord> recs = pivot_records(filePath);
  if (recs.empty()) return -1;

  std::string name, source, tables, items;
  uint32_t pivot_cache_id = 0;
  size_t n_items = 0;
  bool have_def = false;

  for (size_t i = 0; i < recs.size(); ++i) {
    const std::vector<uint8_t>& d = recs[i].data;
    if (debug) Rcpp::Rcout << "slicer cache record " << recs[i].id << ": " << d.size() << std::endl;

    switch (recs[i].id) {
      case BrtBeginSlicerCacheDef: {
        size_t pos = 0;
        name = pivot_string(d, pos);
        source = pivot_string(d, pos);
        have_def = true;
        break;
      }

      case BrtSlicerCachePivotTables: {
        size_t pos = 4;
        for (uint32_t k = 0, n = xls_pv_u32(d, 0); k < n && pos + 8 <= d.size(); ++k) {
          uint32_t tab = xls_pv_u32(d, pos);
          pos += 4;
          std::string pt = pivot_string(d, pos);
          tables += "<pivotTable tabId=\"" + std::to_string(tab) + "\" name=\"" + escape_xml(pt) + "\"/>";
        }
        break;
      }

      case BrtBeginSlicerCacheNative:
        pivot_cache_id = xls_pv_u32(d, 4);
        break;

      case BrtSlicerCacheNativeItem: {
        for (uint32_t k = 0, n = xls_pv_u32(d, 0); k < n && 4 + 5 * static_cast<size_t>(k) + 5 <= d.size(); ++k) {
          size_t at = 4 + 5 * static_cast<size_t>(k);
          items += "<i x=\"" + std::to_string(xls_pv_u32(d, at)) + "\"" + ((d[at + 4] & 0x01) != 0 ? " s=\"1\"" : "") + "/>";
          ++n_items;
        }
        break;
      }

      default:
        break;
    }
  }
  if (!have_def || pivot_cache_id == 0) return -1;

  std::ofstream out(outPath);
  out << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      << "<slicerCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" "
         "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x\" "
         "xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\""
      << escape_xml(name) << "\" sourceName=\"" << escape_xml(source) << "\">";
  if (!tables.empty()) out << "<pivotTables>" << tables << "</pivotTables>";
  out << "<data><tabular pivotCacheId=\"" << pivot_cache_id << "\"><items count=\"" << n_items << "\">" << items
      << "</items></tabular></data></slicerCacheDefinition>\n";
  return 1;
}
