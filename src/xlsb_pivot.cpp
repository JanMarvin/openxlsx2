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
        // flags: 0x1 semi mixed types, 0x2 non date, 0x4 date, 0x8 string, 0x10 blank, 0x40 number,
        // 0x80 integer, 0x100 minimum and maximum. Fields with items get them from the items.
        if (d.size() >= 6 && xls_pv_u32(d, 2) == 0) {
          std::string a;
          if ((fl & 0x01u) == 0) a += " containsSemiMixedTypes=\"0\"";
          if ((fl & 0x02u) == 0) a += " containsNonDate=\"0\"";
          if ((fl & 0x04u) != 0) a += " containsDate=\"1\"";
          if ((fl & 0x08u) == 0) a += " containsString=\"0\"";
          if ((fl & 0x10u) != 0) a += " containsBlank=\"1\"";
          if ((fl & 0x40u) != 0) a += " containsNumber=\"1\"";
          if ((fl & 0x80u) != 0) a += " containsInteger=\"1\"";
          if ((fl & 0x100u) != 0 && d.size() >= 22) {
            uint64_t b1 = static_cast<uint64_t>(xls_pv_u32(d, 6)) | (static_cast<uint64_t>(xls_pv_u32(d, 10)) << 32);
            uint64_t b2 = static_cast<uint64_t>(xls_pv_u32(d, 14)) | (static_cast<uint64_t>(xls_pv_u32(d, 18)) << 32);
            double serial[2] = {0, 0};
            std::memcpy(&serial[0], &b1, sizeof(double));
            std::memcpy(&serial[1], &b2, sizeof(double));
            if ((fl & 0x04u) == 0) {
              a += " minValue=\"" + xls_num(serial[0]) + "\" maxValue=\"" + xls_num(serial[1]) + "\"";
            } else {
              // dates are serial numbers, converted to year, month and day
              std::string iso[2];
              for (int k = 0; k < 2; ++k) {
                double whole = std::floor(serial[k]);
                long z = static_cast<long>(whole) - 25569 + 719468;
                long era = (z >= 0 ? z : z - 146096) / 146097;
                long doe = z - era * 146097;
                long yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
                long doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
                long mp = (5 * doy + 2) / 153;
                long day = doy - (153 * mp + 2) / 5 + 1;
                long month = mp < 10 ? mp + 3 : mp - 9;
                long year = yoe + era * 400 + (month <= 2 ? 1 : 0);
                long secs = std::lround((serial[k] - whole) * 86400.0);
                char buf[128];
                std::snprintf(buf, sizeof(buf), "%04ld-%02ld-%02ldT%02ld:%02ld:%02ld", year, month, day, secs / 3600, (secs / 60) % 60, secs % 60);
                iso[k] = buf;
              }
              a += " minDate=\"" + iso[0] + "\" maxDate=\"" + iso[1] + "\"";
            }
          }
          f.attrs = a;
        }
        break;
      }

      case BrtEndPCDFAtbl:
        in_items = false;
        break;

      case BrtPCDIMissing:
      case BrtPCDINumber:
      case BrtPCDIBoolean:
      case BrtPCDIError:
      case BrtPCDIString:
      case BrtPCDIDatetime:
      case BrtPCDIAMissing:
      case BrtPCDIANumber:
      case BrtPCDIABoolean:
      case BrtPCDIAError:
      case BrtPCDIAString:
      case BrtPCDIADatetime:
      case BrtBeginPCDIRun: {
        if (!in_items || fields.empty()) break;
        XlsPivotCacheField& f = fields.back();
        f.attrs.clear();

        // items are missing 0, number 1, boolean 2, error 3, string 4, datetime 5. The
        // items with attributes carry six more bytes, they are not imported. A run holds
        // many items of the same type.
        int kind = 0;
        size_t n_items = 1, pos = 0;
        if (recs[i].id == BrtBeginPCDIRun) {
          uint32_t run_type = xls_pv_u16(d, 0);
          n_items = xls_pv_u32(d, 2);
          pos = 6;
          if (run_type == 1) kind = 1;
          else if (run_type == 2) kind = 4;
          else if (run_type == 0x20) kind = 5;
          else return -1;
        } else {
          kind = recs[i].id >= BrtPCDIAMissing ? recs[i].id - BrtPCDIAMissing : recs[i].id - BrtPCDIMissing;
        }

        for (size_t k = 0; k < n_items && pos <= d.size(); ++k) {
          if (kind == 4) {
            f.items += "<s v=\"" + escape_xml(pivot_string(d, pos)) + "\"/>";
            ++f.n_str;
          } else if (kind == 1) {
            uint64_t bits = static_cast<uint64_t>(xls_pv_u32(d, pos)) | (static_cast<uint64_t>(xls_pv_u32(d, pos + 4)) << 32);
            pos += 8;
            double x = 0;
            std::memcpy(&x, &bits, sizeof(x));
            if (f.n_num == 0 || x < f.min) f.min = x;
            if (f.n_num == 0 || x > f.max) f.max = x;
            if (x == std::floor(x)) ++f.n_int;
            f.items += "<n v=\"" + xls_num(x) + "\"/>";
            ++f.n_num;
          } else if (kind == 2) {
            f.items += std::string("<b v=\"") + ((pos < d.size() && d[pos] != 0) ? "1" : "0") + "\"/>";
            ++pos;
            ++f.n_other;
          } else if (kind == 3) {
            f.items += "<e v=\"" + xls_error(pos < d.size() ? d[pos] : static_cast<uint8_t>(0)) + "\"/>";
            ++pos;
            ++f.n_other;
          } else if (kind == 5) {
            char buf[32];
            std::snprintf(buf, sizeof(buf), "%04u-%02u-%02uT%02u:%02u:%02u", xls_pv_u16(d, pos), xls_pv_u16(d, pos + 2),
                          pos + 4 < d.size() ? d[pos + 4] : 0u, pos + 5 < d.size() ? d[pos + 5] : 0u,
                          pos + 6 < d.size() ? d[pos + 6] : 0u, pos + 7 < d.size() ? d[pos + 7] : 0u);
            pos += 8;
            std::string dt = buf;
            if (f.n_date == 0 || dt < f.min_date) f.min_date = dt;
            if (f.n_date == 0 || dt > f.max_date) f.max_date = dt;
            f.items += "<d v=\"" + dt + "\"/>";
            ++f.n_date;
          } else {
            f.items += "<m/>";
            ++f.n_blank;
          }
          ++f.count;
        }
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
  int line_axis = -1;
  bool in_format = false;
  v.axes.resize(2);

  for (size_t i = 0; i < recs.size(); ++i) {
    const std::vector<uint8_t>& d = recs[i].data;
    if (debug) Rcpp::Rcout << "pivot table record " << recs[i].id << ": " << d.size() << std::endl;

    switch (recs[i].id) {
      case BrtBeginSXView: {
        cache_id = xls_pv_u32(d, 28);
        if (d.size() > 5) {
          v.row_grand = (d[5] & 0x20) != 0;
          v.col_grand = (d[5] & 0x40) != 0;
        }
        if (d.size() > 15) {
          v.outline = (d[8] & 0x02) != 0;
          v.version = d[14];
        }
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
        f.default_subtotal = d.size() > 1 && (d[1] & 0x01) != 0;
        f.numfmt = xls_pv_u32(d, 4);
        f.multi = d.size() > 10 && (d[10] & 0x08) != 0;
        f.sort = d.size() > 9 ? static_cast<uint8_t>((d[9] & 0x30) == 0x30 ? 1 : ((d[9] & 0x30) == 0x10 ? 2 : 0)) : static_cast<uint8_t>(0);
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

      case BrtBeginSXLIRws:
      case BrtBeginSXLICols:
        line_axis = recs[i].id == BrtBeginSXLIRws ? 0 : 1;
        break;

      case BrtEndSXLIRws:
      case BrtEndSXLICols:
        line_axis = -1;
        break;

      case BrtBeginSXLI:
        if (line_axis >= 0 && d.size() >= 12) {
          XlsPivotLine l;
          l.reuse = static_cast<uint16_t>(xls_pv_u16(d, 0));
          l.type = d[2];
          l.data = xls_pv_u32(d, 8);
          v.lines[line_axis].push_back(l);
        }
        break;

      case BrtBeginISXVIs:
        if (line_axis >= 0 && !v.lines[line_axis].empty()) {
          for (size_t k = 0; k + 4 <= d.size(); k += 4) v.lines[line_axis].back().x.push_back(xls_pv_u32(d, k));
        }
        break;

      case BrtBeginSXFormat:
        if (d.size() >= 6) {
          XlsPivotFormat f;
          f.dxf = xls_pv_u32(d, 2);
          v.formats.push_back(f);
          in_format = true;
        }
        break;

      case BrtEndSXFormat:
        in_format = false;
        break;

      case BrtBeginPRule:
        if (in_format && !v.formats.empty() && d.size() >= 8) {
          XlsPivotFormat& f = v.formats.back();
          f.field = static_cast<int32_t>(xls_pv_u32(d, 0));
          f.type = d[4];
          f.flags = d[5];
          f.axis = static_cast<uint8_t>(xls_pv_u16(d, 6) & 0x0F);
          f.position = static_cast<uint8_t>((xls_pv_u16(d, 6) >> 4) & 0xFF);
          in_format = false;
        }
        break;

      case BrtTableStyleClient:
        if (d.size() >= 6) {
          v.style_flags = static_cast<uint16_t>(xls_pv_u16(d, 0));
          size_t pos = 2;
          std::string style = pivot_string(d, pos);
          if (!style.empty()) v.style = style;
        }
        break;

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
