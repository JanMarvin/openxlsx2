#include "openxlsx2.h"

#include "xlcpp.h"
#include "xlcpp-pimpl.h"
#include "cfbf.h"
#include "utf16.h"
#include "xlsb.h"

using namespace std;

namespace xlcpp {

static void xlsb_walk(span<const uint8_t> data, const function<void(enum xlsb_type, span<const uint8_t>)>& func) {
    while (!data.empty()) {
        uint16_t type;
        uint32_t length;

        type = data[0];
        data = data.subspan(1);

        if (type & 0x80) {
            type = (uint16_t)((type & 0x7f) | ((uint16_t)data[0] << 7));
            data = data.subspan(1);
        }

        length = data[0];
        data = data.subspan(1);

        if (length & 0x80) {
            length = (length & 0x7f) | ((uint32_t)data[0] << 7);
            data = data.subspan(1);
        }

        if (length & 0x4000) {
            length = (length & 0x3fff) | ((uint32_t)data[0] << 14);
            data = data.subspan(1);
        }

        if (length & 0x200000) {
            length = (length & 0x1fffff) | ((uint32_t)(data[0] & 0x7f) << 21);
            data = data.subspan(1);
        }

        func((enum xlsb_type)type, data.subspan(0, length));

        data = data.subspan(length);
    }
}

void workbook_pimpl::load_sheet_binary(string_view name, span<const uint8_t> data, bool visible) {
    auto& s = *sheets.emplace(sheets.end(), *this, name, sheets.size() + 1, visible);
    unsigned int last_index = 0, last_col = 0;
    row* row = nullptr;
    bool in_sheet_data = false;

    xlsb_walk(data, [&](enum xlsb_type type, span<const uint8_t> d) {
        switch (type) {
            case xlsb_type::BrtBeginSheetData:
                in_sheet_data = true;
                break;

            case xlsb_type::BrtEndSheetData:
                in_sheet_data = false;
                break;

            case xlsb_type::BrtRowHdr: {
                if (!in_sheet_data)
                    break;

                if (d.size() < sizeof(brt_row_hdr))
                    Rcpp::stop("Malformed BrtRowHdr record.");

                const auto& h = *(brt_row_hdr*)d.data();

                if (h.rw < last_index)
                    Rcpp::stop("Rows out of order.");

                while (last_index < h.rw) {
                    s.impl->rows.emplace(s.impl->rows.end(), *s.impl, s.impl->rows.size() + 1);
                    last_index++;
                }

                s.impl->rows.emplace(s.impl->rows.end(), *s.impl, s.impl->rows.size() + 1);

                row = &s.impl->rows.back();

                last_index = h.rw + 1;
                last_col = 0;

                break;
            }

            case xlsb_type::BrtCellBlank:
            case xlsb_type::BrtCellError:
            case xlsb_type::BrtFmlaError: {
                if (!in_sheet_data)
                    break;

                if (d.size() < sizeof(xlsb_cell))
                    Rcpp::stop("Malformed cell record.");

                const auto& h = *(xlsb_cell*)d.data();

                if (h.column < last_col)
                    Rcpp::stop("Cells out of order.");

                while (last_col < h.column) {
                    row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, nullptr);
                    last_col++;
                }

                last_col = h.column + 1;

                auto number_format = find_number_format(h.iStyleRef);

                row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, nullptr);
                break;
            }

            case xlsb_type::BrtCellRk: {
                if (!in_sheet_data)
                    break;

                if (d.size() < sizeof(brt_cell_rk))
                    Rcpp::stop("Malformed BrtCellRk record.");

                const auto& h = *(brt_cell_rk*)d.data();

                if (h.cell.column < last_col)
                    Rcpp::stop("Cells out of order.");

                while (last_col < h.cell.column) {
                    row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, nullptr);
                    last_col++;
                }

                last_col = h.cell.column + 1;

                auto number_format = find_number_format(h.cell.iStyleRef);

                double d;

                if (h.value.fInt) {
                    auto num = h.value.num;
                    auto val = *reinterpret_cast<int32_t*>(&num);

                    d = val;
                } else {
                    auto num = ((uint64_t)h.value.num) << 34;
                    d = *reinterpret_cast<double*>(&num);
                }

                if (h.value.fx100)
                    d /= 100.0;

                bool dt = is_date(number_format);
                bool tm = is_time(number_format);
                cell* c;

                // FIXME - we can optimize is_date and is_time if one of the preset number formats

                if (dt && tm) {
                    auto n = (unsigned int)((d - (int)d) * 86400.0);
                    datetime dt(1970y, chrono::January, 1d, chrono::seconds{n});

                    dt.d = number_to_date((int)d, date1904);

                    c = &*row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, dt);
                } else if (dt) {
                    auto ymd = number_to_date((unsigned int)d, date1904);

                    c = &*row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, ymd);
                } else if (tm) {
                    auto n = (unsigned int)(d * 86400.0);
                    chrono::seconds t{n % 86400};

                    c = &*row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, t);
                } else
                    c = &*row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, d);

                c->impl->number_format = number_format;

                break;
            }

            case xlsb_type::BrtCellBool:
            case xlsb_type::BrtFmlaBool: {
                if (!in_sheet_data)
                    break;

                if (d.size() < sizeof(brt_cell_bool))
                    Rcpp::stop("Malformed BrtCellBool record.");

                const auto& h = *(brt_cell_bool*)d.data();

                if (h.cell.column < last_col)
                    Rcpp::stop("Cells out of order.");

                while (last_col < h.cell.column) {
                    row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, nullptr);
                    last_col++;
                }

                last_col = h.cell.column + 1;

                auto number_format = find_number_format(h.cell.iStyleRef);

                auto c = &*row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, h.fBool != 0);

                c->impl->number_format = number_format;

                break;
            }

            case xlsb_type::BrtCellReal:
            case xlsb_type::BrtFmlaNum: {
                if (!in_sheet_data)
                    break;

                if (d.size() < sizeof(brt_cell_real))
                    Rcpp::stop("Malformed BrtCellReal record.");

                const auto& h = *(brt_cell_real*)d.data();

                if (h.cell.column < last_col)
                    Rcpp::stop("Cells out of order.");

                while (last_col < h.cell.column) {
                    row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, nullptr);
                    last_col++;
                }

                last_col = h.cell.column + 1;

                auto number_format = find_number_format(h.cell.iStyleRef);

                bool dt = is_date(number_format);
                bool tm = is_time(number_format);
                cell* c;

                if (dt && tm) {
                    auto n = (unsigned int)((h.xnum - (int)h.xnum) * 86400.0);
                    datetime dt(1970y, chrono::January, 1d, chrono::seconds{n});

                    dt.d = number_to_date((int)h.xnum, date1904);

                    c = &*row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, dt);
                } else if (dt) {
                    auto ymd = number_to_date((unsigned int)h.xnum, date1904);

                    c = &*row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, ymd);
                } else if (tm) {
                    auto n = (unsigned int)(h.xnum * 86400.0);
                    chrono::seconds t{n % 86400};

                    c = &*row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, t);
                } else
                    c = &*row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, h.xnum);

                c->impl->number_format = number_format;

                break;
            }

            case xlsb_type::BrtCellSt:
            case xlsb_type::BrtFmlaString: {
                if (!in_sheet_data)
                    break;

                if (d.size() < offsetof(brt_cell_st, str))
                    Rcpp::stop("Malformed BrtCellSt record.");

                const auto& h = *(brt_cell_st*)d.data();

                if (d.size() < offsetof(brt_cell_st, str) + (h.len * sizeof(char16_t)))
                    Rcpp::stop("Malformed BrtCellSt record.");

                if (h.cell.column < last_col)
                    Rcpp::stop("Cells out of order.");

                while (last_col < h.cell.column) {
                    row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, nullptr);
                    last_col++;
                }

                last_col = h.cell.column + 1;

                auto number_format = find_number_format(h.cell.iStyleRef);

                auto u16sv = u16string_view(h.str, h.len);

                auto c = &*row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, nullptr);

                // so we don't have to expose shared_string publicly
                c->impl->val = utf16_to_utf8(u16sv);
                c->impl->number_format = number_format;

                break;
            }

            case xlsb_type::BrtCellIsst: {
                if (!in_sheet_data)
                    break;

                if (d.size() < sizeof(brt_cell_isst))
                    Rcpp::stop("Malformed BrtCellIsst record.");

                const auto& h = *(brt_cell_isst*)d.data();

                if (h.cell.column < last_col)
                    Rcpp::stop("Cells out of order.");

                while (last_col < h.cell.column) {
                    row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, nullptr);
                    last_col++;
                }

                last_col = h.cell.column + 1;

                auto number_format = find_number_format(h.cell.iStyleRef);

                shared_string ss;
                ss.num = h.isst;

                auto c = &*row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, nullptr);

                // so we don't have to expose shared_string publicly
                delete c->impl;
                c->impl = new cell_pimpl(*row->impl, (unsigned int)row->impl->cells.size(), ss);

                c->impl->number_format = number_format;

                break;
            }

            case xlsb_type::BrtCellRString: {
                if (!in_sheet_data)
                    break;

                if (d.size() < offsetof(brt_cell_rstring, value.str))
                    Rcpp::stop("Malformed BrtCellRString record.");

                const auto& h = *(brt_cell_rstring*)d.data();

                if (d.size() < offsetof(brt_cell_rstring, value.str) + (h.value.len * sizeof(char16_t)))
                    Rcpp::stop("Malformed BrtCellRString record.");

                if (h.cell.column < last_col)
                    Rcpp::stop("Cells out of order.");

                while (last_col < h.cell.column) {
                    row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, nullptr);
                    last_col++;
                }

                last_col = h.cell.column + 1;

                auto number_format = find_number_format(h.cell.iStyleRef);

                auto u16sv = u16string_view(h.value.str, h.value.len);

                auto c = &*row->impl->cells.emplace(row->impl->cells.end(), *row->impl, row->impl->cells.size() + 1, nullptr);

                // so we don't have to expose shared_string publicly
                c->impl->val = utf16_to_utf8(u16sv);

                c->impl->number_format = number_format;

                break;
            }

            default:
                break;
        }
    });
}

void workbook_pimpl::parse_workbook_binary(string_view fn, span<const uint8_t> data, const unordered_map<string, file>& files) {
    struct sheet_info {
        sheet_info(string_view rid, string_view name, bool visible) :
            rid(rid), name(name), visible(visible) { }

        string rid;
        string name;
        bool visible;
    };

    vector<sheet_info> sheets_rels;

    xlsb_walk(data, [&](enum xlsb_type type, span<const uint8_t> d) {
        switch (type) {
            case xlsb_type::BrtBundleSh: {
                if (d.size() < sizeof(brt_bundle_sh))
                    Rcpp::stop("Malformed BrtBundleSh entry.");

                auto& h = *(brt_bundle_sh*)d.data();

                d = d.subspan(sizeof(brt_bundle_sh));

                if (d.size() < sizeof(uint32_t))
                    Rcpp::stop("Malformed BrtBundleSh entry.");

                auto relid_size = *(uint32_t*)d.data();

                d = d.subspan(sizeof(uint32_t));

                if (relid_size == 0xffffffff)
                    relid_size = 0;

                if (d.size() < relid_size * sizeof(char16_t))
                    Rcpp::stop("Malformed BrtBundleSh entry.");

                auto strRelID = u16string_view((char16_t*)d.data(), relid_size);

                d = d.subspan(relid_size * sizeof(char16_t));

                if (d.size() < sizeof(uint32_t))
                    Rcpp::stop("Malformed BrtBundleSh entry.");

                auto name_size = *(uint32_t*)d.data();

                d = d.subspan(sizeof(uint32_t));

                if (d.size() < name_size * sizeof(char16_t))
                    Rcpp::stop("Malformed BrtBundleSh entry.");

                auto strName = u16string_view((char16_t*)d.data(), name_size);

                sheets_rels.emplace_back(utf16_to_utf8(strRelID), utf16_to_utf8(strName), h.hsState == 0);

                break;
            }

            case xlsb_type::BrtWbProp: {
                if (d.size() < sizeof(brt_wb_prop))
                    Rcpp::stop("Malformed BrtWbProp entry.");

                const auto& h = *(brt_wb_prop*)d.data();

                date1904 = h.f1904;

                break;
            }

            default:
                break;
        }
    });

    // FIXME - preserve sheet order

    auto rels = read_relationships(fn, files);

    for (const auto& sr : sheets_rels) {
        for (const auto& r : rels) {
            if (r.first == sr.rid) {
                auto name = filesystem::path(fn);

                // FIXME - can we resolve relative paths properly?

                name.remove_filename();
                name /= r.second;

                auto ns = name.string();

                for (auto& c : ns) {
                    if (c == '\\')
                        c = '/';
                }

                if (files.count(ns) == 0)
                    Rcpp::stop("File {} not found.", ns);

                auto& d = files.at(ns).data;

                load_sheet_binary(sr.name, span((uint8_t*)d.data(), d.size()), sr.visible);
                break;
            }
        }
    }
}

void workbook_pimpl::load_shared_strings_binary(span<const uint8_t> data) {
    xlsb_walk(data, [&](enum xlsb_type type, span<const uint8_t> d) {
        if (type == xlsb_type::BrtSSTItem) {
            if (d.size() < offsetof(brt_sst_item, richStr.str))
                Rcpp::stop("Malformed BrtSSTItem record.");

            const auto& h = *(brt_sst_item*)d.data();

            if (d.size() < offsetof(brt_sst_item, richStr.str) + (sizeof(char16_t) * h.richStr.len))
                Rcpp::stop("Malformed BrtSSTItem record.");

            auto u16sv = u16string_view(h.richStr.str, h.richStr.len);

            shared_strings2.emplace_back(utf16_to_utf8(u16sv));
        }
    });
}

void workbook_pimpl::load_styles_binary(span<const uint8_t> data) {
    bool in_numfmts = false, in_cellxfs = false;

    xlsb_walk(data, [&](enum xlsb_type type, span<const uint8_t> d) {
        switch (type) {
            case xlsb_type::BrtBeginCellXFs:
                in_cellxfs = true;
            break;

            case xlsb_type::BrtEndCellXFs:
                in_cellxfs = false;
            break;

            case xlsb_type::BrtBeginFmts:
                in_numfmts = true;
            break;

            case xlsb_type::BrtEndFmts:
                in_numfmts = false;
            break;

            case xlsb_type::BrtXF:
                if (in_cellxfs) {
                    optional<unsigned int> numfmtid;

                    if (d.size() < sizeof(brt_xf))
                        Rcpp::stop("Malformed BrtXF record.");

                    const auto& h = *(brt_xf*)d.data();

                    numfmtid = h.iFmt;

                    if (h.xfGrbitAtr & 1)
                        numfmtid = nullopt;

                    cell_styles.push_back(numfmtid);
                }
            break;

            case xlsb_type::BrtFmt:
                if (in_numfmts) {
                    if (d.size() < offsetof(brt_fmt, stFmtCode))
                        Rcpp::stop("Malformed BrtFmt record.");

                    const auto& h = *(brt_fmt*)d.data();

                    if (d.size() < offsetof(brt_fmt, stFmtCode) + (h.stFmtCode_len * sizeof(char16_t)))
                        Rcpp::stop("Malformed BrtFmt record.");

                    auto u16sv = u16string_view(h.stFmtCode, h.stFmtCode_len);
                    auto s = utf16_to_utf8(u16sv);

                    // FIXME - removing backslashes?

                    number_formats[h.ifmt] = s;
                }
            break;

            default:
                break;
        }
    });
}

};
