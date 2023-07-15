#pragma once

#include "xlcpp.h"
#include "mmap.h"
#include <list>
#include <variant>
#include <map>
#include <unordered_map>
#include <unordered_set>
#include <functional>
#include <optional>
#include <stack>
#include <span>
#include <archive.h>
#include <archive_entry.h>
#include <fmt/format.h>

namespace xlcpp {

struct shared_string {
    unsigned int num;
};

typedef struct {
    std::string content_type;
    std::string data;
} file;

class font {
public:
    font(std::string_view font_name, unsigned int font_size, bool bold) : font_name(font_name), font_size(font_size), bold(bold) { }

    std::string font_name;
    unsigned int font_size;
    bool bold;
};

class font_hash {
public:
    size_t operator()(const font& f) const {
        return std::hash<std::string>{}(f.font_name) |
        (std::hash<unsigned int>{}(f.font_size) << 1) |
        (std::hash<bool>{}(f.bold) << 2);
    }
};

bool operator==(const font& lhs, const font& rhs) noexcept;

class style {
public:
    style(std::string_view number_format, std::string_view font, unsigned int font_size, bool bold = false) :
        number_format(number_format), font(font, font_size, bold) { }

    void set_font(std::string_view font_name, unsigned int font_size, bool bold);
    void set_number_format(std::string_view fmt);

    std::string number_format;
    xlcpp::font font;

    mutable unsigned int num;
    mutable unsigned int number_format_num;
};

class style_hash {
public:
    size_t operator()(const style& s) const {
        return std::hash<std::string>{}(s.number_format) |
        (font_hash{}(s.font) << 1);
    }
};

bool operator==(const style& lhs, const style& rhs) noexcept;

class workbook_pimpl {
public:
    workbook_pimpl() = default;
    workbook_pimpl(const std::filesystem::path& fn, std::string_view password, std::string_view outfile);
    workbook_pimpl(std::span<const uint8_t> sv, std::string_view password, std::string_view outfile);
    sheet& add_sheet(std::string_view name, bool visible);
    void save(const std::filesystem::path& fn) const;
    std::string data() const;

    void write_workbook_xml(struct archive* a) const;
    void write_content_types_xml(struct archive* a) const;
    void write_rels(struct archive* a) const;
    void write_workbook_rels(struct archive* a) const;
    shared_string get_shared_string(std::string_view s);
    void write_shared_strings(struct archive* a) const;
    void write_styles(struct archive* a) const;
    void write_archive(struct archive* a) const;
    la_ssize_t write_callback(struct archive* a, const void* buffer, size_t length) const;
    void parse_workbook(std::string_view fn, std::string_view data,
                        const std::unordered_map<std::string, file>& files);
    void parse_workbook_binary(std::string_view fn, std::span<const uint8_t> data,
                               const std::unordered_map<std::string, file>& files);
    void load_sheet(std::string_view name, std::string_view data, bool visible);
    void load_sheet_binary(std::string_view name, std::span<const uint8_t> data, bool visible);
    void load_shared_strings2(std::string_view sv);
    void load_shared_strings_binary(std::span<const uint8_t> data);
    void load_shared_strings(const std::unordered_map<std::string, file>& files);
    void load_styles(const std::unordered_map<std::string, file>& files);
    void load_styles2(std::string_view sv);
    void load_styles_binary(std::span<const uint8_t> data);
    std::string find_number_format(unsigned int num);
    void load_archive(struct archive* a);

#ifdef _WIN32
    void rename(const std::filesystem::path& fn) const;
#endif

    template<class... Args>
    const style* find_style(Args&&... args) {
        auto ret = styles.emplace(args...);

        if (ret.second)
            ret.first->num = (unsigned int)(styles.size() - 1);

        return &(*ret.first);
    }

    std::list<sheet> sheets;
    std::map<std::string, shared_string, std::less<>> shared_strings;
    std::vector<std::string> shared_strings2;
    std::unordered_set<style, style_hash> styles;
    std::unordered_map<unsigned int, std::string> number_formats;
    std::vector<std::optional<unsigned int>> cell_styles;
    bool date1904 = false;

    mutable std::string buf;

#ifdef _WIN32
    unique_handle h;
    HANDLE h2;
    uint8_t readbuf[1048576];
#endif

private:
    void load_from_memory(std::span<const uint8_t> sv, std::string_view password, std::string_view outfile);
};

class sheet_pimpl {
public:
    sheet_pimpl(workbook_pimpl& wb, std::string_view name, unsigned int num, bool visible) : parent(wb), name(name), num(num), visible(visible) { }

    void write(struct archive* a) const;
    std::string xml() const;
    row& add_row();

    workbook_pimpl& parent;
    std::string name;
    unsigned int num;
    bool visible;
    std::list<row> rows;
};

class row_pimpl {
public:
    row_pimpl(sheet_pimpl& s, unsigned int num) : parent(s), num(num) { }

    template<typename T>
    cell& add_cell(const T& val) {
        return *cells.emplace(cells.end(), *this, cells.size() + 1, val);
    }

    sheet_pimpl& parent;
    unsigned int num;
    std::list<cell> cells;
};

class cell_pimpl {
public:
    template<typename T>
    cell_pimpl(row_pimpl& r, unsigned int num, const T& t);

    cell_pimpl(row_pimpl& r, unsigned int num, std::string_view t);
    cell_pimpl(row_pimpl& r, unsigned int num, const std::chrono::system_clock::time_point& val) : cell_pimpl(r, num, datetime{val}) { }

    void set_number_format(std::string_view fmt);
    void set_font(std::string_view name, unsigned int size, bool bold = false);

    row_pimpl& parent;

    const style* sty;

    unsigned int num;
    std::variant<int64_t, shared_string, double, std::chrono::year_month_day, std::chrono::seconds, datetime, bool, std::nullptr_t, std::string> val;
    std::string number_format;
};

};

class xml_writer {
public:
    std::string dump() const;
    void start_document();
    void start_element(std::string_view tag, const std::unordered_map<std::string, std::string>& namespaces = {});
    void end_element();
    void text(std::string_view s);
    void attribute(std::string_view name, std::string_view value);

private:
    std::string buf;
    std::stack<std::string> tags;
    bool empty_tag;
};

enum class xml_node {
    unknown,
    text,
    whitespace,
    element,
    end_element,
    processing_instruction,
    comment,
    cdata
};

template<>
struct fmt::formatter<enum xml_node> {
    constexpr auto parse(format_parse_context& ctx) {
        auto it = ctx.begin();

        if (it != ctx.end() && *it != '}')
            throw format_error("invalid format");

        return it;
    }

    template<typename format_context>
    auto format(enum xml_node n, format_context& ctx) const {
        switch (n) {
            case xml_node::unknown:
                return fmt::format_to(ctx.out(), "unknown");
            case xml_node::text:
                return fmt::format_to(ctx.out(), "text");
            case xml_node::whitespace:
                return fmt::format_to(ctx.out(), "whitespace");
            case xml_node::element:
                return fmt::format_to(ctx.out(), "element");
            case xml_node::end_element:
                return fmt::format_to(ctx.out(), "end_element");
            case xml_node::processing_instruction:
                return fmt::format_to(ctx.out(), "processing_instruction");
            case xml_node::comment:
                return fmt::format_to(ctx.out(), "comment");
            case xml_node::cdata:
                return fmt::format_to(ctx.out(), "cdata");
            default:
                return fmt::format_to(ctx.out(), "{:x}", (unsigned int)n);
        }
    }
};

class xml_enc_string_view {
public:
    xml_enc_string_view() { }
    xml_enc_string_view(std::string_view sv) : sv(sv) { }

    bool empty() const noexcept {
        return sv.empty();
    }

    std::string decode() const;
    bool cmp(std::string_view str) const;

private:
    std::string_view sv;
};

using ns_list = std::vector<std::pair<std::string_view, xml_enc_string_view>>;

class xml_reader {
public:
    xml_reader(std::string_view sv) : sv(sv) { }
    bool read();
    enum xml_node node_type() const;
    bool is_empty() const;
    void attributes_loop_raw(const std::function<bool(std::string_view local_name, xml_enc_string_view namespace_uri_raw,
                                                      xml_enc_string_view value_raw)>& func) const;
    std::optional<xml_enc_string_view> get_attribute(std::string_view name, std::string_view ns = "") const;
    xml_enc_string_view namespace_uri_raw() const;
    std::string_view name() const;
    std::string_view local_name() const;
    std::string value() const;

private:
    std::string_view sv, node;
    enum xml_node type = xml_node::unknown;
    bool empty_tag;
    std::vector<ns_list> namespaces;
};

class archive_read_closer {
public:
    typedef archive* pointer;

    void operator()(archive* a) {
        archive_read_free(a);
    }
};

using archive_read_t = std::unique_ptr<archive*, archive_read_closer>;

class archive_write_closer {
public:
    typedef archive* pointer;

    void operator()(archive* a) {
        archive_write_free(a);
    }
};

using archive_write_t = std::unique_ptr<archive*, archive_write_closer>;


static constexpr std::chrono::year_month_day number_to_date(unsigned int num, bool date1904) noexcept {
    unsigned int J = num + 2415019;
    unsigned int f, e, g, h;
    unsigned int day, month;
    int year;

    if (date1904)
        J += 1462;
    else if (num < 61) // Excel's 29/2/1900 bug
        J++;

    f = J;
    f *= 4;
    f += 274277;
    f /= 146097;
    f *= 3;
    f /= 4;
    f += J;
    f += 1363;

    e = (f * 4) + 3;

    g = e % 1461;
    g /= 4;

    h = (5 * g) + 2;

    day = h % 153;
    day /= 5;
    day++;

    month = h;
    month /= 153;
    month += 2;
    month %= 12;
    month++;

    year = 14 - month;
    year /= 12;
    year -= 4716;
    year += e / 1461;

    return std::chrono::year_month_day{std::chrono::year{year}, std::chrono::month{month}, std::chrono::day{day}};
}

// xlcpp.cpp
bool is_date(std::string_view sv);
bool is_time(std::string_view sv);
std::unordered_map<std::string, std::string> read_relationships(std::string_view fn, const std::unordered_map<std::string, xlcpp::file>& files);
