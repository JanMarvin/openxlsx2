#include "openxlsx2.h"

#include "xlcpp-pimpl.h"
#include <charconv>

using namespace std;

static bool __inline is_whitespace(char c) {
    return c == ' ' || c == '\t' || c == '\r' || c == '\n';
}

static void parse_attributes(string_view node, const function<bool(string_view, xml_enc_string_view)>& func) {
    auto s = node.substr(1, node.length() - 2);

    if (!s.empty() && s.back() == '/') {
        s.remove_suffix(1);
    }

    while (!s.empty() && !is_whitespace(s.front())) {
        s.remove_prefix(1);
    }

    while (!s.empty() && is_whitespace(s.front())) {
        s.remove_prefix(1);
    }

    while (!s.empty()) {
        auto av = s;

        auto eq = av.find_first_of('=');
        string_view n, v;

        if (eq == string::npos) {
            n = av;
            v = "";
        } else {
            n = av.substr(0, eq);
            v = av.substr(eq + 1);

            while (!v.empty() && is_whitespace(v.front())) {
                v = v.substr(1);
            }

            if (v.length() >= 2 && (v.front() == '"' || v.front() == '\'')) {
                auto c = v.front();

                v.remove_prefix(1);

                auto end = v.find_first_of(c);

                if (end != string::npos) {
                    v = v.substr(0, end);
                    av = av.substr(0, v.data() + v.length() - av.data() + 1);
                } else {
                    for (size_t i = 0; i < av.length(); i++) {
                        if (is_whitespace(av[i])) {
                            av = av.substr(0, i);
                            break;
                        }
                    }
                }
            }
        }

        while (!n.empty() && is_whitespace(n.back())) {
            n.remove_suffix(1);
        }

        if (!func(n, v))
            return;

        s.remove_prefix(av.length());

        while (!s.empty() && is_whitespace(s.front())) {
            s.remove_prefix(1);
        }
    }
}

bool xml_reader::read() {
    if (sv.empty())
        return false;

    // FIXME - DOCTYPE (<!DOCTYPE greeting SYSTEM "hello.dtd">, <!DOCTYPE greeting [ <!ELEMENT greeting (#PCDATA)> ]>)

    if (type == xml_node::element && empty_tag)
        namespaces.pop_back();

    if (sv.front() != '<') { // text
        auto pos = sv.find_first_of('<');

        if (pos == string::npos) {
            node = sv;
            sv = "";
        } else {
            node = sv.substr(0, pos);
            sv = sv.substr(pos);
        }

        type = xml_node::whitespace;

        for (auto c : sv) {
            if (!is_whitespace(c)) {
                type = xml_node::text;
                break;
            }
        }
    } else {
        if (sv.starts_with("<?")) {
            auto pos = sv.find("?>");

            if (pos == string::npos) {
                node = sv;
                sv = "";
            } else {
                node = sv.substr(0, pos + 2);
                sv = sv.substr(pos + 2);
            }

            type = xml_node::processing_instruction;
        } else if (sv.starts_with("</")) {
            auto pos = sv.find_first_of('>');

            if (pos == string::npos) {
                node = sv;
                sv = "";
            } else {
                node = sv.substr(0, pos + 1);
                sv = sv.substr(pos + 1);
            }

            type = xml_node::end_element;
            namespaces.pop_back();
        } else if (sv.starts_with("<!--")) {
            auto pos = sv.find("-->");

            if (pos == string::npos)
                Rcpp::stop("Malformed comment.");

            node = sv.substr(0, pos + 3);
            sv = sv.substr(pos + 3);

            type = xml_node::comment;
        } else if (sv.starts_with("<![CDATA[")) {
            auto pos = sv.find("]]>");

            if (pos == string::npos)
                Rcpp::stop("Malformed CDATA.");

            node = sv.substr(0, pos + 3);
            sv = sv.substr(pos + 3);

            type = xml_node::cdata;
        } else {
            auto pos = sv.find_first_of('>');

            if (pos == string::npos) {
                node = sv;
                sv = "";
            } else {
                node = sv.substr(0, pos + 1);
                sv = sv.substr(pos + 1);
            }

            type = xml_node::element;
            ns_list ns;

            parse_attributes(node, [&](string_view name, xml_enc_string_view value) {
                if (name.starts_with("xmlns:"))
                    ns.emplace_back(name.substr(6), value);
                else if (name == "xmlns")
                    ns.emplace_back("", value);

                return true;
            });

            namespaces.push_back(ns);

            empty_tag = node.ends_with("/>");
        }
    }

    return true;
}

enum xml_node xml_reader::node_type() const {
    return type;
}

bool xml_reader::is_empty() const {
    return type == xml_node::element && empty_tag;
}

void xml_reader::attributes_loop_raw(const function<bool(string_view local_name, xml_enc_string_view namespace_uri_raw,
                                                         xml_enc_string_view value_raw)>& func) const {
    if (type != xml_node::element)
        return;

    parse_attributes(node, [&](string_view name, xml_enc_string_view value_raw) {
        auto colon = name.find_first_of(':');

        if (colon == string::npos)
            return func(name, xml_enc_string_view{}, value_raw);

        auto prefix = name.substr(0, colon);

        for (auto it = namespaces.rbegin(); it != namespaces.rend(); it++) {
            for (const auto& v : *it) {
                if (v.first == prefix)
                    return func(name.substr(colon + 1), v.second, value_raw);
            }
        }

        return func(name.substr(colon + 1), xml_enc_string_view{}, value_raw);
    });
}

optional<xml_enc_string_view> xml_reader::get_attribute(string_view name, string_view ns) const {
    if (type != xml_node::element)
        return nullopt;

    optional<xml_enc_string_view> xesv;

    attributes_loop_raw([&](string_view local_name, xml_enc_string_view namespace_uri_raw,
                            xml_enc_string_view value_raw) {
        if (local_name == name && namespace_uri_raw.cmp(ns)) {
            xesv = value_raw;
            return false;
        }

        return true;
    });

    return xesv;
}

xml_enc_string_view xml_reader::namespace_uri_raw() const {
    auto tag = name();
    auto colon = tag.find_first_of(':');
    string_view prefix;

    if (colon != string::npos)
        prefix = tag.substr(0, colon);

    for (auto it = namespaces.rbegin(); it != namespaces.rend(); it++) {
        for (const auto& v : *it) {
            if (v.first == prefix)
                return v.second;
        }
    }

    return {};
}

string_view xml_reader::name() const {
    if (type != xml_node::element && type != xml_node::end_element)
        return "";

    auto tag = node.substr(type == xml_node::end_element ? 2 : 1);

    tag.remove_suffix(1);

    for (size_t i = 0; i < tag.length(); i++) {
        if (is_whitespace(tag[i])) {
            tag = tag.substr(0, i);
            break;
        }
    }

    return tag;
}

string_view xml_reader::local_name() const {
    if (type != xml_node::element && type != xml_node::end_element)
        return "";

    auto tag = name();
    auto pos = tag.find_first_of(':');

    if (pos == string::npos)
        return tag;
    else
        return tag.substr(pos + 1);
}

string xml_reader::value() const {
    switch (type) {
        case xml_node::text:
            return xml_enc_string_view{node}.decode();

        case xml_node::cdata:
            return string{node.substr(9, node.length() - 12)};

        default:
            return {};
    }
}

static string esc_char(string_view s) {
    uint32_t c = 0;
    from_chars_result fcr;

    if (s.starts_with("x"))
        fcr = from_chars(s.data() + 1, s.data() + s.length(), c, 16);
    else
        fcr = from_chars(s.data(), s.data() + s.length(), c);

    if (c == 0 || c > 0x10ffff)
        return "";

    if (c < 0x80)
        return string{(char)c, 1};
    else if (c < 0x800) {
        char t[2];

        t[0] = (char)(0xc0 | (c >> 6));
        t[1] = (char)(0x80 | (c & 0x3f));

        return string{string_view(t, 2)};
    } else if (c < 0x10000) {
        char t[3];

        t[0] = (char)(0xe0 | (c >> 12));
        t[1] = (char)(0x80 | ((c >> 6) & 0x3f));
        t[2] = (char)(0x80 | (c & 0x3f));

        return string{string_view(t, 3)};
    } else {
        char t[4];

        t[0] = (char)(0xf0 | (c >> 18));
        t[1] = (char)(0x80 | ((c >> 12) & 0x3f));
        t[2] = (char)(0x80 | ((c >> 6) & 0x3f));
        t[3] = (char)(0x80 | (c & 0x3f));

        return string{string_view(t, 4)};
    }
}

string xml_enc_string_view::decode() const {
    auto v = sv;
    string s;

    s.reserve(v.length());

    while (!v.empty()) {
        if (v.front() == '&') {
            v.remove_prefix(1);

            if (v.starts_with("amp;")) {
                s += "&";
                v.remove_prefix(4);
            } else if (v.starts_with("lt;")) {
                s += "<";
                v.remove_prefix(3);
            } else if (v.starts_with("gt;")) {
                s += ">";
                v.remove_prefix(3);
            } else if (v.starts_with("quot;")) {
                s += "\"";
                v.remove_prefix(5);
            } else if (v.starts_with("apos;")) {
                s += "'";
                v.remove_prefix(5);
            } else if (v.starts_with("#")) {
                string_view bit;

                v.remove_prefix(1);

                auto sc = v.find_first_of(';');
                if (sc == string::npos) {
                    bit = v;
                    v = "";
                } else {
                    bit = v.substr(0, sc);
                    v.remove_prefix(sc + 1);
                }

                s += esc_char(bit);
            } else
                s += "&";
        } else {
            s += v.front();
            v.remove_prefix(1);
        }
    }

    return s;
}

bool xml_enc_string_view::cmp(string_view str) const {
    for (auto c : sv) {
        if (c == '&')
            return decode() == str;
    }

    return sv == str;
}
