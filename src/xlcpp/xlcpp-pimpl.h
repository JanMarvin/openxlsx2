#pragma once

#include "xlcpp.h"
#include <list>
#include <variant>
#include <map>
#include <unordered_map>
#include <unordered_set>
#include <functional>
#include <optional>
#include <stack>
#include <span>

namespace xlcpp {

struct shared_string {
    unsigned int num;
};

typedef struct {
    std::string content_type;
    std::string data;
} file;


class workbook_pimpl {
public:
    workbook_pimpl() = default;
    workbook_pimpl(std::string& fn, std::string_view password, std::string_view outfile);
    workbook_pimpl(std::span<uint8_t> sv, std::string_view password, std::string_view outfile);
    std::string data() const;

    // void write_archive(struct archive* a) const;
    // la_ssize_t write_callback(struct archive* a, const void* buffer, size_t length) const;
    void load_archive(struct archive* a);

private:
    void load_from_memory(std::span<uint8_t> sv, std::string_view password, std::string_view outfile);
};

}; // end namespace xlcpp


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
