#include "openxlsx2.h"
#include <iostream>
#include <fstream>

#include "xlcpp.h"
#include "xlcpp-pimpl.h"
#include "cfbf.h"
#include "utf16.h"
#include <vector>
#include <array>
#include <charconv>

// #ifdef _WIN32
// #include <windows.h>
// #endif

// #define FMT_HEADER_ONLY
// #include <fmt/format.h>
// #include <fmt/compile.h>

#define BLOCK_SIZE 20480

using namespace std;

static const string NS_SPREADSHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
static const string NS_SPREADSHEET_STRICT = "http://purl.oclc.org/ooxml/spreadsheetml/main";
static const string NS_RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
static const string NS_RELATIONSHIPS_STRICT = "http://purl.oclc.org/ooxml/officeDocument/relationships";
static const string NS_PACKAGE_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";
static const string NS_CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types";

#define NUMFMT_OFFSET 165

// static string try_decode(const optional<xml_enc_string_view>& sv) {
//     if (!sv)
//         return "";
//
//     return sv.value().decode();
// }

namespace xlcpp {

/* needed??? */
#ifdef _WIN32
__inline string utf16_to_utf8(const u16string_view& s) {
    string ret;

    if (s.empty())
        return "";

    auto len = WideCharToMultiByte(CP_UTF8, 0, (const wchar_t*)s.data(), (int)s.length(), nullptr, 0,
                                   nullptr, nullptr);

    if (len == 0)
        Rcpp::stop("WideCharToMultiByte 1 failed.");

    ret.resize(len);

    len = WideCharToMultiByte(CP_UTF8, 0, (const wchar_t*)s.data(), (int)s.length(), ret.data(), len,
                              nullptr, nullptr);

    if (len == 0)
        Rcpp::stop("WideCharToMultiByte 2 failed.");

    return ret;
}
#endif

std::vector<uint8_t> loadFile(const std::string& filename) {
  std::ifstream file(filename, std::ios::binary | std::ios::ate);
  if (!file.is_open()) {
    Rcpp::stop("Failed to open file");
  }

  std::streamsize fileSize = file.tellg();
  file.seekg(0, std::ios::beg);

  std::vector<uint8_t> buffer(fileSize);
  if (file.read(reinterpret_cast<char*>(buffer.data()), fileSize)) {
    return buffer;
  } else {
    Rcpp::stop("Failed to read file");
  }
}


workbook_pimpl::workbook_pimpl(const filesystem::path& fn, string_view password, string_view outfile) {

    std::string path = fn;

    std::vector<uint8_t> mem = loadFile(path);

    load_from_memory(mem, password, outfile);
}

workbook_pimpl::workbook_pimpl(span<uint8_t> sv, string_view password, string_view outfile) {
    load_from_memory(sv, password, outfile);
}

void workbook_pimpl::load_from_memory(span<uint8_t> mem, string_view password, string_view outfile) {
    vector<uint8_t> plaintext;

    std::ofstream xlsx((std::string)outfile, ios::out | ios::binary);
    if (mem.size() >= sizeof(uint64_t) && *(uint64_t*)mem.data() == CFBF_SIGNATURE) {
        cfbf c(mem);
        string enc_info, enc_package;

        // FIXME - handle old-style Excel files

        for (unsigned int num = 0; const auto& e : c.entries) {
            if (num == 0) { // root
                num++;
                continue;
            }

            if (e.name == "/EncryptionInfo" || e.name == "/EncryptedPackage") {
                auto& str = e.name == "/EncryptionInfo" ? enc_info : enc_package;

                str.resize(e.get_size());

                uint64_t off = 0;
                auto buf = span((std::byte*)str.data(), str.size());

                while (true) {
                    auto size = e.read(buf, off);

                    if (size == 0)
                        break;

                    off += size;
                }
            }

            num++;
        }

        if (enc_info.empty())
            Rcpp::stop("EncryptionInfo not found.");

        auto u16password = utf8_to_utf16(password);

        c.parse_enc_info(span((uint8_t*)enc_info.data(), enc_info.size()), u16password);
        plaintext = c.decrypt(span((uint8_t*)enc_package.data(), enc_package.size()));

        xlsx.write((char *) plaintext.data(), plaintext.size());
    }
    xlsx.close();

}

workbook::workbook(const filesystem::path& fn, std::string_view password, std::string_view outfile) {
    impl = new workbook_pimpl(fn, password, outfile);
}

workbook::workbook(span<uint8_t> sv, std::string_view password, std::string_view outfile) {
    impl = new workbook_pimpl(sv, password, outfile);
}

workbook::~workbook() {
    delete impl;
}

}
