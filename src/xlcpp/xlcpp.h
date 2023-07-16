#pragma once

#include <filesystem>
#include <chrono>
#include <string>
#include <list>
#include <variant>
#include <span>

// #ifdef _WIN32
//
// #include <windows.h>
//
// #ifdef XLCPP_EXPORT
// #define XLCPP __declspec(dllexport)
// #elif !defined(XLCPP_STATIC)
// #define XLCPP __declspec(dllimport)
// #else
// #define XLCPP
// #endif
//
// #else

#ifdef XLCPP_EXPORT
#define XLCPP __attribute__ ((visibility ("default")))
#elif !defined(XLCPP_STATIC)
#define XLCPP __attribute__ ((dllimport))
#else
#define XLCPP
#endif

// #endif

namespace xlcpp {

class workbook_pimpl;
class sheet;

class XLCPP workbook {
public:
    workbook();
    workbook(const std::filesystem::path& fn, std::string_view password = "", std::string_view outfile = "");
    workbook(std::span<uint8_t> sv, std::string_view password = "", std::string_view outfile = "");
    ~workbook();

    workbook_pimpl* impl;
};


};
