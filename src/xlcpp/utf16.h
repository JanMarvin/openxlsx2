#pragma once

#include <string_view>

template<typename T>
requires (std::is_same_v<T, char16_t> || (sizeof(wchar_t) == 2 && std::is_same_v<T, wchar_t>))
static constexpr size_t utf16_to_utf8_len(std::basic_string_view<T> sv) noexcept {
    size_t ret = 0;

    while (!sv.empty()) {
        if (sv[0] < 0x80)
            ret++;
        else if (sv[0] < 0x800)
            ret += 2;
        else if (sv[0] < 0xd800)
            ret += 3;
        else if (sv[0] < 0xdc00) {
            if (sv.length() < 2 || (sv[1] & 0xdc00) != 0xdc00) {
                ret += 3;
                sv = sv.substr(1);
                continue;
            }

            ret += 4;
            sv = sv.substr(1);
        } else
            ret += 3;

        sv = sv.substr(1);
    }

    return ret;
}

template<typename T, size_t N>
requires (std::is_same_v<T, char16_t> || (sizeof(wchar_t) == 2 && std::is_same_v<T, wchar_t>))
static constexpr size_t utf16_to_utf8_len(const T (&str)[N]) noexcept {
    return utf16_to_utf8_len(std::basic_string_view<T>{str, N - 1});
}

template<typename T, typename U>
requires (std::is_same_v<T, char16_t> || (sizeof(wchar_t) == 2 && std::is_same_v<T, wchar_t>)) &&
((std::ranges::output_range<U, char> && std::is_same_v<std::ranges::range_value_t<U>, char>) ||
(std::ranges::output_range<U, char8_t> && std::is_same_v<std::ranges::range_value_t<U>, char8_t>))
static constexpr void utf16_to_utf8_range(std::basic_string_view<T> sv, U& t) noexcept {
    auto ptr = t.begin();

    if (ptr == t.end())
        return;

    while (!sv.empty()) {
        if (sv[0] < 0x80) {
            *ptr = (uint8_t)sv[0];
            ptr++;

            if (ptr == t.end())
                return;
        } else if (sv[0] < 0x800) {
            *ptr = (uint8_t)(0xc0 | (sv[0] >> 6));
            ptr++;

            if (ptr == t.end())
                return;

            *ptr = (uint8_t)(0x80 | (sv[0] & 0x3f));
            ptr++;

            if (ptr == t.end())
                return;
        } else if (sv[0] < 0xd800) {
            *ptr = (uint8_t)(0xe0 | (sv[0] >> 12));
            ptr++;

            if (ptr == t.end())
                return;

            *ptr = (uint8_t)(0x80 | ((sv[0] >> 6) & 0x3f));
            ptr++;

            if (ptr == t.end())
                return;

            *ptr = (uint8_t)(0x80 | (sv[0] & 0x3f));
            ptr++;

            if (ptr == t.end())
                return;
        } else if (sv[0] < 0xdc00) {
            if (sv.length() < 2 || (sv[1] & 0xdc00) != 0xdc00) {
                *ptr = (uint8_t)0xef;
                ptr++;

                if (ptr == t.end())
                    return;

                *ptr = (uint8_t)0xbf;
                ptr++;

                if (ptr == t.end())
                    return;

                *ptr = (uint8_t)0xbd;
                ptr++;

                if (ptr == t.end())
                    return;

                sv = sv.substr(1);
                continue;
            }

            char32_t cp = 0x10000 | ((sv[0] & ~0xd800) << 10) | (sv[1] & ~0xdc00);

            *ptr = (uint8_t)(0xf0 | (cp >> 18));
            ptr++;

            if (ptr == t.end())
                return;

            *ptr = (uint8_t)(0x80 | ((cp >> 12) & 0x3f));
            ptr++;

            if (ptr == t.end())
                return;

            *ptr = (uint8_t)(0x80 | ((cp >> 6) & 0x3f));
            ptr++;

            if (ptr == t.end())
                return;

            *ptr = (uint8_t)(0x80 | (cp & 0x3f));
            ptr++;

            if (ptr == t.end())
                return;

            sv = sv.substr(1);
        } else if (sv[0] < 0xe000) {
            *ptr = (uint8_t)0xef;
            ptr++;

            if (ptr == t.end())
                return;

            *ptr = (uint8_t)0xbf;
            ptr++;

            if (ptr == t.end())
                return;

            *ptr = (uint8_t)0xbd;
            ptr++;

            if (ptr == t.end())
                return;
        } else {
            *ptr = (uint8_t)(0xe0 | (sv[0] >> 12));
            ptr++;

            if (ptr == t.end())
                return;

            *ptr = (uint8_t)(0x80 | ((sv[0] >> 6) & 0x3f));
            ptr++;

            if (ptr == t.end())
                return;

            *ptr = (uint8_t)(0x80 | (sv[0] & 0x3f));
            ptr++;

            if (ptr == t.end())
                return;
        }

        sv = sv.substr(1);
    }
}

static std::string __inline utf16_to_utf8(std::u16string_view sv) {
    if (sv.empty())
        return "";

    std::string ret(utf16_to_utf8_len(sv), 0);

    utf16_to_utf8_range(sv, ret);

    return ret;
}

static constexpr size_t utf8_to_utf16_len(std::string_view sv) noexcept {
    size_t ret = 0;

    while (!sv.empty()) {
        if ((uint8_t)sv[0] < 0x80) {
            ret++;
            sv = sv.substr(1);
        } else if (((uint8_t)sv[0] & 0xe0) == 0xc0 && (uint8_t)sv.length() >= 2 && ((uint8_t)sv[1] & 0xc0) == 0x80) {
            ret++;
            sv = sv.substr(2);
        } else if (((uint8_t)sv[0] & 0xf0) == 0xe0 && (uint8_t)sv.length() >= 3 && ((uint8_t)sv[1] & 0xc0) == 0x80 && ((uint8_t)sv[2] & 0xc0) == 0x80) {
            ret++;
            sv = sv.substr(3);
        } else if (((uint8_t)sv[0] & 0xf8) == 0xf0 && (uint8_t)sv.length() >= 4 && ((uint8_t)sv[1] & 0xc0) == 0x80 && ((uint8_t)sv[2] & 0xc0) == 0x80 && ((uint8_t)sv[3] & 0xc0) == 0x80) {
            char32_t cp = (char32_t)(((uint8_t)sv[0] & 0x7) << 18) | (char32_t)(((uint8_t)sv[1] & 0x3f) << 12) | (char32_t)(((uint8_t)sv[2] & 0x3f) << 6) | (char32_t)((uint8_t)sv[3] & 0x3f);

            if (cp > 0x10ffff) {
                ret++;
                sv = sv.substr(4);
                continue;
            }

            ret += 2;
            sv = sv.substr(4);
        } else {
            ret++;
            sv = sv.substr(1);
        }
    }

    return ret;
}

static constexpr size_t utf8_to_utf16_len(std::u8string_view sv) noexcept {
    return utf8_to_utf16_len(std::string_view(std::bit_cast<char*>(sv.data()), sv.length()));
}

template<typename T>
requires (std::ranges::output_range<T, char16_t> && std::is_same_v<std::ranges::range_value_t<T>, char16_t>) ||
    (sizeof(wchar_t) == 2 && std::ranges::output_range<T, wchar_t> && std::is_same_v<std::ranges::range_value_t<T>, wchar_t>)
static constexpr void utf8_to_utf16_range(std::string_view sv, T& t) noexcept {
    auto ptr = t.begin();

    if (ptr == t.end())
        return;

    while (!sv.empty()) {
        if ((uint8_t)sv[0] < 0x80) {
            *ptr = (uint8_t)sv[0];
            ptr++;

            if (ptr == t.end())
                return;

            sv = sv.substr(1);
        } else if (((uint8_t)sv[0] & 0xe0) == 0xc0 && (uint8_t)sv.length() >= 2 && ((uint8_t)sv[1] & 0xc0) == 0x80) {
            char16_t cp = (char16_t)(((uint8_t)sv[0] & 0x1f) << 6) | (char16_t)((uint8_t)sv[1] & 0x3f);

            *ptr = cp;
            ptr++;

            if (ptr == t.end())
                return;

            sv = sv.substr(2);
        } else if (((uint8_t)sv[0] & 0xf0) == 0xe0 && (uint8_t)sv.length() >= 3 && ((uint8_t)sv[1] & 0xc0) == 0x80 && ((uint8_t)sv[2] & 0xc0) == 0x80) {
            char16_t cp = (char16_t)(((uint8_t)sv[0] & 0xf) << 12) | (char16_t)(((uint8_t)sv[1] & 0x3f) << 6) | (char16_t)((uint8_t)sv[2] & 0x3f);

            if (cp >= 0xd800 && cp <= 0xdfff) {
                *ptr = 0xfffd;
                ptr++;

                if (ptr == t.end())
                    return;

                sv = sv.substr(3);
                continue;
            }

            *ptr = cp;
            ptr++;

            if (ptr == t.end())
                return;

            sv = sv.substr(3);
        } else if (((uint8_t)sv[0] & 0xf8) == 0xf0 && (uint8_t)sv.length() >= 4 && ((uint8_t)sv[1] & 0xc0) == 0x80 && ((uint8_t)sv[2] & 0xc0) == 0x80 && ((uint8_t)sv[3] & 0xc0) == 0x80) {
            char32_t cp = (char32_t)(((uint8_t)sv[0] & 0x7) << 18) | (char32_t)(((uint8_t)sv[1] & 0x3f) << 12) | (char32_t)(((uint8_t)sv[2] & 0x3f) << 6) | (char32_t)((uint8_t)sv[3] & 0x3f);

            if (cp > 0x10ffff) {
                *ptr = 0xfffd;
                ptr++;

                if (ptr == t.end())
                    return;

                sv = sv.substr(4);
                continue;
            }

            cp -= 0x10000;

            *ptr = (char16_t)(0xd800 | (cp >> 10));
            ptr++;

            if (ptr == t.end())
                return;

            *ptr = (char16_t)(0xdc00 | (cp & 0x3ff));
            ptr++;

            if (ptr == t.end())
                return;

            sv = sv.substr(4);
        } else {
            *ptr = 0xfffd;
            ptr++;

            if (ptr == t.end())
                return;

            sv = sv.substr(1);
        }
    }
}

template<typename T>
requires (std::ranges::output_range<T, char16_t> && std::is_same_v<std::ranges::range_value_t<T>, char16_t>) ||
    (sizeof(wchar_t) == 2 && std::ranges::output_range<T, wchar_t> && std::is_same_v<std::ranges::range_value_t<T>, wchar_t>)
static constexpr void utf8_to_utf16_range(std::u8string_view sv, T& t) noexcept {
    utf8_to_utf16_range(std::string_view((char*)sv.data(), sv.length()), t);
}

static __inline std::u16string utf8_to_utf16(std::string_view sv) {
    if (sv.empty())
        return u"";

    std::u16string ret(utf8_to_utf16_len(sv), 0);

    utf8_to_utf16_range(sv, ret);

    return ret;
}

static __inline std::u16string utf8_to_utf16(std::u8string_view sv) {
    if (sv.empty())
        return u"";

    std::u16string ret(utf8_to_utf16_len(sv), 0);

    utf8_to_utf16_range(sv, ret);

    return ret;
}
