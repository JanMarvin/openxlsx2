#pragma once

#include <string>
#include <string_view>
#include <span>
#include <memory>
#include <fcntl.h>

#ifdef _WIN32
#include <windows.h>
#include "utf16.h"
using fd_t = HANDLE;
#else
#include <unistd.h>
using fd_t = int;
#endif

#ifdef _WIN32
class last_error : public std::exception {
public:
    last_error(std::string_view function, int le) {
        std::string nice_msg;

        {
            char16_t* fm;

            if (FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, nullptr,
                le, 0, reinterpret_cast<LPWSTR>(&fm), 0, nullptr)) {
                try {
                    std::u16string_view s = fm;

                    while (!s.empty() && (s[s.length() - 1] == u'\r' || s[s.length() - 1] == u'\n')) {
                        s.remove_suffix(1);
                    }

                    nice_msg = utf16_to_utf8(s);
                } catch (...) {
                    LocalFree(fm);
                    throw;
                }

                LocalFree(fm);
            }
        }

        msg = std::string(function) + " failed (error " + std::to_string(le) + (!nice_msg.empty() ? (", " + nice_msg) : "") + ").";
    }

    const char* what() const noexcept {
        return msg.c_str();
    }

private:
    std::string msg;
};
#endif

#ifdef _WIN32

class handle_closer {
public:
    typedef HANDLE pointer;

    void operator()(HANDLE h) {
        if (h == INVALID_HANDLE_VALUE)
            return;

        CloseHandle(h);
    }
};

using unique_handle = std::unique_ptr<HANDLE, handle_closer>;

#else

class unique_handle {
public:
    unique_handle() : fd(0) {
    }

    explicit unique_handle(int fd) : fd(fd) {
    }

    unique_handle(unique_handle&& that) noexcept {
        fd = that.fd;
        that.fd = 0;
    }

    unique_handle(const unique_handle&) = delete;
    unique_handle& operator=(const unique_handle&) = delete;

    unique_handle& operator=(unique_handle&& that) noexcept {
        if (fd > 0)
            close(fd);

        fd = that.fd;
        that.fd = 0;

        return *this;
    }

    ~unique_handle() {
        if (fd <= 0)
            return;

        close(fd);
    }

    void reset(int new_fd) noexcept {
        if (fd > 0)
            close(fd);

        fd = new_fd;
    }

    int get() const noexcept {
        return fd;
    }

private:
    int fd;
};
#endif

class errno_error : public std::exception {
public:
    errno_error(std::string_view function, int en);

    const char* what() const noexcept {
        return msg.c_str();
    }

private:
    std::string msg;
};

class mmap {
public:
	explicit mmap(fd_t h);
	~mmap();
	std::span<const uint8_t> map();

	size_t filesize;

private:
#ifdef _WIN32
	HANDLE h = INVALID_HANDLE_VALUE;
	HANDLE mh = INVALID_HANDLE_VALUE;
#endif
	void* ptr = nullptr;
};
