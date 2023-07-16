#pragma once

#include <string>
#include <string_view>
#include <filesystem>
#include <vector>
#include <span>
#include <array>

static const uint64_t CFBF_SIGNATURE = 0xe11ab1a1e011cfd0;

#define formatted_error(s, ...) _formatted_error(FMT_COMPILE(s), ##__VA_ARGS__)

class cfbf;
struct dirent;

class cfbf_entry {
public:
    cfbf_entry(cfbf& file, const dirent& de, std::string_view name);
    size_t read(std::span<std::byte> buf, uint64_t off) const;
    size_t get_size() const;

    cfbf& file;
    const dirent& de;
    std::string name;
};

enum class hash_algorithm {
    sha1,
    sha512
};

class cfbf {
public:
    cfbf(std::span<const uint8_t> s);
    uint32_t next_sector(uint32_t sector) const;
    uint32_t next_mini_sector(uint32_t sector) const;
    void parse_enc_info(std::span<const uint8_t> enc_info, std::u16string_view password);
    void parse_enc_info_44(std::span<const uint8_t> enc_info, std::u16string_view password);
    std::vector<uint8_t> decrypt(std::span<uint8_t> enc_package);

    std::vector<cfbf_entry> entries;
    std::span<const uint8_t> s;

private:
    void add_entry(std::string_view path, uint32_t num, bool ignore_right);
    void check_password(std::u16string_view password, std::span<const uint8_t> salt,
                        std::span<const uint8_t> encrypted_verifier,
                        std::span<const uint8_t> encrypted_verifier_hash);
    const dirent& find_dirent(uint32_t num);
    std::vector<uint8_t> decrypt44(std::span<uint8_t> enc_package);

    std::array<uint8_t, 32> key;
    unsigned int key_size;
    std::array<uint8_t, 16> salt;
    bool agile_enc = false;
    enum hash_algorithm hashalgo;
};
