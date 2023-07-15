#include "openxlsx2.h"

#include <fmt/format.h>
#include <fstream>
#include <charconv>
#include "cfbf.h"
#include "utf16.h"
#include "sha1.h"
#include "sha512.h"
#include "aes.h"
#include "b64.h"
#include "xlcpp-pimpl.h"

using namespace std;

static const uint32_t NOSTREAM = 0xffffffff;

struct structured_storage_header {
    uint64_t sig;
    uint8_t clsid[16];
    uint16_t minor_version;
    uint16_t major_version;
    uint16_t byte_order;
    uint16_t sector_shift;
    uint16_t mini_sector_shift;
    uint16_t reserved1;
    uint32_t reserved2;
    uint32_t num_sect_dir;
    uint32_t num_sect_fat;
    uint32_t sect_dir_start;
    uint32_t transaction_signature;
    uint32_t mini_sector_cutoff;
    uint32_t mini_fat_start;
    uint32_t num_sect_mini_fat;
    uint32_t sect_dif_start;
    uint32_t num_sect_dif;
    uint32_t sect_dif[109];
};

static_assert(sizeof(structured_storage_header) == 0x200);

enum class obj_type : uint8_t {
    STGTY_INVALID = 0,
    STGTY_STORAGE = 1,
    STGTY_STREAM = 2,
    STGTY_LOCKBYTES = 3,
    STGTY_PROPERTY = 4,
    STGTY_ROOT = 5
};

enum class tree_colour : uint8_t {
    red = 0,
    black = 1
};

#pragma pack(push,1)

struct dirent {
    char16_t name[32];
    uint16_t name_len;
    obj_type type;
    tree_colour colour;
    uint32_t sid_left_sibling;
    uint32_t sid_right_sibling;
    uint32_t sid_child;
    uint8_t clsid[16];
    uint32_t user_flags;
    uint64_t create_time;
    uint64_t modify_time;
    uint32_t sect_start;
    uint64_t size;
};

#pragma pack(pop)

static_assert(sizeof(dirent) == 0x80);

static const string_view NS_ENCRYPTION = "http://schemas.microsoft.com/office/2006/encryption";
static const string_view NS_PASSWORD = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

cfbf::cfbf(span<const uint8_t> s) : s(s) {
    auto& ssh = *(structured_storage_header*)s.data();

    if (ssh.sig != CFBF_SIGNATURE)
        Rcpp::stop("Incorrect signature.");

    auto& de = *(dirent*)(s.data() + (ssh.sect_dir_start + 1) * (1 << ssh.sector_shift));

    if (de.type != obj_type::STGTY_ROOT)
        Rcpp::stop("Root directory entry did not have type STGTY_ROOT.");

    add_entry("", 0, false);
}

const dirent& cfbf::find_dirent(uint32_t num) {
    auto& ssh = *(structured_storage_header*)s.data();

    auto dirents_per_sector = (1 << ssh.sector_shift) / sizeof(dirent);
    auto sector_skip = num / dirents_per_sector;
    auto sector = ssh.sect_dir_start;

    while (sector_skip > 0) {
        sector = next_sector(sector);
        sector_skip--;
    }

    return *(dirent*)(s.data() + ((sector + 1) << ssh.sector_shift) + ((num % dirents_per_sector) * sizeof(dirent)));
}

void cfbf::add_entry(string_view path, uint32_t num, bool ignore_right) {
    const auto& de = find_dirent(num);

    if (de.sid_left_sibling != NOSTREAM)
        add_entry(path, de.sid_left_sibling, true);

    auto name = de.name_len >= sizeof(char16_t) && num != 0 ? utf16_to_utf8(u16string_view(de.name, (de.name_len / sizeof(char16_t)) - 1)) : "";

    entries.emplace_back(*this, de, string(path) + name);

    if (de.sid_child != NOSTREAM)
        add_entry(string(path) + string(name) + "/", de.sid_child, false);

    if (!ignore_right && de.sid_right_sibling != NOSTREAM)
        add_entry(path, de.sid_right_sibling, false);
}

cfbf_entry::cfbf_entry(cfbf& file, const dirent& de, string_view name) : file(file), de(de), name(name) {
}

uint32_t cfbf::next_sector(uint32_t sector) const {
    auto& ssh = *(structured_storage_header*)s.data();
    auto sectors_per_dif = (1 << ssh.sector_shift) / sizeof(uint32_t);
    auto fat = (uint32_t*)(s.data() + ((ssh.sect_dif[sector / sectors_per_dif] + 1) << ssh.sector_shift));

    return fat[sector % sectors_per_dif];
}

uint32_t cfbf::next_mini_sector(uint32_t sector) const {
    auto& ssh = *(structured_storage_header*)s.data();
    auto mini_fat = (uint32_t*)(s.data() + ((ssh.mini_fat_start + 1) << ssh.sector_shift));

    return mini_fat[sector];
}

size_t cfbf_entry::read(span<std::byte> buf, uint64_t off) const {
    auto& ssh = *(structured_storage_header*)file.s.data();

    if (off >= de.size)
        return 0;

    if (off + buf.size() > de.size)
        buf = buf.subspan(0, de.size - off);

    size_t read = 0;

    if (de.size < ssh.mini_sector_cutoff) {
        auto mini_sector = de.sect_start;
        auto mini_sector_skip = off >> ssh.mini_sector_shift;

        for (unsigned int i = 0; i < mini_sector_skip; i++) {
            mini_sector = file.next_mini_sector(mini_sector);
        }

        auto mini_sectors_per_sector = 1 << (ssh.sector_shift - ssh.mini_sector_shift);

        do {
            auto mini_stream_sector = mini_sector / mini_sectors_per_sector;
            auto sector = file.entries[0].de.sect_start;

            while (mini_stream_sector > 0) {
                sector = file.next_sector(sector);
                mini_stream_sector--;
            }

            auto src = file.s.subspan(((sector + 1) << ssh.sector_shift) + ((mini_sector % mini_sectors_per_sector) << ssh.mini_sector_shift), 1 << ssh.mini_sector_shift);
            auto to_copy = min(src.size(), buf.size());

            memcpy(buf.data(), src.data(), to_copy);

            read += to_copy;
            buf = buf.subspan(to_copy);

            if (buf.empty())
                break;

            mini_sector = file.next_mini_sector(mini_sector);
        } while (true);
    } else {
        auto sector = de.sect_start;
        auto sector_skip = off >> ssh.sector_shift;

        for (unsigned int i = 0; i < sector_skip; i++) {
            sector = file.next_sector(sector);
        }

        do {
            auto src = file.s.subspan((sector + 1) << ssh.sector_shift, 1 << ssh.sector_shift);
            auto to_copy = min(src.size(), buf.size());

            memcpy(buf.data(), src.data(), to_copy);

            read += to_copy;
            buf = buf.subspan(to_copy);

            if (buf.empty())
                break;

            sector = file.next_sector(sector);
        } while (true);
    }

    return read;
}

size_t cfbf_entry::get_size() const {
    return de.size;
}

static array<uint8_t, 16> generate_key(u16string_view password, span<const uint8_t> salt, unsigned int spin_count) {
    array<uint8_t, 20> h;

    {
        SHA1_CTX ctx;

        ctx.update(salt);
        ctx.update(span((uint8_t*)password.data(), password.size() * sizeof(char16_t)));

        ctx.finalize(h);
    }

    for (uint32_t i = 0; i < spin_count; i++) {
        SHA1_CTX ctx;

        ctx.update(span((uint8_t*)&i, sizeof(uint32_t)));
        ctx.update(h);

        ctx.finalize(h);
    }

    {
        SHA1_CTX ctx;
        uint32_t block = 0;

        ctx.update(h);
        ctx.update(span((uint8_t*)&block, sizeof(uint32_t)));

        ctx.finalize(h);
    }

    array<uint8_t, 64> buf1 = {
        0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36,
        0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36,
        0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36,
        0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36,
        0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36,
        0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36,
        0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36,
        0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36, 0x36
    };

    for (unsigned int i = 0; auto c : h) {
        buf1[i] ^= c;
        i++;
    }

    auto x1 = sha1(buf1);

    array<uint8_t, 16> ret;
    memcpy(ret.data(), x1.data(), ret.size());

    return ret;
}

static void generate_key44_sha1(u16string_view password, span<const uint8_t> salt, unsigned int spin_count,
                                span<const uint8_t> block_key, span<uint8_t> ret) {
    array<uint8_t, 20> h;

    {
        SHA1_CTX ctx;

        ctx.update(salt);
        ctx.update(span((uint8_t*)password.data(), password.size() * sizeof(char16_t)));

        ctx.finalize(h);
    }

    for (uint32_t i = 0; i < spin_count; i++) {
        SHA1_CTX ctx;

        ctx.update(span((uint8_t*)&i, sizeof(uint32_t)));
        ctx.update(h);

        ctx.finalize(h);
    }

    {
        SHA1_CTX ctx;

        ctx.update(h);
        ctx.update(block_key);

        ctx.finalize(h);
    }

    memcpy(ret.data(), h.data(), 16);
}

static void generate_key44_sha512(u16string_view password, span<const uint8_t> salt, unsigned int spin_count,
                                  span<const uint8_t> block_key, span<uint8_t> ret) {
    array<uint8_t, 64> h;

    {
        sha512_state ctx;

        sha_init(ctx);
        sha_process(ctx, salt.data(), (uint32_t)salt.size());
        sha_process(ctx, password.data(), (uint32_t)(password.size() * sizeof(char16_t)));
        sha_done(ctx, h.data());
    }

    for (uint32_t i = 0; i < spin_count; i++) {
        sha512_state ctx;

        sha_init(ctx);
        sha_process(ctx, &i, sizeof(uint32_t));
        sha_process(ctx, h.data(), h.size());
        sha_done(ctx, h.data());
    }

    {
        sha512_state ctx;

        sha_init(ctx);
        sha_process(ctx, h.data(), h.size());
        sha_process(ctx, block_key.data(), (uint32_t)block_key.size());
        sha_done(ctx, h.data());
    }

    memcpy(ret.data(), h.data(), 64);
}

#pragma pack(push, 1)

struct encryption_info {
    uint16_t major;
    uint16_t minor;
    uint32_t flags;
    uint32_t header_size;
};

struct encryption_header {
    uint32_t flags;
    uint32_t size_extra;
    uint32_t alg_id;
    uint32_t alg_id_hash;
    uint32_t key_size;
    uint32_t provider_type;
    uint32_t reserved1;
    uint32_t reserved2;
    char16_t csp_name[0];
};

#pragma pack(pop)

static const uint32_t ALG_ID_AES_128 = 0x660e;
static const uint32_t ALG_ID_SHA_1 = 0x8004;

void cfbf::check_password(u16string_view password, span<const uint8_t> salt,
                          span<const uint8_t> encrypted_verifier,
                          span<const uint8_t> encrypted_verifier_hash) {
    auto key = generate_key(password, salt, 50000);
    AES_ctx ctx;
    array<uint8_t, 16> verifier;
    array<uint8_t, 32> verifier_hash;

    if (encrypted_verifier.size() != verifier.size())
        Rcpp::stop("encrypted_verifier.size() was {}, expected {}", encrypted_verifier.size(), verifier.size());

    if (encrypted_verifier_hash.size() != verifier_hash.size())
        Rcpp::stop("encrypted_verifier_hash.size() was {}, expected {}", encrypted_verifier_hash.size(), verifier_hash.size());

    AES128_init_ctx(&ctx, key.data());

    memcpy(verifier.data(), encrypted_verifier.data(), encrypted_verifier.size());

    AES128_ECB_decrypt(&ctx, verifier.data());

#if 0
    fmt::print("verifier = ");
    for (auto c : verifier) {
        fmt::print("{:02x} ", c);
    }
    fmt::print("\n");
#endif

    memcpy(verifier_hash.data(), encrypted_verifier_hash.data(), encrypted_verifier_hash.size());

    AES128_ECB_decrypt(&ctx, verifier_hash.data());
    AES128_ECB_decrypt(&ctx, verifier_hash.data() + 16);

#if 0
    fmt::print("verifier hash = ");
    for (auto c : verifier_hash) {
        fmt::print("{:02x} ", c);
    }
    fmt::print("\n");
#endif

    auto hash = sha1(verifier);

    if (memcmp(hash.data(), verifier_hash.data(), hash.size()))
        Rcpp::stop("Incorrect password.");

    key_size = 16;
    memcpy(this->key.data(), key.data(), key_size);
}

void cfbf::parse_enc_info_44(span<const uint8_t> enc_info, u16string_view password) {
    enc_info = enc_info.subspan(sizeof(uint32_t));

    if (enc_info.size() < sizeof(uint32_t) || *(uint32_t*)enc_info.data() != 0x40)
        Rcpp::stop("EncryptionInfo reserved value was not 0x40.");

    enc_info = enc_info.subspan(sizeof(uint32_t));

    xml_reader r(string_view((char*)enc_info.data(), enc_info.size()));

    bool found_root = false, found_key_data = false, found_password = false;

    while (r.read()) {
        if (r.node_type() == xml_node::element) {
            if (!found_root) {
                if (r.local_name() != "encryption" || !r.namespace_uri_raw().cmp(NS_ENCRYPTION))
                    Rcpp::stop("Root tag was {{{}}}{}, expected {{{}}}encryption.",
                                          r.namespace_uri_raw().decode(), r.local_name(), NS_ENCRYPTION);

                found_root = true;
            } else {
                if (r.local_name() == "keyData" && r.namespace_uri_raw().cmp(NS_ENCRYPTION)) {
                    string salt_value_b64, cipher_algorithm, key_bits_str, cipher_chaining, hash_algorithm;
                    unsigned int key_bits;

                    r.attributes_loop_raw([&](string_view local_name, xml_enc_string_view namespace_uri_raw,
                                              xml_enc_string_view value_raw) {

                        if (local_name == "saltValue")
                            salt_value_b64 = value_raw.decode();
                        else if (local_name == "cipherAlgorithm")
                            cipher_algorithm = value_raw.decode();
                        else if (local_name == "keyBits")
                            key_bits_str = value_raw.decode();
                        else if (local_name == "cipherChaining")
                            cipher_chaining = value_raw.decode();
                        else if (local_name == "hashAlgorithm")
                            hash_algorithm = value_raw.decode();

                        return true;
                    });

                    if (salt_value_b64.empty())
                        Rcpp::stop("saltValue not set");

                    if (cipher_algorithm.empty())
                        Rcpp::stop("cipherAlgorithm not set");

                    if (key_bits_str.empty())
                        Rcpp::stop("keyBits not set");

                    if (cipher_chaining.empty())
                        Rcpp::stop("cipherChaining not set");

                    if (hash_algorithm.empty())
                        Rcpp::stop("hashAlgorithm not set");

                    auto salt_value = b64decode(salt_value_b64);

                    if (cipher_algorithm != "AES")
                        Rcpp::stop("cipherAlgorithm was {}, expected AES", cipher_algorithm);

                    {
                        auto [ptr, ec] = from_chars(key_bits_str.data(), key_bits_str.data() + key_bits_str.size(), key_bits);

                        if (ptr != key_bits_str.data() + key_bits_str.size())
                            Rcpp::stop("Could not convert \"{}\" to integer.", key_bits_str);
                    }

                    if (key_bits != 128 && key_bits != 256)
                        Rcpp::stop("keyBits was {}, expected 128 or 256", key_bits);

                    if (cipher_chaining != "ChainingModeCBC")
                        Rcpp::stop("cipherChaining was {}, expected ChainingModeCBC", cipher_chaining);

                    if (hash_algorithm != "SHA1" && hash_algorithm != "SHA512")
                        Rcpp::stop("hashAlgorithm was {}, expected SHA1 or SHA512", hash_algorithm);

                    memcpy(salt.data(), salt_value.data(), min(salt_value.size(), salt.size()));
                    found_key_data = true;
                } else if (r.local_name() == "encryptedKey" && r.namespace_uri_raw().cmp(NS_PASSWORD)) {
                    string spin_count_str, salt_value_b64, cipher_algorithm, key_bits_str, cipher_chaining, hash_algorithm,
                           encrypted_verifier_hash_input_b64, encrypted_verifier_hash_value_b64, encrypted_key_value_b64;
                    unsigned int spin_count, key_bits;

                    r.attributes_loop_raw([&](string_view local_name, xml_enc_string_view namespace_uri_raw,
                                              xml_enc_string_view value_raw) {

                        if (local_name == "spinCount")
                            spin_count_str = value_raw.decode();
                        else if (local_name == "saltValue")
                            salt_value_b64 = value_raw.decode();
                        else if (local_name == "cipherAlgorithm")
                            cipher_algorithm = value_raw.decode();
                        else if (local_name == "keyBits")
                            key_bits_str = value_raw.decode();
                        else if (local_name == "cipherChaining")
                            cipher_chaining = value_raw.decode();
                        else if (local_name == "hashAlgorithm")
                            hash_algorithm = value_raw.decode();
                        else if (local_name == "encryptedVerifierHashInput")
                            encrypted_verifier_hash_input_b64 = value_raw.decode();
                        else if (local_name == "encryptedVerifierHashValue")
                            encrypted_verifier_hash_value_b64 = value_raw.decode();
                        else if (local_name == "encryptedKeyValue")
                            encrypted_key_value_b64 = value_raw.decode();

                        return true;
                    });

                    if (spin_count_str.empty())
                        Rcpp::stop("spinCount not set");

                    if (salt_value_b64.empty())
                        Rcpp::stop("saltValue not set");

                    if (cipher_algorithm.empty())
                        Rcpp::stop("cipherAlgorithm not set");

                    if (key_bits_str.empty())
                        Rcpp::stop("keyBits not set");

                    if (cipher_chaining.empty())
                        Rcpp::stop("cipherChaining not set");

                    if (hash_algorithm.empty())
                        Rcpp::stop("hashAlgorithm not set");

                    if (encrypted_verifier_hash_input_b64.empty())
                        Rcpp::stop("encryptedVerifierHashInput not set");

                    if (encrypted_verifier_hash_value_b64.empty())
                        Rcpp::stop("encryptedVerifierHashValue not set");

                    if (encrypted_key_value_b64.empty())
                        Rcpp::stop("encryptedKeyValue not set");

                    {
                        auto [ptr, ec] = from_chars(spin_count_str.data(), spin_count_str.data() + spin_count_str.size(), spin_count);

                        if (ptr != spin_count_str.data() + spin_count_str.size())
                            Rcpp::stop("Could not convert \"{}\" to integer.", spin_count_str);
                    }

                    auto salt_value = b64decode(salt_value_b64);

                    if (cipher_algorithm != "AES")
                        Rcpp::stop("cipherAlgorithm was {}, expected AES", cipher_algorithm);

                    {
                        auto [ptr, ec] = from_chars(key_bits_str.data(), key_bits_str.data() + key_bits_str.size(), key_bits);

                        if (ptr != key_bits_str.data() + key_bits_str.size())
                            Rcpp::stop("Could not convert \"{}\" to integer.", key_bits_str);
                    }

                    if (key_bits != 128 && key_bits != 256)
                        Rcpp::stop("keyBits was {}, expected 128 or 256", key_bits);

                    if (cipher_chaining != "ChainingModeCBC")
                        Rcpp::stop("cipherChaining was {}, expected ChainingModeCBC", cipher_chaining);

                    if (hash_algorithm == "SHA1")
                        hashalgo = hash_algorithm::sha1;
                    else if (hash_algorithm == "SHA512")
                        hashalgo = hash_algorithm::sha512;
                    else
                        Rcpp::stop("hashAlgorithm was {}, expected SHA1 or SHA512", hash_algorithm);

                    auto encrypted_verifier_hash_input = b64decode(encrypted_verifier_hash_input_b64);

                    auto encrypted_verifier_hash_value = b64decode(encrypted_verifier_hash_value_b64);

                    auto encrypted_key_value = b64decode(encrypted_key_value_b64);

                    static const array<uint8_t, 8> block1 = { 0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79 };
                    static const array<uint8_t, 8> block2 = { 0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e };
                    static const array<uint8_t, 8> block3 = { 0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6 };

                    // FIXME - we can save time by saving the partial hash for key1, key2, and key3

                    array<uint8_t, 64> key1, key2, key3;

                    if (hashalgo == hash_algorithm::sha512)
                        generate_key44_sha512(password, span((uint8_t*)salt_value.data(), salt_value.size()), spin_count, block1, key1);
                    else
                        generate_key44_sha1(password, span((uint8_t*)salt_value.data(), salt_value.size()), spin_count, block1, key1);

                    // FIXME - extend key if short

                    AES_ctx ctx;
                    array<uint8_t, 16> verifier;
                    array<uint8_t, 64> verifier_hash;

                    if (encrypted_verifier_hash_input.size() != verifier.size())
                        Rcpp::stop("encrypted_verifier_hash_input.size() was {}, expected {}", encrypted_verifier_hash_input.size(), verifier.size());

                    if (encrypted_verifier_hash_value.size() > verifier_hash.size())
                        Rcpp::stop("encrypted_verifier_hash_value.size() was {}, expected at most {}", encrypted_verifier_hash_value.size(), verifier_hash.size());

                    array<uint8_t, 16> iv;

                    memcpy(iv.data(), salt_value.data(), min(salt_value.size(), sizeof(iv)));

                    if (salt_value.size() < sizeof(iv))
                        memset(&iv[salt_value.size()], 0, sizeof(iv) - salt_value.size());

                    memcpy(verifier.data(), encrypted_verifier_hash_input.data(), encrypted_verifier_hash_input.size());

                    if (key_bits == 256) {
                        AES256_init_ctx_iv(&ctx, key1.data(), iv.data());
                        AES256_CBC_decrypt_buffer(&ctx, verifier.data(), verifier.size());
                    } else {
                        AES128_init_ctx_iv(&ctx, key1.data(), iv.data());
                        AES128_CBC_decrypt_buffer(&ctx, verifier.data(), verifier.size());
                    }

                    memcpy(verifier_hash.data(), encrypted_verifier_hash_value.data(), encrypted_verifier_hash_value.size());

                    if (hashalgo == hash_algorithm::sha512)
                        generate_key44_sha512(password, span((uint8_t*)salt_value.data(), salt_value.size()), spin_count, block2, key2);
                    else
                        generate_key44_sha1(password, span((uint8_t*)salt_value.data(), salt_value.size()), spin_count, block2, key2);

                    if (key_bits == 256) {
                        AES256_init_ctx_iv(&ctx, key2.data(), iv.data());
                        AES256_CBC_decrypt_buffer(&ctx, verifier_hash.data(), verifier_hash.size());
                    } else {
                        AES128_init_ctx_iv(&ctx, key2.data(), iv.data());
                        AES128_CBC_decrypt_buffer(&ctx, verifier_hash.data(), verifier_hash.size());
                    }

                    if (hashalgo == hash_algorithm::sha512) {
                        array<uint8_t, 64> hash;
                        sha512_state ctx;

                        sha_init(ctx);
                        sha_process(ctx, verifier.data(), (uint32_t)verifier.size());
                        sha_done(ctx, hash.data());

                        if (memcmp(hash.data(), verifier_hash.data(), hash.size()))
                            Rcpp::stop("Incorrect password.");
                    } else {
                        auto hash = sha1(verifier);
                        if (memcmp(hash.data(), verifier_hash.data(), hash.size()))
                            Rcpp::stop("Incorrect password.");
                    }

                    if (hashalgo == hash_algorithm::sha512)
                        generate_key44_sha512(password, span((uint8_t*)salt_value.data(), salt_value.size()), spin_count, block3, key3);
                    else
                        generate_key44_sha1(password, span((uint8_t*)salt_value.data(), salt_value.size()), spin_count, block3, key3);

                    if (key_bits == 256) {
                        AES256_init_ctx_iv(&ctx, key3.data(), iv.data());
                        AES256_CBC_decrypt_buffer(&ctx, (uint8_t*)encrypted_key_value.data(), encrypted_key_value.size());
                    } else {
                        AES128_init_ctx_iv(&ctx, key3.data(), iv.data());
                        AES128_CBC_decrypt_buffer(&ctx, (uint8_t*)encrypted_key_value.data(), encrypted_key_value.size());
                    }

                    key_size = key_bits / 8;
                    memcpy(key.data(), encrypted_key_value.data(), min((size_t)key_size, encrypted_key_value.size()));

                    found_password = true;
                }
            }
        }
    }

    if (!found_key_data)
        Rcpp::stop("keyData not found");

    if (!found_password)
        Rcpp::stop("encryptedKey not found");

    agile_enc = true;
}

void cfbf::parse_enc_info(span<const uint8_t> enc_info, u16string_view password) {
    if (enc_info.size() < sizeof(encryption_info))
        Rcpp::stop("EncryptionInfo was {} bytes, expected at least {}", enc_info.size(), sizeof(encryption_info));

    auto& ei = *(encryption_info*)enc_info.data();

    if (ei.major == 4 && ei.minor == 4) {
        parse_enc_info_44(enc_info, password);
        return;
    } else if (ei.major != 3 || ei.minor != 2)
        Rcpp::stop("Unsupported EncryptionInfo version {}.{}", ei.major, ei.minor);

    if (ei.flags != 0x24) // AES
        Rcpp::stop("Unsupported EncryptionInfo flags {:x}", ei.flags);

    if (ei.header_size < offsetof(encryption_header, csp_name))
        Rcpp::stop("Encryption header was {} bytes, expected at least {}", ei.header_size, offsetof(encryption_header, csp_name));

    if (ei.header_size > enc_info.size() - sizeof(encryption_info))
        Rcpp::stop("Encryption header was {} bytes, but only {} remaining", ei.header_size, enc_info.size() - sizeof(encryption_info));

    auto& h = *(encryption_header*)(enc_info.data() + sizeof(encryption_info));

    if (h.alg_id != ALG_ID_AES_128)
        Rcpp::stop("Unsupported algorithm ID {:x}", h.alg_id);

    if (h.alg_id_hash != ALG_ID_SHA_1 && h.alg_id_hash != 0)
        Rcpp::stop("Unsupported hash algorithm ID {:x}", h.alg_id_hash);

    if (h.key_size != 128)
        Rcpp::stop("Key size was {}, expected 128", h.key_size);

    auto sp = enc_info.subspan(sizeof(encryption_info) + ei.header_size);

    if (sp.size() < sizeof(uint32_t))
        Rcpp::stop("Malformed EncryptionInfo");

    auto salt_size = *(uint32_t*)sp.data();
    sp = sp.subspan(sizeof(uint32_t));

    if (sp.size() < salt_size)
        Rcpp::stop("Malformed EncryptionInfo");

    auto salt = sp.subspan(0, salt_size);
    sp = sp.subspan(salt_size);

    if (sp.size() < 16)
        Rcpp::stop("Malformed EncryptionInfo");

    auto encrypted_verifier = sp.subspan(0, 16);
    sp = sp.subspan(16);

    if (sp.size() < sizeof(uint32_t))
        Rcpp::stop("Malformed EncryptionInfo");

    // skip verifier_hash_size
    sp = sp.subspan(sizeof(uint32_t));

    if (sp.size() < 32)
        Rcpp::stop("Malformed EncryptionInfo");

    auto encrypted_verifier_hash = sp.subspan(0, 32);

    check_password(password, salt, encrypted_verifier, encrypted_verifier_hash); // throws if wrong

    memcpy(this->salt.data(), salt.data(), this->salt.size());
}

vector<uint8_t> cfbf::decrypt44(span<uint8_t> enc_package) {
    uint32_t segment_no = 0;
    vector<uint8_t> ret;

    static const size_t SEGMENT_LENGTH = 0x1000;

    ret.resize(enc_package.size());
    auto ptr = ret.data();

    while (true) {
        array<uint8_t, 64> h;
        AES_ctx ctx;

        auto seg = enc_package.subspan(0, min(SEGMENT_LENGTH, enc_package.size()));

        memcpy(ptr, seg.data(), seg.size());

        if (hashalgo == hash_algorithm::sha512) {
            sha512_state ctx;

            sha_init(ctx);
            sha_process(ctx, salt.data(), (uint32_t)salt.size());
            sha_process(ctx, &segment_no, sizeof(segment_no));
            sha_done(ctx, h.data());
        } else {
            SHA1_CTX ctx;

            ctx.update(salt);
            ctx.update(span((uint8_t*)&segment_no, sizeof(segment_no)));

            ctx.finalize(h);
        }

        if (key_size == 32) {
            AES256_init_ctx_iv(&ctx, key.data(), h.data());
            AES256_CBC_decrypt_buffer(&ctx, ptr, seg.size());
        } else {
            AES128_init_ctx_iv(&ctx, key.data(), h.data());
            AES128_CBC_decrypt_buffer(&ctx, ptr, seg.size());
        }

        if (enc_package.size() == seg.size())
            break;

        enc_package = enc_package.subspan(seg.size());
        segment_no++;
        ptr += seg.size();
    }

    return ret;
}

vector<uint8_t> cfbf::decrypt(span<uint8_t> enc_package) {
    if (enc_package.size() < sizeof(uint64_t))
        Rcpp::stop("EncryptedPackage was {} bytes, expected at least {}", enc_package.size(), sizeof(uint64_t));

    auto size = *(uint64_t*)enc_package.data();

    enc_package = enc_package.subspan(sizeof(uint64_t));

    if (enc_package.size() < size)
        Rcpp::stop("EncryptedPackage was {} bytes, expected at least {}", enc_package.size() + sizeof(uint64_t), size + sizeof(uint64_t));

    if (agile_enc)
        return decrypt44(enc_package);

    AES_ctx ctx;
    auto buf = enc_package;

    AES128_init_ctx(&ctx, key.data());

    while (!buf.empty()) {
        AES128_ECB_decrypt(&ctx, buf.data());

        buf = buf.subspan(16);
    }

    vector<uint8_t> ret;

    ret.assign(enc_package.begin(), enc_package.end());

    return ret;
}
