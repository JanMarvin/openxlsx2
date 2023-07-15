/*
SHA-1 in C
By Steve Reid <steve@edmweb.com>
100% Public Domain
*/

#pragma once

#include <array>
#include <span>

struct SHA1_CTX {
    constexpr SHA1_CTX() {
        /* SHA1 initialization constants */
        state[0] = 0x67452301;
        state[1] = 0xEFCDAB89;
        state[2] = 0x98BADCFE;
        state[3] = 0x10325476;
        state[4] = 0xC3D2E1F0;
        count[0] = count[1] = 0;

        for (unsigned int i = 0; i < sizeof(buffer); i++) {
            buffer[i] = 0;
        }
    }

    void update(std::span<const uint8_t> data);
    void finalize(std::span<uint8_t> digest);

    uint32_t state[5];
    uint32_t count[2];
    unsigned char buffer[64];
};

std::array<uint8_t, 20> sha1(std::span<const uint8_t> s);
