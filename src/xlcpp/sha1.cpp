/*
SHA-1 in C
By Steve Reid <steve@edmweb.com>
100% Public Domain

Test Vectors (from FIPS PUB 180-1)
"abc"
A9993E36 4706816A BA3E2571 7850C26C 9CD0D89D
"abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq"
84983E44 1C3BD26E BAAE4AA1 F95129E5 E54670F1
A million repetitions of "a"
34AA973C D4C4DAA4 F61EEB2B DBAD2731 6534016F
*/

/* #define LITTLE_ENDIAN * This should be #define'd already, if true. */
/* #define SHA1HANDSOFF * Copies data before messing with it. */

#define SHA1HANDSOFF

#include <stdio.h>
#include <string.h>
#include <stdint.h>
#include <bit>

#include "sha1.h"

using namespace std;

constexpr void R0(uint32_t v, uint32_t& w, uint32_t x, uint32_t y, uint32_t& z, uint32_t& bl) {
    z += (w & (x^y)) ^ y;

    bl = (rotl(bl,24) & 0xff00ff00) | (rotl(bl,8) & 0x00ff00ff);
    z += bl;

    z += 0x5a827999;
    z += rotl(v, 5);

    w = rotl(w, 30);
}

constexpr void R1(uint32_t v, uint32_t& w, uint32_t x, uint32_t y, uint32_t& z, uint32_t i, uint32_t* l) {
    z += (w & (x ^ y)) ^ y;

    l[i&15] = rotl(l[(i+13)&15] ^ l[(i+8)&15] ^ l[(i+2)&15] ^ l[i&15], 1);
    z += l[i&15];

    z += 0x5a827999;
    z += rotl(v, 5);

    w = rotl(w, 30);
}

constexpr void R2(uint32_t v, uint32_t& w, uint32_t x, uint32_t y, uint32_t& z, uint32_t i, uint32_t* l) {
    z += w ^ x ^ y;

    l[i&15] = rotl(l[(i+13)&15] ^ l[(i+8)&15] ^ l[(i+2)&15] ^ l[i&15], 1);
    z += l[i&15];

    z += 0x6ed9eba1;
    z += rotl(v, 5);

    w = rotl(w, 30);
}

constexpr void R3(uint32_t v, uint32_t& w, uint32_t x, uint32_t y, uint32_t& z, uint32_t i, uint32_t* l) {
    z += ((w | x) & y) | (w & x);

    l[i&15] = rotl(l[(i+13)&15] ^ l[(i+8)&15] ^ l[(i+2)&15] ^ l[i&15], 1);
    z += l[i&15];

    z += 0x8f1bbcdc;
    z += rotl(v, 5);

    w = rotl(w, 30);
}

constexpr void R4(uint32_t v, uint32_t& w, uint32_t x, uint32_t y, uint32_t& z, uint32_t i, uint32_t* l) {
    z += w ^ x ^ y;

    l[i&15] = rotl(l[(i+13)&15] ^ l[(i+8)&15] ^ l[(i+2)&15] ^ l[i&15], 1);
    z += l[i&15];

    z += 0xca62c1d6;
    z += rotl(v,5);

    w = rotl(w,30);
}


/* Hash a single 512-bit block. This is the core of the algorithm. */

static void SHA1Transform(uint32_t state[5], uint8_t buffer[64]) {
    uint32_t a, b, c, d, e;
    uint32_t l[16];

    memcpy(l, buffer, 64);

    /* Copy context->state[] to working vars */
    a = state[0];
    b = state[1];
    c = state[2];
    d = state[3];
    e = state[4];

    /* 4 rounds of 20 operations each. Loop unrolled. */
    R0(a,b,c,d,e, l[0]);
    R0(e,a,b,c,d, l[1]);
    R0(d,e,a,b,c, l[2]);
    R0(c,d,e,a,b, l[3]);
    R0(b,c,d,e,a, l[4]);
    R0(a,b,c,d,e, l[5]);
    R0(e,a,b,c,d, l[6]);
    R0(d,e,a,b,c, l[7]);
    R0(c,d,e,a,b, l[8]);
    R0(b,c,d,e,a, l[9]);
    R0(a,b,c,d,e, l[10]);
    R0(e,a,b,c,d, l[11]);
    R0(d,e,a,b,c, l[12]);
    R0(c,d,e,a,b, l[13]);
    R0(b,c,d,e,a, l[14]);
    R0(a,b,c,d,e, l[15]);

    R1(e,a,b,c,d, 16, l);
    R1(d,e,a,b,c, 17, l);
    R1(c,d,e,a,b, 18, l);
    R1(b,c,d,e,a, 19, l);

    R2(a,b,c,d,e,20, l);
    R2(e,a,b,c,d,21, l);
    R2(d,e,a,b,c,22, l);
    R2(c,d,e,a,b,23, l);
    R2(b,c,d,e,a,24, l);
    R2(a,b,c,d,e,25, l);
    R2(e,a,b,c,d,26, l);
    R2(d,e,a,b,c,27, l);
    R2(c,d,e,a,b,28, l);
    R2(b,c,d,e,a,29, l);
    R2(a,b,c,d,e,30, l);
    R2(e,a,b,c,d,31, l);
    R2(d,e,a,b,c,32, l);
    R2(c,d,e,a,b,33, l);
    R2(b,c,d,e,a,34, l);
    R2(a,b,c,d,e,35, l);
    R2(e,a,b,c,d,36, l);
    R2(d,e,a,b,c,37, l);
    R2(c,d,e,a,b,38, l);
    R2(b,c,d,e,a,39, l);

    R3(a,b,c,d,e,40, l);
    R3(e,a,b,c,d,41, l);
    R3(d,e,a,b,c,42, l);
    R3(c,d,e,a,b,43, l);
    R3(b,c,d,e,a,44, l);
    R3(a,b,c,d,e,45, l);
    R3(e,a,b,c,d,46, l);
    R3(d,e,a,b,c,47, l);
    R3(c,d,e,a,b,48, l);
    R3(b,c,d,e,a,49, l);
    R3(a,b,c,d,e,50, l);
    R3(e,a,b,c,d,51, l);
    R3(d,e,a,b,c,52, l);
    R3(c,d,e,a,b,53, l);
    R3(b,c,d,e,a,54, l);
    R3(a,b,c,d,e,55, l);
    R3(e,a,b,c,d,56, l);
    R3(d,e,a,b,c,57, l);
    R3(c,d,e,a,b,58, l);
    R3(b,c,d,e,a,59, l);

    R4(a,b,c,d,e,60, l);
    R4(e,a,b,c,d,61, l);
    R4(d,e,a,b,c,62, l);
    R4(c,d,e,a,b,63, l);
    R4(b,c,d,e,a,64, l);
    R4(a,b,c,d,e,65, l);
    R4(e,a,b,c,d,66, l);
    R4(d,e,a,b,c,67, l);
    R4(c,d,e,a,b,68, l);
    R4(b,c,d,e,a,69, l);
    R4(a,b,c,d,e,70, l);
    R4(e,a,b,c,d,71, l);
    R4(d,e,a,b,c,72, l);
    R4(c,d,e,a,b,73, l);
    R4(b,c,d,e,a,74, l);
    R4(a,b,c,d,e,75, l);
    R4(e,a,b,c,d,76, l);
    R4(d,e,a,b,c,77, l);
    R4(c,d,e,a,b,78, l);
    R4(b,c,d,e,a,79, l);
    /* Add the working vars back into context.state[] */
    state[0] += a;
    state[1] += b;
    state[2] += c;
    state[3] += d;
    state[4] += e;
}

/* Run your data through this. */

void SHA1_CTX::update(std::span<const uint8_t> data) {
    uint32_t i;

    auto j = count[0];

    count[0] += (uint32_t)(data.size() << 3);

    if (count[0] < j)
        count[1]++;

    count[1] += (uint32_t)(data.size() >> 29);

    j = (j >> 3) & 63;

    if (j + data.size() > 63) {
        i = 64 - j;
        memcpy(&buffer[j], data.data(), i);
        SHA1Transform(state, buffer);
        for ( ; i + 63 < data.size(); i += 64) {
            SHA1Transform(state, (uint8_t*)&data[i]);
        }
        j = 0;
    } else
        i = 0;

    memcpy(&buffer[j], &data[i], data.size() - i);
}


/* Add padding and return the message digest. */

void SHA1_CTX::finalize(span<uint8_t> digest) {
    unsigned i;
    unsigned char finalcount[8];
    unsigned char c;

    for (i = 0; i < 8; i++) {
        finalcount[i] = (unsigned char)((count[(i >= 4 ? 0 : 1)]
            >> ((3-(i & 3)) * 8) ) & 255);  /* Endian independent */
    }

    c = 0200;
    update(span(&c, 1));
    while ((count[0] & 504) != 448) {
        c = 0000;
        update(span(&c, 1));
    }
    update(span(finalcount, 8));  /* Should cause a SHA1Transform() */
    for (i = 0; i < 20; i++) {
        digest[i] = (unsigned char)
            ((state[i>>2] >> ((3-(i & 3)) * 8) ) & 255);
    }
}

array<uint8_t, 20> sha1(span<const uint8_t> s) {
    array<uint8_t, 20> digest;
    SHA1_CTX ctx;

    ctx.update(s);
    ctx.finalize(digest);

    return digest;
}

