#pragma once

#include <stdint.h>
#include <stddef.h>

#define AES_BLOCKLEN 16 // Block length in bytes - AES is 128b block only
#define AES_keyExpSize 240

struct AES_ctx {
    uint8_t RoundKey[AES_keyExpSize];
    uint8_t Iv[AES_BLOCKLEN];
};

void AES128_init_ctx(struct AES_ctx* ctx, const uint8_t* key);
void AES128_init_ctx_iv(struct AES_ctx* ctx, const uint8_t* key, const uint8_t* iv);
void AES128_ECB_encrypt(const struct AES_ctx* ctx, uint8_t* buf);
void AES128_ECB_decrypt(const struct AES_ctx* ctx, uint8_t* buf);
void AES128_CBC_encrypt_buffer(struct AES_ctx* ctx, uint8_t* buf, size_t length);
void AES128_CBC_decrypt_buffer(struct AES_ctx* ctx, uint8_t* buf, size_t length);

void AES256_init_ctx(struct AES_ctx* ctx, const uint8_t* key);
void AES256_init_ctx_iv(struct AES_ctx* ctx, const uint8_t* key, const uint8_t* iv);
void AES256_ECB_encrypt(const struct AES_ctx* ctx, uint8_t* buf);
void AES256_ECB_decrypt(const struct AES_ctx* ctx, uint8_t* buf);
void AES256_CBC_encrypt_buffer(struct AES_ctx* ctx, uint8_t* buf, size_t length);
void AES256_CBC_decrypt_buffer(struct AES_ctx* ctx, uint8_t* buf, size_t length);
