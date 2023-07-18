#include <openssl/opensslv.h>
#if OPENSSL_VERSION_NUMBER < 0x10000000L
#error OpenSSL version too old
#endif
