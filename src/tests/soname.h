#include "openssl/opensslv.h"

#define XSTR(x) STR(x)
#define STR(x) #x
#ifdef SHLIB_VERSION_NUMBER
echo XSTR(SHLIB_VERSION_NUMBER)
#else
  echo XSTR(OPENSSL_SHLIB_VERSION)
#endif
