#ifndef OPENXLSX2_BIT_CAST_HPP
#define OPENXLSX2_BIT_CAST_HPP

#if defined(__cpp_lib_bit_cast) && __cpp_lib_bit_cast >= 201806L

#include <bit>

namespace openxlsx2 {
using std::bit_cast;
}

#else

namespace openxlsx2 {

template <typename To, typename From>
typename std::enable_if<
  sizeof(To) == sizeof(From) &&
  std::is_trivially_copyable<From>::value &&
  std::is_trivially_copyable<To>::value,
  To
>::type
bit_cast(const From& src) noexcept {
  To dst;
  std::memcpy(&dst, &src, sizeof(To));
  return dst;
}

} // namespace openxlsx2

#endif // __cpp_lib_bit_cast

#endif // OPENXLSX2_BIT_CAST_HPP
