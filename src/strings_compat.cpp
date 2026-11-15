#include <cstring>
#include <cstdint>
#include <climits>
#include <algorithm>
#include <string>
#include <unordered_set>
#include <vector>

#include "openxlsx2.h"

// Helpers for R/compat-stringi.R. All strings are handled as UTF-8: elements
// are converted with Rf_translateCharUTF8() (ASCII and UTF-8 strings are used
// as they are) and results are marked as UTF-8.

static SEXP make_utf8(const std::string& x) {
  if (x.size() > static_cast<size_t>(INT_MAX)) Rcpp::stop("string is too long");
  return Rf_mkCharLenCE(x.data(), static_cast<int>(x.size()), CE_UTF8);
}

// as.character() for anything that is not a character vector yet
static Rcpp::CharacterVector as_chr(SEXP x) {
  if (TYPEOF(x) == STRSXP) return Rcpp::CharacterVector(x);
  Rcpp::Function as_character("as.character");
  return Rcpp::CharacterVector(as_character(x));
}

static bool is_ascii(const char* x) {
  for (; *x != '\0'; ++x) {
    if (static_cast<unsigned char>(*x) > 127) return false;
  }
  return true;
}

// ASCII check of a string with known length, 8 bytes at a time
static bool is_ascii_n(const char* x, size_t len) {
  size_t i = 0;
  for (; i + 8 <= len; i += 8) {
    uint64_t word;
    std::memcpy(&word, x + i, 8);
    if ((word & 0x8080808080808080ULL) != 0) return false;
  }
  for (; i < len; ++i) {
    if (static_cast<unsigned char>(x[i]) > 127) return false;
  }
  return true;
}

// byte offset that is nchars characters after the byte offset start
static size_t advance_chars(const char* x, size_t len, size_t start, int64_t nchars) {
  size_t pos = start;
  int64_t seen = 0;
  while (pos < len && seen < nchars) {
    ++pos;
    while (pos < len && (static_cast<unsigned char>(x[pos]) & 0xC0) == 0x80) ++pos;
    ++seen;
  }
  return pos;
}

// number of characters in the first nbytes bytes
static int64_t count_chars(const char* x, size_t nbytes) {
  int64_t n = 0;
  for (size_t i = 0; i < nbytes; ++i) {
    if ((static_cast<unsigned char>(x[i]) & 0xC0) != 0x80) ++n;
  }
  return n;
}

// paste(..., sep = sep) with NA as result if any element is NA. a zero length
// element gives a zero length result
// [[Rcpp::export(rng = false)]]
Rcpp::CharacterVector stri_join_cpp(Rcpp::List args, Rcpp::CharacterVector sep) {

  int64_t k = static_cast<int64_t>(args.size());
  std::string sep_utf8 = Rf_translateCharUTF8(STRING_ELT(sep, 0));

  std::vector<Rcpp::CharacterVector> cols;
  cols.reserve(static_cast<size_t>(k));
  std::vector<int64_t> lens;
  int64_t n = 0;
  for (int64_t j = 0; j < k; ++j) {
    cols.push_back(as_chr(args[j]));
    lens.push_back(static_cast<int64_t>(cols[static_cast<size_t>(j)].size()));
    if (lens[static_cast<size_t>(j)] == 0) return Rcpp::CharacterVector(0);
    if (lens[static_cast<size_t>(j)] > n) n = lens[static_cast<size_t>(j)];
  }

  Rcpp::CharacterVector res(n);
  std::string buf;

  for (int64_t i = 0; i < n; ++i) {
    const void* vmax = vmaxget();
    buf.clear();
    bool is_na = false;

    for (int64_t j = 0; j < k; ++j) {
      size_t col = static_cast<size_t>(j);
      SEXP el = STRING_ELT(cols[col], lens[col] == n ? i : i % lens[col]);
      if (el == NA_STRING) {
        is_na = true;
        break;
      }
      if (j > 0) buf += sep_utf8;
      buf += Rf_translateCharUTF8(el);
    }

    if (is_na) {
      SET_STRING_ELT(res, i, NA_STRING);
    } else {
      SET_STRING_ELT(res, i, make_utf8(buf));
    }
    vmaxset(vmax);
  }

  return res;
}

// paste(x, collapse = collapse) with NA as result if any element is NA
// [[Rcpp::export(rng = false)]]
Rcpp::CharacterVector stri_collapse_cpp(Rcpp::CharacterVector x, Rcpp::CharacterVector collapse) {

  std::string sep = Rf_translateCharUTF8(STRING_ELT(collapse, 0));
  int64_t n = static_cast<int64_t>(x.size());
  std::string buf;

  const void* vmax = vmaxget();
  for (int64_t i = 0; i < n; ++i) {
    SEXP el = STRING_ELT(x, i);
    if (el == NA_STRING) {
      vmaxset(vmax);
      return Rcpp::CharacterVector::create(NA_STRING);
    }
    if (i > 0) buf += sep;
    buf += Rf_translateCharUTF8(el);
    vmaxset(vmax);
  }

  Rcpp::CharacterVector res(1);
  SET_STRING_ELT(res, 0, make_utf8(buf));
  return res;
}

// every pattern is replaced in every string, one after the other. pattern and
// replacement have the same length
// [[Rcpp::export(rng = false)]]
Rcpp::CharacterVector stri_replace_all_fixed_cpp(
    Rcpp::CharacterVector str,
    Rcpp::CharacterVector pattern,
    Rcpp::CharacterVector replacement
) {

  size_t m = static_cast<size_t>(pattern.size());
  std::vector<std::string> pat(m);
  std::vector<std::string> rep(m);
  for (size_t j = 0; j < m; ++j) {
    pat[j] = Rf_translateCharUTF8(STRING_ELT(pattern, static_cast<R_xlen_t>(j)));
    rep[j] = Rf_translateCharUTF8(STRING_ELT(replacement, static_cast<R_xlen_t>(j)));
    if (pat[j].empty()) Rcpp::stop("empty search patterns are not supported");
  }

  int64_t n = static_cast<int64_t>(str.size());
  Rcpp::CharacterVector res(n);

  for (int64_t i = 0; i < n; ++i) {
    SEXP el = STRING_ELT(str, i);
    if (el == NA_STRING) {
      SET_STRING_ELT(res, i, NA_STRING);
      continue;
    }

    const void* vmax = vmaxget();
    const char* chars = Rf_translateCharUTF8(el);
    std::string cur;
    bool changed = false;

    for (size_t j = 0; j < m; ++j) {
      const char* src = changed ? cur.c_str() : chars;
      const char* hit = std::strstr(src, pat[j].c_str());
      if (hit == nullptr) continue;

      std::string out;
      while (hit != nullptr) {
        out.append(src, static_cast<size_t>(hit - src));
        out += rep[j];
        src = hit + pat[j].size();
        hit = std::strstr(src, pat[j].c_str());
      }
      out += src;

      cur.swap(out);
      changed = true;
    }

    if (changed) {
      SET_STRING_ELT(res, i, make_utf8(cur));
    } else if (Rf_getCharCE(el) == CE_UTF8 || is_ascii(chars)) {
      SET_STRING_ELT(res, i, el);
    } else {
      SET_STRING_ELT(res, i, make_utf8(chars));
    }
    vmaxset(vmax);
  }

  return res;
}

// replaces the characters from to to (counted in characters, starting at 1)
// with value. to = from - 1 inserts value in front of from. the positions must
// be valid (from >= 1, to >= from - 1)
// [[Rcpp::export(rng = false)]]
Rcpp::CharacterVector stri_sub_replace_cpp(
    Rcpp::CharacterVector str,
    Rcpp::IntegerVector from,
    Rcpp::IntegerVector to,
    Rcpp::CharacterVector value
) {

  int64_t n_str = static_cast<int64_t>(str.size());
  int64_t n_from = static_cast<int64_t>(from.size());
  int64_t n_to = static_cast<int64_t>(to.size());
  int64_t n_value = static_cast<int64_t>(value.size());

  int64_t n = n_str;
  if (n_from > n) n = n_from;
  if (n_to > n) n = n_to;
  if (n_value > n) n = n_value;

  Rcpp::CharacterVector res(n);

  for (int64_t i = 0; i < n; ++i) {
    SEXP el = STRING_ELT(str, i % n_str);
    SEXP val = STRING_ELT(value, i % n_value);
    int f = from[i % n_from];
    int t = to[i % n_to];

    if (el == NA_STRING || val == NA_STRING || f == NA_INTEGER || t == NA_INTEGER) {
      SET_STRING_ELT(res, i, NA_STRING);
      continue;
    }

    const void* vmax = vmaxget();
    const char* chars = Rf_translateCharUTF8(el);
    const char* ins = Rf_translateCharUTF8(val);
    size_t len = std::strlen(chars);

    size_t head;
    size_t tail;
    if (is_ascii_n(chars, len)) {
      // one byte per character
      head = std::min(static_cast<size_t>(f) - 1, len);
      tail = std::min(static_cast<size_t>(t), len);
      if (tail < head) tail = head;
    } else {
      head = advance_chars(chars, len, 0, static_cast<int64_t>(f) - 1);
      tail = advance_chars(chars, len, head, static_cast<int64_t>(t) - static_cast<int64_t>(f) + 1);
    }

    std::string out;
    out.reserve(len + std::strlen(ins));
    out.append(chars, head);
    out += ins;
    out.append(chars + tail, len - tail);

    SET_STRING_ELT(res, i, make_utf8(out));
    vmaxset(vmax);
  }

  return res;
}

// start and end (in characters) of the last occurrence of pattern, including
// overlapping ones. NA if there is none
// [[Rcpp::export(rng = false)]]
Rcpp::IntegerMatrix stri_locate_last_fixed_cpp(Rcpp::CharacterVector str, Rcpp::CharacterVector pattern) {

  std::string pat = Rf_translateCharUTF8(STRING_ELT(pattern, 0));
  if (pat.empty()) Rcpp::stop("empty search patterns are not supported");
  int64_t pat_chars = count_chars(pat.c_str(), pat.size());

  int64_t n = static_cast<int64_t>(str.size());
  Rcpp::IntegerMatrix res(static_cast<int>(n), 2);

  for (int64_t i = 0; i < n; ++i) {
    SEXP el = STRING_ELT(str, i);
    int row = static_cast<int>(i);
    res(row, 0) = NA_INTEGER;
    res(row, 1) = NA_INTEGER;
    if (el == NA_STRING) continue;

    const void* vmax = vmaxget();
    const char* chars = Rf_translateCharUTF8(el);
    size_t len = std::strlen(chars);
    size_t plen = pat.size();
    if (len < plen) {
      vmaxset(vmax);
      continue;
    }

    // search backwards, this finds the last occurrence including overlapping ones
    const char* last = nullptr;
    for (size_t pos = len - plen + 1; pos-- > 0; ) {
      if (chars[pos] == pat[0] && std::memcmp(chars + pos, pat.data(), plen) == 0) {
        last = chars + pos;
        break;
      }
    }
    if (last == nullptr) {
      vmaxset(vmax);
      continue;
    }

    size_t offset = static_cast<size_t>(last - chars);
    int64_t start = (is_ascii_n(chars, offset) ? static_cast<int64_t>(offset) : count_chars(chars, offset)) + 1;
    res(row, 0) = static_cast<int>(start);
    res(row, 1) = static_cast<int>(start + pat_chars - 1);
    vmaxset(vmax);
  }

  return res;
}

// elements that contain pattern. NA for NA
// [[Rcpp::export(rng = false)]]
Rcpp::LogicalVector stri_detect_fixed_cpp(SEXP str_in, SEXP pattern_in) {

  Rcpp::CharacterVector str = as_chr(str_in);
  Rcpp::CharacterVector pattern = as_chr(pattern_in);
  if (str.size() == 0 || pattern.size() == 0) return Rcpp::LogicalVector(0);

  std::string pat = Rf_translateCharUTF8(STRING_ELT(pattern, 0));
  if (pat.empty()) Rcpp::stop("empty search patterns are not supported");

  int64_t n = static_cast<int64_t>(str.size());
  Rcpp::LogicalVector res(n);

  for (int64_t i = 0; i < n; ++i) {
    SEXP el = STRING_ELT(str, i);
    if (el == NA_STRING) {
      res[i] = NA_LOGICAL;
      continue;
    }
    const void* vmax = vmaxget();
    res[i] = std::strstr(Rf_translateCharUTF8(el), pat.c_str()) != nullptr;
    vmaxset(vmax);
  }

  return res;
}

// unique strings in order of their first occurrence. strings with the same
// bytes are the same CHARSXP once they are UTF-8, so pointers can be compared
// [[Rcpp::export(rng = false)]]
Rcpp::CharacterVector stri_unique_cpp(Rcpp::CharacterVector str) {

  int64_t n = static_cast<int64_t>(str.size());

  // only strings that are neither ASCII nor UTF-8 need a conversion. the
  // converted copy keeps the new strings protected
  bool convert = false;
  for (int64_t i = 0; i < n && !convert; ++i) {
    SEXP el = STRING_ELT(str, i);
    convert = el != NA_STRING && Rf_getCharCE(el) != CE_UTF8 && !is_ascii(CHAR(el));
  }

  Rcpp::CharacterVector src = str;
  if (convert) {
    src = Rcpp::CharacterVector(n);
    for (int64_t i = 0; i < n; ++i) {
      SEXP el = STRING_ELT(str, i);
      if (el != NA_STRING && Rf_getCharCE(el) != CE_UTF8 && !is_ascii(CHAR(el))) {
        const void* vmax = vmaxget();
        el = make_utf8(Rf_translateCharUTF8(el));
        vmaxset(vmax);
      }
      SET_STRING_ELT(src, i, el);
    }
  }

  std::unordered_set<SEXP> seen;
  std::vector<int64_t> keep;
  for (int64_t i = 0; i < n; ++i) {
    if (seen.insert(STRING_ELT(src, i)).second) keep.push_back(i);
  }

  Rcpp::CharacterVector out(static_cast<int64_t>(keep.size()));
  for (size_t j = 0; j < keep.size(); ++j) {
    SET_STRING_ELT(out, static_cast<R_xlen_t>(j), STRING_ELT(src, keep[j]));
  }

  return out;
}

// pads on the left with the single code point pad until a string has width characters
// [[Rcpp::export(rng = false)]]
Rcpp::CharacterVector stri_pad_left_cpp(Rcpp::CharacterVector str, Rcpp::IntegerVector width, Rcpp::CharacterVector pad) {

  std::string pad_utf8 = Rf_translateCharUTF8(STRING_ELT(pad, 0));
  if (count_chars(pad_utf8.c_str(), pad_utf8.size()) != 1) Rcpp::stop("pad must be a single code point");

  int64_t n_str = static_cast<int64_t>(str.size());
  int64_t n_width = static_cast<int64_t>(width.size());
  int64_t n = n_str > n_width ? n_str : n_width;
  Rcpp::CharacterVector res(n);
  std::string buf;

  for (int64_t i = 0; i < n; ++i) {
    SEXP el = STRING_ELT(str, i % n_str);
    int w = width[i % n_width];
    if (el == NA_STRING || w == NA_INTEGER) {
      SET_STRING_ELT(res, i, NA_STRING);
      continue;
    }

    const void* vmax = vmaxget();
    const char* chars = Rf_translateCharUTF8(el);
    int64_t len = count_chars(chars, std::strlen(chars));

    buf.clear();
    for (int64_t k = len; k < w; ++k) buf += pad_utf8;
    buf += chars;
    SET_STRING_ELT(res, i, make_utf8(buf));
    vmaxset(vmax);
  }

  return res;
}

// [[Rcpp::export(rng = false)]]
Rcpp::CharacterVector stri_extract_first_cpp(Rcpp::CharacterVector str, Rcpp::IntegerVector start, Rcpp::IntegerVector len) {
  R_xlen_t n = str.size();
  Rcpp::CharacterVector res(n, NA_STRING);
  const void* vmax = vmaxget();
  for (R_xlen_t i = 0; i < n; ++i) {
    int s = start[i];
    int l = len[i];
    if (s == NA_INTEGER || s < 1 || l == NA_INTEGER || l < 0) continue;
    SEXP el = STRING_ELT(str, i);
    if (el == NA_STRING) continue;
    const char* chars = Rf_translateCharUTF8(el);
    size_t nbytes = std::strlen(chars);
    size_t head = advance_chars(chars, nbytes, 0, static_cast<int64_t>(s) - 1);
    size_t tail = advance_chars(chars, nbytes, head, static_cast<int64_t>(l));
    SET_STRING_ELT(res, i, Rf_mkCharLenCE(chars + head, static_cast<int>(tail - head), CE_UTF8));
    vmaxset(vmax);
  }
  return res;
}

// order of the strings like sorting by code points, but digit sequences are
// compared by their value. NA is last. equal strings keep their order
static int natural_compare(const char* a, const char* b) {
  while (*a != '\0' && *b != '\0') {
    if (*a >= '0' && *a <= '9' && *b >= '0' && *b <= '9') {
      const char* a_end = a;
      while (*a_end >= '0' && *a_end <= '9') ++a_end;
      const char* b_end = b;
      while (*b_end >= '0' && *b_end <= '9') ++b_end;
      while (a < a_end - 1 && *a == '0') ++a;
      while (b < b_end - 1 && *b == '0') ++b;
      ptrdiff_t a_len = a_end - a;
      ptrdiff_t b_len = b_end - b;
      if (a_len != b_len) return a_len < b_len ? -1 : 1;
      int cmp = std::memcmp(a, b, static_cast<size_t>(a_len));
      if (cmp != 0) return cmp < 0 ? -1 : 1;
      a = a_end;
      b = b_end;
    } else {
      unsigned char ca = static_cast<unsigned char>(*a);
      unsigned char cb = static_cast<unsigned char>(*b);
      if (ca != cb) return ca < cb ? -1 : 1;
      ++a;
      ++b;
    }
  }
  if (*a != '\0') return 1;
  if (*b != '\0') return -1;
  return 0;
}

// [[Rcpp::export(rng = false)]]
Rcpp::IntegerVector stri_order_numeric_cpp(Rcpp::CharacterVector str) {

  int64_t n = static_cast<int64_t>(str.size());
  std::vector<const char*> chars(static_cast<size_t>(n));
  for (int64_t i = 0; i < n; ++i) {
    SEXP el = STRING_ELT(str, i);
    chars[static_cast<size_t>(i)] = el == NA_STRING ? nullptr : Rf_translateCharUTF8(el);
  }

  std::vector<int> idx(static_cast<size_t>(n));
  for (int64_t i = 0; i < n; ++i) idx[static_cast<size_t>(i)] = static_cast<int>(i);

  std::stable_sort(idx.begin(), idx.end(), [&chars](int x, int y) {
    const char* cx = chars[static_cast<size_t>(x)];
    const char* cy = chars[static_cast<size_t>(y)];
    if (cx == nullptr || cy == nullptr) return cx != nullptr && cy == nullptr;
    return natural_compare(cx, cy) < 0;
  });

  Rcpp::IntegerVector res(n);
  for (int64_t i = 0; i < n; ++i) res[i] = idx[static_cast<size_t>(i)] + 1;
  return res;
}
