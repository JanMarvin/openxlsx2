#ifndef XLS_CFB_H
#define XLS_CFB_H

#include <algorithm>
#include <cstdint>
#include <string>
#include <vector>

// An xls file is an OLE2 "compound file": a small file system inside a file.
// The workbook is the stream called "Workbook". This is a minimal reader for it.
//
// The file is cut into sectors. A sector allocation table (FAT) says which sector
// follows which, so a stream is a chain of sectors. Streams smaller than the
// cutoff (4096 bytes) are stored in 64 byte "mini sectors" inside the root
// entry's stream and chained by a second table, the mini FAT.

inline uint32_t cfb_u16(const std::vector<uint8_t>& f, size_t pos) {
  if (pos + 2 > f.size()) Rcpp::stop("xls: file is truncated");
  return static_cast<uint32_t>(f[pos]) | (static_cast<uint32_t>(f[pos + 1]) << 8);
}

inline uint32_t cfb_u32(const std::vector<uint8_t>& f, size_t pos) {
  return cfb_u16(f, pos) | (cfb_u16(f, pos + 2) << 16);
}

// follow a chain of sectors in the FAT and return their content
inline std::vector<uint8_t> cfb_chain(const std::vector<uint8_t>& file, const std::vector<uint32_t>& fat,
                                      size_t sec_size, uint32_t start) {
  std::vector<uint8_t> out;
  uint32_t sec = start;
  size_t steps = 0;
  while (sec < 0xFFFFFFFAu) {
    size_t off = (static_cast<size_t>(sec) + 1) * sec_size;
    // #nocov start: corrupt file
    if (sec >= fat.size() || ++steps > file.size() / sec_size || off + sec_size > file.size())
      Rcpp::stop("xls: corrupt sector chain");
    // #nocov end
    out.insert(out.end(), file.begin() + static_cast<std::ptrdiff_t>(off),
               file.begin() + static_cast<std::ptrdiff_t>(off + sec_size));
    sec = fat[sec];
  }
  return out;
}

// return the content of the stream with the given name, empty if it does not exist
inline std::vector<uint8_t> cfb_stream(const std::vector<uint8_t>& file, const std::string& name) {
  static const uint8_t magic[8] = {0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1};
  // #nocov start: not an OLE2 file
  if (file.size() < 512 || !std::equal(magic, magic + 8, file.begin()))
    Rcpp::stop("xls: not an OLE2 compound file");
  // #nocov end

  uint32_t sec_shift = cfb_u16(file, 0x1E);
  uint32_t mini_shift = cfb_u16(file, 0x20);
  // #nocov start: corrupt file
  if (sec_shift < 7 || sec_shift > 16 || mini_shift < 1 || mini_shift > sec_shift)
    Rcpp::stop("xls: invalid compound file header");
  // #nocov end
  size_t sec_size = static_cast<size_t>(1) << sec_shift;
  size_t mini_size = static_cast<size_t>(1) << mini_shift;

  uint32_t dir_start = cfb_u32(file, 0x30);
  uint32_t mini_cutoff = cfb_u32(file, 0x38);
  uint32_t minifat_start = cfb_u32(file, 0x3C);
  uint32_t difat_start = cfb_u32(file, 0x44);
  uint32_t difat_count = cfb_u32(file, 0x48);

  // the header lists the first 109 FAT sectors, more are listed in DIFAT sectors
  std::vector<uint32_t> fat_sectors;
  for (size_t i = 0; i < 109; ++i) {
    uint32_t sec = cfb_u32(file, 0x4C + 4 * i);
    if (sec < 0xFFFFFFFAu) fat_sectors.push_back(sec);
  }
  uint32_t difat = difat_start;
  if (difat_count > file.size() / sec_size) Rcpp::stop("xls: corrupt DIFAT");
  for (uint32_t n = 0; n < difat_count && difat < 0xFFFFFFFAu; ++n) {
    size_t off = (static_cast<size_t>(difat) + 1) * sec_size;
    for (size_t i = 0; i < sec_size / 4 - 1; ++i) {
      uint32_t sec = cfb_u32(file, off + 4 * i);
      if (sec < 0xFFFFFFFAu) fat_sectors.push_back(sec);
    }
    difat = cfb_u32(file, off + sec_size - 4);
  }

  std::vector<uint32_t> fat;
  for (size_t i = 0; i < fat_sectors.size(); ++i) {
    size_t off = (static_cast<size_t>(fat_sectors[i]) + 1) * sec_size;
    for (size_t j = 0; j < sec_size / 4; ++j) fat.push_back(cfb_u32(file, off + 4 * j));
  }

  // directory: 128 byte entries. name: utf-16 (64 bytes) at 0, length at 64, type
  // at 66 (2 stream, 5 root), first sector at 116, size at 120
  std::vector<uint8_t> dir = cfb_chain(file, fat, sec_size, dir_start);

  // the entries form a tree, entry 0 is the root. children are found through
  // the child entry and the left and right siblings of it. Streams are looked
  // up by their path, so a stream with the same name in an embedded object is
  // not mistaken for it. Names are compared without regard to case.
  size_t n_entries = dir.size() / 128;
  std::vector<std::string> parts;
  for (size_t at = 0; at <= name.size();) {
    size_t slash = name.find('/', at);
    if (slash == std::string::npos) slash = name.size();
    parts.push_back(name.substr(at, slash - at));
    at = slash + 1;
  }

  auto upper = [](std::string x) {
    for (size_t i = 0; i < x.size(); ++i)
      if (x[i] >= 'a' && x[i] <= 'z') x[i] = static_cast<char>(x[i] - 'a' + 'A');
    return x;
  };

  uint32_t found_start = 0, found_size = 0;
  bool found = n_entries > 0;
  uint32_t root_start = 0, root_size = 0;
  if (found) {
    root_start = cfb_u32(dir, 116);
    root_size = cfb_u32(dir, 120);
  }
  size_t current = 0;
  for (size_t p = 0; found && p < parts.size(); ++p) {
    std::vector<size_t> todo;
    todo.push_back(cfb_u32(dir, current * 128 + 76));
    size_t visited = 0;
    found = false;
    while (!todo.empty()) {
      size_t id = todo.back();
      todo.pop_back();
      if (id >= n_entries || ++visited > n_entries) continue;
      size_t off = id * 128;

      std::string entry;
      size_t len = cfb_u16(dir, off + 64);
      for (size_t i = 0; i + 2 <= len && i < 64 && dir[off + i] != 0; i += 2) entry.push_back(static_cast<char>(dir[off + i]));
      if (upper(entry) == upper(parts[p])) {
        current = id;
        found = dir[off + 66] == (p + 1 < parts.size() ? 1 : 2);
        break;
      }
      todo.push_back(cfb_u32(dir, off + 68));
      todo.push_back(cfb_u32(dir, off + 72));
    }
  }
  if (!found) return std::vector<uint8_t>();
  found_start = cfb_u32(dir, current * 128 + 116);
  found_size = cfb_u32(dir, current * 128 + 120);

  if (found_size >= mini_cutoff) {
    std::vector<uint8_t> out = cfb_chain(file, fat, sec_size, found_start);
    if (out.size() > found_size) out.resize(found_size);
    return out;
  }

  std::vector<uint8_t> mini_stream = cfb_chain(file, fat, sec_size, root_start);
  if (mini_stream.size() > root_size) mini_stream.resize(root_size);

  std::vector<uint8_t> minifat_bytes = cfb_chain(file, fat, sec_size, minifat_start);
  std::vector<uint32_t> minifat;
  for (size_t i = 0; i + 4 <= minifat_bytes.size(); i += 4) minifat.push_back(cfb_u32(minifat_bytes, i));

  std::vector<uint8_t> out;
  uint32_t sec = found_start;
  size_t steps = 0;
  while (sec < 0xFFFFFFFAu) {
    size_t off = static_cast<size_t>(sec) * mini_size;
    // #nocov start: corrupt file
    if (sec >= minifat.size() || ++steps > mini_stream.size() / mini_size || off + mini_size > mini_stream.size())
      Rcpp::stop("xls: corrupt mini sector chain");
    // #nocov end
    out.insert(out.end(), mini_stream.begin() + static_cast<std::ptrdiff_t>(off),
               mini_stream.begin() + static_cast<std::ptrdiff_t>(off + mini_size));
    sec = minifat[sec];
  }
  if (out.size() > found_size) out.resize(found_size);
  return out;
}

#endif
