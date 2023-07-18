#pragma once
/**
	@file
	@brief bit vector
	@author MITSUNARI Shigeo(@herumi)
	@license modified new BSD license
	http://opensource.org/licenses/BSD-3-Clause
*/
#include <cybozu/exception.hpp>
#include <algorithm>
#include <vector>
#include <assert.h>

namespace cybozu {

template<class T>
size_t RoundupBit(size_t bitLen)
{
	const size_t unitBitSize = sizeof(T) * 8;
	return (bitLen + unitBitSize - 1) / unitBitSize;
}

template<class T>
T GetMaskBit(size_t bitLen)
{
	assert(bitLen < sizeof(T) * 8);
	return (T(1) << bitLen) - 1;
}

template<class T>
void SetBlockBit(T *buf, size_t bitLen)
{
	const size_t unitBitSize = sizeof(T) * 8;
	const size_t q = bitLen / unitBitSize;
	const size_t r = bitLen % unitBitSize;
	buf[q] |= T(1) << r;
}
template<class T>
void ResetBlockBit(T *buf, size_t bitLen)
{
	const size_t unitBitSize = sizeof(T) * 8;
	const size_t q = bitLen / unitBitSize;
	const size_t r = bitLen % unitBitSize;
	buf[q] &= ~(T(1) << r);
}
template<class T>
bool GetBlockBit(const T *buf, size_t bitLen)
{
	const size_t unitBitSize = sizeof(T) * 8;
	const size_t q = bitLen / unitBitSize;
	const size_t r = bitLen % unitBitSize;
	return (buf[q] & (T(1) << r)) != 0;
}

template<class T>
void CopyBit(T* dst, const T* src, size_t bitLen)
{
	const size_t unitBitSize = sizeof(T) * 8;
	const size_t q = bitLen / unitBitSize;
	const size_t r = bitLen % unitBitSize;
	for (size_t i = 0; i < q; i++) dst[i] = src[i];
	if (r == 0) return;
	dst[q] = src[q] & GetMaskBit<T>(r);
}
/*
	dst[] = (src[] << shift) | ext
	@param dst [out] dst[0..bitLen)
	@param src [in] src[0..bitLen)
	@param bitLen [in] length of src, dst
	@param shift [in] 0 <= shift < unitBitSize
	@param ext [in] or bit
*/
template<class T>
T ShiftLeftBit(T* dst, const T* src, size_t bitLen, size_t shift, T ext = 0)
{
	if (bitLen == 0) return 0;
	const size_t unitBitSize = sizeof(T) * 8;
	if (shift >= unitBitSize) {
		throw cybozu::Exception("ShiftLeftBit:large shift") << shift;
	}
	const size_t n = RoundupBit<T>(bitLen); // n >= 1 because bitLen > 0
	const size_t r = bitLen % unitBitSize;
	const T mask = r > 0 ? GetMaskBit<T>(r) : T(-1);
	if (shift == 0) {
		if (n == 1) {
			dst[0] = (src[0] & mask) | ext;
		} else {
			dst[n - 1] = src[n - 1] & mask;
			for (size_t i = n - 2; i > 0; i--) {
				dst[i] = src[i];
			}
			dst[0] = src[0] | ext;
		}
		return 0;
	}
	const size_t revShift = unitBitSize - shift;
	T prev = src[n - 1] & mask;
	const T ret = prev >> revShift;
	for (size_t i = n - 1; i > 0; i--) {
		T v = src[i - 1];
		dst[i] = (prev << shift) | (v >> revShift);
		prev = v;
	}
	dst[0] = (prev << shift) | ext;
	return ret;
}

namespace bitvector_local {

/*
	dst[] = src[] << shift
	dst[0..shift) does not change

	@param dst [out] dst[shift..bitLen + shift)
	@param src [in] src[0..bitLen)
	@param bitLen [in] read bit size
	@param shift [in] 0 <= shift < unitBitSize
*/
template<class T>
void shiftLeftBit(T* dst, const T* src, size_t bitLen, size_t shift)
{
	const size_t unitBitSize = sizeof(T) * 8;

	assert(bitLen);
	assert(0 < shift && shift < unitBitSize);

	const size_t dstN = RoundupBit<T>(bitLen + shift);
	const size_t srcN = RoundupBit<T>(bitLen);
	const size_t r = bitLen % unitBitSize;
	const T mask = r ? GetMaskBit<T>(r) : T(-1);
	const size_t revShift = unitBitSize - shift;

	T prev = src[srcN - 1] & mask;
	if (dstN > srcN) {
		dst[dstN - 1] = prev >> revShift;
	}
	for (size_t i = srcN - 1; i > 0; i--) {
		T v = src[i - 1];
		dst[i] = (prev << shift) | (v >> revShift);
		prev = v;
	}
	T ext = dst[0] & GetMaskBit<T>(shift);
	dst[0] = (prev << shift) | ext;
}

/*
	dst[] = src[] >> shift

	@param dst [out] dst[0..bitLen)
	@param src [in] src[shift..bitLen + shift)
	@param bitLen [in] write bit size
	@param shift [in] 0 <= shift < unitBitSize
	@note src[bitLen + shift - 1] is accessable
*/
template<class T>
void shiftRightBit(T* dst, const T* src, size_t bitLen, size_t shift)
{
	const size_t unitBitSize = sizeof(T) * 8;

	assert(bitLen);
	assert(0 < shift && shift < unitBitSize);

	const size_t dstN = RoundupBit<T>(bitLen);
	const size_t srcN = RoundupBit<T>(bitLen + shift);// srcN = dstN, dstN + 1
	const size_t r = (bitLen + shift) % unitBitSize;
	const T mask = r ? GetMaskBit<T>(r) : T(-1);
	if (srcN == 1) {
		dst[0] = (src[0] & mask) >> shift;
		return;
	}
	const size_t revShift = unitBitSize - shift;
	T prev = src[0];
	for (size_t i = 0; i < srcN - 2; i++) {
		T v = src[i + 1];
		dst[i] = (prev >> shift) | (v << revShift);
		prev = v;
	}
	// i = srcN - 1
	T v = src[srcN - 1] & mask;
	dst[srcN - 2] = (prev >> shift) | (v << revShift);
	if (srcN == dstN) {
		dst[srcN - 1] = v >> shift;
	}
}

} // cybozu::bitvector_local

template<class T>
class BitVectorT {
	static const size_t unitBitSize = sizeof(T) * 8;
	size_t bitLen_;
	std::vector<T> v_;
public:
	typedef T value_type;
	BitVectorT() : bitLen_(0) {}
	BitVectorT(const T *buf, size_t bitLen)
	{
		init(buf, bitLen);
	}
	void init(const T *buf, size_t bitLen)
	{
		resize(bitLen);
		std::copy(buf, buf + v_.size(), &v_[0]);
	}
	void resize(size_t bitLen)
	{
		bitLen_ = bitLen;
		const size_t n = RoundupBit<T>(bitLen);
		const size_t r = bitLen % unitBitSize;
		v_.resize(n);
		if (r) {
			v_[n - 1] &= GetMaskBit<T>(r);
		}
	}
	void reserve(size_t bitLen)
	{
		v_.reserve(RoundupBit<T>(bitLen));
	}
	bool get(size_t idx) const
	{
		if (idx >= bitLen_) throw cybozu::Exception("BitVectorT:get:bad idx") << idx;
		return GetBlockBit(v_.data(), idx);
	}
	void clear()
	{
		bitLen_ = 0;
		v_.clear();
	}
	void set(size_t idx, bool b)
	{
		if (b) {
			set(idx);
		} else {
			reset(idx);
		}
	}
	// set(idx, true);
	void set(size_t idx)
	{
		if (idx >= bitLen_) throw cybozu::Exception("BitVectorT:set:bad idx") << idx;
		SetBlockBit(v_.data(), idx);
	}
	// set(idx, false);
	void reset(size_t idx)
	{
		if (idx >= bitLen_) throw cybozu::Exception("BitVectorT:reset:bad idx") << idx;
		ResetBlockBit(v_.data(), idx);
	}
	size_t size() const { return bitLen_; }
	const T *getBlock() const { return &v_[0]; }
	T *getBlock() { return &v_[0]; }
	size_t getBlockSize() const { return v_.size(); }
	/*
		append src[0, bitLen)
	*/
	void append(const T* src, size_t bitLen)
	{
		if (bitLen == 0) return;
		const size_t q = bitLen_ / unitBitSize;
		const size_t r = bitLen_ % unitBitSize;
		resize(bitLen_ + bitLen);
		if (r == 0) {
			CopyBit<T>(&v_[q], src, bitLen);
			return;
		}
		bitvector_local::shiftLeftBit<T>(&v_[q], src, bitLen, r);
	}
	/*
		append src & mask(bitLen)
	*/
	void append(uint64_t src, size_t bitLen)
	{
		if (bitLen == 0) return;
		if (bitLen > unitBitSize) {
			throw cybozu::Exception("BitVectorT:append:bad bitLen") << bitLen;
		}
		if (bitLen < unitBitSize) {
			src &= GetMaskBit<T>(bitLen);
		}
		const size_t q = bitLen_ / unitBitSize;
		const size_t r = bitLen_ % unitBitSize;
		resize(bitLen_ + bitLen);
		if (r == 0) {
			v_[q] = T(src);
			return;
		}
		v_[q] |= T(src << r);
		if (r + bitLen > unitBitSize) {
			v_[q + 1] = T(src >> (unitBitSize - r));
		}
	}
	/*
		append bitVector
	*/
	void append(const BitVectorT<T>& v)
	{
		append(v.getBlock(), v.size());
	}
	/*
		dst[0, bitLen) = vec[pos, pos + bitLen)
	*/
	void extract(T* dst, size_t pos, size_t bitLen) const
	{
		if (bitLen == 0) return;
		if (pos + bitLen > bitLen_) {
			throw cybozu::Exception("BitVectorT:extract:bad range") << bitLen << pos << bitLen_;
		}
		const size_t q = pos / unitBitSize;
		const size_t r = pos % unitBitSize;
		if (r == 0) {
			CopyBit<T>(dst, &v_[q], bitLen);
			return;
		}
		bitvector_local::shiftRightBit<T>(dst, &v_[q], bitLen, r);
	}
	/*
		dst = vec[pos, pos + bitLen)
	*/
	void extract(BitVectorT<T>& dst, size_t pos, size_t bitLen) const
	{
		dst.resize(bitLen);
		extract(dst.getBlock(), pos, bitLen);
	}
	/*
		return vec[pos, pos + bitLen)
	*/
	T extract(size_t pos, size_t bitLen) const
	{
		if (bitLen == 0) return 0;
		if (bitLen > unitBitSize || pos + bitLen > bitLen_) {
			throw cybozu::Exception("BitVectorT:extract:bad range") << bitLen << pos << bitLen_;
		}
		const size_t q = pos / unitBitSize;
		const size_t r = pos % unitBitSize;
		T v;
		if (r == 0) {
			v = v_[q];
		} else if (q == v_.size() - 1) {
			v = v_[q] >> r;
		} else {
			v = (v_[q] >> r) | v_[q + 1] << (unitBitSize - r);
		}
		if (bitLen < unitBitSize) {
			v &= GetMaskBit<T>(bitLen);
		}
		return v;
	}
	bool operator==(const BitVectorT<T>& rhs) const { return v_ == rhs.v_; }
	bool operator!=(const BitVectorT<T>& rhs) const { return v_ != rhs.v_; }
};

#if (CYBOZU_OS_BIT == 32)
typedef BitVectorT<uint32_t> BitVector;
#else
typedef BitVectorT<uint64_t> BitVector;
#endif

} // cybozu
