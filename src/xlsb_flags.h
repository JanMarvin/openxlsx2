#ifndef XLSB_FLAGS_H
#define XLSB_FLAGS_H

#include <cstdint>  // For uint8_t, uint16_t, uint32_t, uintmax_t
#include <type_traits>  // For std::enable_if, std::is_integral

#if __cplusplus >= 201103L
  #if defined(__GNUC__) && !defined(__clang__) && (__GNUC__ < 6)
    #define XL_CONSTEXPR /* empty */
  #else
    #define XL_CONSTEXPR constexpr
  #endif
#else
  #define XL_CONSTEXPR /* empty */
#endif


// --- Base Helper Struct: FlagProxy ---
struct FlagProxy {
    const uintmax_t raw_flags_value;
    const uintmax_t bit_mask_storage;

    template<typename T, typename = typename std::enable_if<std::is_integral<T>::value>::type>
    XL_CONSTEXPR explicit FlagProxy(const T flags, uintmax_t mask)
        : raw_flags_value(static_cast<uintmax_t>(flags)),
          bit_mask_storage(mask) {}

    XL_CONSTEXPR operator bool() const {
        return (raw_flags_value & bit_mask_storage) != 0;
    }
};

// --- Base Helper Struct: ValueProxy ---
template<typename T_Return>
struct ValueProxy {
    const uintmax_t raw_flags_value;
    const uintmax_t shift_amount;
    const uintmax_t width_mask;

    template<typename T_Raw, typename = typename std::enable_if<std::is_integral<T_Raw>::value>::type>
    XL_CONSTEXPR explicit ValueProxy(const T_Raw flags, uintmax_t shift, uintmax_t width)
        : raw_flags_value(static_cast<uintmax_t>(flags)),
          shift_amount(shift),
          width_mask((width >= (sizeof(uintmax_t) * 8)) ? ~static_cast<uintmax_t>(0) : ((1ULL << width) - 1)) {}

    XL_CONSTEXPR operator T_Return() const {
        return static_cast<T_Return>((raw_flags_value >> shift_amount) & width_mask);
    }
};

// --- Conversion for StyleFlagsUnion ---
struct PaneFlags {
  const uint8_t raw_flags;

  explicit PaneFlags(uint8_t initial_flags = 0) : raw_flags(initial_flags) {}

  XL_CONSTEXPR FlagProxy fFrozen() const    { return FlagProxy(raw_flags, 1 << 0); }
  XL_CONSTEXPR FlagProxy fFrozenNoSplit() const   { return FlagProxy(raw_flags, 1 << 1); }
  XL_CONSTEXPR ValueProxy<uint8_t> reserved() const { return ValueProxy<uint8_t>(raw_flags, 2, 6); }
};

// --- Conversion for StyleFlagsUnion ---
struct StyleFlags {
    const uint16_t raw_flags;

    explicit StyleFlags(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    XL_CONSTEXPR FlagProxy fBuiltIn() const  { return FlagProxy(raw_flags, 1 << 0); }
    XL_CONSTEXPR FlagProxy fHidden() const   { return FlagProxy(raw_flags, 1 << 1); }
    XL_CONSTEXPR FlagProxy fCustom() const   { return FlagProxy(raw_flags, 1 << 2); }
    XL_CONSTEXPR ValueProxy<uint16_t> unused() const { return ValueProxy<uint16_t>(raw_flags, 3, 13); }
};

// --- Conversion for xfGrbitAtrUnion ---
struct xfGrbitAtrFlags {
    const uint8_t raw_flags;

    explicit xfGrbitAtrFlags(uint8_t initial_flags = 0) : raw_flags(initial_flags) {}

    XL_CONSTEXPR FlagProxy bit1() const { return FlagProxy(raw_flags, 1 << 0); }
    XL_CONSTEXPR FlagProxy bit2() const { return FlagProxy(raw_flags, 1 << 1); }
    XL_CONSTEXPR FlagProxy bit3() const { return FlagProxy(raw_flags, 1 << 2); }
    XL_CONSTEXPR FlagProxy bit4() const { return FlagProxy(raw_flags, 1 << 3); }
    XL_CONSTEXPR FlagProxy bit5() const { return FlagProxy(raw_flags, 1 << 4); }
    XL_CONSTEXPR FlagProxy bit6() const { return FlagProxy(raw_flags, 1 << 5); }
};

// --- Conversion for GrbitFmlaUnion ---
struct GrbitFmlaFlags {
    const uint16_t raw_flags;

    explicit GrbitFmlaFlags(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    XL_CONSTEXPR FlagProxy reserved() const    { return FlagProxy(raw_flags, 1 << 0); }
    XL_CONSTEXPR FlagProxy fAlwaysCalc() const { return FlagProxy(raw_flags, 1 << 1); }
};

// --- Conversion for ColRelShortUnion ---
struct ColRelShortFlags {
    const uint16_t raw_flags;

    explicit ColRelShortFlags(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    XL_CONSTEXPR ValueProxy<uint16_t> col() const { return ValueProxy<uint16_t>(raw_flags, 0, 14); }
    XL_CONSTEXPR FlagProxy fColRel() const        { return FlagProxy(raw_flags, 1 << 14); }
    XL_CONSTEXPR FlagProxy fRwRel() const         { return FlagProxy(raw_flags, 1 << 15); }
};

// --- Conversion for BrtBeginCsView ---
struct BrtBeginCsViewFlags {
  const uint16_t raw_flags;

  explicit BrtBeginCsViewFlags(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

  XL_CONSTEXPR FlagProxy fSelected() const        { return FlagProxy(raw_flags, 1 << 0); }
  XL_CONSTEXPR ValueProxy<uint16_t> unused() const { return ValueProxy<uint16_t>(raw_flags, 1, 15); }
};

// --- Conversion for BrtOleObject ---
struct BrtOleObjectFlags {
  const uint16_t raw_flags;

  explicit BrtOleObjectFlags(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

  XL_CONSTEXPR FlagProxy fLinked() const        { return FlagProxy(raw_flags, 1 << 0); }
  XL_CONSTEXPR FlagProxy fAutoLoad() const         { return FlagProxy(raw_flags, 1 << 1); }
  XL_CONSTEXPR ValueProxy<uint16_t> reserved() const { return ValueProxy<uint16_t>(raw_flags, 2, 14); }
};

// --- Conversion for BrtRowHdrUnion ---
struct BrtRowHdrFlags {
    const uint16_t raw_flags;

    explicit BrtRowHdrFlags(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    XL_CONSTEXPR FlagProxy fExtraAsc() const   { return FlagProxy(raw_flags, 1 << 0); }
    XL_CONSTEXPR FlagProxy fExtraDsc() const   { return FlagProxy(raw_flags, 1 << 1); }
    XL_CONSTEXPR ValueProxy<uint8_t> reserved1() const { return ValueProxy<uint8_t>(raw_flags, 2, 6); }
    XL_CONSTEXPR ValueProxy<uint8_t> iOutLevel() const { return ValueProxy<uint8_t>(raw_flags, 8, 3); }
    XL_CONSTEXPR FlagProxy fCollapsed() const  { return FlagProxy(raw_flags, 1 << 11); }
    XL_CONSTEXPR FlagProxy fDyZero() const     { return FlagProxy(raw_flags, 1 << 12); }
    XL_CONSTEXPR FlagProxy fUnsynced() const   { return FlagProxy(raw_flags, 1 << 13); }
    XL_CONSTEXPR FlagProxy fGhostDirty() const { return FlagProxy(raw_flags, 1 << 14); }
    XL_CONSTEXPR FlagProxy fReserved() const   { return FlagProxy(raw_flags, 1 << 15); }
};

// --- Conversion for BrtColInfoUnion ---
struct BrtColInfoFlags {
    const uint16_t raw_flags;

    explicit BrtColInfoFlags(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    XL_CONSTEXPR FlagProxy fHidden() const    { return FlagProxy(raw_flags, 1 << 0); }
    XL_CONSTEXPR FlagProxy fUserSet() const   { return FlagProxy(raw_flags, 1 << 1); }
    XL_CONSTEXPR FlagProxy fBestFit() const   { return FlagProxy(raw_flags, 1 << 2); }
    XL_CONSTEXPR FlagProxy fPhonetic() const  { return FlagProxy(raw_flags, 1 << 3); }
    XL_CONSTEXPR ValueProxy<uint8_t> reserved1() const { return ValueProxy<uint8_t>(raw_flags, 4, 4); }
    XL_CONSTEXPR ValueProxy<uint8_t> iOutLevel() const { return ValueProxy<uint8_t>(raw_flags, 8, 3); }
    XL_CONSTEXPR FlagProxy unused() const     { return FlagProxy(raw_flags, 1 << 11); }
    XL_CONSTEXPR FlagProxy fCollapsed() const { return FlagProxy(raw_flags, 1 << 12); }
    XL_CONSTEXPR ValueProxy<uint8_t> reserved2() const { return ValueProxy<uint8_t>(raw_flags, 13, 3); }
};

// --- Conversion for BrtNameUnion ---
struct BrtNameFlags {
    const uint16_t raw_flags;

    explicit BrtNameFlags(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    XL_CONSTEXPR FlagProxy fHidden() const    { return FlagProxy(raw_flags, 1 << 0); }
    XL_CONSTEXPR FlagProxy fFunc() const      { return FlagProxy(raw_flags, 1 << 1); }
    XL_CONSTEXPR FlagProxy fOB() const        { return FlagProxy(raw_flags, 1 << 2); }
    XL_CONSTEXPR FlagProxy fProc() const      { return FlagProxy(raw_flags, 1 << 3); }
    XL_CONSTEXPR FlagProxy fCalcExp() const   { return FlagProxy(raw_flags, 1 << 4); }
    XL_CONSTEXPR FlagProxy fBuiltin() const   { return FlagProxy(raw_flags, 1 << 5); }
    XL_CONSTEXPR ValueProxy<uint16_t> fgrp() const { return ValueProxy<uint16_t>(raw_flags, 6, 9); }
    XL_CONSTEXPR FlagProxy fPublished() const { return FlagProxy(raw_flags, 1 << 15); }
};

// --- Conversion for BrtName2Union ---
struct BrtName2Flags {
    const uint16_t raw_flags;

    explicit BrtName2Flags(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    XL_CONSTEXPR FlagProxy fWorkbookParam() const { return FlagProxy(raw_flags, 1 << 0); }
    XL_CONSTEXPR FlagProxy fFutureFunction() const { return FlagProxy(raw_flags, 1 << 1); }
    XL_CONSTEXPR ValueProxy<uint16_t> reserved() const { return ValueProxy<uint16_t>(raw_flags, 2, 14); }
};

// --- Conversion for FontFlagsUnion ---
struct FontFlags {
    const uint16_t raw_flags;

    explicit FontFlags(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    XL_CONSTEXPR FlagProxy unused1() const   { return FlagProxy(raw_flags, 1 << 0); }
    XL_CONSTEXPR FlagProxy fItalic() const   { return FlagProxy(raw_flags, 1 << 1); }
    XL_CONSTEXPR FlagProxy unused2() const   { return FlagProxy(raw_flags, 1 << 2); }
    XL_CONSTEXPR FlagProxy fStrikeout() const { return FlagProxy(raw_flags, 1 << 3); }
    XL_CONSTEXPR FlagProxy fOutline() const  { return FlagProxy(raw_flags, 1 << 4); }
    XL_CONSTEXPR FlagProxy fShadow() const   { return FlagProxy(raw_flags, 1 << 5); }
    XL_CONSTEXPR FlagProxy fCondense() const { return FlagProxy(raw_flags, 1 << 6); }
    XL_CONSTEXPR FlagProxy fExtend() const   { return FlagProxy(raw_flags, 1 << 7); }
};

// --- Conversion for XFUnion ---
struct XFFlags {
    const uint32_t raw_flags;

    explicit XFFlags(uint32_t initial_flags = 0) : raw_flags(initial_flags) {}

    XL_CONSTEXPR ValueProxy<uint8_t> alc() const          { return ValueProxy<uint8_t>(raw_flags, 0, 3); }
    XL_CONSTEXPR ValueProxy<uint8_t> alcv() const         { return ValueProxy<uint8_t>(raw_flags, 3, 3); }
    XL_CONSTEXPR FlagProxy fWrap() const                  { return FlagProxy(raw_flags, 1 << 6); }
    XL_CONSTEXPR FlagProxy fJustLast() const              { return FlagProxy(raw_flags, 1 << 7); }
    XL_CONSTEXPR FlagProxy fShrinkToFit() const           { return FlagProxy(raw_flags, 1 << 8); }
    XL_CONSTEXPR FlagProxy fMergeCell() const             { return FlagProxy(raw_flags, 1 << 9); }
    XL_CONSTEXPR ValueProxy<uint8_t> iReadingOrder() const { return ValueProxy<uint8_t>(raw_flags, 10, 2); }
    XL_CONSTEXPR FlagProxy fLocked() const                { return FlagProxy(raw_flags, 1 << 12); }
    XL_CONSTEXPR FlagProxy fHidden() const                { return FlagProxy(raw_flags, 1 << 13); }
    XL_CONSTEXPR FlagProxy fSxButton() const              { return FlagProxy(raw_flags, 1 << 14); }
    XL_CONSTEXPR FlagProxy f123Prefix() const             { return FlagProxy(raw_flags, 1 << 15); }
    XL_CONSTEXPR ValueProxy<uint8_t> xfGrbitAtr() const   { return ValueProxy<uint8_t>(raw_flags, 16, 6); }
    XL_CONSTEXPR ValueProxy<uint16_t> unused() const      { return ValueProxy<uint16_t>(raw_flags, 22, 10); }
};

// --- Conversion for BrtWsProp1Union ---
struct BrtWsProp1Flags {
    const uint16_t raw_flags;

    explicit BrtWsProp1Flags(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    XL_CONSTEXPR FlagProxy fShowAutoBreaks() const    { return FlagProxy(raw_flags, 1 << 0); }
    XL_CONSTEXPR ValueProxy<uint8_t> rserved1() const { return ValueProxy<uint8_t>(raw_flags, 1, 2); }
    XL_CONSTEXPR FlagProxy fPublish() const           { return FlagProxy(raw_flags, 1 << 3); }
    XL_CONSTEXPR FlagProxy fDialog() const            { return FlagProxy(raw_flags, 1 << 4); }
    XL_CONSTEXPR FlagProxy fApplyStyles() const       { return FlagProxy(raw_flags, 1 << 5); }
    XL_CONSTEXPR FlagProxy fRowSumsBelow() const      { return FlagProxy(raw_flags, 1 << 6); }
    XL_CONSTEXPR FlagProxy fColSumsRight() const      { return FlagProxy(raw_flags, 1 << 7); }
    XL_CONSTEXPR FlagProxy fFitToPage() const         { return FlagProxy(raw_flags, 1 << 8); }
    XL_CONSTEXPR ValueProxy<uint8_t> reserved2() const { return ValueProxy<uint8_t>(raw_flags, 9, 1); }
    XL_CONSTEXPR FlagProxy fShowOutlineSymbols() const { return FlagProxy(raw_flags, 1 << 10); }
    XL_CONSTEXPR ValueProxy<uint8_t> reserved3() const { return ValueProxy<uint8_t>(raw_flags, 11, 1); }
    XL_CONSTEXPR FlagProxy fSyncHoriz() const         { return FlagProxy(raw_flags, 1 << 12); }
    XL_CONSTEXPR FlagProxy fSyncVert() const          { return FlagProxy(raw_flags, 1 << 13); }
    XL_CONSTEXPR FlagProxy fAltExprEval() const       { return FlagProxy(raw_flags, 1 << 14); }
    XL_CONSTEXPR FlagProxy fAltFormulaEntry() const   { return FlagProxy(raw_flags, 1 << 15); }
};

// --- Conversion for BrtWsProp2Union ---
struct BrtWsProp2Flags {
    const uint8_t raw_flags;

    explicit BrtWsProp2Flags(uint8_t initial_flags = 0) : raw_flags(initial_flags) {}

    XL_CONSTEXPR FlagProxy fFilterMode() const    { return FlagProxy(raw_flags, 1 << 0); }
    XL_CONSTEXPR FlagProxy fCondFmtCalc() const   { return FlagProxy(raw_flags, 1 << 1); }
    XL_CONSTEXPR ValueProxy<uint8_t> reserved4() const { return ValueProxy<uint8_t>(raw_flags, 2, 6); }
};

// --- Conversion for BrtBeginWsViewUnion ---
struct BrtBeginWsViewFlags {
    const uint16_t raw_flags;

    explicit BrtBeginWsViewFlags(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    XL_CONSTEXPR FlagProxy fWnProt() const           { return FlagProxy(raw_flags, 1 << 0); }
    XL_CONSTEXPR FlagProxy fDspFmla() const          { return FlagProxy(raw_flags, 1 << 1); }
    XL_CONSTEXPR FlagProxy fDspGrid() const          { return FlagProxy(raw_flags, 1 << 2); }
    XL_CONSTEXPR FlagProxy fDspRwCol() const         { return FlagProxy(raw_flags, 1 << 3); }
    XL_CONSTEXPR FlagProxy fDspZeros() const         { return FlagProxy(raw_flags, 1 << 4); }
    XL_CONSTEXPR FlagProxy fRightToLeft() const      { return FlagProxy(raw_flags, 1 << 5); }
    XL_CONSTEXPR FlagProxy fSelected() const         { return FlagProxy(raw_flags, 1 << 6); }
    XL_CONSTEXPR FlagProxy fDspRuler() const         { return FlagProxy(raw_flags, 1 << 7); }
    XL_CONSTEXPR FlagProxy fDspGuts() const          { return FlagProxy(raw_flags, 1 << 8); }
    XL_CONSTEXPR FlagProxy fDefaultHdr() const       { return FlagProxy(raw_flags, 1 << 9); }
    XL_CONSTEXPR FlagProxy fWhitespaceHidden() const { return FlagProxy(raw_flags, 1 << 10); }
    XL_CONSTEXPR ValueProxy<uint8_t> reserved1() const { return ValueProxy<uint8_t>(raw_flags, 11, 5); }
};

// --- Conversion for BrtTableStyleClientUnion ---
struct BrtTableStyleClientFlags {
    const uint16_t raw_flags;

    explicit BrtTableStyleClientFlags(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    XL_CONSTEXPR FlagProxy fFirstColumn() const    { return FlagProxy(raw_flags, 1 << 0); }
    XL_CONSTEXPR FlagProxy fLastColumn() const     { return FlagProxy(raw_flags, 1 << 1); }
    XL_CONSTEXPR FlagProxy fRowStripes() const     { return FlagProxy(raw_flags, 1 << 2); }
    XL_CONSTEXPR FlagProxy fColumnStripes() const  { return FlagProxy(raw_flags, 1 << 3); }
    XL_CONSTEXPR FlagProxy fRowHeaders() const     { return FlagProxy(raw_flags, 1 << 4); }
    XL_CONSTEXPR FlagProxy fColumnHeaders() const  { return FlagProxy(raw_flags, 1 << 5); }
    XL_CONSTEXPR ValueProxy<uint16_t> reserved() const { return ValueProxy<uint16_t>(raw_flags, 6, 10); }
};

// --- Conversion for BrtBeginCFRuleUnion ---
struct BrtBeginCFRuleFlags {
    const uint16_t raw_flags;

    explicit BrtBeginCFRuleFlags(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    XL_CONSTEXPR FlagProxy reserved3() const  { return FlagProxy(raw_flags, 1 << 0); }
    XL_CONSTEXPR FlagProxy fStopTrue() const  { return FlagProxy(raw_flags, 1 << 1); }
    XL_CONSTEXPR FlagProxy fAbove() const     { return FlagProxy(raw_flags, 1 << 2); }
    XL_CONSTEXPR FlagProxy fBottom() const    { return FlagProxy(raw_flags, 1 << 3); }
    XL_CONSTEXPR FlagProxy fPercent() const   { return FlagProxy(raw_flags, 1 << 4); }
    XL_CONSTEXPR ValueProxy<uint16_t> reserved4() const { return ValueProxy<uint16_t>(raw_flags, 5, 11); }
};

// --- Conversion for BrtWbPropUnion ---
struct BrtWbPropFlags {
    const uint32_t raw_flags;

    explicit BrtWbPropFlags(uint32_t initial_flags = 0) : raw_flags(initial_flags) {}

    XL_CONSTEXPR FlagProxy f1904() const                { return FlagProxy(raw_flags, 1 << 0); }
    XL_CONSTEXPR FlagProxy reserved1() const            { return FlagProxy(raw_flags, 1 << 1); }
    XL_CONSTEXPR FlagProxy fHideBorderUnselLists() const { return FlagProxy(raw_flags, 1 << 2); }
    XL_CONSTEXPR FlagProxy fFilterPrivacy() const       { return FlagProxy(raw_flags, 1 << 3); }
    XL_CONSTEXPR FlagProxy fBuggedUserAboutSolution() const { return FlagProxy(raw_flags, 1 << 4); }
    XL_CONSTEXPR FlagProxy fShowInkAnnotation() const   { return FlagProxy(raw_flags, 1 << 5); }
    XL_CONSTEXPR FlagProxy fBackup() const              { return FlagProxy(raw_flags, 1 << 6); }
    XL_CONSTEXPR FlagProxy fNoSaveSup() const           { return FlagProxy(raw_flags, 1 << 7); }
    XL_CONSTEXPR ValueProxy<uint8_t> grbitUpdateLinks() const { return ValueProxy<uint8_t>(raw_flags, 8, 2); }
    XL_CONSTEXPR FlagProxy fHidePivotTableFList() const { return FlagProxy(raw_flags, 1 << 10); }
    XL_CONSTEXPR FlagProxy fPublishedBookItems() const  { return FlagProxy(raw_flags, 1 << 11); }
    XL_CONSTEXPR FlagProxy fCheckCompat() const         { return FlagProxy(raw_flags, 1 << 12); }
    XL_CONSTEXPR ValueProxy<uint8_t> mdDspObj() const   { return ValueProxy<uint8_t>(raw_flags, 13, 2); }
    XL_CONSTEXPR FlagProxy fShowPivotChartFilter() const { return FlagProxy(raw_flags, 1 << 15); }
    XL_CONSTEXPR FlagProxy fAutoCompressPictures() const { return FlagProxy(raw_flags, 1 << 16); }
    XL_CONSTEXPR FlagProxy reserved2() const            { return FlagProxy(raw_flags, 1 << 17); }
    XL_CONSTEXPR FlagProxy fRefreshAll() const          { return FlagProxy(raw_flags, 1 << 18); }
    XL_CONSTEXPR ValueProxy<uint16_t> unused() const    { return ValueProxy<uint16_t>(raw_flags, 19, 13); }
};

// --- Conversion for BrtDValUnion ---
struct BrtDValFlags {
    const uint32_t raw_flags;

    explicit BrtDValFlags(uint32_t initial_flags = 0) : raw_flags(initial_flags) {}

    XL_CONSTEXPR ValueProxy<uint8_t> valType() const       { return ValueProxy<uint8_t>(raw_flags, 0, 4); }
    XL_CONSTEXPR ValueProxy<uint8_t> errStyle() const      { return ValueProxy<uint8_t>(raw_flags, 4, 3); }
    XL_CONSTEXPR FlagProxy unused() const          { return FlagProxy(raw_flags, 1 << 7); }
    XL_CONSTEXPR FlagProxy fAllowBlank() const     { return FlagProxy(raw_flags, 1 << 8); }
    XL_CONSTEXPR FlagProxy fSuppressCombo() const  { return FlagProxy(raw_flags, 1 << 9); }
    XL_CONSTEXPR ValueProxy<uint8_t> mdImeMode() const     { return ValueProxy<uint8_t>(raw_flags, 10, 8); }
    XL_CONSTEXPR FlagProxy fShowInputMsg() const   { return FlagProxy(raw_flags, 1 << 18); }
    XL_CONSTEXPR FlagProxy fShowErrorMsg() const   { return FlagProxy(raw_flags, 1 << 19); }
    XL_CONSTEXPR ValueProxy<uint8_t> typOperator() const   { return ValueProxy<uint8_t>(raw_flags, 20, 4); }
    XL_CONSTEXPR FlagProxy fDVMinFmla() const      { return FlagProxy(raw_flags, 1 << 24); }
    XL_CONSTEXPR FlagProxy fDVMaxFmla() const      { return FlagProxy(raw_flags, 1 << 25); }
    XL_CONSTEXPR ValueProxy<uint8_t> reserved() const      { return ValueProxy<uint8_t>(raw_flags, 26, 6); }
};

// --- Conversion for FRTVersionUnion ---
struct FRTVersionFlags {
    const uint16_t raw_flags;

    explicit FRTVersionFlags(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    XL_CONSTEXPR ValueProxy<uint16_t> product() const { return ValueProxy<uint16_t>(raw_flags, 0, 15); }
    XL_CONSTEXPR FlagProxy reserved() const { return FlagProxy(raw_flags, 1 << 15); }
};

// --- Conversion for FRTHeaderUnion ---
struct FRTHeaderFlags {
    const uint32_t raw_flags;

    explicit FRTHeaderFlags(uint32_t initial_flags = 0) : raw_flags(initial_flags) {}

    XL_CONSTEXPR FlagProxy fRef() const     { return FlagProxy(raw_flags, 1 << 0); }
    XL_CONSTEXPR FlagProxy fSqref() const   { return FlagProxy(raw_flags, 1 << 1); }
    XL_CONSTEXPR FlagProxy fFormula() const { return FlagProxy(raw_flags, 1 << 2); }
    XL_CONSTEXPR FlagProxy fRelID() const   { return FlagProxy(raw_flags, 1 << 3); }
    XL_CONSTEXPR ValueProxy<uint32_t> reserved() const { return ValueProxy<uint32_t>(raw_flags, 4, 28); }
};

// --- Conversion for PtgListUnion ---
struct PtgListFlags {
    const uint16_t raw_flags;

    explicit PtgListFlags(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    XL_CONSTEXPR ValueProxy<uint8_t> columns() const          { return ValueProxy<uint8_t>(raw_flags, 0, 2); }
    XL_CONSTEXPR ValueProxy<uint8_t> rowType() const          { return ValueProxy<uint8_t>(raw_flags, 2, 5); }
    XL_CONSTEXPR FlagProxy squareBracketSpace() const { return FlagProxy(raw_flags, 1 << 7); }
    XL_CONSTEXPR FlagProxy commaSpace() const         { return FlagProxy(raw_flags, 1 << 8); }
    XL_CONSTEXPR FlagProxy unused() const             { return FlagProxy(raw_flags, 1 << 9); } // Bit 9
    XL_CONSTEXPR ValueProxy<uint8_t> type() const     { return ValueProxy<uint8_t>(raw_flags, 10, 2); } // Bits 10-11
    XL_CONSTEXPR FlagProxy invalid() const            { return FlagProxy(raw_flags, 1 << 12); } // Bit 12
    XL_CONSTEXPR FlagProxy nonresident() const        { return FlagProxy(raw_flags, 1 << 13); } // Bit 13
    XL_CONSTEXPR ValueProxy<uint8_t> reserved2() const { return ValueProxy<uint8_t>(raw_flags, 14, 2); } // Bits 14-15
};

// --- Conversion for BrtBeginUserShViewUnion ---
struct BrtBeginUserShViewFlags {
    const uint32_t raw_flags;

    explicit BrtBeginUserShViewFlags(uint32_t initial_flags = 0) : raw_flags(initial_flags) {}

    XL_CONSTEXPR FlagProxy fShowBrks() const           { return FlagProxy(raw_flags, 1 << 0); }
    XL_CONSTEXPR FlagProxy fDspFmlaSv() const          { return FlagProxy(raw_flags, 1 << 1); }
    XL_CONSTEXPR FlagProxy fDspGridSv() const          { return FlagProxy(raw_flags, 1 << 2); }
    XL_CONSTEXPR FlagProxy fDspRwColSv() const         { return FlagProxy(raw_flags, 1 << 3); }
    XL_CONSTEXPR FlagProxy fDspGutsSv() const          { return FlagProxy(raw_flags, 1 << 4); }
    XL_CONSTEXPR FlagProxy fDspZerosSv() const         { return FlagProxy(raw_flags, 1 << 5); }
    XL_CONSTEXPR FlagProxy fHorizontal() const         { return FlagProxy(raw_flags, 1 << 6); }
    XL_CONSTEXPR FlagProxy fVertical() const           { return FlagProxy(raw_flags, 1 << 7); }
    XL_CONSTEXPR FlagProxy fPrintRwCol() const         { return FlagProxy(raw_flags, 1 << 8); }
    XL_CONSTEXPR FlagProxy fPrintGrid() const          { return FlagProxy(raw_flags, 1 << 9); }
    XL_CONSTEXPR FlagProxy fFitToPage() const          { return FlagProxy(raw_flags, 1 << 10); }
    XL_CONSTEXPR FlagProxy fPrintArea() const          { return FlagProxy(raw_flags, 1 << 11); }
    XL_CONSTEXPR FlagProxy fOnePrintArea() const       { return FlagProxy(raw_flags, 1 << 12); }
    XL_CONSTEXPR FlagProxy fFilterMode() const         { return FlagProxy(raw_flags, 1 << 13); }
    XL_CONSTEXPR FlagProxy fEzFilter() const           { return FlagProxy(raw_flags, 1 << 14); }
    XL_CONSTEXPR FlagProxy reserved1() const           { return FlagProxy(raw_flags, 1 << 15); }
    XL_CONSTEXPR FlagProxy reserved2() const           { return FlagProxy(raw_flags, 1 << 16); }
    XL_CONSTEXPR FlagProxy fSplitV() const             { return FlagProxy(raw_flags, 1 << 17); }
    XL_CONSTEXPR FlagProxy fSplitH() const             { return FlagProxy(raw_flags, 1 << 18); }
    XL_CONSTEXPR ValueProxy<uint8_t> fHiddenRw() const { return ValueProxy<uint8_t>(raw_flags, 19, 2); }
    XL_CONSTEXPR FlagProxy fHiddenCol() const          { return FlagProxy(raw_flags, 1 << 21); }
    XL_CONSTEXPR ValueProxy<uint8_t> hsState() const   { return ValueProxy<uint8_t>(raw_flags, 22, 2); }
    XL_CONSTEXPR FlagProxy reserved3() const           { return FlagProxy(raw_flags, 1 << 24); }
    XL_CONSTEXPR FlagProxy fFilterUnique() const       { return FlagProxy(raw_flags, 1 << 25); }
    XL_CONSTEXPR FlagProxy fSheetLayoutView() const    { return FlagProxy(raw_flags, 1 << 26); }
    XL_CONSTEXPR FlagProxy fPageLayoutView() const     { return FlagProxy(raw_flags, 1 << 27); }
    XL_CONSTEXPR FlagProxy reserved4() const           { return FlagProxy(raw_flags, 1 << 28); }
    XL_CONSTEXPR FlagProxy fRuler() const              { return FlagProxy(raw_flags, 1 << 29); }
    XL_CONSTEXPR FlagProxy reserved5() const           { return FlagProxy(raw_flags, 1 << 30); }
    XL_CONSTEXPR FlagProxy reserved6() const           { return FlagProxy(raw_flags, 1u << 31); }
};

// --- Conversion for BrtUserBookViewUnion ---
struct BrtUserBookViewFlags {
    const uint32_t raw_flags;

    explicit BrtUserBookViewFlags(uint32_t initial_flags = 0) : raw_flags(initial_flags) {}

    XL_CONSTEXPR FlagProxy fIconic() const       { return FlagProxy(raw_flags, 1 << 0); }
    XL_CONSTEXPR FlagProxy fDspHScroll() const   { return FlagProxy(raw_flags, 1 << 1); }
    XL_CONSTEXPR FlagProxy fDspVScroll() const   { return FlagProxy(raw_flags, 1 << 2); }
    XL_CONSTEXPR FlagProxy fBotAdornment() const { return FlagProxy(raw_flags, 1 << 3); }
    XL_CONSTEXPR FlagProxy fZoom() const         { return FlagProxy(raw_flags, 1 << 4); }
    XL_CONSTEXPR FlagProxy fDspFmlaBar() const   { return FlagProxy(raw_flags, 1 << 5); }
    XL_CONSTEXPR FlagProxy fDspStatus() const    { return FlagProxy(raw_flags, 1 << 6); }
    XL_CONSTEXPR ValueProxy<uint8_t> mdDspNote() const { return ValueProxy<uint8_t>(raw_flags, 7, 2); }
    XL_CONSTEXPR ValueProxy<uint8_t> mdHideObj() const { return ValueProxy<uint8_t>(raw_flags, 9, 2); }
    XL_CONSTEXPR FlagProxy fPrintIncl() const    { return FlagProxy(raw_flags, 1 << 11); }
    XL_CONSTEXPR FlagProxy fRowColIncl() const   { return FlagProxy(raw_flags, 1 << 12); }
    XL_CONSTEXPR FlagProxy fTimedUpdate() const  { return FlagProxy(raw_flags, 1 << 13); }
    XL_CONSTEXPR FlagProxy fAllMemChanges() const { return FlagProxy(raw_flags, 1 << 14); }
    XL_CONSTEXPR FlagProxy fOnlySync() const     { return FlagProxy(raw_flags, 1 << 15); }
    XL_CONSTEXPR FlagProxy fPersonalView() const { return FlagProxy(raw_flags, 1 << 16); }
    XL_CONSTEXPR ValueProxy<uint16_t> unused() const { return ValueProxy<uint16_t>(raw_flags, 17, 15); }
};

// --- Conversion for BrtBeginHeaderFooterUnion ---
struct BrtBeginHeaderFooterFlags {
    const uint16_t raw_flags;

    explicit BrtBeginHeaderFooterFlags(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    XL_CONSTEXPR FlagProxy fHFDiffOddEven() const  { return FlagProxy(raw_flags, 1 << 0); }
    XL_CONSTEXPR FlagProxy fHFDiffFirst() const    { return FlagProxy(raw_flags, 1 << 1); }
    XL_CONSTEXPR FlagProxy fHFScaleWithDoc() const { return FlagProxy(raw_flags, 1 << 2); }
    XL_CONSTEXPR FlagProxy fHFAlignMargins() const { return FlagProxy(raw_flags, 1 << 3); }
    XL_CONSTEXPR ValueProxy<uint16_t> reserved() const { return ValueProxy<uint16_t>(raw_flags, 4, 12); }
};

#endif // XLSB_UNPACK_H
