#ifndef XLSB_FLAGS_H
#define XLSB_FLAGS_H

#include <cstdint>  // For uint8_t, uint16_t, uint32_t, uintmax_t
#include <type_traits>  // For std::enable_if, std::is_integral

// --- Base Helper Struct: FlagProxy ---
struct FlagProxy {
    const uintmax_t raw_flags_value;
    const uintmax_t bit_mask_storage;

    template<typename T, typename = typename std::enable_if<std::is_integral<T>::value>::type>
    constexpr explicit FlagProxy(const T flags, uintmax_t mask)
        : raw_flags_value(static_cast<uintmax_t>(flags)),
          bit_mask_storage(mask) {}

    constexpr operator bool() const {
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
    constexpr explicit ValueProxy(const T_Raw flags, uintmax_t shift, uintmax_t width)
        : raw_flags_value(static_cast<uintmax_t>(flags)),
          shift_amount(shift),
          width_mask((width >= (sizeof(uintmax_t) * 8)) ? ~static_cast<uintmax_t>(0) : ((1ULL << width) - 1)) {}

    constexpr operator T_Return() const {
        return static_cast<T_Return>((raw_flags_value >> shift_amount) & width_mask);
    }
};


// --- Conversion for StyleFlagsUnion ---
struct StyleFlags {
    const uint16_t raw_flags;

    explicit StyleFlags(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    constexpr FlagProxy fBuiltIn() const  { return FlagProxy(raw_flags, 1 << 0); }
    constexpr FlagProxy fHidden() const   { return FlagProxy(raw_flags, 1 << 1); }
    constexpr FlagProxy fCustom() const   { return FlagProxy(raw_flags, 1 << 2); }
};

// --- Conversion for xfGrbitAtrUnion ---
struct xfGrbitAtr {
    const uint8_t raw_flags;

    explicit xfGrbitAtr(uint8_t initial_flags = 0) : raw_flags(initial_flags) {}

    constexpr FlagProxy bit1() const { return FlagProxy(raw_flags, 1 << 0); }
    constexpr FlagProxy bit2() const { return FlagProxy(raw_flags, 1 << 1); }
    constexpr FlagProxy bit3() const { return FlagProxy(raw_flags, 1 << 2); }
    constexpr FlagProxy bit4() const { return FlagProxy(raw_flags, 1 << 3); }
    constexpr FlagProxy bit5() const { return FlagProxy(raw_flags, 1 << 4); }
    constexpr FlagProxy bit6() const { return FlagProxy(raw_flags, 1 << 5); }
};

// --- Conversion for GrbitFmlaUnion ---
struct GrbitFmla {
    const uint16_t raw_flags;

    explicit GrbitFmla(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    constexpr FlagProxy reserved() const    { return FlagProxy(raw_flags, 1 << 0); }
    constexpr FlagProxy fAlwaysCalc() const { return FlagProxy(raw_flags, 1 << 1); }
};

// --- Conversion for ColRelShortUnion ---
struct ColRelShort {
    const uint16_t raw_flags;

    explicit ColRelShort(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    constexpr ValueProxy<uint16_t> col() const { return ValueProxy<uint16_t>(raw_flags, 0, 14); }
    constexpr FlagProxy fColRel() const        { return FlagProxy(raw_flags, 1 << 14); }
    constexpr FlagProxy fRwRel() const         { return FlagProxy(raw_flags, 1 << 15); }
};

// --- Conversion for BrtRowHdrUnion ---
struct BrtRowHdr {
    const uint16_t raw_flags;

    explicit BrtRowHdr(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    constexpr FlagProxy fExtraAsc() const   { return FlagProxy(raw_flags, 1 << 0); }
    constexpr FlagProxy fExtraDsc() const   { return FlagProxy(raw_flags, 1 << 1); }
    constexpr ValueProxy<uint8_t> reserved1() const { return ValueProxy<uint8_t>(raw_flags, 2, 6); }
    constexpr ValueProxy<uint8_t> iOutLevel() const { return ValueProxy<uint8_t>(raw_flags, 8, 3); }
    constexpr FlagProxy fCollapsed() const  { return FlagProxy(raw_flags, 1 << 11); }
    constexpr FlagProxy fDyZero() const     { return FlagProxy(raw_flags, 1 << 12); }
    constexpr FlagProxy fUnsynced() const   { return FlagProxy(raw_flags, 1 << 13); }
    constexpr FlagProxy fGhostDirty() const { return FlagProxy(raw_flags, 1 << 14); }
    constexpr FlagProxy fReserved() const   { return FlagProxy(raw_flags, 1 << 15); }
};

// --- Conversion for BrtColInfoUnion ---
struct BrtColInfo {
    const uint16_t raw_flags;

    explicit BrtColInfo(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    constexpr FlagProxy fHidden() const    { return FlagProxy(raw_flags, 1 << 0); }
    constexpr FlagProxy fUserSet() const   { return FlagProxy(raw_flags, 1 << 1); }
    constexpr FlagProxy fBestFit() const   { return FlagProxy(raw_flags, 1 << 2); }
    constexpr FlagProxy fPhonetic() const  { return FlagProxy(raw_flags, 1 << 3); }
    constexpr ValueProxy<uint8_t> reserved1() const { return ValueProxy<uint8_t>(raw_flags, 4, 4); }
    constexpr ValueProxy<uint8_t> iOutLevel() const { return ValueProxy<uint8_t>(raw_flags, 8, 3); }
    constexpr FlagProxy unused() const     { return FlagProxy(raw_flags, 1 << 11); }
    constexpr FlagProxy fCollapsed() const { return FlagProxy(raw_flags, 1 << 12); }
    constexpr ValueProxy<uint8_t> reserved2() const { return ValueProxy<uint8_t>(raw_flags, 13, 3); }
};

// --- Conversion for BrtNameUnion ---
struct BrtName {
    const uint16_t raw_flags;

    explicit BrtName(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    constexpr FlagProxy fHidden() const    { return FlagProxy(raw_flags, 1 << 0); }
    constexpr FlagProxy fFunc() const      { return FlagProxy(raw_flags, 1 << 1); }
    constexpr FlagProxy fOB() const        { return FlagProxy(raw_flags, 1 << 2); }
    constexpr FlagProxy fProc() const      { return FlagProxy(raw_flags, 1 << 3); }
    constexpr FlagProxy fCalcExp() const   { return FlagProxy(raw_flags, 1 << 4); }
    constexpr FlagProxy fBuiltin() const   { return FlagProxy(raw_flags, 1 << 5); }
    constexpr ValueProxy<uint16_t> fgrp() const { return ValueProxy<uint16_t>(raw_flags, 6, 9); }
    constexpr FlagProxy fPublished() const { return FlagProxy(raw_flags, 1 << 15); }
};

// --- Conversion for BrtName2Union ---
struct BrtName2 {
    const uint16_t raw_flags;

    explicit BrtName2(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    constexpr FlagProxy fWorkbookParam() const { return FlagProxy(raw_flags, 1 << 0); }
    constexpr FlagProxy fFutureFunction() const { return FlagProxy(raw_flags, 1 << 1); }
    constexpr ValueProxy<uint16_t> reserved() const { return ValueProxy<uint16_t>(raw_flags, 2, 14); }
};

// --- Conversion for FontFlagsUnion ---
struct FontFlags {
    const uint16_t raw_flags;

    explicit FontFlags(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    constexpr FlagProxy unused1() const   { return FlagProxy(raw_flags, 1 << 0); }
    constexpr FlagProxy fItalic() const   { return FlagProxy(raw_flags, 1 << 1); }
    constexpr FlagProxy unused2() const   { return FlagProxy(raw_flags, 1 << 2); }
    constexpr FlagProxy fStrikeout() const { return FlagProxy(raw_flags, 1 << 3); }
    constexpr FlagProxy fOutline() const  { return FlagProxy(raw_flags, 1 << 4); }
    constexpr FlagProxy fShadow() const   { return FlagProxy(raw_flags, 1 << 5); }
    constexpr FlagProxy fCondense() const { return FlagProxy(raw_flags, 1 << 6); }
    constexpr FlagProxy fExtend() const   { return FlagProxy(raw_flags, 1 << 7); }
};

// --- Conversion for XFUnion ---
struct XF {
    const uint32_t raw_flags;

    explicit XF(uint32_t initial_flags = 0) : raw_flags(initial_flags) {}

    constexpr ValueProxy<uint8_t> alc() const          { return ValueProxy<uint8_t>(raw_flags, 0, 3); }
    constexpr ValueProxy<uint8_t> alcv() const         { return ValueProxy<uint8_t>(raw_flags, 3, 3); }
    constexpr FlagProxy fWrap() const                  { return FlagProxy(raw_flags, 1 << 6); }
    constexpr FlagProxy fJustLast() const              { return FlagProxy(raw_flags, 1 << 7); }
    constexpr FlagProxy fShrinkToFit() const           { return FlagProxy(raw_flags, 1 << 8); }
    constexpr FlagProxy fMergeCell() const             { return FlagProxy(raw_flags, 1 << 9); }
    constexpr ValueProxy<uint8_t> iReadingOrder() const { return ValueProxy<uint8_t>(raw_flags, 10, 2); }
    constexpr FlagProxy fLocked() const                { return FlagProxy(raw_flags, 1 << 12); }
    constexpr FlagProxy fHidden() const                { return FlagProxy(raw_flags, 1 << 13); }
    constexpr FlagProxy fSxButton() const              { return FlagProxy(raw_flags, 1 << 14); }
    constexpr FlagProxy f123Prefix() const             { return FlagProxy(raw_flags, 1 << 15); }
    constexpr ValueProxy<uint8_t> xfGrbitAtr() const   { return ValueProxy<uint8_t>(raw_flags, 16, 6); }
    constexpr ValueProxy<uint16_t> unused() const      { return ValueProxy<uint16_t>(raw_flags, 22, 10); }
};

// --- Conversion for BrtWsProp1Union ---
struct BrtWsProp1 {
    const uint16_t raw_flags;

    explicit BrtWsProp1(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    constexpr FlagProxy fShowAutoBreaks() const    { return FlagProxy(raw_flags, 1 << 0); }
    constexpr ValueProxy<uint8_t> rserved1() const { return ValueProxy<uint8_t>(raw_flags, 1, 2); }
    constexpr FlagProxy fPublish() const           { return FlagProxy(raw_flags, 1 << 3); }
    constexpr FlagProxy fDialog() const            { return FlagProxy(raw_flags, 1 << 4); }
    constexpr FlagProxy fApplyStyles() const       { return FlagProxy(raw_flags, 1 << 5); }
    constexpr FlagProxy fRowSumsBelow() const      { return FlagProxy(raw_flags, 1 << 6); }
    constexpr FlagProxy fColSumsRight() const      { return FlagProxy(raw_flags, 1 << 7); }
    constexpr FlagProxy fFitToPage() const         { return FlagProxy(raw_flags, 1 << 8); }
    constexpr ValueProxy<uint8_t> reserved2() const { return ValueProxy<uint8_t>(raw_flags, 9, 1); }
    constexpr FlagProxy fShowOutlineSymbols() const { return FlagProxy(raw_flags, 1 << 10); }
    constexpr ValueProxy<uint8_t> reserved3() const { return ValueProxy<uint8_t>(raw_flags, 11, 1); }
    constexpr FlagProxy fSyncHoriz() const         { return FlagProxy(raw_flags, 1 << 12); }
    constexpr FlagProxy fSyncVert() const          { return FlagProxy(raw_flags, 1 << 13); }
    constexpr FlagProxy fAltExprEval() const       { return FlagProxy(raw_flags, 1 << 14); }
    constexpr FlagProxy fAltFormulaEntry() const   { return FlagProxy(raw_flags, 1 << 15); }
};

// --- Conversion for BrtWsProp2Union ---
struct BrtWsProp2 {
    const uint8_t raw_flags;

    explicit BrtWsProp2(uint8_t initial_flags = 0) : raw_flags(initial_flags) {}

    constexpr FlagProxy fFilterMode() const    { return FlagProxy(raw_flags, 1 << 0); }
    constexpr FlagProxy fCondFmtCalc() const   { return FlagProxy(raw_flags, 1 << 1); }
    constexpr ValueProxy<uint8_t> reserved4() const { return ValueProxy<uint8_t>(raw_flags, 2, 6); }
};

// --- Conversion for BrtBeginWsViewUnion ---
struct BrtBeginWsView {
    const uint16_t raw_flags;

    explicit BrtBeginWsView(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    constexpr FlagProxy fWnProt() const           { return FlagProxy(raw_flags, 1 << 0); }
    constexpr FlagProxy fDspFmla() const          { return FlagProxy(raw_flags, 1 << 1); }
    constexpr FlagProxy fDspGrid() const          { return FlagProxy(raw_flags, 1 << 2); }
    constexpr FlagProxy fDspRwCol() const         { return FlagProxy(raw_flags, 1 << 3); }
    constexpr FlagProxy fDspZeros() const         { return FlagProxy(raw_flags, 1 << 4); }
    constexpr FlagProxy fRightToLeft() const      { return FlagProxy(raw_flags, 1 << 5); }
    constexpr FlagProxy fSelected() const         { return FlagProxy(raw_flags, 1 << 6); }
    constexpr FlagProxy fDspRuler() const         { return FlagProxy(raw_flags, 1 << 7); }
    constexpr FlagProxy fDspGuts() const          { return FlagProxy(raw_flags, 1 << 8); }
    constexpr FlagProxy fDefaultHdr() const       { return FlagProxy(raw_flags, 1 << 9); }
    constexpr FlagProxy fWhitespaceHidden() const { return FlagProxy(raw_flags, 1 << 10); }
    constexpr ValueProxy<uint8_t> reserved1() const { return ValueProxy<uint8_t>(raw_flags, 11, 5); }
};

// --- Conversion for BrtTableStyleClientUnion ---
struct BrtTableStyleClient {
    const uint16_t raw_flags;

    explicit BrtTableStyleClient(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    constexpr FlagProxy fFirstColumn() const    { return FlagProxy(raw_flags, 1 << 0); }
    constexpr FlagProxy fLastColumn() const     { return FlagProxy(raw_flags, 1 << 1); }
    constexpr FlagProxy fRowStripes() const     { return FlagProxy(raw_flags, 1 << 2); }
    constexpr FlagProxy fColumnStripes() const  { return FlagProxy(raw_flags, 1 << 3); }
    constexpr FlagProxy fRowHeaders() const     { return FlagProxy(raw_flags, 1 << 4); }
    constexpr FlagProxy fColumnHeaders() const  { return FlagProxy(raw_flags, 1 << 5); }
    constexpr ValueProxy<uint16_t> reserved() const { return ValueProxy<uint16_t>(raw_flags, 6, 10); }
};

// --- Conversion for BrtBeginCFRuleUnion ---
struct BrtBeginCFRule {
    const uint16_t raw_flags;

    explicit BrtBeginCFRule(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    constexpr FlagProxy reserved3() const  { return FlagProxy(raw_flags, 1 << 0); }
    constexpr FlagProxy fStopTrue() const  { return FlagProxy(raw_flags, 1 << 1); }
    constexpr FlagProxy fAbove() const     { return FlagProxy(raw_flags, 1 << 2); }
    constexpr FlagProxy fBottom() const    { return FlagProxy(raw_flags, 1 << 3); }
    constexpr FlagProxy fPercent() const   { return FlagProxy(raw_flags, 1 << 4); }
    constexpr ValueProxy<uint16_t> reserved4() const { return ValueProxy<uint16_t>(raw_flags, 5, 11); }
};

// --- Conversion for BrtWbPropUnion ---
struct BrtWbProp {
    const uint32_t raw_flags;

    explicit BrtWbProp(uint32_t initial_flags = 0) : raw_flags(initial_flags) {}

    constexpr FlagProxy f1904() const                { return FlagProxy(raw_flags, 1 << 0); }
    constexpr FlagProxy reserved1() const            { return FlagProxy(raw_flags, 1 << 1); }
    constexpr FlagProxy fHideBorderUnselLists() const { return FlagProxy(raw_flags, 1 << 2); }
    constexpr FlagProxy fFilterPrivacy() const       { return FlagProxy(raw_flags, 1 << 3); }
    constexpr FlagProxy fBuggedUserAboutSolution() const { return FlagProxy(raw_flags, 1 << 4); }
    constexpr FlagProxy fShowInkAnnotation() const   { return FlagProxy(raw_flags, 1 << 5); }
    constexpr FlagProxy fBackup() const              { return FlagProxy(raw_flags, 1 << 6); }
    constexpr FlagProxy fNoSaveSup() const           { return FlagProxy(raw_flags, 1 << 7); }
    constexpr ValueProxy<uint8_t> grbitUpdateLinks() const { return ValueProxy<uint8_t>(raw_flags, 8, 2); }
    constexpr FlagProxy fHidePivotTableFList() const { return FlagProxy(raw_flags, 1 << 10); }
    constexpr FlagProxy fPublishedBookItems() const  { return FlagProxy(raw_flags, 1 << 11); }
    constexpr FlagProxy fCheckCompat() const         { return FlagProxy(raw_flags, 1 << 12); }
    constexpr ValueProxy<uint8_t> mdDspObj() const   { return ValueProxy<uint8_t>(raw_flags, 13, 2); }
    constexpr FlagProxy fShowPivotChartFilter() const { return FlagProxy(raw_flags, 1 << 15); }
    constexpr FlagProxy fAutoCompressPictures() const { return FlagProxy(raw_flags, 1 << 16); }
    constexpr FlagProxy reserved2() const            { return FlagProxy(raw_flags, 1 << 17); }
    constexpr FlagProxy fRefreshAll() const          { return FlagProxy(raw_flags, 1 << 18); }
    constexpr ValueProxy<uint16_t> unused() const    { return ValueProxy<uint16_t>(raw_flags, 19, 13); }
};

// --- Conversion for BrtDValUnion ---
struct BrtDVal {
    const uint32_t raw_flags;

    explicit BrtDVal(uint32_t initial_flags = 0) : raw_flags(initial_flags) {}

    constexpr ValueProxy<uint8_t> valType() const       { return ValueProxy<uint8_t>(raw_flags, 0, 4); }
    constexpr ValueProxy<uint8_t> errStyle() const      { return ValueProxy<uint8_t>(raw_flags, 4, 3); }
    constexpr FlagProxy unused() const          { return FlagProxy(raw_flags, 1 << 7); }
    constexpr FlagProxy fAllowBlank() const     { return FlagProxy(raw_flags, 1 << 8); }
    constexpr FlagProxy fSuppressCombo() const  { return FlagProxy(raw_flags, 1 << 9); }
    constexpr ValueProxy<uint8_t> mdImeMode() const     { return ValueProxy<uint8_t>(raw_flags, 10, 8); }
    constexpr FlagProxy fShowInputMsg() const   { return FlagProxy(raw_flags, 1 << 18); }
    constexpr FlagProxy fShowErrorMsg() const   { return FlagProxy(raw_flags, 1 << 19); }
    constexpr ValueProxy<uint8_t> typOperator() const   { return ValueProxy<uint8_t>(raw_flags, 20, 4); }
    constexpr FlagProxy fDVMinFmla() const      { return FlagProxy(raw_flags, 1 << 24); }
    constexpr FlagProxy fDVMaxFmla() const      { return FlagProxy(raw_flags, 1 << 25); }
    constexpr ValueProxy<uint8_t> reserved() const      { return ValueProxy<uint8_t>(raw_flags, 26, 6); }
};

// --- Conversion for FRTVersionUnion ---
struct FRTVersion {
    const uint16_t raw_flags;

    explicit FRTVersion(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    constexpr ValueProxy<uint16_t> product() const { return ValueProxy<uint16_t>(raw_flags, 0, 15); }
    constexpr FlagProxy reserved() const { return FlagProxy(raw_flags, 1 << 15); }
};

// --- Conversion for FRTHeaderUnion ---
struct FRTHeader {
    const uint32_t raw_flags;

    explicit FRTHeader(uint32_t initial_flags = 0) : raw_flags(initial_flags) {}

    constexpr FlagProxy fRef() const     { return FlagProxy(raw_flags, 1 << 0); }
    constexpr FlagProxy fSqref() const   { return FlagProxy(raw_flags, 1 << 1); }
    constexpr FlagProxy fFormula() const { return FlagProxy(raw_flags, 1 << 2); }
    constexpr FlagProxy fRelID() const   { return FlagProxy(raw_flags, 1 << 3); }
    constexpr ValueProxy<uint32_t> reserved() const { return ValueProxy<uint32_t>(raw_flags, 4, 28); }
};

// --- Conversion for PtgListUnion ---
struct PtgList {
    const uint16_t raw_flags;

    explicit PtgList(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    constexpr ValueProxy<uint8_t> columns() const          { return ValueProxy<uint8_t>(raw_flags, 0, 2); }
    constexpr ValueProxy<uint8_t> rowType() const          { return ValueProxy<uint8_t>(raw_flags, 2, 5); }
    constexpr FlagProxy squareBracketSpace() const { return FlagProxy(raw_flags, 1 << 7); }
    constexpr FlagProxy commaSpace() const         { return FlagProxy(raw_flags, 1 << 8); }
    constexpr FlagProxy unused() const             { return FlagProxy(raw_flags, 1 << 9); } // Bit 9
    constexpr ValueProxy<uint8_t> type() const     { return ValueProxy<uint8_t>(raw_flags, 10, 2); } // Bits 10-11
    constexpr FlagProxy invalid() const            { return FlagProxy(raw_flags, 1 << 12); } // Bit 12
    constexpr FlagProxy nonresident() const        { return FlagProxy(raw_flags, 1 << 13); } // Bit 13
    constexpr ValueProxy<uint8_t> reserved2() const { return ValueProxy<uint8_t>(raw_flags, 14, 2); } // Bits 14-15
};

// --- Conversion for BrtBeginUserShViewUnion ---
struct BrtBeginUserShView {
    const uint32_t raw_flags;

    explicit BrtBeginUserShView(uint32_t initial_flags = 0) : raw_flags(initial_flags) {}

    constexpr FlagProxy fShowBrks() const           { return FlagProxy(raw_flags, 1 << 0); }
    constexpr FlagProxy fDspFmlaSv() const          { return FlagProxy(raw_flags, 1 << 1); }
    constexpr FlagProxy fDspGridSv() const          { return FlagProxy(raw_flags, 1 << 2); }
    constexpr FlagProxy fDspRwColSv() const         { return FlagProxy(raw_flags, 1 << 3); }
    constexpr FlagProxy fDspGutsSv() const          { return FlagProxy(raw_flags, 1 << 4); }
    constexpr FlagProxy fDspZerosSv() const         { return FlagProxy(raw_flags, 1 << 5); }
    constexpr FlagProxy fHorizontal() const         { return FlagProxy(raw_flags, 1 << 6); }
    constexpr FlagProxy fVertical() const           { return FlagProxy(raw_flags, 1 << 7); }
    constexpr FlagProxy fPrintRwCol() const         { return FlagProxy(raw_flags, 1 << 8); }
    constexpr FlagProxy fPrintGrid() const          { return FlagProxy(raw_flags, 1 << 9); }
    constexpr FlagProxy fFitToPage() const          { return FlagProxy(raw_flags, 1 << 10); }
    constexpr FlagProxy fPrintArea() const          { return FlagProxy(raw_flags, 1 << 11); }
    constexpr FlagProxy fOnePrintArea() const       { return FlagProxy(raw_flags, 1 << 12); }
    constexpr FlagProxy fFilterMode() const         { return FlagProxy(raw_flags, 1 << 13); }
    constexpr FlagProxy fEzFilter() const           { return FlagProxy(raw_flags, 1 << 14); }
    constexpr FlagProxy reserved1() const           { return FlagProxy(raw_flags, 1 << 15); }
    constexpr FlagProxy reserved2() const           { return FlagProxy(raw_flags, 1 << 16); }
    constexpr FlagProxy fSplitV() const             { return FlagProxy(raw_flags, 1 << 17); }
    constexpr FlagProxy fSplitH() const             { return FlagProxy(raw_flags, 1 << 18); }
    constexpr ValueProxy<uint8_t> fHiddenRw() const { return ValueProxy<uint8_t>(raw_flags, 19, 2); }
    constexpr FlagProxy fHiddenCol() const          { return FlagProxy(raw_flags, 1 << 21); }
    constexpr ValueProxy<uint8_t> hsState() const   { return ValueProxy<uint8_t>(raw_flags, 22, 2); }
    constexpr FlagProxy reserved3() const           { return FlagProxy(raw_flags, 1 << 24); }
    constexpr FlagProxy fFilterUnique() const       { return FlagProxy(raw_flags, 1 << 25); }
    constexpr FlagProxy fSheetLayoutView() const    { return FlagProxy(raw_flags, 1 << 26); }
    constexpr FlagProxy fPageLayoutView() const     { return FlagProxy(raw_flags, 1 << 27); }
    constexpr FlagProxy reserved4() const           { return FlagProxy(raw_flags, 1 << 28); }
    constexpr FlagProxy fRuler() const              { return FlagProxy(raw_flags, 1 << 29); }
    constexpr FlagProxy reserved5() const           { return FlagProxy(raw_flags, 1 << 30); }
    constexpr FlagProxy reserved6() const           { return FlagProxy(raw_flags, 1 << 31); }
};

// --- Conversion for BrtUserBookViewUnion ---
struct BrtUserBookView {
    const uint32_t raw_flags;

    explicit BrtUserBookView(uint32_t initial_flags = 0) : raw_flags(initial_flags) {}

    constexpr FlagProxy fIconic() const       { return FlagProxy(raw_flags, 1 << 0); }
    constexpr FlagProxy fDspHScroll() const   { return FlagProxy(raw_flags, 1 << 1); }
    constexpr FlagProxy fDspVScroll() const   { return FlagProxy(raw_flags, 1 << 2); }
    constexpr FlagProxy fBotAdornment() const { return FlagProxy(raw_flags, 1 << 3); }
    constexpr FlagProxy fZoom() const         { return FlagProxy(raw_flags, 1 << 4); }
    constexpr FlagProxy fDspFmlaBar() const   { return FlagProxy(raw_flags, 1 << 5); }
    constexpr FlagProxy fDspStatus() const    { return FlagProxy(raw_flags, 1 << 6); }
    constexpr ValueProxy<uint8_t> mdDspNote() const { return ValueProxy<uint8_t>(raw_flags, 7, 2); }
    constexpr ValueProxy<uint8_t> mdHideObj() const { return ValueProxy<uint8_t>(raw_flags, 9, 2); }
    constexpr FlagProxy fPrintIncl() const    { return FlagProxy(raw_flags, 1 << 11); }
    constexpr FlagProxy fRowColIncl() const   { return FlagProxy(raw_flags, 1 << 12); }
    constexpr FlagProxy fTimedUpdate() const  { return FlagProxy(raw_flags, 1 << 13); }
    constexpr FlagProxy fAllMemChanges() const { return FlagProxy(raw_flags, 1 << 14); }
    constexpr FlagProxy fOnlySync() const     { return FlagProxy(raw_flags, 1 << 15); }
    constexpr FlagProxy fPersonalView() const { return FlagProxy(raw_flags, 1 << 16); }
    constexpr ValueProxy<uint16_t> unused() const { return ValueProxy<uint16_t>(raw_flags, 17, 15); }
};

// --- Conversion for BrtBeginHeaderFooterUnion ---
struct BrtBeginHeaderFooter {
    const uint16_t raw_flags;

    explicit BrtBeginHeaderFooter(uint16_t initial_flags = 0) : raw_flags(initial_flags) {}

    constexpr FlagProxy fHFDiffOddEven() const  { return FlagProxy(raw_flags, 1 << 0); }
    constexpr FlagProxy fHFDiffFirst() const    { return FlagProxy(raw_flags, 1 << 1); }
    constexpr FlagProxy fHFScaleWithDoc() const { return FlagProxy(raw_flags, 1 << 2); }
    constexpr FlagProxy fHFAlignMargins() const { return FlagProxy(raw_flags, 1 << 3); }
    constexpr ValueProxy<uint16_t> reserved() const { return ValueProxy<uint16_t>(raw_flags, 4, 12); }
};

#endif // XLSB_UNPACK_H
