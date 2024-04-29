#ifndef XLSB_DEFINES_H
#define XLSB_DEFINES_H

// defines
typedef struct {
  bool fBuiltIn : 1;
  bool fHidden : 1;
  bool fCustom : 1;
  uint16_t unused : 13;
} StyleFlagsFields;

typedef struct {
  bool bit1 : 1;
  bool bit2 : 1;
  bool bit3 : 1;
  bool bit4 : 1;
  bool bit5 : 1;
  bool bit6 : 1;
} xfGrbitAtrFields;

typedef struct {
  bool reserved : 1;
  bool fAlwaysCalc : 1;
  uint16_t unused : 14;
} GrbitFmlaFields;

typedef struct {
  uint16_t col : 14;
  bool fColRel : 1;
  bool fRwRel : 1;
} ColRelShortFields;

typedef struct {
  bool fExtraAsc : 1;
  bool fExtraDsc : 1;
  uint8_t reserved1 : 6;
  uint8_t iOutLevel : 3;
  bool fCollapsed : 1;
  bool fDyZero : 1;
  bool fUnsynced : 1;
  bool fGhostDirty : 1;
  bool fReserved : 1;
} BrtRowHdrFields;

typedef struct {
  bool fHidden : 1;
  bool fUserSet : 1;
  bool fBestFit : 1;
  bool fPhonetic : 1;
  uint8_t reserved1 : 4;
  uint8_t iOutLevel : 3;
  bool unused : 1;
  bool fCollapsed : 1;
  uint8_t reserved2 : 3;
} BrtColInfoFields;

typedef struct {
  bool fHidden : 1;
  bool fFunc : 1;
  bool fOB : 1;
  bool fProc : 1;
  bool fCalcExp : 1;
  bool fBuiltin : 1;
  uint16_t fgrp : 9;
  bool fPublished : 1;
} BrtNameFields;

typedef struct {
  bool fWorkbookParam : 1;
  bool fFutureFunction : 1;
  uint16_t reserved : 14;
} BrtNameFields2;

typedef struct {
  bool unused1 : 1;
  bool fItalic : 1;
  bool unused2 : 1;
  bool fStrikeout : 1;
  bool fOutline : 1;
  bool fShadow : 1;
  bool fCondense : 1;
  bool fExtend : 1;
  uint8_t unused3 : 8;
} FontFlagsFields;

typedef struct {
  uint8_t alc : 3;
  uint8_t alcv : 3;
  bool fWrap : 1;
  bool fJustLast : 1;
  bool fShrinkToFit : 1;
  bool fMergeCell : 1;
  uint8_t iReadingOrder : 2;
  bool fLocked : 1;
  bool fHidden : 1;
  bool fSxButton : 1;
  bool f123Prefix : 1;
  uint8_t xfGrbitAtr : 6;
  uint16_t unused : 10;
} XFFields;


typedef struct {
  bool fShowAutoBreaks : 1;
  uint8_t rserved1 : 2;
  bool fPublish : 1;
  bool fDialog : 1;
  bool fApplyStyles : 1;
  bool fRowSumsBelow : 1;
  bool fColSumsRight : 1;
  bool fFitToPage : 1;
  uint8_t reserved2 : 1;
  bool fShowOutlineSymbols : 1;
  uint8_t reserved3 : 1;
  bool fSyncHoriz : 1;
  bool fSyncVert : 1;
  bool fAltExprEval : 1;
  bool fAltFormulaEntry : 1;
} BrtWsPropFields1;

typedef struct {
  bool fFilterMode : 1;
  bool fCondFmtCalc : 1;
  uint8_t reserved4 :6;
} BrtWsPropFields2;

typedef struct {
  bool fWnProt : 1;
  bool fDspFmla : 1;
  bool fDspGrid : 1;
  bool fDspRwCol : 1;
  bool fDspZeros : 1;
  bool fRightToLeft : 1;
  bool fSelected : 1;
  bool fDspRuler : 1;
  bool fDspGuts : 1;
  bool fDefaultHdr : 1;
  bool fWhitespaceHidden : 1;
  uint8_t reserved1 : 5;
} BrtBeginWsViewFields;

typedef struct {
  bool fFirstColumn : 1;
  bool fLastColumn : 1;
  bool fRowStripes : 1;
  bool fColumnStripes : 1;
  bool fRowHeaders : 1;
  bool fColumnHeaders : 1;
  uint16_t reserved : 10;
} BrtTableStyleClientFields;

typedef struct {
  bool reserved3 : 1;
  bool fStopTrue : 1;
  bool fAbove : 1;
  bool fBottom : 1;
  bool fPercent : 1;
  uint16_t reserved4 : 11;
} BrtBeginCFRuleFields;

typedef struct {
  bool f1904 : 1;
  bool reserved1 : 1;
  bool fHideBorderUnselLists : 1;
  bool fFilterPrivacy : 1;
  bool fBuggedUserAboutSolution : 1;
  bool fShowInkAnnotation : 1;
  bool fBackup : 1;
  bool fNoSaveSup : 1;
  uint8_t grbitUpdateLinks : 2;
  bool fHidePivotTableFList : 1;
  bool fPublishedBookItems : 1;
  bool fCheckCompat : 1;
  uint8_t mdDspObj : 2;
  bool fShowPivotChartFilter : 1;
  bool fAutoCompressPictures : 1;
  bool reserved2 : 1;
  bool fRefreshAll : 1;
  uint16_t unused : 13;
} BrtWbPropFields;

enum RgbExtra
{
  PtgExtraArray = 0,
  PtgExtraMem = 1,
  PtgExtraCol = 2,
  PtgExtraList = 3,
  RevNameTabid = 4,
  RevName = 5,
  RevExtern = 6
};

enum RecordTypes
{
  BrtRowHdr = 0,
  BrtCellBlank = 1,
  BrtCellRk = 2,
  BrtCellError = 3,
  BrtCellBool = 4,
  BrtCellReal = 5,
  BrtCellSt = 6,
  BrtCellIsst = 7,
  BrtFmlaString = 8,
  BrtFmlaNum = 9,
  BrtFmlaBool = 10,
  BrtFmlaError = 11,
  BrtSSTItem = 19,
  BrtPCDIMissing = 20,
  BrtPCDINumber = 21,
  BrtPCDIBoolean = 22,
  BrtPCDIError = 23,
  BrtPCDIString = 24,
  BrtPCDIDatetime = 25,
  BrtPCDIIndex = 26,
  BrtPCDIAMissing = 27,
  BrtPCDIANumber = 28,
  BrtPCDIABoolean = 29,
  BrtPCDIAError = 30,
  BrtPCDIAString = 31,
  BrtPCDIADatetime = 32,
  BrtPCRRecord = 33,
  BrtPCRRecordDt = 34,
  BrtFRTBegin = 35,
  BrtFRTEnd = 36,
  BrtACBegin = 37,
  BrtACEnd = 38,
  BrtName = 39,
  BrtIndexRowBlock = 40,
  BrtIndexBlock = 42,
  BrtFont = 43,
  BrtFmt = 44,
  BrtFill = 45,
  BrtBorder = 46,
  BrtXF = 47,
  BrtStyle = 48,
  BrtCellMeta = 49,
  BrtValueMeta = 50,
  BrtMdb = 51,
  BrtBeginFmd = 52,
  BrtEndFmd = 53,
  BrtBeginMdx = 54,
  BrtEndMdx = 55,
  BrtBeginMdxTuple = 56,
  BrtEndMdxTuple = 57,
  BrtMdxMbrIstr = 58,
  BrtStr = 59,
  BrtColInfo = 60,
  BrtCellRString = 62,
  BrtDVal = 64,
  BrtSxvcellNum = 65,
  BrtSxvcellStr = 66,
  BrtSxvcellBool = 67,
  BrtSxvcellErr = 68,
  BrtSxvcellDate = 69,
  BrtSxvcellNil = 70,
  BrtFileVersion = 128,
  BrtBeginSheet = 129,
  BrtEndSheet = 130,
  BrtBeginBook = 131,
  BrtEndBook = 132,
  BrtBeginWsViews = 133,
  BrtEndWsViews = 134,
  BrtBeginBookViews = 135,
  BrtEndBookViews = 136,
  BrtBeginWsView = 137,
  BrtEndWsView = 138,
  BrtBeginCsViews = 139,
  BrtEndCsViews = 140,
  BrtBeginCsView = 141,
  BrtEndCsView = 142,
  BrtBeginBundleShs = 143,
  BrtEndBundleShs = 144,
  BrtBeginSheetData = 145,
  BrtEndSheetData = 146,
  BrtWsProp = 147,
  BrtWsDim = 148,
  BrtPane = 151,
  BrtSel = 152,
  BrtWbProp = 153,
  BrtWbFactoid = 154,
  BrtFileRecover = 155,
  BrtBundleSh = 156,
  BrtCalcProp = 157,
  BrtBookView = 158,
  BrtBeginSst = 159,
  BrtEndSst = 160,
  BrtBeginAFilter = 161,
  BrtEndAFilter = 162,
  BrtBeginFilterColumn = 163,
  BrtEndFilterColumn = 164,
  BrtBeginFilters = 165,
  BrtEndFilters = 166,
  BrtFilter = 167,
  BrtColorFilter = 168,
  BrtIconFilter = 169,
  BrtTop10Filter = 170,
  BrtDynamicFilter = 171,
  BrtBeginCustomFilters = 172,
  BrtEndCustomFilters = 173,
  BrtCustomFilter = 174,
  BrtAFilterDateGroupItem = 175,
  BrtMergeCell = 176,
  BrtBeginMergeCells = 177,
  BrtEndMergeCells = 178,
  BrtBeginPivotCacheDef = 179,
  BrtEndPivotCacheDef = 180,
  BrtBeginPCDFields = 181,
  BrtEndPCDFields = 182,
  BrtBeginPCDField = 183,
  BrtEndPCDField = 184,
  BrtBeginPCDSource = 185,
  BrtEndPCDSource = 186,
  BrtBeginPCDSRange = 187,
  BrtEndPCDSRange = 188,
  BrtBeginPCDFAtbl = 189,
  BrtEndPCDFAtbl = 190,
  BrtBeginPCDIRun = 191,
  BrtEndPCDIRun = 192,
  BrtBeginPivotCacheRecords = 193,
  BrtEndPivotCacheRecords = 194,
  BrtBeginPCDHierarchies = 195,
  BrtEndPCDHierarchies = 196,
  BrtBeginPCDHierarchy = 197,
  BrtEndPCDHierarchy = 198,
  BrtBeginPCDHFieldsUsage = 199,
  BrtEndPCDHFieldsUsage = 200,
  BrtBeginExtConnection = 201,
  BrtEndExtConnection = 202,
  BrtBeginECDbProps = 203,
  BrtEndECDbProps = 204,
  BrtBeginECOlapProps = 205,
  BrtEndECOlapProps = 206,
  BrtBeginPCDSConsol = 207,
  BrtEndPCDSConsol = 208,
  BrtBeginPCDSCPages = 209,
  BrtEndPCDSCPages = 210,
  BrtBeginPCDSCPage = 211,
  BrtEndPCDSCPage = 212,
  BrtBeginPCDSCPItem = 213,
  BrtEndPCDSCPItem = 214,
  BrtBeginPCDSCSets = 215,
  BrtEndPCDSCSets = 216,
  BrtBeginPCDSCSet = 217,
  BrtEndPCDSCSet = 218,
  BrtBeginPCDFGroup = 219,
  BrtEndPCDFGroup = 220,
  BrtBeginPCDFGItems = 221,
  BrtEndPCDFGItems = 222,
  BrtBeginPCDFGRange = 223,
  BrtEndPCDFGRange = 224,
  BrtBeginPCDFGDiscrete = 225,
  BrtEndPCDFGDiscrete = 226,
  BrtBeginPCDSDTupleCache = 227,
  BrtEndPCDSDTupleCache = 228,
  BrtBeginPCDSDTCEntries = 229,
  BrtEndPCDSDTCEntries = 230,
  BrtBeginPCDSDTCEMembers = 231,
  BrtEndPCDSDTCEMembers = 232,
  BrtBeginPCDSDTCEMember = 233,
  BrtEndPCDSDTCEMember = 234,
  BrtBeginPCDSDTCQueries = 235,
  BrtEndPCDSDTCQueries = 236,
  BrtBeginPCDSDTCQuery = 237,
  BrtEndPCDSDTCQuery = 238,
  BrtBeginPCDSDTCSets = 239,
  BrtEndPCDSDTCSets = 240,
  BrtBeginPCDSDTCSet = 241,
  BrtEndPCDSDTCSet = 242,
  BrtBeginPCDCalcItems = 243,
  BrtEndPCDCalcItems = 244,
  BrtBeginPCDCalcItem = 245,
  BrtEndPCDCalcItem = 246,
  BrtBeginPRule = 247,
  BrtEndPRule = 248,
  BrtBeginPRFilters = 249,
  BrtEndPRFilters = 250,
  BrtBeginPRFilter = 251,
  BrtEndPRFilter = 252,
  BrtBeginPNames = 253,
  BrtEndPNames = 254,
  BrtBeginPName = 255,
  BrtEndPName = 256,
  BrtBeginPNPairs = 257,
  BrtEndPNPairs = 258,
  BrtBeginPNPair = 259,
  BrtEndPNPair = 260,
  BrtBeginECWebProps = 261,
  BrtEndECWebProps = 262,
  BrtBeginEcWpTables = 263,
  BrtEndECWPTables = 264,
  BrtBeginECParams = 265,
  BrtEndECParams = 266,
  BrtBeginECParam = 267,
  BrtEndECParam = 268,
  BrtBeginPCDKPIs = 269,
  BrtEndPCDKPIs = 270,
  BrtBeginPCDKPI = 271,
  BrtEndPCDKPI = 272,
  BrtBeginDims = 273,
  BrtEndDims = 274,
  BrtBeginDim = 275,
  BrtEndDim = 276,
  BrtIndexPartEnd = 277,
  BrtBeginStyleSheet = 278,
  BrtEndStyleSheet = 279,
  BrtBeginSXView = 280,
  BrtEndSXVI = 281,
  BrtBeginSXVI = 282,
  BrtBeginSXVIs = 283,
  BrtEndSXVIs = 284,
  BrtBeginSXVD = 285,
  BrtEndSXVD = 286,
  BrtBeginSXVDs = 287,
  BrtEndSXVDs = 288,
  BrtBeginSXPI = 289,
  BrtEndSXPI = 290,
  BrtBeginSXPIs = 291,
  BrtEndSXPIs = 292,
  BrtBeginSXDI = 293,
  BrtEndSXDI = 294,
  BrtBeginSXDIs = 295,
  BrtEndSXDIs = 296,
  BrtBeginSXLI = 297,
  BrtEndSXLI = 298,
  BrtBeginSXLIRws = 299,
  BrtEndSXLIRws = 300,
  BrtBeginSXLICols = 301,
  BrtEndSXLICols = 302,
  BrtBeginSXFormat = 303,
  BrtEndSXFormat = 304,
  BrtBeginSXFormats = 305,
  BrtEndSxFormats = 306,
  BrtBeginSxSelect = 307,
  BrtEndSxSelect = 308,
  BrtBeginISXVDRws = 309,
  BrtEndISXVDRws = 310,
  BrtBeginISXVDCols = 311,
  BrtEndISXVDCols = 312,
  BrtEndSXLocation = 313,
  BrtBeginSXLocation = 314,
  BrtEndSXView = 315,
  BrtBeginSXTHs = 316,
  BrtEndSXTHs = 317,
  BrtBeginSXTH = 318,
  BrtEndSXTH = 319,
  BrtBeginISXTHRws = 320,
  BrtEndISXTHRws = 321,
  BrtBeginISXTHCols = 322,
  BrtEndISXTHCols = 323,
  BrtBeginSXTDMPS = 324,
  BrtEndSXTDMPs = 325,
  BrtBeginSXTDMP = 326,
  BrtEndSXTDMP = 327,
  BrtBeginSXTHItems = 328,
  BrtEndSXTHItems = 329,
  BrtBeginSXTHItem = 330,
  BrtEndSXTHItem = 331,
  BrtBeginMetadata = 332,
  BrtEndMetadata = 333,
  BrtBeginEsmdtinfo = 334,
  BrtMdtinfo = 335,
  BrtEndEsmdtinfo = 336,
  BrtBeginEsmdb = 337,
  BrtEndEsmdb = 338,
  BrtBeginEsfmd = 339,
  BrtEndEsfmd = 340,
  BrtBeginSingleCells = 341,
  BrtEndSingleCells = 342,
  BrtBeginList = 343,
  BrtEndList = 344,
  BrtBeginListCols = 345,
  BrtEndListCols = 346,
  BrtBeginListCol = 347,
  BrtEndListCol = 348,
  BrtBeginListXmlCPr = 349,
  BrtEndListXmlCPr = 350,
  BrtListCCFmla = 351,
  BrtListTrFmla = 352,
  BrtBeginExternals = 353,
  BrtEndExternals = 354,
  BrtSupBookSrc = 355,
  BrtSupSelf = 357,
  BrtSupSame = 358,
  BrtSupTabs = 359,
  BrtBeginSupBook = 360,
  BrtPlaceholderName = 361,
  BrtExternSheet = 362,
  BrtExternTableStart = 363,
  BrtExternTableEnd = 364,
  BrtExternRowHdr = 366,
  BrtExternCellBlank = 367,
  BrtExternCellReal = 368,
  BrtExternCellBool = 369,
  BrtExternCellError = 370,
  BrtExternCellString = 371,
  BrtBeginEsmdx = 372,
  BrtEndEsmdx = 373,
  BrtBeginMdxSet = 374,
  BrtEndMdxSet = 375,
  BrtBeginMdxMbrProp = 376,
  BrtEndMdxMbrProp = 377,
  BrtBeginMdxKPI = 378,
  BrtEndMdxKPI = 379,
  BrtBeginEsstr = 380,
  BrtEndEsstr = 381,
  BrtBeginPRFItem = 382,
  BrtEndPRFItem = 383,
  BrtBeginPivotCacheIDs = 384,
  BrtEndPivotCacheIDs = 385,
  BrtBeginPivotCacheID = 386,
  BrtEndPivotCacheID = 387,
  BrtBeginISXVIs = 388,
  BrtEndISXVIs = 389,
  BrtBeginColInfos = 390,
  BrtEndColInfos = 391,
  BrtBeginRwBrk = 392,
  BrtEndRwBrk = 393,
  BrtBeginColBrk = 394,
  BrtEndColBrk = 395,
  BrtBrk = 396,
  BrtUserBookView = 397,
  BrtInfo = 398,
  BrtCUsr = 399,
  BrtUsr = 400,
  BrtBeginUsers = 401,
  BrtEOF = 403,
  BrtUCR = 404,
  BrtRRInsDel = 405,
  BrtRREndInsDel = 406,
  BrtRRMove = 407,
  BrtRREndMove = 408,
  BrtRRChgCell = 409,
  BrtRREndChgCell = 410,
  BrtRRHeader = 411,
  BrtRRUserView = 412,
  BrtRRRenSheet = 413,
  BrtRRInsertSh = 414,
  BrtRRDefName = 415,
  BrtRRNote = 416,
  BrtRRConflict = 417,
  BrtRRTQSIF = 418,
  BrtRRFormat = 419,
  BrtRREndFormat = 420,
  BrtRRAutoFmt = 421,
  BrtBeginUserShViews = 422,
  BrtBeginUserShView = 423,
  BrtEndUserShView = 424,
  BrtEndUserShViews = 425,
  BrtArrFmla = 426,
  BrtShrFmla = 427,
  BrtTable = 428,
  BrtBeginExtConnections = 429,
  BrtEndExtConnections = 430,
  BrtBeginPCDCalcMems = 431,
  BrtEndPCDCalcMems = 432,
  BrtBeginPCDCalcMem = 433,
  BrtEndPCDCalcMem = 434,
  BrtBeginPCDHGLevels = 435,
  BrtEndPCDHGLevels = 436,
  BrtBeginPCDHGLevel = 437,
  BrtEndPCDHGLevel = 438,
  BrtBeginPCDHGLGroups = 439,
  BrtEndPCDHGLGroups = 440,
  BrtBeginPCDHGLGroup = 441,
  BrtEndPCDHGLGroup = 442,
  BrtBeginPCDHGLGMembers = 443,
  BrtEndPCDHGLGMembers = 444,
  BrtBeginPCDHGLGMember = 445,
  BrtEndPCDHGLGMember = 446,
  BrtBeginQSI = 447,
  BrtEndQSI = 448,
  BrtBeginQSIR = 449,
  BrtEndQSIR = 450,
  BrtBeginDeletedNames = 451,
  BrtEndDeletedNames = 452,
  BrtBeginDeletedName = 453,
  BrtEndDeletedName = 454,
  BrtBeginQSIFs = 455,
  BrtEndQSIFs = 456,
  BrtBeginQSIF = 457,
  BrtEndQSIF = 458,
  BrtBeginAutoSortScope = 459,
  BrtEndAutoSortScope = 460,
  BrtBeginConditionalFormatting = 461,
  BrtEndConditionalFormatting = 462,
  BrtBeginCFRule = 463,
  BrtEndCFRule = 464,
  BrtBeginIconSet = 465,
  BrtEndIconSet = 466,
  BrtBeginDatabar = 467,
  BrtEndDatabar = 468,
  BrtBeginColorScale = 469,
  BrtEndColorScale = 470,
  BrtCFVO = 471,
  BrtExternValueMeta = 472,
  BrtBeginColorPalette = 473,
  BrtEndColorPalette = 474,
  BrtIndexedColor = 475,
  BrtMargins = 476,
  BrtPrintOptions = 477,
  BrtPageSetup = 478,
  BrtBeginHeaderFooter = 479,
  BrtEndHeaderFooter = 480,
  BrtBeginSXCrtFormat = 481,
  BrtEndSXCrtFormat = 482,
  BrtBeginSXCrtFormats = 483,
  BrtEndSXCrtFormats = 484,
  BrtWsFmtInfo = 485,
  BrtBeginMgs = 486,
  BrtEndMGs = 487,
  BrtBeginMGMaps = 488,
  BrtEndMGMaps = 489,
  BrtBeginMG = 490,
  BrtEndMG = 491,
  BrtBeginMap = 492,
  BrtEndMap = 493,
  BrtHLink = 494,
  BrtBeginDCon = 495,
  BrtEndDCon = 496,
  BrtBeginDRefs = 497,
  BrtEndDRefs = 498,
  BrtDRef = 499,
  BrtBeginScenMan = 500,
  BrtEndScenMan = 501,
  BrtBeginSct = 502,
  BrtEndSct = 503,
  BrtSlc = 504,
  BrtBeginDXFs = 505,
  BrtEndDXFs = 506,
  BrtDXF = 507,
  BrtBeginTableStyles = 508,
  BrtEndTableStyles = 509,
  BrtBeginTableStyle = 510,
  BrtEndTableStyle = 511,
  BrtTableStyleElement = 512,
  BrtTableStyleClient = 513,
  BrtBeginVolDeps = 514,
  BrtEndVolDeps = 515,
  BrtBeginVolType = 516,
  BrtEndVolType = 517,
  BrtBeginVolMain = 518,
  BrtEndVolMain = 519,
  BrtBeginVolTopic = 520,
  BrtEndVolTopic = 521,
  BrtVolSubtopic = 522,
  BrtVolRef = 523,
  BrtVolNum = 524,
  BrtVolErr = 525,
  BrtVolStr = 526,
  BrtVolBool = 527,
  BrtBeginSortState = 530,
  BrtEndSortState = 531,
  BrtBeginSortCond = 532,
  BrtEndSortCond = 533,
  BrtBookProtection = 534,
  BrtSheetProtection = 535,
  BrtRangeProtection = 536,
  BrtPhoneticInfo = 537,
  BrtBeginECTxtWiz = 538,
  BrtEndECTxtWiz = 539,
  BrtBeginECTWFldInfoLst = 540,
  BrtEndECTWFldInfoLst = 541,
  BrtBeginECTwFldInfo = 542,
  BrtFileSharing = 548,
  BrtOleSize = 549,
  BrtDrawing = 550,
  BrtLegacyDrawing = 551,
  BrtLegacyDrawingHF = 552,
  BrtWebOpt = 553,
  BrtBeginWebPubItems = 554,
  BrtEndWebPubItems = 555,
  BrtBeginWebPubItem = 556,
  BrtEndWebPubItem = 557,
  BrtBeginSXCondFmt = 558,
  BrtEndSXCondFmt = 559,
  BrtBeginSXCondFmts = 560,
  BrtEndSXCondFmts = 561,
  BrtBkHim = 562,
  BrtColor = 564,
  BrtBeginIndexedColors = 565,
  BrtEndIndexedColors = 566,
  BrtBeginMRUColors = 569,
  BrtEndMRUColors = 570,
  BrtMRUColor = 572,
  BrtBeginDVals = 573,
  BrtEndDVals = 574,
  BrtSupNameStart = 577,
  BrtSupNameValueStart = 578,
  BrtSupNameValueEnd = 579,
  BrtSupNameNum = 580,
  BrtSupNameErr = 581,
  BrtSupNameSt = 582,
  BrtSupNameNil = 583,
  BrtSupNameBool = 584,
  BrtSupNameFmla = 585,
  BrtSupNameBits = 586,
  BrtSupNameEnd = 587,
  BrtEndSupBook = 588,
  BrtCellSmartTagProperty = 589,
  BrtBeginCellSmartTag = 590,
  BrtEndCellSmartTag = 591,
  BrtBeginCellSmartTags = 592,
  BrtEndCellSmartTags = 593,
  BrtBeginSmartTags = 594,
  BrtEndSmartTags = 595,
  BrtSmartTagType = 596,
  BrtBeginSmartTagTypes = 597,
  BrtEndSmartTagTypes = 598,
  BrtBeginSXFilters = 599,
  BrtEndSXFilters = 600,
  BrtBeginSXFILTER = 601,
  BrtEndSXFilter = 602,
  BrtBeginFills = 603,
  BrtEndFills = 604,
  BrtBeginCellWatches = 605,
  BrtEndCellWatches = 606,
  BrtCellWatch = 607,
  BrtBeginCRErrs = 608,
  BrtEndCRErrs = 609,
  BrtCrashRecErr = 610,
  BrtBeginFonts = 611,
  BrtEndFonts = 612,
  BrtBeginBorders = 613,
  BrtEndBorders = 614,
  BrtBeginFmts = 615,
  BrtEndFmts = 616,
  BrtBeginCellXFs = 617,
  BrtEndCellXFs = 618,
  BrtBeginStyles = 619,
  BrtEndStyles = 620,
  BrtBigName = 625,
  BrtBeginCellStyleXFs = 626,
  BrtEndCellStyleXFs = 627,
  BrtBeginComments = 628,
  BrtEndComments = 629,
  BrtBeginCommentAuthors = 630,
  BrtEndCommentAuthors = 631,
  BrtCommentAuthor = 632,
  BrtBeginCommentList = 633,
  BrtEndCommentList = 634,
  BrtBeginComment = 635,
  BrtEndComment = 636,
  BrtCommentText = 637,
  BrtBeginOleObjects = 638,
  BrtOleObject = 639,
  BrtEndOleObjects = 640,
  BrtBeginSxrules = 641,
  BrtEndSxRules = 642,
  BrtBeginActiveXControls = 643,
  BrtActiveX = 644,
  BrtEndActiveXControls = 645,
  BrtBeginPCDSDTCEMembersSortBy = 646,
  BrtBeginCellIgnoreECs = 648,
  BrtCellIgnoreEC = 649,
  BrtEndCellIgnoreECs = 650,
  BrtCsProp = 651,
  BrtCsPageSetup = 652,
  BrtBeginUserCsViews = 653,
  BrtEndUserCsViews = 654,
  BrtBeginUserCsView = 655,
  BrtEndUserCsView = 656,
  BrtBeginPcdSFCIEntries = 657,
  BrtEndPCDSFCIEntries = 658,
  BrtPCDSFCIEntry = 659,
  BrtBeginListParts = 660,
  BrtListPart = 661,
  BrtEndListParts = 662,
  BrtSheetCalcProp = 663,
  BrtBeginFnGroup = 664,
  BrtFnGroup = 665,
  BrtEndFnGroup = 666,
  BrtSupAddin = 667,
  BrtSXTDMPOrder = 668,
  BrtCsProtection = 669,
  BrtBeginWsSortMap = 671,
  BrtEndWsSortMap = 672,
  BrtBeginRRSort = 673,
  BrtEndRRSort = 674,
  BrtRRSortItem = 675,
  BrtFileSharingIso = 676,
  BrtBookProtectionIso = 677,
  BrtSheetProtectionIso = 678,
  BrtCsProtectionIso = 679,
  BrtRangeProtectionIso = 680,
  BrtDValList = 681,
  BrtRwDescent = 1024,
  BrtKnownFonts = 1025,
  BrtBeginSXTupleSet = 1026,
  BrtEndSXTupleSet = 1027,
  BrtBeginSXTupleSetHeader = 1028,
  BrtEndSXTupleSetHeader = 1029,
  BrtSXTupleSetHeaderItem = 1030,
  BrtBeginSXTupleSetData = 1031,
  BrtEndSXTupleSetData = 1032,
  BrtBeginSXTupleSetRow = 1033,
  BrtEndSXTupleSetRow = 1034,
  BrtSXTupleSetRowItem = 1035,
  BrtNameExt = 1036,
  BrtPCDH14 = 1037,
  BrtBeginPCDCalcMem14 = 1038,
  BrtEndPCDCalcMem14 = 1039,
  BrtSXTH14 = 1040,
  BrtBeginSparklineGroup = 1041,
  BrtEndSparklineGroup = 1042,
  BrtSparkline = 1043,
  BrtSXDI14 = 1044,
  BrtWsFmtInfoEx14 = 1045,
  BrtBeginConditionalFormatting14 = 1046,
  BrtEndConditionalFormatting14 = 1047,
  BrtBeginCFRule14 = 1048,
  BrtEndCFRule14 = 1049,
  BrtCFVO14 = 1050,
  BrtBeginDatabar14 = 1051,
  BrtBeginIconSet14 = 1052,
  BrtDVal14 = 1053,
  BrtBeginDVals14 = 1054,
  BrtColor14 = 1055,
  BrtBeginSparklines = 1056,
  BrtEndSparklines = 1057,
  BrtBeginSparklineGroups = 1058,
  BrtEndSparklineGroups = 1059,
  BrtSXVD14 = 1061,
  BrtBeginSxView14 = 1062,
  BrtEndSxView14 = 1063,
  BrtBeginSXView16 = 1064,
  BrtEndSXView16 = 1065,
  BrtBeginPCD14 = 1066,
  BrtEndPCD14 = 1067,
  BrtBeginExtConn14 = 1068,
  BrtEndExtConn14 = 1069,
  BrtBeginSlicerCacheIDs = 1070,
  BrtEndSlicerCacheIDs = 1071,
  BrtBeginSlicerCacheID = 1072,
  BrtEndSlicerCacheID = 1073,
  BrtBeginSlicerCache = 1075,
  BrtEndSlicerCache = 1076,
  BrtBeginSlicerCacheDef = 1077,
  BrtEndSlicerCacheDef = 1078,
  BrtBeginSlicersEx = 1079,
  BrtEndSlicersEx = 1080,
  BrtBeginSlicerEx = 1081,
  BrtEndSlicerEx = 1082,
  BrtBeginSlicer = 1083,
  BrtEndSlicer = 1084,
  BrtSlicerCachePivotTables = 1085,
  BrtBeginSlicerCacheOlapImpl = 1086,
  BrtEndSlicerCacheOlapImpl = 1087,
  BrtBeginSlicerCacheLevelsData = 1088,
  BrtEndSlicerCacheLevelsData = 1089,
  BrtBeginSlicerCacheLevelData = 1090,
  BrtEndSlicerCacheLevelData = 1091,
  BrtBeginSlicerCacheSiRanges = 1092,
  BrtEndSlicerCacheSiRanges = 1093,
  BrtBeginSlicerCacheSiRange = 1094,
  BrtEndSlicerCacheSiRange = 1095,
  BrtSlicerCacheOlapItem = 1096,
  BrtBeginSlicerCacheSelections = 1097,
  BrtSlicerCacheSelection = 1098,
  BrtEndSlicerCacheSelections = 1099,
  BrtBeginSlicerCacheNative = 1100,
  BrtEndSlicerCacheNative = 1101,
  BrtSlicerCacheNativeItem = 1102,
  BrtRangeProtection14 = 1103,
  BrtRangeProtectionIso14 = 1104,
  BrtCellIgnoreEC14 = 1105,
  BrtList14 = 1111,
  BrtCFIcon = 1112,
  BrtBeginSlicerCachesPivotCacheIDs = 1113,
  BrtEndSlicerCachesPivotCacheIDs = 1114,
  BrtBeginSlicers = 1115,
  BrtEndSlicers = 1116,
  BrtWbProp14 = 1117,
  BrtBeginSXEdit = 1118,
  BrtEndSXEdit = 1119,
  BrtBeginSXEdits = 1120,
  BrtEndSXEdits = 1121,
  BrtBeginSXChange = 1122,
  BrtEndSXChange = 1123,
  BrtBeginSXChanges = 1124,
  BrtEndSXChanges = 1125,
  BrtSXTupleItems = 1126,
  BrtBeginSlicerStyle = 1128,
  BrtEndSlicerStyle = 1129,
  BrtSlicerStyleElement = 1130,
  BrtBeginStyleSheetExt14 = 1131,
  BrtEndStyleSheetExt14 = 1132,
  BrtBeginSlicerCachesPivotCacheID = 1133,
  BrtEndSlicerCachesPivotCacheID = 1134,
  BrtBeginConditionalFormattings = 1135,
  BrtEndConditionalFormattings = 1136,
  BrtBeginPCDCalcMemExt = 1137,
  BrtEndPCDCalcMemExt = 1138,
  BrtBeginPCDCalcMemsExt = 1139,
  BrtEndPCDCalcMemsExt = 1140,
  BrtPCDField14 = 1141,
  BrtBeginSlicerStyles = 1142,
  BrtEndSlicerStyles = 1143,
  BrtBeginSlicerStyleElements = 1144,
  BrtEndSlicerStyleElements = 1145,
  BrtCFRuleExt = 1146,
  BrtBeginSXCondFmt14 = 1147,
  BrtEndSXCondFmt14 = 1148,
  BrtBeginSXCondFmts14 = 1149,
  BrtEndSXCondFmts14 = 1150,
  BrtBeginSortCond14 = 1152,
  BrtEndSortCond14 = 1153,
  BrtEndDVals14 = 1154,
  BrtEndIconSet14 = 1155,
  BrtEndDatabar14 = 1156,
  BrtBeginColorScale14 = 1157,
  BrtEndColorScale14 = 1158,
  BrtBeginSxrules14 = 1159,
  BrtEndSxrules14 = 1160,
  BrtBeginPRule14 = 1161,
  BrtEndPRule14 = 1162,
  BrtBeginPRFilters14 = 1163,
  BrtEndPRFilters14 = 1164,
  BrtBeginPRFilter14 = 1165,
  BrtEndPRFilter14 = 1166,
  BrtBeginPRFItem14 = 1167,
  BrtEndPRFItem14 = 1168,
  BrtBeginCellIgnoreECs14 = 1169,
  BrtEndCellIgnoreECs14 = 1170,
  BrtDxf14 = 1171,
  BrtBeginDxF14s = 1172,
  BrtEndDxf14s = 1173,
  BrtFilter14 = 1177,
  BrtBeginCustomFilters14 = 1178,
  BrtCustomFilter14 = 1180,
  BrtIconFilter14 = 1181,
  BrtPivotCacheConnectionName = 1182,
  BrtBeginDecoupledPivotCacheIDs = 2048,
  BrtEndDecoupledPivotCacheIDs = 2049,
  BrtDecoupledPivotCacheID = 2050,
  BrtBeginPivotTableRefs = 2051,
  BrtEndPivotTableRefs = 2052,
  BrtPivotTableRef = 2053,
  BrtSlicerCacheBookPivotTables = 2054,
  BrtBeginSxvcells = 2055,
  BrtEndSxvcells = 2056,
  BrtBeginSxRow = 2057,
  BrtEndSxRow = 2058,
  BrtPcdCalcMem15 = 2060,
  BrtQsi15 = 2067,
  BrtBeginWebExtensions = 2068,
  BrtEndWebExtensions = 2069,
  BrtWebExtension = 2070,
  BrtAbsPath15 = 2071,
  BrtBeginPivotTableUISettings = 2072,
  BrtEndPivotTableUISettings = 2073,
  BrtTableSlicerCacheIDs = 2075,
  BrtTableSlicerCacheID = 2076,
  BrtBeginTableSlicerCache = 2077,
  BrtEndTableSlicerCache = 2078,
  BrtSxFilter15 = 2079,
  BrtBeginTimelineCachePivotCacheIDs = 2080,
  BrtEndTimelineCachePivotCacheIDs = 2081,
  BrtTimelineCachePivotCacheID = 2082,
  BrtBeginTimelineCacheIDs = 2083,
  BrtEndTimelineCacheIDs = 2084,
  BrtBeginTimelineCacheID = 2085,
  BrtEndTimelineCacheID = 2086,
  BrtBeginTimelinesEx = 2087,
  BrtEndTimelinesEx = 2088,
  BrtBeginTimelineEx = 2089,
  BrtEndTimelineEx = 2090,
  BrtWorkBookPr15 = 2091,
  BrtPCDH15 = 2092,
  BrtBeginTimelineStyle = 2093,
  BrtEndTimelineStyle = 2094,
  BrtTimelineStyleElement = 2095,
  BrtBeginTimelineStylesheetExt15 = 2096,
  BrtEndTimelineStylesheetExt15 = 2097,
  BrtBeginTimelineStyles = 2098,
  BrtEndTimelineStyles = 2099,
  BrtBeginTimelineStyleElements = 2100,
  BrtEndTimelineStyleElements = 2101,
  BrtDxf15 = 2102,
  BrtBeginDxfs15 = 2103,
  BrtEndDXFs15 = 2104,
  BrtSlicerCacheHideItemsWithNoData = 2105,
  BrtBeginItemUniqueNames = 2106,
  BrtEndItemUniqueNames = 2107,
  BrtItemUniqueName = 2108,
  BrtBeginExtConn15 = 2109,
  BrtEndExtConn15 = 2110,
  BrtBeginOledbPr15 = 2111,
  BrtEndOledbPr15 = 2112,
  BrtBeginDataFeedPr15 = 2113,
  BrtEndDataFeedPr15 = 2114,
  BrtTextPr15 = 2115,
  BrtRangePr15 = 2116,
  BrtDbCommand15 = 2117,
  BrtBeginDbTables15 = 2118,
  BrtEndDbTables15 = 2119,
  BrtDbTable15 = 2120,
  BrtBeginDataModel = 2121,
  BrtEndDataModel = 2122,
  BrtBeginModelTables = 2123,
  BrtEndModelTables = 2124,
  BrtModelTable = 2125,
  BrtBeginModelRelationships = 2126,
  BrtEndModelRelationships = 2127,
  BrtModelRelationship = 2128,
  BrtBeginECTxtWiz15 = 2129,
  BrtEndECTxtWiz15 = 2130,
  BrtBeginECTWFldInfoLst15 = 2131,
  BrtEndECTWFldInfoLst15 = 2132,
  BrtBeginECTWFldInfo15 = 2133,
  BrtFieldListActiveItem = 2134,
  BrtPivotCacheIdVersion = 2135,
  BrtSXDI15 = 2136,
  BrtBeginModelTimeGroupings = 2137,
  BrtEndModelTimeGroupings = 2138,
  BrtBeginModelTimeGrouping = 2139,
  BrtEndModelTimeGrouping = 2140,
  BrtModelTimeGroupingCalcCol = 2141,
  BrtUID = 3072, /* No longer part of 2.3.2 By Number */
  BrtRevisionPtr = 3073,
  BrtBeginDynamicArrayPr = 4096,
  BrtEndDynamicArrayPr = 4097,
  BrtBeginRichValueBlock = 5002,
  BrtEndRichValueBlock = 5003,
  BrtBeginRichFilters = 5081,
  BrtEndRichFilters = 5082,
  BrtRichFilter = 5083,
  BrtBeginRichFilterColumn = 5084,
  BrtEndRichFilterColumn = 5085,
  BrtBeginCustomRichFilters = 5086,
  BrtEndCustomRichFilters = 5087,
  BRTCustomRichFilter = 5088,
  BrtTop10RichFilter = 5089,
  BrtDynamicRichFilter = 5090,
  BrtBeginRichSortCondition = 5092,
  BrtEndRichSortCondition = 5093,
  BrtRichFilterDateGroupItem = 5094,
  BrtBeginCalcFeatures = 5095,
  BrtEndCalcFeatures = 5096,
  BrtCalcFeature = 5097,
  BrtExternalLinksPr = 5099,
  BrtPivotCacheImplicitMeasureSupport = 5100,
  BrtPivotFieldIgnorableAfter = 5101,
  BrtPivotHierarchyIgnorableAfter = 5102,
  BrtPivotDataFieldFutureData = 5103,
  BrtPivotCacheRichData = 5105,
  BrtExternalLinksAlternateUrls = 5108,
  BrtBeginPivotVersionInfo = 5109,
  BrtEndPivotVersionInfo = 5110,
  BrtBeginCacheVersionInfo = 5111,
  BrtEndCacheVersionInfo = 5112, /* possible typo in 2.3.2 By Number */
  BrtPivotRequiredFeature = 5113,
  BrtPivotLastUsedFeature = 5114,
  BrtExternalCodeService = 5117,
  BrtSXDIAggregation = 5130,
  BrtPivotFieldFeatureSupportInfo = 5131,
  BrtPivotCacheAutoRefresh = 5132
};

enum PtgTypes
{
  PtgExp = 0x01,
  PtgAdd = 0x03,
  PtgSub = 0x04,
  PtgMul = 0x05,
  PtgDiv = 0x06,
  PtgPower = 0x07,
  PtgConcat = 0x08,
  PtgLt = 0x09,
  PtgLe = 0x0A,
  PtgEq = 0x0B,
  PtgGe = 0x0C,
  PtgGt = 0x0D,
  PtgNe = 0x0E,
  PtgIsect = 0x0F,
  PtgUnion = 0x10,
  PtgRange = 0x11,
  PtgUPlus = 0x12,
  PtgUMinus = 0x13,
  PtgPercent = 0x14,
  PtgParen = 0x15,
  PtgMissArg = 0x16,
  PtgStr = 0x17,
  PtgList_PtgSxName = 0x18,
  PtgAttr = 0x19,
  PtgErr = 0x1C,
  PtgBool = 0x1D,
  PtgInt = 0x1E,
  PtgNum = 0x1F,
  PtgArray = 0x20,
  PtgFunc = 0x21,
  PtgFuncVar = 0x22,
  PtgName = 0x23,
  PtgRef = 0x24,
  PtgArea = 0x25,
  PtgMemArea = 0x26,
  PtgMemErr = 0x27,
  PtgMemNoMem = 0x28,
  PtgMemFunc = 0x29,
  PtgRefErr = 0x2A,
  PtgAreaErr = 0x2B,
  PtgRefN = 0x2C,
  PtgAreaN = 0x2D,
  PtgNameX = 0x39,
  PtgRef3d = 0x3A,
  PtgArea3d = 0x3B,
  PtgRefErr3d = 0x3C,
  PtgAreaErr3d = 0x3D,
  PtgArray2 = 0x40,
  PtgFunc2 = 0x41,
  PtgFuncVar2 = 0x42,
  PtgName2 = 0x43,
  PtgRef2 = 0x44,
  PtgArea2 = 0x45,
  PtgMemArea2 = 0x46,
  PtgMemErr2 = 0x47,
  PtgMemNoMem2 = 0x48,
  PtgMemFunc2 = 0x49,
  PtgRefErr2 = 0x4A,
  PtgAreaErr2 = 0x4B,
  PtgRefN2 = 0x4C,
  PtgAreaN2 = 0x4D,
  PtgNameX2 = 0x59,
  PtgRef3d2 = 0x5A,
  PtgArea3d2 = 0x5B,
  PtgRefErr3d2 = 0x5C,
  PtgAreaErr3d2 = 0x5D,
  PtgArray3 = 0x60,
  PtgFunc3 = 0x61,
  PtgFuncVar3 = 0x62,
  PtgName3 = 0x63,
  PtgRef3 = 0x64,
  PtgArea3 = 0x65,
  PtgMemArea3 = 0x66,
  PtgMemErr3 = 0x67,
  PtgMemNoMem3 = 0x68,
  PtgMemFunc3 = 0x69,
  PtgRefErr3 = 0x6A,
  PtgAreaErr3 = 0x6B,
  PtgRefN3 = 0x6C,
  PtgAreaN3 = 0x6D,
  PtgNameX3 = 0x79,
  PtgRef3d3 = 0x7A,
  PtgArea3d3 = 0x7B,
  PtgRefErr3d3 = 0x7C,
  PtgAreaErr3d3 = 0x7D
};

enum PtgStructure2
{
  PtgList = 0x19,
  PtgSxName = 0x1D,
  PtgAttrSemi = 0x01,
  PtgAttrIf = 0x02,
  PtgAttrChoose = 0x04,
  PtgAttrGoTo = 0x08,
  PtgAttrSum = 0x10,
  PtgAttrBaxcel = 0x20,
  PtgAttrBaxcel2 = 0x21,
  PtgAttrSpace = 0x40,
  PtgAttrSpaceSemi = 0x41,
  PtgAttrIfError = 0x80
};

// #nocov start
// copied from the website
std::string Cetab(const uint16_t val) {
  switch(val) {
  case	0x0000: return "BEEP";
  case	0x0001: return "OPEN";
  case	0x0002: return "OPEN.LINKS";
  case	0x0003: return "CLOSE.ALL";
  case	0x0004: return "SAVE";
  case	0x0005: return "SAVE.AS";
  case	0x0006: return "FILE.DELETE";
  case	0x0007: return "PAGE.SETUP";
  case	0x0008: return "PRINT";
  case	0x0009: return "PRINTER.SETUP";
  case	0x000A: return "QUIT";
  case	0x000B: return "NEW.WINDOW";
  case	0x000C: return "ARRANGE.ALL";
  case	0x000D: return "WINDOW.SIZE";
  case	0x000E: return "WINDOW.MOVE";
  case	0x000F: return "FULL";
  case	0x0010: return "CLOSE";
  case	0x0011: return "RUN";
  case	0x0016: return "SET.PRINT.AREA";
  case	0x0017: return "SET.PRINT.TITLES";
  case	0x0018: return "SET.PAGE.BREAK";
  case	0x0019: return "REMOVE.PAGE.BREAK";
  case	0x001A: return "FONT";
  case	0x001B: return "DISPLAY";
  case	0x001C: return "PROTECT.DOCUMENT";
  case	0x001D: return "PRECISION";
  case	0x001E: return "A1.R1C1";
  case	0x001F: return "CALCULATE.NOW";
  case	0x0020: return "CALCULATION";
  case	0x0022: return "DATA.FIND";
  case	0x0023: return "EXTRACT";
  case	0x0024: return "DATA.DELETE";
  case	0x0025: return "SET.DATABASE";
  case	0x0026: return "SET.CRITERIA";
  case	0x0027: return "SORT";
  case	0x0028: return "DATA.SERIES";
  case	0x0029: return "TABLE";
  case	0x002A: return "FORMAT.NUMBER";
  case	0x002B: return "ALIGNMENT";
  case	0x002C: return "STYLE";
  case	0x002D: return "BORDER";
  case	0x002E: return "CELL.PROTECTION";
  case	0x002F: return "COLUMN.WIDTH";
  case	0x0030: return "UNDO";
  case	0x0031: return "CUT";
  case	0x0032: return "COPY";
  case	0x0033: return "PASTE";
  case	0x0034: return "CLEAR";
  case	0x0035: return "PASTE.SPECIAL";
  case	0x0036: return "EDIT.DELETE";
  case	0x0037: return "INSERT";
  case	0x0038: return "FILL.RIGHT";
  case	0x0039: return "FILL.DOWN";
  case	0x003D: return "DEFINE.NAME";
  case	0x003E: return "CREATE.NAMES";
  case	0x003F: return "FORMULA.GOTO";
  case	0x0040: return "FORMULA.FIND";
  case	0x0041: return "SELECT.LAST.CELL";
  case	0x0042: return "SHOW.ACTIVE.CELL";
  case	0x0043: return "GALLERY.AREA";
  case	0x0044: return "GALLERY.BAR";
  case	0x0045: return "GALLERY.COLUMN";
  case	0x0046: return "GALLERY.LINE";
  case	0x0047: return "GALLERY.PIE";
  case	0x0048: return "GALLERY.SCATTER";
  case	0x0049: return "COMBINATION";
  case	0x004A: return "PREFERRED";
  case	0x004B: return "ADD.OVERLAY";
  case	0x004C: return "GRIDLINES";
  case	0x004D: return "SET.PREFERRED";
  case	0x004E: return "AXES";
  case	0x004F: return "LEGEND";
  case	0x0050: return "ATTACH.TEXT";
  case	0x0051: return "ADD.ARROW";
  case	0x0052: return "SELECT.CHART";
  case	0x0053: return "SELECT.PLOT.AREA";
  case	0x0054: return "PATTERNS";
  case	0x0055: return "MAIN.CHART";
  case	0x0056: return "OVERLAY";
  case	0x0057: return "SCALE";
  case	0x0058: return "FORMAT.LEGEND";
  case	0x0059: return "FORMAT.TEXT";
  case	0x005A: return "EDIT.REPEAT";
  case	0x005B: return "PARSE";
  case	0x005C: return "JUSTIFY";
  case	0x005D: return "HIDE";
  case	0x005E: return "UNHIDE";
  case	0x005F: return "WORKSPACE";
  case	0x0060: return "FORMULA";
  case	0x0061: return "FORMULA.FILL";
  case	0x0062: return "FORMULA.ARRAY";
  case	0x0063: return "DATA.FIND.NEXT";
  case	0x0064: return "DATA.FIND.PREV";
  case	0x0065: return "FORMULA.FIND.NEXT";
  case	0x0066: return "FORMULA.FIND.PREV";
  case	0x0067: return "ACTIVATE";
  case	0x0068: return "ACTIVATE.NEXT";
  case	0x0069: return "ACTIVATE.PREV";
  case	0x006A: return "UNLOCKED.NEXT";
  case	0x006B: return "UNLOCKED.PREV";
  case	0x006C: return "COPY.PICTURE";
  case	0x006D: return "SELECT";
  case	0x006E: return "DELETE.NAME";
  case	0x006F: return "DELETE.FORMAT";
  case	0x0070: return "VLINE";
  case	0x0071: return "HLINE";
  case	0x0072: return "VPAGE";
  case	0x0073: return "HPAGE";
  case	0x0074: return "VSCROLL";
  case	0x0075: return "HSCROLL";
  case	0x0076: return "ALERT";
  case	0x0077: return "NEW";
  case	0x0078: return "CANCEL.COPY";
  case	0x0079: return "SHOW.CLIPBOARD";
  case	0x007A: return "MESSAGE";
  case	0x007C: return "PASTE.LINK";
  case	0x007D: return "APP.ACTIVATE";
  case	0x007E: return "DELETE.ARROW";
  case	0x007F: return "ROW.HEIGHT";
  case	0x0080: return "FORMAT.MOVE";
  case	0x0081: return "FORMAT.SIZE";
  case	0x0082: return "FORMULA.REPLACE";
  case	0x0083: return "SEND.KEYS";
  case	0x0084: return "SELECT.SPECIAL";
  case	0x0085: return "APPLY.NAMES";
  case	0x0086: return "REPLACE.FONT";
  case	0x0087: return "FREEZE.PANES";
  case	0x0088: return "SHOW.INFO";
  case	0x0089: return "SPLIT";
  case	0x008A: return "ON.WINDOW";
  case	0x008B: return "ON.DATA";
  case	0x008C: return "DISABLE.INPUT";
  case	0x008E: return "OUTLINE";
  case	0x008F: return "LIST.NAMES";
  case	0x0090: return "FILE.CLOSE";
  case	0x0091: return "SAVE.WORKBOOK";
  case	0x0092: return "DATA.FORM";
  case	0x0093: return "COPY.CHART";
  case	0x0094: return "ON.TIME";
  case	0x0095: return "WAIT";
  case	0x0096: return "FORMAT.FONT";
  case	0x0097: return "FILL.UP";
  case	0x0098: return "FILL.LEFT";
  case	0x0099: return "DELETE.OVERLAY";
  case	0x009B: return "SHORT.MENUS";
  case	0x009F: return "SET.UPDATE.STATUS";
  case	0x00A1: return "COLOR.PALETTE";
  case	0x00A2: return "DELETE.STYLE";
  case	0x00A3: return "WINDOW.RESTORE";
  case	0x00A4: return "WINDOW.MAXIMIZE";
  case	0x00A6: return "CHANGE.LINK";
  case	0x00A7: return "CALCULATE.DOCUMENT";
  case	0x00A8: return "ON.KEY";
  case	0x00A9: return "APP.RESTORE";
  case	0x00AA: return "APP.MOVE";
  case	0x00AB: return "APP.SIZE";
  case	0x00AC: return "APP.MINIMIZE";
  case	0x00AD: return "APP.MAXIMIZE";
  case	0x00AE: return "BRING.TO.FRONT";
  case	0x00AF: return "SEND.TO.BACK";
  case	0x00B9: return "MAIN.CHART.TYPE";
  case	0x00BA: return "OVERLAY.CHART.TYPE";
  case	0x00BB: return "SELECT.END";
  case	0x00BC: return "OPEN.MAIL";
  case	0x00BD: return "SEND.MAIL";
  case	0x00BE: return "STANDARD.FONT";
  case	0x00BF: return "CONSOLIDATE";
  case	0x00C0: return "SORT.SPECIAL";
  case	0x00C1: return "GALLERY.3D.AREA";
  case	0x00C2: return "GALLERY.3D.COLUMN";
  case	0x00C3: return "GALLERY.3D.LINE";
  case	0x00C4: return "GALLERY.3D.PIE";
  case	0x00C5: return "VIEW.3D";
  case	0x00C6: return "GOAL.SEEK";
  case	0x00C7: return "WORKGROUP";
  case	0x00C8: return "FILL.GROUP";
  case	0x00C9: return "UPDATE.LINK";
  case	0x00CA: return "PROMOTE";
  case	0x00CB: return "DEMOTE";
  case	0x00CC: return "SHOW.DETAIL";
  case	0x00CE: return "UNGROUP";
  case	0x00CF: return "OBJECT.PROPERTIES";
  case	0x00D0: return "SAVE.NEW.OBJECT";
  case	0x00D1: return "SHARE";
  case	0x00D2: return "SHARE.NAME";
  case	0x00D3: return "DUPLICATE";
  case	0x00D4: return "APPLY.STYLE";
  case	0x00D5: return "ASSIGN.TO.OBJECT";
  case	0x00D6: return "OBJECT.PROTECTION";
  case	0x00D7: return "HIDE.OBJECT";
  case	0x00D8: return "SET.EXTRACT";
  case	0x00D9: return "CREATE.PUBLISHER";
  case	0x00DA: return "SUBSCRIBE.TO";
  case	0x00DB: return "ATTRIBUTES";
  case	0x00DC: return "SHOW.TOOLBAR";
  case	0x00DE: return "PRINT.PREVIEW";
  case	0x00DF: return "EDIT.COLOR";
  case	0x00E0: return "SHOW.LEVELS";
  case	0x00E1: return "FORMAT.MAIN";
  case	0x00E2: return "FORMAT.OVERLAY";
  case	0x00E3: return "ON.RECALC";
  case	0x00E4: return "EDIT.SERIES";
  case	0x00E5: return "DEFINE.STYLE";
  case	0x00F0: return "LINE.PRINT";
  case	0x00F3: return "ENTER.DATA";
  case	0x00F9: return "GALLERY.RADAR";
  case	0x00FA: return "MERGE.STYLES";
  case	0x00FB: return "EDITION.OPTIONS";
  case	0x00FC: return "PASTE.PICTURE";
  case	0x00FD: return "PASTE.PICTURE.LINK";
  case	0x00FE: return "SPELLING";
  case	0x0100: return "ZOOM";
  case	0x0103: return "INSERT.OBJECT";
  case	0x0104: return "WINDOW.MINIMIZE";
  case	0x0109: return "SOUND.NOTE";
  case	0x010A: return "SOUND.PLAY";
  case	0x010B: return "FORMAT.SHAPE";
  case	0x010C: return "EXTEND.POLYGON";
  case	0x010D: return "FORMAT.AUTO";
  case	0x0110: return "GALLERY.3D.BAR";
  case	0x0111: return "GALLERY.3D.SURFACE";
  case	0x0112: return "FILL.AUTO";
  case	0x0114: return "CUSTOMIZE.TOOLBAR";
  case	0x0115: return "ADD.TOOL";
  case	0x0116: return "EDIT.OBJECT";
  case	0x0117: return "ON.DOUBLECLICK";
  case	0x0118: return "ON.ENTRY";
  case	0x0119: return "WORKBOOK.ADD";
  case	0x011A: return "WORKBOOK.MOVE";
  case	0x011B: return "WORKBOOK.COPY";
  case	0x011C: return "WORKBOOK.OPTIONS";
  case	0x011D: return "SAVE.WORKSPACE";
  case	0x0120: return "CHART.WIZARD";
  case	0x0121: return "DELETE.TOOL";
  case	0x0122: return "MOVE.TOOL";
  case	0x0123: return "WORKBOOK.SELECT";
  case	0x0124: return "WORKBOOK.ACTIVATE";
  case	0x0125: return "ASSIGN.TO.TOOL";
  case	0x0127: return "COPY.TOOL";
  case	0x0128: return "RESET.TOOL";
  case	0x0129: return "CONSTRAIN.NUMERIC";
  case	0x012A: return "PASTE.TOOL";
  case	0x012E: return "WORKBOOK.NEW";
  case	0x0131: return "SCENARIO.CELLS";
  case	0x0132: return "SCENARIO.DELETE";
  case	0x0133: return "SCENARIO.ADD";
  case	0x0134: return "SCENARIO.EDIT";
  case	0x0135: return "SCENARIO.SHOW";
  case	0x0136: return "SCENARIO.SHOW.NEXT";
  case	0x0137: return "SCENARIO.SUMMARY";
  case	0x0138: return "PIVOT.TABLE.WIZARD";
  case	0x0139: return "PIVOT.FIELD.PROPERTIES";
  case	0x013A: return "PIVOT.FIELD";
  case	0x013B: return "PIVOT.ITEM";
  case	0x013C: return "PIVOT.ADD.FIELDS";
  case	0x013E: return "OPTIONS.CALCULATION";
  case	0x013F: return "OPTIONS.EDIT";
  case	0x0140: return "OPTIONS.VIEW";
  case	0x0141: return "ADDIN.MANAGER";
  case	0x0142: return "MENU.EDITOR";
  case	0x0143: return "ATTACH.TOOLBARS";
  case	0x0144: return "VBAActivate";
  case	0x0145: return "OPTIONS.CHART";
  case	0x0148: return "VBA.INSERT.FILE";
  case	0x014A: return "VBA.PROCEDURE.DEFINITION";
  case	0x0150: return "ROUTING.SLIP";
  case	0x0152: return "ROUTE.DOCUMENT";
  case	0x0153: return "MAIL.LOGON";
  case	0x0156: return "INSERT.PICTURE";
  case	0x0157: return "EDIT.TOOL";
  case	0x0158: return "GALLERY.DOUGHNUT";
  case	0x015E: return "CHART.TREND";
  case	0x0160: return "PIVOT.ITEM.PROPERTIES";
  case	0x0162: return "WORKBOOK.INSERT";
  case	0x0163: return "OPTIONS.TRANSITION";
  case	0x0164: return "OPTIONS.GENERAL";
  case	0x0172: return "FILTER.ADVANCED";
  case	0x0175: return "MAIL.ADD.MAILER";
  case	0x0176: return "MAIL.DELETE.MAILER";
  case	0x0177: return "MAIL.REPLY";
  case	0x0178: return "MAIL.REPLY.ALL";
  case	0x0179: return "MAIL.FORWARD";
  case	0x017A: return "MAIL.NEXT.LETTER";
  case	0x017B: return "DATA.LABEL";
  case	0x017C: return "INSERT.TITLE";
  case	0x017D: return "FONT.PROPERTIES";
  case	0x017E: return "MACRO.OPTIONS";
  case	0x017F: return "WORKBOOK.HIDE";
  case	0x0180: return "WORKBOOK.UNHIDE";
  case	0x0181: return "WORKBOOK.DELETE";
  case	0x0182: return "WORKBOOK.NAME";
  case	0x0184: return "GALLERY.CUSTOM";
  case	0x0186: return "ADD.CHART.AUTOFORMAT";
  case	0x0187: return "DELETE.CHART.AUTOFORMAT";
  case	0x0188: return "CHART.ADD.DATA";
  case	0x0189: return "AUTO.OUTLINE";
  case	0x018A: return "TAB.ORDER";
  case	0x018B: return "SHOW.DIALOG";
  case	0x018C: return "SELECT.ALL";
  case	0x018D: return "UNGROUP.SHEETS";
  case	0x018E: return "SUBTOTAL.CREATE";
  case	0x018F: return "SUBTOTAL.REMOVE";
  case	0x0190: return "RENAME.OBJECT";
  case	0x019C: return "WORKBOOK.SCROLL";
  case	0x019D: return "WORKBOOK.NEXT";
  case	0x019E: return "WORKBOOK.PREV";
  case	0x019F: return "WORKBOOK.TAB.SPLIT";
  case	0x01A0: return "FULL.SCREEN";
  case	0x01A1: return "WORKBOOK.PROTECT";
  case	0x01A4: return "SCROLLBAR.PROPERTIES";
  case	0x01A5: return "PIVOT.SHOW.PAGES";
  case	0x01A6: return "TEXT.TO.COLUMNS";
  case	0x01A7: return "FORMAT.CHARTTYPE";
  case	0x01A8: return "LINK.FORMAT";
  case	0x01A9: return "TRACER.DISPLAY";
  case	0x01AE: return "TRACER.NAVIGATE";
  case	0x01AF: return "TRACER.CLEAR";
  case	0x01B0: return "TRACER.ERROR";
  case	0x01B1: return "PIVOT.FIELD.GROUP";
  case	0x01B2: return "PIVOT.FIELD.UNGROUP";
  case	0x01B3: return "CHECKBOX.PROPERTIES";
  case	0x01B4: return "LABEL.PROPERTIES";
  case	0x01B5: return "LISTBOX.PROPERTIES";
  case	0x01B6: return "EDITBOX.PROPERTIES";
  case	0x01B7: return "PIVOT.REFRESH";
  case	0x01B8: return "LINK.COMBO";
  case	0x01B9: return "OPEN.TEXT";
  case	0x01BA: return "HIDE.DIALOG";
  case	0x01BB: return "SET.DIALOG.FOCUS";
  case	0x01BC: return "ENABLE.OBJECT";
  case	0x01BD: return "PUSHBUTTON.PROPERTIES";
  case	0x01BE: return "SET.DIALOG.DEFAULT";
  case	0x01BF: return "FILTER";
  case	0x01C0: return "FILTER.SHOW.ALL";
  case	0x01C1: return "CLEAR.OUTLINE";
  case	0x01C2: return "FUNCTION.WIZARD";
  case	0x01C3: return "ADD.LIST.ITEM";
  case	0x01C4: return "SET.LIST.ITEM";
  case	0x01C5: return "REMOVE.LIST.ITEM";
  case	0x01C6: return "SELECT.LIST.ITEM";
  case	0x01C7: return "SET.CONTROL.VALUE";
  case	0x01C8: return "SAVE.COPY.AS";
  case	0x01CA: return "OPTIONS.LISTS.ADD";
  case	0x01CB: return "OPTIONS.LISTS.DELETE";
  case	0x01CC: return "SERIES.AXES";
  case	0x01CD: return "SERIES.X";
  case	0x01CE: return "SERIES.Y";
  case	0x01CF: return "ERRORBAR.X";
  case	0x01D0: return "ERRORBAR.Y";
  case	0x01D1: return "FORMAT.CHART";
  case	0x01D2: return "SERIES.ORDER";
  case	0x01D3: return "MAIL.LOGOFF";
  case	0x01D4: return "CLEAR.ROUTING.SLIP";
  case	0x01D5: return "APP.ACTIVATE.MICROSOFT";
  case	0x01D6: return "MAIL.EDIT.MAILER";
  case	0x01D7: return "ON.SHEET";
  case	0x01D8: return "STANDARD.WIDTH";
  case	0x01D9: return "SCENARIO.MERGE";
  case	0x01DA: return "SUMMARY.INFO";
  case	0x01DB: return "FIND.FILE";
  case	0x01DC: return "ACTIVE.CELL.FONT";
  case	0x01DD: return "ENABLE.TIPWIZARD";
  case	0x01DE: return "VBA.MAKE.ADDIN";
  case	0x01E0: return "INSERTDATATABLE";
  case	0x01E1: return "WORKGROUP.OPTIONS";
  case	0x01E2: return "MAIL.SEND.MAILER";
  case	0x01E5: return "AUTOCORRECT";
  case	0x01E9: return "POST.DOCUMENT";
  case	0x01EB: return "PICKLIST";
  case	0x01ED: return "VIEW.SHOW";
  case	0x01EE: return "VIEW.DEFINE";
  case	0x01EF: return "VIEW.DELETE";
  case	0x01FD: return "SHEET.BACKGROUND";
  case	0x01FE: return "INSERT.MAP.OBJECT";
  case	0x01FF: return "OPTIONS.MENONO";
  case	0x0205: return "MSOCHECKS";
  case	0x0206: return "NORMAL";
  case	0x0207: return "LAYOUT";
  case	0x0208: return "RM.PRINT.AREA";
  case	0x0209: return "CLEAR.PRINT.AREA";
  case	0x020A: return "ADD.PRINT.AREA";
  case	0x020B: return "MOVE.BRK";
  case	0x0221: return "HIDECURR.NOTE";
  case	0x0222: return "HIDEALL.NOTES";
  case	0x0223: return "DELETE.NOTE";
  case	0x0224: return "TRAVERSE.NOTES";
  case	0x0225: return "ACTIVATE.NOTES";
  case	0x026C: return "PROTECT.REVISIONS";
  case	0x026D: return "UNPROTECT.REVISIONS";
  case	0x0287: return "OPTIONS.ME";
  case	0x028D: return "WEB.PUBLISH";
  case	0x029B: return "NEWWEBQUERY";
  case	0x02A1: return "PIVOT.TABLE.CHART";
  case	0x02F1: return "OPTIONS.SAVE";
  case	0x02F3: return "OPTIONS.SPELL";
  case	0x0328: return "HIDEALL.INKANNOTS";
  }
  return "";
}

// some functions provide a fixed number of arguments
std::string Ftab(const uint16_t val) {
  switch(val) {
  case	0x0000:	return "COUNT";
  case	0x0001:	return "IF";
  case	0x0002:	return "ISNA(%s)";
  case	0x0003:	return "ISERROR(%s)";
  case	0x0004:	return "SUM";
  case	0x0005:	return "AVERAGE";
  case	0x0006:	return "MIN";
  case	0x0007:	return "MAX";
  case	0x0008:	return "ROW"; // (%s)
  case	0x0009:	return "COLUMN"; // (%s)
  case	0x000A:	return "NA()";
  case	0x000B:	return "NPV";
  case	0x000C:	return "STDEV";
  case	0x000D:	return "DOLLAR"; // (%s,%s)
  case	0x000E:	return "FIXED"; // (%s,%s,%s)
  case	0x000F:	return "SIN(%s)";
  case	0x0010:	return "COS(%s)";
  case	0x0011:	return "TAN(%s)";
  case	0x0012:	return "ATAN(%s)";
  case	0x0013:	return "PI()";
  case	0x0014:	return "SQRT(%s)";
  case	0x0015:	return "EXP(%s)";
  case	0x0016:	return "LN(%s)";
  case	0x0017:	return "LOG10(%s)";
  case	0x0018:	return "ABS(%s)";
  case	0x0019:	return "INT(%s)";
  case	0x001A:	return "SIGN(%s)";
  case	0x001B:	return "ROUND(%s,%s)";
  case	0x001C:	return "LOOKUP"; // (%s,%s,%s,%s,%s)
  case	0x001D:	return "INDEX";
  case	0x001E:	return "REPT(%s,%s)";
  case	0x001F:	return "MID(%s,%s,%s)";
  case	0x0020:	return "LEN(%s)";
  case	0x0021:	return "VALUE(%s)";
  case	0x0022:	return "TRUE()";
  case	0x0023:	return "FALSE()";
  case	0x0024:	return "AND";
  case	0x0025:	return "OR";
  case	0x0026:	return "NOT(%s)";
  case	0x0027:	return "MOD(%s,%s)";
  case	0x0028:	return "DCOUNT";
  case	0x0029:	return "DSUM";
  case	0x002A:	return "DAVERAGE";
  case	0x002B:	return "DMIN";
  case	0x002C:	return "DMAX";
  case	0x002D:	return "DSTDEV";
  case	0x002E:	return "VAR";
  case	0x002F:	return "DVAR";
  case	0x0030:	return "TEXT(%s,%s)";
  case	0x0031:	return "LINEST";
  case	0x0032:	return "TREND";
  case	0x0033:	return "LOGEST";
  case	0x0034:	return "GROWTH";
  case	0x0035:	return "GOTO(%s)";
  case	0x0036:	return "HALT"; // (%s)
  case	0x0037:	return "RETURN";
  case	0x0038:	return "PV"; // mb fixed, val,val,val * 2val ???
  case	0x0039:	return "FV";
  case	0x003A:	return "NPER";
  case	0x003B:	return "PMT";
  case	0x003C:	return "RATE";
  case	0x003D:	return "MIRR";
  case	0x003E:	return "IRR";
  case	0x003F:	return "RAND()";
  case	0x0040:	return "MATCH";
  case	0x0041:	return "DATE(%s,%s,%s)";
  case	0x0042:	return "TIME(%s,%s,%s)";
  case	0x0043:	return "DAY(%s)";
  case	0x0044:	return "MONTH(%s)";
  case	0x0045:	return "YEAR(%s)";
  case	0x0046:	return "WEEKDAY";
  case	0x0047:	return "HOUR(%s)";
  case	0x0048:	return "MINUTE(%s)";
  case	0x0049:	return "SECOND(%s)";
  case	0x004A:	return "NOW()";
  case	0x004B:	return "AREAS";
  case	0x004C:	return "ROWS(%s)"; // array
  case	0x004D:	return "COLUMNS(%s)"; // array
  case	0x004E:	return "OFFSET";
  case	0x004F:	return "ABSREF(%s,%s)";
  case	0x0050:	return "RELREF(%s,%s)";
  case	0x0051:	return "ARGUMENT";
  case	0x0052:	return "SEARCH";
  case	0x0053:	return "TRANSPOSE(%s)";
  case	0x0054:	return "ERROR";
  case	0x0055:	return "STEP()";
  case	0x0056:	return "TYPE(%s)";
  case	0x0057:	return "ECHO";
  case	0x0058:	return "SET.NAME";
  case	0x0059:	return "CALLER()";
  case	0x005A:	return "DEREF(%s)";
  case	0x005B:	return "WINDOWS";
  case	0x005C:	return "SERIES"; /* no longer in 2.5.98.10 Ftab */
  case	0x005D:	return "DOCUMENTS";
  case	0x005E:	return "ACTIVE.CELL()";
  case	0x005F:	return "SELECTION()";
  case	0x0060:	return "RESULT";
  case	0x0061:	return "ATAN2(%s,%s)";
  case	0x0062:	return "ASIN(%s)";
  case	0x0063:	return "ACOS(%s)";
  case	0x0064:	return "CHOOSE";
  case	0x0065:	return "HLOOKUP";
  case	0x0066:	return "VLOOKUP";
  case	0x0067:	return "LINKS";
  case	0x0068:	return "INPUT";
  case	0x0069:	return "ISREF(%s)";
  case	0x006A:	return "GET.FORMULA";
  case	0x006B:	return "GET.NAME";
  case	0x006C:	return "SET.VALUE(%s,%s)";
  case	0x006D:	return "LOG";
  case	0x006E:	return "EXEC";
  case	0x006F:	return "CHAR(%s)";
  case	0x0070:	return "LOWER(%s)";
  case	0x0071:	return "UPPER(%s)";
  case	0x0072:	return "PROPER(%s)";
  case	0x0073:	return "LEFT";
  case	0x0074:	return "RIGHT";
  case	0x0075:	return "EXACT";
  case	0x0076:	return "TRIM(%s)";
  case	0x0077:	return "REPLACE(%s,%s,%s,%s)";
  case	0x0078:	return "SUBSTITUTE";
  case	0x0079:	return "CODE(%s)";
  case	0x007A:	return "NAMES";
  case	0x007B:	return "DIRECTORY";
  case	0x007C:	return "FIND";
  case	0x007D:	return "CELL";
  case	0x007E:	return "ISERR(%s)";
  case	0x007F:	return "ISTEXT(%s)";
  case	0x0080:	return "ISNUMBER(%s)";
  case	0x0081:	return "ISBLANK(%s)";
  case	0x0082:	return "T(%s)";
  case	0x0083:	return "N(%s)";
  case	0x0084:	return "FOPEN";
  case	0x0085:	return "FCLOSE(%s)";
  case	0x0086:	return "FSIZE(%s)";
  case	0x0087:	return "FREADLN(%s)";
  case	0x0088:	return "FREAD(%s,%s)";
  case	0x0089:	return "FWRITELN(%s,%s)";
  case	0x008A:	return "FWRITE(%s,%s)";
  case	0x008B:	return "FPOS";
  case	0x008C:	return "DATEVALUE(%s)";
  case	0x008D:	return "TIMEVALUE(%s)";
  case	0x008E:	return "SLN(%s,%s,%s)";
  case	0x008F:	return "SYD(%s,%s,%s,%s)";
  case	0x0090:	return "DDB";
  case	0x0091:	return "GET.DEF";
  case	0x0092:	return "REFTEXT";
  case	0x0093:	return "TEXTREF";
  case	0x0094:	return "INDIRECT";
  case	0x0095:	return "REGISTER";
  case	0x0096:	return "CALL";
  case	0x0097:	return "ADD.BAR";
  case	0x0098:	return "ADD.MENU";
  case	0x0099:	return "ADD.COMMAND";
  case	0x009A:	return "ENABLE.COMMAND";
  case	0x009B:	return "CHECK.COMMAND";
  case	0x009C:	return "RENAME.COMMAND";
  case	0x009D:	return "SHOW.BAR";
  case	0x009E:	return "DELETE.MENU";
  case	0x009F:	return "DELETE.COMMAND";
  case	0x00A0:	return "GET.CHART.ITEM";
  case	0x00A1:	return "DIALOG.BOX";
  case	0x00A2:	return "CLEAN(%s)";
  case	0x00A3:	return "MDETERM(%s)";
  case	0x00A4:	return "MINVERSE(%s)";
  case	0x00A5:	return "MMULT(%s,%s)";
  case	0x00A6:	return "FILES";
  case	0x00A7:	return "IPMT";
  case	0x00A8:	return "PPMT";
  case	0x00A9:	return "COUNTA";
  case	0x00AA:	return "CANCEL.KEY";
  case	0x00AB:	return "FOR";
  case	0x00AC:	return "WHILE(%s)";
  case	0x00AD:	return "BREAK()";
  case	0x00AE:	return "NEXT()";
  case	0x00AF:	return "INITIATE(%s,%s)";
  case	0x00B0:	return "REQUEST(%s,%s)";
  case	0x00B1:	return "POKE(%s,%s,%s)";
  case	0x00B2:	return "EXECUTE(%s,%s)";
  case	0x00B3:	return "TERMINATE(%s)";
  case	0x00B4:	return "RESTART";
  case	0x00B5:	return "HELP";
  case	0x00B6:	return "GET.BAR";
  case	0x00B7:	return "PRODUCT";
  case	0x00B8:	return "FACT(%s)";
  case	0x00B9:	return "GET.CELL";
  case	0x00BA:	return "GET.WORKSPACE(%s)";
  case	0x00BB:	return "GET.WINDOW";
  case	0x00BC:	return "GET.DOCUMENT";
  case	0x00BD:	return "DPRODUCT";
  case	0x00BE:	return "ISNONTEXT(%s)";
  case	0x00BF:	return "GET.NOTE";
  case	0x00C0:	return "NOTE";
  case	0x00C1:	return "STDEVP";
  case	0x00C2:	return "VARP";
  case	0x00C3:	return "DSTDEVP(%s,%s,%s)";
  case	0x00C4:	return "DVARP(%s,%s,%s)";
  case	0x00C5:	return "TRUNC";
  case	0x00C6:	return "ISLOGICAL(%s)";
  case	0x00C7:	return "DCOUNTA(%s,%s,%s)";
  case	0x00C8:	return "DELETE.BAR(%s)";
  case	0x00C9:	return "UNREGISTER(%s)";
  case	0x00CC:	return "USDOLLAR";
  case	0x00CD:	return "FINDB";
  case	0x00CE:	return "SEARCHB";
  case	0x00CF:	return "REPLACEB(%s,%s,%s,%s)";
  case	0x00D0:	return "LEFTB";
  case	0x00D1:	return "RIGHTB";
  case	0x00D2:	return "MIDB(%s,%s,%s)";
  case	0x00D3:	return "LENB(%s,%s,%s)";
  case	0x00D4:	return "ROUNDUP(%s,%s)";
  case	0x00D5:	return "ROUNDDOWN(%s,%s)";
  case	0x00D6:	return "ASC(%s)";
  case	0x00D7:	return "DBCS(%s)";
  case	0x00D8:	return "RANK";
  case	0x00DB:	return "ADDRESS";
  case	0x00DC:	return "DAYS360";
  case	0x00DD:	return "TODAY()";
  case	0x00DE:	return "VDB";
  case	0x00DF:	return "ELSE()";
  case	0x00E0:	return "ELSE.IF(%s)";
  case	0x00E1:	return "END.IF()";
  case	0x00E2:	return "FOR.CELL";
  case	0x00E3:	return "MEDIAN";
  case	0x00E4:	return "SUMPRODUCT";
  case	0x00E5:	return "SINH(%s)";
  case	0x00E6:	return "COSH(%s)";
  case	0x00E7:	return "TANH(%s)";
  case	0x00E8:	return "ASINH(%s)";
  case	0x00E9:	return "ACOSH(%s)";
  case	0x00EA:	return "ATANH(%s)";
  case	0x00EB:	return "DGET(%s,%s,%s)";
  case	0x00EC:	return "CREATE.OBJECT";
  case	0x00ED:	return "VOLATILE";
  case	0x00EE:	return "LAST.ERROR()";
  case	0x00EF:	return "CUSTOM.UNDO";
  case	0x00F0:	return "CUSTOM.REPEAT";
  case	0x00F1:	return "FORMULA.CONVERT";
  case	0x00F2:	return "GET.LINK.INFO";
  case	0x00F3:	return "TEXT.BOX";
  case	0x00F4:	return "INFO(%s)";
  case	0x00F5:	return "GROUP()";
  case	0x00F6:	return "GET.OBJECT";
  case	0x00F7:	return "DB";
  case	0x00F8:	return "PAUSE";
  case	0x00FB:	return "RESUME";
  case	0x00FC:	return "FREQUENCY";
  case	0x00FD:	return "ADD.TOOLBAR";
  case	0x00FE:	return "DELETE.TOOLBAR(%s)";
  case	0x00FF:	return "User_Defined_Function";
  case	0x0100:	return "RESET.TOOLBAR(%s)";
  case	0x0101:	return "EVALUATE(%s)";
  case	0x0102:	return "GET.TOOLBAR";
  case	0x0103:	return "GET.TOOL";
  case	0x0104:	return "SPELLING.CHECK";
  case	0x0105:	return "ERROR.TYPE(%s)";
  case	0x0106:	return "APP.TITLE";
  case	0x0107:	return "WINDOW.TITLE";
  case	0x0108:	return "SAVE.TOOLBAR";
  case	0x0109:	return "ENABLE.TOOL(%s,%s,%s)";
  case	0x010A:	return "PRESS.TOOL(%s,%s,%s)";
  case	0x010B:	return "REGISTER.ID";
  case	0x010C:	return "GET.WORKBOOK";
  case	0x010D:	return "AVEDEV";
  case	0x010E:	return "BETADIST";
  case	0x010F:	return "GAMMALN(%s)";
  case	0x0110:	return "BETAINV";
  case	0x0111:	return "BINOMDIST(%s,%s,%s,%s)";
  case	0x0112:	return "CHIDIST(%s,%s)";
  case	0x0113:	return "CHIINV(%s,%s)";
  case	0x0114:	return "COMBIN(%s,%s)";
  case	0x0115:	return "CONFIDENCE(%s,%s,%s)";
  case	0x0116:	return "CRITBINOM(%s,%s,%s)";
  case	0x0117:	return "EVEN(%s)";
  case	0x0118:	return "EXPONDIST(%s,%s,%s)";
  case	0x0119:	return "FDIST(%s,%s,%s)";
  case	0x011A:	return "FINV(%s,%s,%s)";
  case	0x011B:	return "FISHER(%s)";
  case	0x011C:	return "FISHERINV(%s)";
  case	0x011D:	return "FLOOR(%s,%s)";
  case	0x011E:	return "GAMMADIST(%s,%s,%s,%s)";
  case	0x011F:	return "GAMMAINV(%s,%s,%s)";
  case	0x0120:	return "CEILING(%s,%s)";
  case	0x0121:	return "HYPGEOMDIST(%s,%s,%s,%s)";
  case	0x0122:	return "LOGNORMDIST(%s,%s,%s)";
  case	0x0123:	return "LOGINV(%s,%s,%s)";
  case	0x0124:	return "NEGBINOMDIST(%s,%s,%s)";
  case	0x0125:	return "NORMDIST(%s,%s,%s,%s)";
  case	0x0126:	return "NORMSDIST(%s)";
  case	0x0127:	return "NORMINV(%s,%s,%s)";
  case	0x0128:	return "NORMSINV(%s)";
  case	0x0129:	return "STANDARDIZE(%s,%s,%s)";
  case	0x012A:	return "ODD(%s)";
  case	0x012B:	return "PERMUT(%s,%s)";
  case	0x012C:	return "POISSON(%s,%s,%s)";
  case	0x012D:	return "TDIST(%s,%s,%s)";
  case	0x012E:	return "WEIBULL(%s,%s,%s,%s)";
  case	0x012F:	return "SUMXMY2(%s,%s)";
  case	0x0130:	return "SUMX2MY2(%s,%s)";
  case	0x0131:	return "SUMX2PY2(%s,%s)";
  case	0x0132:	return "CHITEST(%s,%s)";
  case	0x0133:	return "CORREL(%s,%s)";
  case	0x0134:	return "COVAR(%s,%s)";
  case	0x0135:	return "FORECAST(%s,%s,%s)";
  case	0x0136:	return "FTEST(%s,%s)";
  case	0x0137:	return "INTERCEPT(%s,%s)";
  case	0x0138:	return "PEARSON(%s,%s)";
  case	0x0139:	return "RSQ(%s,%s)";
  case	0x013A:	return "STEYX(%s,%s)";
  case	0x013B:	return "SLOPE(%s,%s)";
  case	0x013C:	return "TTEST(%s,%s,%s,%s)";
  case	0x013D:	return "PROB";
  case	0x013E:	return "DEVSQ";
  case	0x013F:	return "GEOMEAN";
  case	0x0140:	return "HARMEAN";
  case	0x0141:	return "SUMSQ";
  case	0x0142:	return "KURT";
  case	0x0143:	return "SKEW";
  case	0x0144:	return "ZTEST";
  case	0x0145:	return "LARGE(%s,%s)";
  case	0x0146:	return "SMALL(%s,%s)";
  case	0x0147:	return "QUARTILE(%s,%s)";
  case	0x0148:	return "PERCENTILE(%s,%s)";
  case	0x0149:	return "PERCENTRANK";
  case	0x014A:	return "MODE";
  case	0x014B:	return "TRIMMEAN(%s,%s)";
  case	0x014C:	return "TINV(%s,%s)";
  case	0x014E:	return "MOVIE.COMMAND";
  case	0x014F:	return "GET.MOVIE";
  case	0x0150:	return "CONCATENATE";
  case	0x0151:	return "POWER(%s,%s)";
  case	0x0152:	return "PIVOT.ADD.DATA";
  case	0x0153:	return "GET.PIVOT.TABLE";
  case	0x0154:	return "GET.PIVOT.FIELD";
  case	0x0155:	return "GET.PIVOT.ITEM";
  case	0x0156:	return "RADIANS(%s)";
  case	0x0157:	return "DEGREES(%s)";
  case	0x0158:	return "SUBTOTAL";
  case	0x0159:	return "SUMIF";
  case	0x015A:	return "COUNTIF(%s,%s)";
  case	0x015B:	return "COUNTBLANK(%s)";
  case	0x015C:	return "SCENARIO.GET";
  case	0x015D:	return "OPTIONS.LISTS.GET(%s)";
  case	0x015E:	return "ISPMT(%s,%s,%s,%s)";
  case	0x015F:	return "DATEDIF(%s,%s,%s)";
  case	0x0160:	return "DATESTRING(%s)";
  case	0x0161:	return "NUMBERSTRING(%s,%s)";
  case	0x0162:	return "ROMAN";
  case	0x0163:	return "OPEN.DIALOG";
  case	0x0164:	return "SAVE.DIALOG";
  case	0x0165:	return "VIEW.GET";
  case	0x0166:	return "GETPIVOTDATA";
  case	0x0167:	return "HYPERLINK";
  case	0x0168:	return "PHONETIC(%s)";
  case	0x0169:	return "AVERAGEA";
  case	0x016A:	return "MAXA";
  case	0x016B:	return "MINA";
  case	0x016C:	return "STDEVPA";
  case	0x016D:	return "VARPA";
  case	0x016E:	return "STDEVA";
  case	0x016F:	return "VARA";
  case	0x0170:	return "BAHTTEXT(%s)";
  case	0x0171:	return "THAIDAYOFWEEK(%s)";
  case	0x0172:	return "THAIDIGIT(%s)";
  case	0x0173:	return "THAIMONTHOFYEAR(%s)";
  case	0x0174:	return "THAINUMSOUND(%s)";
  case	0x0175:	return "THAINUMSTRING(%s)";
  case	0x0176:	return "THAISTRINGLENGTH(%s)";
  case	0x0177:	return "ISTHAIDIGIT(%s)";
  case	0x0178:	return "ROUNDBAHTDOWN(%s)";
  case	0x0179:	return "ROUNDBAHTUP(%s)";
  case	0x017A:	return "THAIYEAR(%s)";
  case	0x017B:	return "RTD";
  case	0x017C: return "CUBEVALUE";
  case	0x017D: return "CUBEMEMBER";
  case	0x017E: return "CUBEMEMBERPROPERTY(%s,%s,%s)";
  case	0x017F: return "CUBERANKEDMEMBER";
  case	0x0180: return "HEX2BIN";
  case	0x0181: return "HEX2DEC(%s)";
  case	0x0182: return "HEX2OCT";
  case	0x0183: return "DEC2BIN";
  case	0x0184: return "DEC2HEX";
  case	0x0185: return "DEC2OCT";
  case	0x0186: return "OCT2BIN";
  case	0x0187: return "OCT2HEX";
  case	0x0188: return "OCT2DEC(%s)";
  case	0x0189: return "BIN2DEC(%s)";
  case	0x018A: return "BIN2OCT";
  case	0x018B: return "BIN2HEX";
  case	0x018C: return "IMSUB(%s,%s)";
  case	0x018D: return "IMDIV(%s,%s)";
  case	0x018E: return "IMPOWER(%s,%s)";
  case	0x018F: return "IMABS(%s)";
  case	0x0190: return "IMSQRT(%s)";
  case	0x0191: return "IMLN(%s)";
  case	0x0192: return "IMLOG2(%s)";
  case	0x0193: return "IMLOG10(%s)";
  case	0x0194: return "IMSIN(%s)";
  case	0x0195: return "IMCOS(%s)";
  case	0x0196: return "IMEXP(%s)";
  case	0x0197: return "IMARGUMENT(%s)";
  case	0x0198: return "IMCONJUGATE(%s)";
  case	0x0199: return "IMAGINARY(%s)";
  case	0x019A: return "IMREAL(%s)";
  case	0x019B: return "COMPLEX";
  case	0x019C: return "IMSUM";
  case	0x019D: return "IMPRODUCT";
  case	0x019E: return "SERIESSUM(%s,%s,%s,%s)";
  case	0x019F: return "FACTDOUBLE(%s)";
  case	0x01A0: return "SQRTPI(%s)";
  case	0x01A1: return "QUOTIENT(%s,%s)";
  case	0x01A2: return "DELTA";
  case	0x01A3: return "GESTEP";
  case	0x01A4: return "ISEVEN(%s)";
  case	0x01A5: return "ISODD(%s)";
  case	0x01A6: return "MROUND(%s,%s)";
  case	0x01A7: return "ERF";
  case	0x01A8: return "ERFC(%s)";
  case	0x01A9: return "BESSELJ(%s,%s)";
  case	0x01AA: return "BESSELK(%s,%s)";
  case	0x01AB: return "BESSELY(%s,%s)";
  case	0x01AC: return "BESSELI(%s,%s)";
  case	0x01AD: return "XIRR(%s,%s,%s)";
  case	0x01AE: return "XNPV";
  case	0x01AF: return "PRICEMAT";
  case	0x01B0: return "YIELDMAT";
  case	0x01B1: return "INTRATE";
  case	0x01B2: return "RECEIVED";
  case	0x01B3: return "DISC";
  case	0x01B4: return "PRICEDISC";
  case	0x01B5: return "YIELDDISC";
  case	0x01B6: return "TBILLEQ(%s,%s,%s)";
  case	0x01B7: return "TBILLPRICE(%s,%s,%s)";
  case	0x01B8: return "TBILLYIELD(%s,%s,%s)";
  case	0x01B9: return "PRICE";
  case	0x01BA: return "YIELD";
  case	0x01BB: return "DOLLARDE(%s,%s)";
  case	0x01BC: return "DOLLARFR(%s,%s)";
  case	0x01BD: return "NOMINAL(%s,%s)";
  case	0x01BE: return "EFFECT(%s,%s)";
  case	0x01BF: return "CUMPRINC(%s,%s,%s,%s,%s,%s)";
  case	0x01C0: return "CUMIPMT(%s,%s,%s,%s,%s,%s)";
  case	0x01C1: return "EDATE(%s,%s)";
  case	0x01C2: return "EOMONTH(%s,%s)";
  case	0x01C3: return "YEARFRAC";
  case	0x01C4: return "COUPDAYBS";
  case	0x01C5: return "COUPDAYS";
  case	0x01C6: return "COUPDAYSNC";
  case	0x01C7: return "COUPNCD";
  case	0x01C8: return "COUPNUM";
  case	0x01C9: return "COUPPCD";
  case	0x01CA: return "DURATION";
  case	0x01CB: return "MDURATION";
  case	0x01CC: return "ODDLPRICE";
  case	0x01CD: return "ODDLYIELD";
  case	0x01CE: return "ODDFPRICE";
  case	0x01CF: return "ODDFYIELD";
  case	0x01D0: return "RANDBETWEEN(%s,%s)";
  case	0x01D1: return "WEEKNUM";
  case	0x01D2: return "AMORDEGRC";
  case	0x01D3: return "AMORLINC";
  case	0x01D5: return "ACCRINT";
  case	0x01D6: return "ACCRINTM";
  case	0x01D7: return "WORKDAY";
  case	0x01D8: return "NETWORKDAYS";
  case	0x01D9: return "GCD";
  case	0x01DA: return "MULTINOMIAL";
  case	0x01DB: return "LCM";
  case	0x01DC: return "FVSCHEDULE(%s,%s)";
  case	0x01DD: return "CUBEKPIMEMBER";
  case	0x01DE: return "CUBESET";
  case	0x01DF: return "CUBESETCOUNT(%s)";
  case	0x01E0: return "IFERROR(%s,%s)";
  case	0x01E1: return "COUNTIFS";
  case	0x01E2: return "SUMIFS";
  case	0x01E3: return "AVERAGEIF";
  case	0x01E4: return "AVERAGEIFS";
  default: return "Unknow function: " + std::to_string(val);
  }
  return "";
}
// #nocov end

#endif
