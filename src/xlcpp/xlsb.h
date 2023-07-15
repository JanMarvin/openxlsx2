#pragma once

enum class xlsb_type {
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
    brtBeginModelTimeGroupings = 2137,
    brtEndModelTimeGroupings = 2138,
    brtBeginModelTimeGrouping = 2139,
    brtEndModelTimeGrouping = 2140,
    brtModelTimeGroupingCalcCol = 2141,
    brtRevisionPtr = 3073,
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
    BrtExternalLinksPr = 5099
};

template<>
struct fmt::formatter<enum xlsb_type> {
    constexpr auto parse(format_parse_context& ctx) {
        auto it = ctx.begin();

        if (it != ctx.end() && *it != '}')
            throw format_error("invalid format");

        return it;
    }

    template<typename format_context>
    auto format(enum xlsb_type t, format_context& ctx) const {
        switch (t) {
            case xlsb_type::BrtRowHdr:
                return fmt::format_to(ctx.out(), "BrtRowHdr");
            case xlsb_type::BrtCellBlank:
                return fmt::format_to(ctx.out(), "BrtCellBlank");
            case xlsb_type::BrtCellRk:
                return fmt::format_to(ctx.out(), "BrtCellRk");
            case xlsb_type::BrtCellError:
                return fmt::format_to(ctx.out(), "BrtCellError");
            case xlsb_type::BrtCellBool:
                return fmt::format_to(ctx.out(), "BrtCellBool");
            case xlsb_type::BrtCellReal:
                return fmt::format_to(ctx.out(), "BrtCellReal");
            case xlsb_type::BrtCellSt:
                return fmt::format_to(ctx.out(), "BrtCellSt");
            case xlsb_type::BrtCellIsst:
                return fmt::format_to(ctx.out(), "BrtCellIsst");
            case xlsb_type::BrtFmlaString:
                return fmt::format_to(ctx.out(), "BrtFmlaString");
            case xlsb_type::BrtFmlaNum:
                return fmt::format_to(ctx.out(), "BrtFmlaNum");
            case xlsb_type::BrtFmlaBool:
                return fmt::format_to(ctx.out(), "BrtFmlaBool");
            case xlsb_type::BrtFmlaError:
                return fmt::format_to(ctx.out(), "BrtFmlaError");
            case xlsb_type::BrtSSTItem:
                return fmt::format_to(ctx.out(), "BrtSSTItem");
            case xlsb_type::BrtPCDIMissing:
                return fmt::format_to(ctx.out(), "BrtPCDIMissing");
            case xlsb_type::BrtPCDINumber:
                return fmt::format_to(ctx.out(), "BrtPCDINumber");
            case xlsb_type::BrtPCDIBoolean:
                return fmt::format_to(ctx.out(), "BrtPCDIBoolean");
            case xlsb_type::BrtPCDIError:
                return fmt::format_to(ctx.out(), "BrtPCDIError");
            case xlsb_type::BrtPCDIString:
                return fmt::format_to(ctx.out(), "BrtPCDIString");
            case xlsb_type::BrtPCDIDatetime:
                return fmt::format_to(ctx.out(), "BrtPCDIDatetime");
            case xlsb_type::BrtPCDIIndex:
                return fmt::format_to(ctx.out(), "BrtPCDIIndex");
            case xlsb_type::BrtPCDIAMissing:
                return fmt::format_to(ctx.out(), "BrtPCDIAMissing");
            case xlsb_type::BrtPCDIANumber:
                return fmt::format_to(ctx.out(), "BrtPCDIANumber");
            case xlsb_type::BrtPCDIABoolean:
                return fmt::format_to(ctx.out(), "BrtPCDIABoolean");
            case xlsb_type::BrtPCDIAError:
                return fmt::format_to(ctx.out(), "BrtPCDIAError");
            case xlsb_type::BrtPCDIAString:
                return fmt::format_to(ctx.out(), "BrtPCDIAString");
            case xlsb_type::BrtPCDIADatetime:
                return fmt::format_to(ctx.out(), "BrtPCDIADatetime");
            case xlsb_type::BrtPCRRecord:
                return fmt::format_to(ctx.out(), "BrtPCRRecord");
            case xlsb_type::BrtPCRRecordDt:
                return fmt::format_to(ctx.out(), "BrtPCRRecordDt");
            case xlsb_type::BrtFRTBegin:
                return fmt::format_to(ctx.out(), "BrtFRTBegin");
            case xlsb_type::BrtFRTEnd:
                return fmt::format_to(ctx.out(), "BrtFRTEnd");
            case xlsb_type::BrtACBegin:
                return fmt::format_to(ctx.out(), "BrtACBegin");
            case xlsb_type::BrtACEnd:
                return fmt::format_to(ctx.out(), "BrtACEnd");
            case xlsb_type::BrtName:
                return fmt::format_to(ctx.out(), "BrtName");
            case xlsb_type::BrtIndexRowBlock:
                return fmt::format_to(ctx.out(), "BrtIndexRowBlock");
            case xlsb_type::BrtIndexBlock:
                return fmt::format_to(ctx.out(), "BrtIndexBlock");
            case xlsb_type::BrtFont:
                return fmt::format_to(ctx.out(), "BrtFont");
            case xlsb_type::BrtFmt:
                return fmt::format_to(ctx.out(), "BrtFmt");
            case xlsb_type::BrtFill:
                return fmt::format_to(ctx.out(), "BrtFill");
            case xlsb_type::BrtBorder:
                return fmt::format_to(ctx.out(), "BrtBorder");
            case xlsb_type::BrtXF:
                return fmt::format_to(ctx.out(), "BrtXF");
            case xlsb_type::BrtStyle:
                return fmt::format_to(ctx.out(), "BrtStyle");
            case xlsb_type::BrtCellMeta:
                return fmt::format_to(ctx.out(), "BrtCellMeta");
            case xlsb_type::BrtValueMeta:
                return fmt::format_to(ctx.out(), "BrtValueMeta");
            case xlsb_type::BrtMdb:
                return fmt::format_to(ctx.out(), "BrtMdb");
            case xlsb_type::BrtBeginFmd:
                return fmt::format_to(ctx.out(), "BrtBeginFmd");
            case xlsb_type::BrtEndFmd:
                return fmt::format_to(ctx.out(), "BrtEndFmd");
            case xlsb_type::BrtBeginMdx:
                return fmt::format_to(ctx.out(), "BrtBeginMdx");
            case xlsb_type::BrtEndMdx:
                return fmt::format_to(ctx.out(), "BrtEndMdx");
            case xlsb_type::BrtBeginMdxTuple:
                return fmt::format_to(ctx.out(), "BrtBeginMdxTuple");
            case xlsb_type::BrtEndMdxTuple:
                return fmt::format_to(ctx.out(), "BrtEndMdxTuple");
            case xlsb_type::BrtMdxMbrIstr:
                return fmt::format_to(ctx.out(), "BrtMdxMbrIstr");
            case xlsb_type::BrtStr:
                return fmt::format_to(ctx.out(), "BrtStr");
            case xlsb_type::BrtColInfo:
                return fmt::format_to(ctx.out(), "BrtColInfo");
            case xlsb_type::BrtCellRString:
                return fmt::format_to(ctx.out(), "BrtCellRString");
            case xlsb_type::BrtDVal:
                return fmt::format_to(ctx.out(), "BrtDVal");
            case xlsb_type::BrtSxvcellNum:
                return fmt::format_to(ctx.out(), "BrtSxvcellNum");
            case xlsb_type::BrtSxvcellStr:
                return fmt::format_to(ctx.out(), "BrtSxvcellStr");
            case xlsb_type::BrtSxvcellBool:
                return fmt::format_to(ctx.out(), "BrtSxvcellBool");
            case xlsb_type::BrtSxvcellErr:
                return fmt::format_to(ctx.out(), "BrtSxvcellErr");
            case xlsb_type::BrtSxvcellDate:
                return fmt::format_to(ctx.out(), "BrtSxvcellDate");
            case xlsb_type::BrtSxvcellNil:
                return fmt::format_to(ctx.out(), "BrtSxvcellNil");
            case xlsb_type::BrtFileVersion:
                return fmt::format_to(ctx.out(), "BrtFileVersion");
            case xlsb_type::BrtBeginSheet:
                return fmt::format_to(ctx.out(), "BrtBeginSheet");
            case xlsb_type::BrtEndSheet:
                return fmt::format_to(ctx.out(), "BrtEndSheet");
            case xlsb_type::BrtBeginBook:
                return fmt::format_to(ctx.out(), "BrtBeginBook");
            case xlsb_type::BrtEndBook:
                return fmt::format_to(ctx.out(), "BrtEndBook");
            case xlsb_type::BrtBeginWsViews:
                return fmt::format_to(ctx.out(), "BrtBeginWsViews");
            case xlsb_type::BrtEndWsViews:
                return fmt::format_to(ctx.out(), "BrtEndWsViews");
            case xlsb_type::BrtBeginBookViews:
                return fmt::format_to(ctx.out(), "BrtBeginBookViews");
            case xlsb_type::BrtEndBookViews:
                return fmt::format_to(ctx.out(), "BrtEndBookViews");
            case xlsb_type::BrtBeginWsView:
                return fmt::format_to(ctx.out(), "BrtBeginWsView");
            case xlsb_type::BrtEndWsView:
                return fmt::format_to(ctx.out(), "BrtEndWsView");
            case xlsb_type::BrtBeginCsViews:
                return fmt::format_to(ctx.out(), "BrtBeginCsViews");
            case xlsb_type::BrtEndCsViews:
                return fmt::format_to(ctx.out(), "BrtEndCsViews");
            case xlsb_type::BrtBeginCsView:
                return fmt::format_to(ctx.out(), "BrtBeginCsView");
            case xlsb_type::BrtEndCsView:
                return fmt::format_to(ctx.out(), "BrtEndCsView");
            case xlsb_type::BrtBeginBundleShs:
                return fmt::format_to(ctx.out(), "BrtBeginBundleShs");
            case xlsb_type::BrtEndBundleShs:
                return fmt::format_to(ctx.out(), "BrtEndBundleShs");
            case xlsb_type::BrtBeginSheetData:
                return fmt::format_to(ctx.out(), "BrtBeginSheetData");
            case xlsb_type::BrtEndSheetData:
                return fmt::format_to(ctx.out(), "BrtEndSheetData");
            case xlsb_type::BrtWsProp:
                return fmt::format_to(ctx.out(), "BrtWsProp");
            case xlsb_type::BrtWsDim:
                return fmt::format_to(ctx.out(), "BrtWsDim");
            case xlsb_type::BrtPane:
                return fmt::format_to(ctx.out(), "BrtPane");
            case xlsb_type::BrtSel:
                return fmt::format_to(ctx.out(), "BrtSel");
            case xlsb_type::BrtWbProp:
                return fmt::format_to(ctx.out(), "BrtWbProp");
            case xlsb_type::BrtWbFactoid:
                return fmt::format_to(ctx.out(), "BrtWbFactoid");
            case xlsb_type::BrtFileRecover:
                return fmt::format_to(ctx.out(), "BrtFileRecover");
            case xlsb_type::BrtBundleSh:
                return fmt::format_to(ctx.out(), "BrtBundleSh");
            case xlsb_type::BrtCalcProp:
                return fmt::format_to(ctx.out(), "BrtCalcProp");
            case xlsb_type::BrtBookView:
                return fmt::format_to(ctx.out(), "BrtBookView");
            case xlsb_type::BrtBeginSst:
                return fmt::format_to(ctx.out(), "BrtBeginSst");
            case xlsb_type::BrtEndSst:
                return fmt::format_to(ctx.out(), "BrtEndSst");
            case xlsb_type::BrtBeginAFilter:
                return fmt::format_to(ctx.out(), "BrtBeginAFilter");
            case xlsb_type::BrtEndAFilter:
                return fmt::format_to(ctx.out(), "BrtEndAFilter");
            case xlsb_type::BrtBeginFilterColumn:
                return fmt::format_to(ctx.out(), "BrtBeginFilterColumn");
            case xlsb_type::BrtEndFilterColumn:
                return fmt::format_to(ctx.out(), "BrtEndFilterColumn");
            case xlsb_type::BrtBeginFilters:
                return fmt::format_to(ctx.out(), "BrtBeginFilters");
            case xlsb_type::BrtEndFilters:
                return fmt::format_to(ctx.out(), "BrtEndFilters");
            case xlsb_type::BrtFilter:
                return fmt::format_to(ctx.out(), "BrtFilter");
            case xlsb_type::BrtColorFilter:
                return fmt::format_to(ctx.out(), "BrtColorFilter");
            case xlsb_type::BrtIconFilter:
                return fmt::format_to(ctx.out(), "BrtIconFilter");
            case xlsb_type::BrtTop10Filter:
                return fmt::format_to(ctx.out(), "BrtTop10Filter");
            case xlsb_type::BrtDynamicFilter:
                return fmt::format_to(ctx.out(), "BrtDynamicFilter");
            case xlsb_type::BrtBeginCustomFilters:
                return fmt::format_to(ctx.out(), "BrtBeginCustomFilters");
            case xlsb_type::BrtEndCustomFilters:
                return fmt::format_to(ctx.out(), "BrtEndCustomFilters");
            case xlsb_type::BrtCustomFilter:
                return fmt::format_to(ctx.out(), "BrtCustomFilter");
            case xlsb_type::BrtAFilterDateGroupItem:
                return fmt::format_to(ctx.out(), "BrtAFilterDateGroupItem");
            case xlsb_type::BrtMergeCell:
                return fmt::format_to(ctx.out(), "BrtMergeCell");
            case xlsb_type::BrtBeginMergeCells:
                return fmt::format_to(ctx.out(), "BrtBeginMergeCells");
            case xlsb_type::BrtEndMergeCells:
                return fmt::format_to(ctx.out(), "BrtEndMergeCells");
            case xlsb_type::BrtBeginPivotCacheDef:
                return fmt::format_to(ctx.out(), "BrtBeginPivotCacheDef");
            case xlsb_type::BrtEndPivotCacheDef:
                return fmt::format_to(ctx.out(), "BrtEndPivotCacheDef");
            case xlsb_type::BrtBeginPCDFields:
                return fmt::format_to(ctx.out(), "BrtBeginPCDFields");
            case xlsb_type::BrtEndPCDFields:
                return fmt::format_to(ctx.out(), "BrtEndPCDFields");
            case xlsb_type::BrtBeginPCDField:
                return fmt::format_to(ctx.out(), "BrtBeginPCDField");
            case xlsb_type::BrtEndPCDField:
                return fmt::format_to(ctx.out(), "BrtEndPCDField");
            case xlsb_type::BrtBeginPCDSource:
                return fmt::format_to(ctx.out(), "BrtBeginPCDSource");
            case xlsb_type::BrtEndPCDSource:
                return fmt::format_to(ctx.out(), "BrtEndPCDSource");
            case xlsb_type::BrtBeginPCDSRange:
                return fmt::format_to(ctx.out(), "BrtBeginPCDSRange");
            case xlsb_type::BrtEndPCDSRange:
                return fmt::format_to(ctx.out(), "BrtEndPCDSRange");
            case xlsb_type::BrtBeginPCDFAtbl:
                return fmt::format_to(ctx.out(), "BrtBeginPCDFAtbl");
            case xlsb_type::BrtEndPCDFAtbl:
                return fmt::format_to(ctx.out(), "BrtEndPCDFAtbl");
            case xlsb_type::BrtBeginPCDIRun:
                return fmt::format_to(ctx.out(), "BrtBeginPCDIRun");
            case xlsb_type::BrtEndPCDIRun:
                return fmt::format_to(ctx.out(), "BrtEndPCDIRun");
            case xlsb_type::BrtBeginPivotCacheRecords:
                return fmt::format_to(ctx.out(), "BrtBeginPivotCacheRecords");
            case xlsb_type::BrtEndPivotCacheRecords:
                return fmt::format_to(ctx.out(), "BrtEndPivotCacheRecords");
            case xlsb_type::BrtBeginPCDHierarchies:
                return fmt::format_to(ctx.out(), "BrtBeginPCDHierarchies");
            case xlsb_type::BrtEndPCDHierarchies:
                return fmt::format_to(ctx.out(), "BrtEndPCDHierarchies");
            case xlsb_type::BrtBeginPCDHierarchy:
                return fmt::format_to(ctx.out(), "BrtBeginPCDHierarchy");
            case xlsb_type::BrtEndPCDHierarchy:
                return fmt::format_to(ctx.out(), "BrtEndPCDHierarchy");
            case xlsb_type::BrtBeginPCDHFieldsUsage:
                return fmt::format_to(ctx.out(), "BrtBeginPCDHFieldsUsage");
            case xlsb_type::BrtEndPCDHFieldsUsage:
                return fmt::format_to(ctx.out(), "BrtEndPCDHFieldsUsage");
            case xlsb_type::BrtBeginExtConnection:
                return fmt::format_to(ctx.out(), "BrtBeginExtConnection");
            case xlsb_type::BrtEndExtConnection:
                return fmt::format_to(ctx.out(), "BrtEndExtConnection");
            case xlsb_type::BrtBeginECDbProps:
                return fmt::format_to(ctx.out(), "BrtBeginECDbProps");
            case xlsb_type::BrtEndECDbProps:
                return fmt::format_to(ctx.out(), "BrtEndECDbProps");
            case xlsb_type::BrtBeginECOlapProps:
                return fmt::format_to(ctx.out(), "BrtBeginECOlapProps");
            case xlsb_type::BrtEndECOlapProps:
                return fmt::format_to(ctx.out(), "BrtEndECOlapProps");
            case xlsb_type::BrtBeginPCDSConsol:
                return fmt::format_to(ctx.out(), "BrtBeginPCDSConsol");
            case xlsb_type::BrtEndPCDSConsol:
                return fmt::format_to(ctx.out(), "BrtEndPCDSConsol");
            case xlsb_type::BrtBeginPCDSCPages:
                return fmt::format_to(ctx.out(), "BrtBeginPCDSCPages");
            case xlsb_type::BrtEndPCDSCPages:
                return fmt::format_to(ctx.out(), "BrtEndPCDSCPages");
            case xlsb_type::BrtBeginPCDSCPage:
                return fmt::format_to(ctx.out(), "BrtBeginPCDSCPage");
            case xlsb_type::BrtEndPCDSCPage:
                return fmt::format_to(ctx.out(), "BrtEndPCDSCPage");
            case xlsb_type::BrtBeginPCDSCPItem:
                return fmt::format_to(ctx.out(), "BrtBeginPCDSCPItem");
            case xlsb_type::BrtEndPCDSCPItem:
                return fmt::format_to(ctx.out(), "BrtEndPCDSCPItem");
            case xlsb_type::BrtBeginPCDSCSets:
                return fmt::format_to(ctx.out(), "BrtBeginPCDSCSets");
            case xlsb_type::BrtEndPCDSCSets:
                return fmt::format_to(ctx.out(), "BrtEndPCDSCSets");
            case xlsb_type::BrtBeginPCDSCSet:
                return fmt::format_to(ctx.out(), "BrtBeginPCDSCSet");
            case xlsb_type::BrtEndPCDSCSet:
                return fmt::format_to(ctx.out(), "BrtEndPCDSCSet");
            case xlsb_type::BrtBeginPCDFGroup:
                return fmt::format_to(ctx.out(), "BrtBeginPCDFGroup");
            case xlsb_type::BrtEndPCDFGroup:
                return fmt::format_to(ctx.out(), "BrtEndPCDFGroup");
            case xlsb_type::BrtBeginPCDFGItems:
                return fmt::format_to(ctx.out(), "BrtBeginPCDFGItems");
            case xlsb_type::BrtEndPCDFGItems:
                return fmt::format_to(ctx.out(), "BrtEndPCDFGItems");
            case xlsb_type::BrtBeginPCDFGRange:
                return fmt::format_to(ctx.out(), "BrtBeginPCDFGRange");
            case xlsb_type::BrtEndPCDFGRange:
                return fmt::format_to(ctx.out(), "BrtEndPCDFGRange");
            case xlsb_type::BrtBeginPCDFGDiscrete:
                return fmt::format_to(ctx.out(), "BrtBeginPCDFGDiscrete");
            case xlsb_type::BrtEndPCDFGDiscrete:
                return fmt::format_to(ctx.out(), "BrtEndPCDFGDiscrete");
            case xlsb_type::BrtBeginPCDSDTupleCache:
                return fmt::format_to(ctx.out(), "BrtBeginPCDSDTupleCache");
            case xlsb_type::BrtEndPCDSDTupleCache:
                return fmt::format_to(ctx.out(), "BrtEndPCDSDTupleCache");
            case xlsb_type::BrtBeginPCDSDTCEntries:
                return fmt::format_to(ctx.out(), "BrtBeginPCDSDTCEntries");
            case xlsb_type::BrtEndPCDSDTCEntries:
                return fmt::format_to(ctx.out(), "BrtEndPCDSDTCEntries");
            case xlsb_type::BrtBeginPCDSDTCEMembers:
                return fmt::format_to(ctx.out(), "BrtBeginPCDSDTCEMembers");
            case xlsb_type::BrtEndPCDSDTCEMembers:
                return fmt::format_to(ctx.out(), "BrtEndPCDSDTCEMembers");
            case xlsb_type::BrtBeginPCDSDTCEMember:
                return fmt::format_to(ctx.out(), "BrtBeginPCDSDTCEMember");
            case xlsb_type::BrtEndPCDSDTCEMember:
                return fmt::format_to(ctx.out(), "BrtEndPCDSDTCEMember");
            case xlsb_type::BrtBeginPCDSDTCQueries:
                return fmt::format_to(ctx.out(), "BrtBeginPCDSDTCQueries");
            case xlsb_type::BrtEndPCDSDTCQueries:
                return fmt::format_to(ctx.out(), "BrtEndPCDSDTCQueries");
            case xlsb_type::BrtBeginPCDSDTCQuery:
                return fmt::format_to(ctx.out(), "BrtBeginPCDSDTCQuery");
            case xlsb_type::BrtEndPCDSDTCQuery:
                return fmt::format_to(ctx.out(), "BrtEndPCDSDTCQuery");
            case xlsb_type::BrtBeginPCDSDTCSets:
                return fmt::format_to(ctx.out(), "BrtBeginPCDSDTCSets");
            case xlsb_type::BrtEndPCDSDTCSets:
                return fmt::format_to(ctx.out(), "BrtEndPCDSDTCSets");
            case xlsb_type::BrtBeginPCDSDTCSet:
                return fmt::format_to(ctx.out(), "BrtBeginPCDSDTCSet");
            case xlsb_type::BrtEndPCDSDTCSet:
                return fmt::format_to(ctx.out(), "BrtEndPCDSDTCSet");
            case xlsb_type::BrtBeginPCDCalcItems:
                return fmt::format_to(ctx.out(), "BrtBeginPCDCalcItems");
            case xlsb_type::BrtEndPCDCalcItems:
                return fmt::format_to(ctx.out(), "BrtEndPCDCalcItems");
            case xlsb_type::BrtBeginPCDCalcItem:
                return fmt::format_to(ctx.out(), "BrtBeginPCDCalcItem");
            case xlsb_type::BrtEndPCDCalcItem:
                return fmt::format_to(ctx.out(), "BrtEndPCDCalcItem");
            case xlsb_type::BrtBeginPRule:
                return fmt::format_to(ctx.out(), "BrtBeginPRule");
            case xlsb_type::BrtEndPRule:
                return fmt::format_to(ctx.out(), "BrtEndPRule");
            case xlsb_type::BrtBeginPRFilters:
                return fmt::format_to(ctx.out(), "BrtBeginPRFilters");
            case xlsb_type::BrtEndPRFilters:
                return fmt::format_to(ctx.out(), "BrtEndPRFilters");
            case xlsb_type::BrtBeginPRFilter:
                return fmt::format_to(ctx.out(), "BrtBeginPRFilter");
            case xlsb_type::BrtEndPRFilter:
                return fmt::format_to(ctx.out(), "BrtEndPRFilter");
            case xlsb_type::BrtBeginPNames:
                return fmt::format_to(ctx.out(), "BrtBeginPNames");
            case xlsb_type::BrtEndPNames:
                return fmt::format_to(ctx.out(), "BrtEndPNames");
            case xlsb_type::BrtBeginPName:
                return fmt::format_to(ctx.out(), "BrtBeginPName");
            case xlsb_type::BrtEndPName:
                return fmt::format_to(ctx.out(), "BrtEndPName");
            case xlsb_type::BrtBeginPNPairs:
                return fmt::format_to(ctx.out(), "BrtBeginPNPairs");
            case xlsb_type::BrtEndPNPairs:
                return fmt::format_to(ctx.out(), "BrtEndPNPairs");
            case xlsb_type::BrtBeginPNPair:
                return fmt::format_to(ctx.out(), "BrtBeginPNPair");
            case xlsb_type::BrtEndPNPair:
                return fmt::format_to(ctx.out(), "BrtEndPNPair");
            case xlsb_type::BrtBeginECWebProps:
                return fmt::format_to(ctx.out(), "BrtBeginECWebProps");
            case xlsb_type::BrtEndECWebProps:
                return fmt::format_to(ctx.out(), "BrtEndECWebProps");
            case xlsb_type::BrtBeginEcWpTables:
                return fmt::format_to(ctx.out(), "BrtBeginEcWpTables");
            case xlsb_type::BrtEndECWPTables:
                return fmt::format_to(ctx.out(), "BrtEndECWPTables");
            case xlsb_type::BrtBeginECParams:
                return fmt::format_to(ctx.out(), "BrtBeginECParams");
            case xlsb_type::BrtEndECParams:
                return fmt::format_to(ctx.out(), "BrtEndECParams");
            case xlsb_type::BrtBeginECParam:
                return fmt::format_to(ctx.out(), "BrtBeginECParam");
            case xlsb_type::BrtEndECParam:
                return fmt::format_to(ctx.out(), "BrtEndECParam");
            case xlsb_type::BrtBeginPCDKPIs:
                return fmt::format_to(ctx.out(), "BrtBeginPCDKPIs");
            case xlsb_type::BrtEndPCDKPIs:
                return fmt::format_to(ctx.out(), "BrtEndPCDKPIs");
            case xlsb_type::BrtBeginPCDKPI:
                return fmt::format_to(ctx.out(), "BrtBeginPCDKPI");
            case xlsb_type::BrtEndPCDKPI:
                return fmt::format_to(ctx.out(), "BrtEndPCDKPI");
            case xlsb_type::BrtBeginDims:
                return fmt::format_to(ctx.out(), "BrtBeginDims");
            case xlsb_type::BrtEndDims:
                return fmt::format_to(ctx.out(), "BrtEndDims");
            case xlsb_type::BrtBeginDim:
                return fmt::format_to(ctx.out(), "BrtBeginDim");
            case xlsb_type::BrtEndDim:
                return fmt::format_to(ctx.out(), "BrtEndDim");
            case xlsb_type::BrtIndexPartEnd:
                return fmt::format_to(ctx.out(), "BrtIndexPartEnd");
            case xlsb_type::BrtBeginStyleSheet:
                return fmt::format_to(ctx.out(), "BrtBeginStyleSheet");
            case xlsb_type::BrtEndStyleSheet:
                return fmt::format_to(ctx.out(), "BrtEndStyleSheet");
            case xlsb_type::BrtBeginSXView:
                return fmt::format_to(ctx.out(), "BrtBeginSXView");
            case xlsb_type::BrtEndSXVI:
                return fmt::format_to(ctx.out(), "BrtEndSXVI");
            case xlsb_type::BrtBeginSXVI:
                return fmt::format_to(ctx.out(), "BrtBeginSXVI");
            case xlsb_type::BrtBeginSXVIs:
                return fmt::format_to(ctx.out(), "BrtBeginSXVIs");
            case xlsb_type::BrtEndSXVIs:
                return fmt::format_to(ctx.out(), "BrtEndSXVIs");
            case xlsb_type::BrtBeginSXVD:
                return fmt::format_to(ctx.out(), "BrtBeginSXVD");
            case xlsb_type::BrtEndSXVD:
                return fmt::format_to(ctx.out(), "BrtEndSXVD");
            case xlsb_type::BrtBeginSXVDs:
                return fmt::format_to(ctx.out(), "BrtBeginSXVDs");
            case xlsb_type::BrtEndSXVDs:
                return fmt::format_to(ctx.out(), "BrtEndSXVDs");
            case xlsb_type::BrtBeginSXPI:
                return fmt::format_to(ctx.out(), "BrtBeginSXPI");
            case xlsb_type::BrtEndSXPI:
                return fmt::format_to(ctx.out(), "BrtEndSXPI");
            case xlsb_type::BrtBeginSXPIs:
                return fmt::format_to(ctx.out(), "BrtBeginSXPIs");
            case xlsb_type::BrtEndSXPIs:
                return fmt::format_to(ctx.out(), "BrtEndSXPIs");
            case xlsb_type::BrtBeginSXDI:
                return fmt::format_to(ctx.out(), "BrtBeginSXDI");
            case xlsb_type::BrtEndSXDI:
                return fmt::format_to(ctx.out(), "BrtEndSXDI");
            case xlsb_type::BrtBeginSXDIs:
                return fmt::format_to(ctx.out(), "BrtBeginSXDIs");
            case xlsb_type::BrtEndSXDIs:
                return fmt::format_to(ctx.out(), "BrtEndSXDIs");
            case xlsb_type::BrtBeginSXLI:
                return fmt::format_to(ctx.out(), "BrtBeginSXLI");
            case xlsb_type::BrtEndSXLI:
                return fmt::format_to(ctx.out(), "BrtEndSXLI");
            case xlsb_type::BrtBeginSXLIRws:
                return fmt::format_to(ctx.out(), "BrtBeginSXLIRws");
            case xlsb_type::BrtEndSXLIRws:
                return fmt::format_to(ctx.out(), "BrtEndSXLIRws");
            case xlsb_type::BrtBeginSXLICols:
                return fmt::format_to(ctx.out(), "BrtBeginSXLICols");
            case xlsb_type::BrtEndSXLICols:
                return fmt::format_to(ctx.out(), "BrtEndSXLICols");
            case xlsb_type::BrtBeginSXFormat:
                return fmt::format_to(ctx.out(), "BrtBeginSXFormat");
            case xlsb_type::BrtEndSXFormat:
                return fmt::format_to(ctx.out(), "BrtEndSXFormat");
            case xlsb_type::BrtBeginSXFormats:
                return fmt::format_to(ctx.out(), "BrtBeginSXFormats");
            case xlsb_type::BrtEndSxFormats:
                return fmt::format_to(ctx.out(), "BrtEndSxFormats");
            case xlsb_type::BrtBeginSxSelect:
                return fmt::format_to(ctx.out(), "BrtBeginSxSelect");
            case xlsb_type::BrtEndSxSelect:
                return fmt::format_to(ctx.out(), "BrtEndSxSelect");
            case xlsb_type::BrtBeginISXVDRws:
                return fmt::format_to(ctx.out(), "BrtBeginISXVDRws");
            case xlsb_type::BrtEndISXVDRws:
                return fmt::format_to(ctx.out(), "BrtEndISXVDRws");
            case xlsb_type::BrtBeginISXVDCols:
                return fmt::format_to(ctx.out(), "BrtBeginISXVDCols");
            case xlsb_type::BrtEndISXVDCols:
                return fmt::format_to(ctx.out(), "BrtEndISXVDCols");
            case xlsb_type::BrtEndSXLocation:
                return fmt::format_to(ctx.out(), "BrtEndSXLocation");
            case xlsb_type::BrtBeginSXLocation:
                return fmt::format_to(ctx.out(), "BrtBeginSXLocation");
            case xlsb_type::BrtEndSXView:
                return fmt::format_to(ctx.out(), "BrtEndSXView");
            case xlsb_type::BrtBeginSXTHs:
                return fmt::format_to(ctx.out(), "BrtBeginSXTHs");
            case xlsb_type::BrtEndSXTHs:
                return fmt::format_to(ctx.out(), "BrtEndSXTHs");
            case xlsb_type::BrtBeginSXTH:
                return fmt::format_to(ctx.out(), "BrtBeginSXTH");
            case xlsb_type::BrtEndSXTH:
                return fmt::format_to(ctx.out(), "BrtEndSXTH");
            case xlsb_type::BrtBeginISXTHRws:
                return fmt::format_to(ctx.out(), "BrtBeginISXTHRws");
            case xlsb_type::BrtEndISXTHRws:
                return fmt::format_to(ctx.out(), "BrtEndISXTHRws");
            case xlsb_type::BrtBeginISXTHCols:
                return fmt::format_to(ctx.out(), "BrtBeginISXTHCols");
            case xlsb_type::BrtEndISXTHCols:
                return fmt::format_to(ctx.out(), "BrtEndISXTHCols");
            case xlsb_type::BrtBeginSXTDMPS:
                return fmt::format_to(ctx.out(), "BrtBeginSXTDMPS");
            case xlsb_type::BrtEndSXTDMPs:
                return fmt::format_to(ctx.out(), "BrtEndSXTDMPs");
            case xlsb_type::BrtBeginSXTDMP:
                return fmt::format_to(ctx.out(), "BrtBeginSXTDMP");
            case xlsb_type::BrtEndSXTDMP:
                return fmt::format_to(ctx.out(), "BrtEndSXTDMP");
            case xlsb_type::BrtBeginSXTHItems:
                return fmt::format_to(ctx.out(), "BrtBeginSXTHItems");
            case xlsb_type::BrtEndSXTHItems:
                return fmt::format_to(ctx.out(), "BrtEndSXTHItems");
            case xlsb_type::BrtBeginSXTHItem:
                return fmt::format_to(ctx.out(), "BrtBeginSXTHItem");
            case xlsb_type::BrtEndSXTHItem:
                return fmt::format_to(ctx.out(), "BrtEndSXTHItem");
            case xlsb_type::BrtBeginMetadata:
                return fmt::format_to(ctx.out(), "BrtBeginMetadata");
            case xlsb_type::BrtEndMetadata:
                return fmt::format_to(ctx.out(), "BrtEndMetadata");
            case xlsb_type::BrtBeginEsmdtinfo:
                return fmt::format_to(ctx.out(), "BrtBeginEsmdtinfo");
            case xlsb_type::BrtMdtinfo:
                return fmt::format_to(ctx.out(), "BrtMdtinfo");
            case xlsb_type::BrtEndEsmdtinfo:
                return fmt::format_to(ctx.out(), "BrtEndEsmdtinfo");
            case xlsb_type::BrtBeginEsmdb:
                return fmt::format_to(ctx.out(), "BrtBeginEsmdb");
            case xlsb_type::BrtEndEsmdb:
                return fmt::format_to(ctx.out(), "BrtEndEsmdb");
            case xlsb_type::BrtBeginEsfmd:
                return fmt::format_to(ctx.out(), "BrtBeginEsfmd");
            case xlsb_type::BrtEndEsfmd:
                return fmt::format_to(ctx.out(), "BrtEndEsfmd");
            case xlsb_type::BrtBeginSingleCells:
                return fmt::format_to(ctx.out(), "BrtBeginSingleCells");
            case xlsb_type::BrtEndSingleCells:
                return fmt::format_to(ctx.out(), "BrtEndSingleCells");
            case xlsb_type::BrtBeginList:
                return fmt::format_to(ctx.out(), "BrtBeginList");
            case xlsb_type::BrtEndList:
                return fmt::format_to(ctx.out(), "BrtEndList");
            case xlsb_type::BrtBeginListCols:
                return fmt::format_to(ctx.out(), "BrtBeginListCols");
            case xlsb_type::BrtEndListCols:
                return fmt::format_to(ctx.out(), "BrtEndListCols");
            case xlsb_type::BrtBeginListCol:
                return fmt::format_to(ctx.out(), "BrtBeginListCol");
            case xlsb_type::BrtEndListCol:
                return fmt::format_to(ctx.out(), "BrtEndListCol");
            case xlsb_type::BrtBeginListXmlCPr:
                return fmt::format_to(ctx.out(), "BrtBeginListXmlCPr");
            case xlsb_type::BrtEndListXmlCPr:
                return fmt::format_to(ctx.out(), "BrtEndListXmlCPr");
            case xlsb_type::BrtListCCFmla:
                return fmt::format_to(ctx.out(), "BrtListCCFmla");
            case xlsb_type::BrtListTrFmla:
                return fmt::format_to(ctx.out(), "BrtListTrFmla");
            case xlsb_type::BrtBeginExternals:
                return fmt::format_to(ctx.out(), "BrtBeginExternals");
            case xlsb_type::BrtEndExternals:
                return fmt::format_to(ctx.out(), "BrtEndExternals");
            case xlsb_type::BrtSupBookSrc:
                return fmt::format_to(ctx.out(), "BrtSupBookSrc");
            case xlsb_type::BrtSupSelf:
                return fmt::format_to(ctx.out(), "BrtSupSelf");
            case xlsb_type::BrtSupSame:
                return fmt::format_to(ctx.out(), "BrtSupSame");
            case xlsb_type::BrtSupTabs:
                return fmt::format_to(ctx.out(), "BrtSupTabs");
            case xlsb_type::BrtBeginSupBook:
                return fmt::format_to(ctx.out(), "BrtBeginSupBook");
            case xlsb_type::BrtPlaceholderName:
                return fmt::format_to(ctx.out(), "BrtPlaceholderName");
            case xlsb_type::BrtExternSheet:
                return fmt::format_to(ctx.out(), "BrtExternSheet");
            case xlsb_type::BrtExternTableStart:
                return fmt::format_to(ctx.out(), "BrtExternTableStart");
            case xlsb_type::BrtExternTableEnd:
                return fmt::format_to(ctx.out(), "BrtExternTableEnd");
            case xlsb_type::BrtExternRowHdr:
                return fmt::format_to(ctx.out(), "BrtExternRowHdr");
            case xlsb_type::BrtExternCellBlank:
                return fmt::format_to(ctx.out(), "BrtExternCellBlank");
            case xlsb_type::BrtExternCellReal:
                return fmt::format_to(ctx.out(), "BrtExternCellReal");
            case xlsb_type::BrtExternCellBool:
                return fmt::format_to(ctx.out(), "BrtExternCellBool");
            case xlsb_type::BrtExternCellError:
                return fmt::format_to(ctx.out(), "BrtExternCellError");
            case xlsb_type::BrtExternCellString:
                return fmt::format_to(ctx.out(), "BrtExternCellString");
            case xlsb_type::BrtBeginEsmdx:
                return fmt::format_to(ctx.out(), "BrtBeginEsmdx");
            case xlsb_type::BrtEndEsmdx:
                return fmt::format_to(ctx.out(), "BrtEndEsmdx");
            case xlsb_type::BrtBeginMdxSet:
                return fmt::format_to(ctx.out(), "BrtBeginMdxSet");
            case xlsb_type::BrtEndMdxSet:
                return fmt::format_to(ctx.out(), "BrtEndMdxSet");
            case xlsb_type::BrtBeginMdxMbrProp:
                return fmt::format_to(ctx.out(), "BrtBeginMdxMbrProp");
            case xlsb_type::BrtEndMdxMbrProp:
                return fmt::format_to(ctx.out(), "BrtEndMdxMbrProp");
            case xlsb_type::BrtBeginMdxKPI:
                return fmt::format_to(ctx.out(), "BrtBeginMdxKPI");
            case xlsb_type::BrtEndMdxKPI:
                return fmt::format_to(ctx.out(), "BrtEndMdxKPI");
            case xlsb_type::BrtBeginEsstr:
                return fmt::format_to(ctx.out(), "BrtBeginEsstr");
            case xlsb_type::BrtEndEsstr:
                return fmt::format_to(ctx.out(), "BrtEndEsstr");
            case xlsb_type::BrtBeginPRFItem:
                return fmt::format_to(ctx.out(), "BrtBeginPRFItem");
            case xlsb_type::BrtEndPRFItem:
                return fmt::format_to(ctx.out(), "BrtEndPRFItem");
            case xlsb_type::BrtBeginPivotCacheIDs:
                return fmt::format_to(ctx.out(), "BrtBeginPivotCacheIDs");
            case xlsb_type::BrtEndPivotCacheIDs:
                return fmt::format_to(ctx.out(), "BrtEndPivotCacheIDs");
            case xlsb_type::BrtBeginPivotCacheID:
                return fmt::format_to(ctx.out(), "BrtBeginPivotCacheID");
            case xlsb_type::BrtEndPivotCacheID:
                return fmt::format_to(ctx.out(), "BrtEndPivotCacheID");
            case xlsb_type::BrtBeginISXVIs:
                return fmt::format_to(ctx.out(), "BrtBeginISXVIs");
            case xlsb_type::BrtEndISXVIs:
                return fmt::format_to(ctx.out(), "BrtEndISXVIs");
            case xlsb_type::BrtBeginColInfos:
                return fmt::format_to(ctx.out(), "BrtBeginColInfos");
            case xlsb_type::BrtEndColInfos:
                return fmt::format_to(ctx.out(), "BrtEndColInfos");
            case xlsb_type::BrtBeginRwBrk:
                return fmt::format_to(ctx.out(), "BrtBeginRwBrk");
            case xlsb_type::BrtEndRwBrk:
                return fmt::format_to(ctx.out(), "BrtEndRwBrk");
            case xlsb_type::BrtBeginColBrk:
                return fmt::format_to(ctx.out(), "BrtBeginColBrk");
            case xlsb_type::BrtEndColBrk:
                return fmt::format_to(ctx.out(), "BrtEndColBrk");
            case xlsb_type::BrtBrk:
                return fmt::format_to(ctx.out(), "BrtBrk");
            case xlsb_type::BrtUserBookView:
                return fmt::format_to(ctx.out(), "BrtUserBookView");
            case xlsb_type::BrtInfo:
                return fmt::format_to(ctx.out(), "BrtInfo");
            case xlsb_type::BrtCUsr:
                return fmt::format_to(ctx.out(), "BrtCUsr");
            case xlsb_type::BrtUsr:
                return fmt::format_to(ctx.out(), "BrtUsr");
            case xlsb_type::BrtBeginUsers:
                return fmt::format_to(ctx.out(), "BrtBeginUsers");
            case xlsb_type::BrtEOF:
                return fmt::format_to(ctx.out(), "BrtEOF");
            case xlsb_type::BrtUCR:
                return fmt::format_to(ctx.out(), "BrtUCR");
            case xlsb_type::BrtRRInsDel:
                return fmt::format_to(ctx.out(), "BrtRRInsDel");
            case xlsb_type::BrtRREndInsDel:
                return fmt::format_to(ctx.out(), "BrtRREndInsDel");
            case xlsb_type::BrtRRMove:
                return fmt::format_to(ctx.out(), "BrtRRMove");
            case xlsb_type::BrtRREndMove:
                return fmt::format_to(ctx.out(), "BrtRREndMove");
            case xlsb_type::BrtRRChgCell:
                return fmt::format_to(ctx.out(), "BrtRRChgCell");
            case xlsb_type::BrtRREndChgCell:
                return fmt::format_to(ctx.out(), "BrtRREndChgCell");
            case xlsb_type::BrtRRHeader:
                return fmt::format_to(ctx.out(), "BrtRRHeader");
            case xlsb_type::BrtRRUserView:
                return fmt::format_to(ctx.out(), "BrtRRUserView");
            case xlsb_type::BrtRRRenSheet:
                return fmt::format_to(ctx.out(), "BrtRRRenSheet");
            case xlsb_type::BrtRRInsertSh:
                return fmt::format_to(ctx.out(), "BrtRRInsertSh");
            case xlsb_type::BrtRRDefName:
                return fmt::format_to(ctx.out(), "BrtRRDefName");
            case xlsb_type::BrtRRNote:
                return fmt::format_to(ctx.out(), "BrtRRNote");
            case xlsb_type::BrtRRConflict:
                return fmt::format_to(ctx.out(), "BrtRRConflict");
            case xlsb_type::BrtRRTQSIF:
                return fmt::format_to(ctx.out(), "BrtRRTQSIF");
            case xlsb_type::BrtRRFormat:
                return fmt::format_to(ctx.out(), "BrtRRFormat");
            case xlsb_type::BrtRREndFormat:
                return fmt::format_to(ctx.out(), "BrtRREndFormat");
            case xlsb_type::BrtRRAutoFmt:
                return fmt::format_to(ctx.out(), "BrtRRAutoFmt");
            case xlsb_type::BrtBeginUserShViews:
                return fmt::format_to(ctx.out(), "BrtBeginUserShViews");
            case xlsb_type::BrtBeginUserShView:
                return fmt::format_to(ctx.out(), "BrtBeginUserShView");
            case xlsb_type::BrtEndUserShView:
                return fmt::format_to(ctx.out(), "BrtEndUserShView");
            case xlsb_type::BrtEndUserShViews:
                return fmt::format_to(ctx.out(), "BrtEndUserShViews");
            case xlsb_type::BrtArrFmla:
                return fmt::format_to(ctx.out(), "BrtArrFmla");
            case xlsb_type::BrtShrFmla:
                return fmt::format_to(ctx.out(), "BrtShrFmla");
            case xlsb_type::BrtTable:
                return fmt::format_to(ctx.out(), "BrtTable");
            case xlsb_type::BrtBeginExtConnections:
                return fmt::format_to(ctx.out(), "BrtBeginExtConnections");
            case xlsb_type::BrtEndExtConnections:
                return fmt::format_to(ctx.out(), "BrtEndExtConnections");
            case xlsb_type::BrtBeginPCDCalcMems:
                return fmt::format_to(ctx.out(), "BrtBeginPCDCalcMems");
            case xlsb_type::BrtEndPCDCalcMems:
                return fmt::format_to(ctx.out(), "BrtEndPCDCalcMems");
            case xlsb_type::BrtBeginPCDCalcMem:
                return fmt::format_to(ctx.out(), "BrtBeginPCDCalcMem");
            case xlsb_type::BrtEndPCDCalcMem:
                return fmt::format_to(ctx.out(), "BrtEndPCDCalcMem");
            case xlsb_type::BrtBeginPCDHGLevels:
                return fmt::format_to(ctx.out(), "BrtBeginPCDHGLevels");
            case xlsb_type::BrtEndPCDHGLevels:
                return fmt::format_to(ctx.out(), "BrtEndPCDHGLevels");
            case xlsb_type::BrtBeginPCDHGLevel:
                return fmt::format_to(ctx.out(), "BrtBeginPCDHGLevel");
            case xlsb_type::BrtEndPCDHGLevel:
                return fmt::format_to(ctx.out(), "BrtEndPCDHGLevel");
            case xlsb_type::BrtBeginPCDHGLGroups:
                return fmt::format_to(ctx.out(), "BrtBeginPCDHGLGroups");
            case xlsb_type::BrtEndPCDHGLGroups:
                return fmt::format_to(ctx.out(), "BrtEndPCDHGLGroups");
            case xlsb_type::BrtBeginPCDHGLGroup:
                return fmt::format_to(ctx.out(), "BrtBeginPCDHGLGroup");
            case xlsb_type::BrtEndPCDHGLGroup:
                return fmt::format_to(ctx.out(), "BrtEndPCDHGLGroup");
            case xlsb_type::BrtBeginPCDHGLGMembers:
                return fmt::format_to(ctx.out(), "BrtBeginPCDHGLGMembers");
            case xlsb_type::BrtEndPCDHGLGMembers:
                return fmt::format_to(ctx.out(), "BrtEndPCDHGLGMembers");
            case xlsb_type::BrtBeginPCDHGLGMember:
                return fmt::format_to(ctx.out(), "BrtBeginPCDHGLGMember");
            case xlsb_type::BrtEndPCDHGLGMember:
                return fmt::format_to(ctx.out(), "BrtEndPCDHGLGMember");
            case xlsb_type::BrtBeginQSI:
                return fmt::format_to(ctx.out(), "BrtBeginQSI");
            case xlsb_type::BrtEndQSI:
                return fmt::format_to(ctx.out(), "BrtEndQSI");
            case xlsb_type::BrtBeginQSIR:
                return fmt::format_to(ctx.out(), "BrtBeginQSIR");
            case xlsb_type::BrtEndQSIR:
                return fmt::format_to(ctx.out(), "BrtEndQSIR");
            case xlsb_type::BrtBeginDeletedNames:
                return fmt::format_to(ctx.out(), "BrtBeginDeletedNames");
            case xlsb_type::BrtEndDeletedNames:
                return fmt::format_to(ctx.out(), "BrtEndDeletedNames");
            case xlsb_type::BrtBeginDeletedName:
                return fmt::format_to(ctx.out(), "BrtBeginDeletedName");
            case xlsb_type::BrtEndDeletedName:
                return fmt::format_to(ctx.out(), "BrtEndDeletedName");
            case xlsb_type::BrtBeginQSIFs:
                return fmt::format_to(ctx.out(), "BrtBeginQSIFs");
            case xlsb_type::BrtEndQSIFs:
                return fmt::format_to(ctx.out(), "BrtEndQSIFs");
            case xlsb_type::BrtBeginQSIF:
                return fmt::format_to(ctx.out(), "BrtBeginQSIF");
            case xlsb_type::BrtEndQSIF:
                return fmt::format_to(ctx.out(), "BrtEndQSIF");
            case xlsb_type::BrtBeginAutoSortScope:
                return fmt::format_to(ctx.out(), "BrtBeginAutoSortScope");
            case xlsb_type::BrtEndAutoSortScope:
                return fmt::format_to(ctx.out(), "BrtEndAutoSortScope");
            case xlsb_type::BrtBeginConditionalFormatting:
                return fmt::format_to(ctx.out(), "BrtBeginConditionalFormatting");
            case xlsb_type::BrtEndConditionalFormatting:
                return fmt::format_to(ctx.out(), "BrtEndConditionalFormatting");
            case xlsb_type::BrtBeginCFRule:
                return fmt::format_to(ctx.out(), "BrtBeginCFRule");
            case xlsb_type::BrtEndCFRule:
                return fmt::format_to(ctx.out(), "BrtEndCFRule");
            case xlsb_type::BrtBeginIconSet:
                return fmt::format_to(ctx.out(), "BrtBeginIconSet");
            case xlsb_type::BrtEndIconSet:
                return fmt::format_to(ctx.out(), "BrtEndIconSet");
            case xlsb_type::BrtBeginDatabar:
                return fmt::format_to(ctx.out(), "BrtBeginDatabar");
            case xlsb_type::BrtEndDatabar:
                return fmt::format_to(ctx.out(), "BrtEndDatabar");
            case xlsb_type::BrtBeginColorScale:
                return fmt::format_to(ctx.out(), "BrtBeginColorScale");
            case xlsb_type::BrtEndColorScale:
                return fmt::format_to(ctx.out(), "BrtEndColorScale");
            case xlsb_type::BrtCFVO:
                return fmt::format_to(ctx.out(), "BrtCFVO");
            case xlsb_type::BrtExternValueMeta:
                return fmt::format_to(ctx.out(), "BrtExternValueMeta");
            case xlsb_type::BrtBeginColorPalette:
                return fmt::format_to(ctx.out(), "BrtBeginColorPalette");
            case xlsb_type::BrtEndColorPalette:
                return fmt::format_to(ctx.out(), "BrtEndColorPalette");
            case xlsb_type::BrtIndexedColor:
                return fmt::format_to(ctx.out(), "BrtIndexedColor");
            case xlsb_type::BrtMargins:
                return fmt::format_to(ctx.out(), "BrtMargins");
            case xlsb_type::BrtPrintOptions:
                return fmt::format_to(ctx.out(), "BrtPrintOptions");
            case xlsb_type::BrtPageSetup:
                return fmt::format_to(ctx.out(), "BrtPageSetup");
            case xlsb_type::BrtBeginHeaderFooter:
                return fmt::format_to(ctx.out(), "BrtBeginHeaderFooter");
            case xlsb_type::BrtEndHeaderFooter:
                return fmt::format_to(ctx.out(), "BrtEndHeaderFooter");
            case xlsb_type::BrtBeginSXCrtFormat:
                return fmt::format_to(ctx.out(), "BrtBeginSXCrtFormat");
            case xlsb_type::BrtEndSXCrtFormat:
                return fmt::format_to(ctx.out(), "BrtEndSXCrtFormat");
            case xlsb_type::BrtBeginSXCrtFormats:
                return fmt::format_to(ctx.out(), "BrtBeginSXCrtFormats");
            case xlsb_type::BrtEndSXCrtFormats:
                return fmt::format_to(ctx.out(), "BrtEndSXCrtFormats");
            case xlsb_type::BrtWsFmtInfo:
                return fmt::format_to(ctx.out(), "BrtWsFmtInfo");
            case xlsb_type::BrtBeginMgs:
                return fmt::format_to(ctx.out(), "BrtBeginMgs");
            case xlsb_type::BrtEndMGs:
                return fmt::format_to(ctx.out(), "BrtEndMGs");
            case xlsb_type::BrtBeginMGMaps:
                return fmt::format_to(ctx.out(), "BrtBeginMGMaps");
            case xlsb_type::BrtEndMGMaps:
                return fmt::format_to(ctx.out(), "BrtEndMGMaps");
            case xlsb_type::BrtBeginMG:
                return fmt::format_to(ctx.out(), "BrtBeginMG");
            case xlsb_type::BrtEndMG:
                return fmt::format_to(ctx.out(), "BrtEndMG");
            case xlsb_type::BrtBeginMap:
                return fmt::format_to(ctx.out(), "BrtBeginMap");
            case xlsb_type::BrtEndMap:
                return fmt::format_to(ctx.out(), "BrtEndMap");
            case xlsb_type::BrtHLink:
                return fmt::format_to(ctx.out(), "BrtHLink");
            case xlsb_type::BrtBeginDCon:
                return fmt::format_to(ctx.out(), "BrtBeginDCon");
            case xlsb_type::BrtEndDCon:
                return fmt::format_to(ctx.out(), "BrtEndDCon");
            case xlsb_type::BrtBeginDRefs:
                return fmt::format_to(ctx.out(), "BrtBeginDRefs");
            case xlsb_type::BrtEndDRefs:
                return fmt::format_to(ctx.out(), "BrtEndDRefs");
            case xlsb_type::BrtDRef:
                return fmt::format_to(ctx.out(), "BrtDRef");
            case xlsb_type::BrtBeginScenMan:
                return fmt::format_to(ctx.out(), "BrtBeginScenMan");
            case xlsb_type::BrtEndScenMan:
                return fmt::format_to(ctx.out(), "BrtEndScenMan");
            case xlsb_type::BrtBeginSct:
                return fmt::format_to(ctx.out(), "BrtBeginSct");
            case xlsb_type::BrtEndSct:
                return fmt::format_to(ctx.out(), "BrtEndSct");
            case xlsb_type::BrtSlc:
                return fmt::format_to(ctx.out(), "BrtSlc");
            case xlsb_type::BrtBeginDXFs:
                return fmt::format_to(ctx.out(), "BrtBeginDXFs");
            case xlsb_type::BrtEndDXFs:
                return fmt::format_to(ctx.out(), "BrtEndDXFs");
            case xlsb_type::BrtDXF:
                return fmt::format_to(ctx.out(), "BrtDXF");
            case xlsb_type::BrtBeginTableStyles:
                return fmt::format_to(ctx.out(), "BrtBeginTableStyles");
            case xlsb_type::BrtEndTableStyles:
                return fmt::format_to(ctx.out(), "BrtEndTableStyles");
            case xlsb_type::BrtBeginTableStyle:
                return fmt::format_to(ctx.out(), "BrtBeginTableStyle");
            case xlsb_type::BrtEndTableStyle:
                return fmt::format_to(ctx.out(), "BrtEndTableStyle");
            case xlsb_type::BrtTableStyleElement:
                return fmt::format_to(ctx.out(), "BrtTableStyleElement");
            case xlsb_type::BrtTableStyleClient:
                return fmt::format_to(ctx.out(), "BrtTableStyleClient");
            case xlsb_type::BrtBeginVolDeps:
                return fmt::format_to(ctx.out(), "BrtBeginVolDeps");
            case xlsb_type::BrtEndVolDeps:
                return fmt::format_to(ctx.out(), "BrtEndVolDeps");
            case xlsb_type::BrtBeginVolType:
                return fmt::format_to(ctx.out(), "BrtBeginVolType");
            case xlsb_type::BrtEndVolType:
                return fmt::format_to(ctx.out(), "BrtEndVolType");
            case xlsb_type::BrtBeginVolMain:
                return fmt::format_to(ctx.out(), "BrtBeginVolMain");
            case xlsb_type::BrtEndVolMain:
                return fmt::format_to(ctx.out(), "BrtEndVolMain");
            case xlsb_type::BrtBeginVolTopic:
                return fmt::format_to(ctx.out(), "BrtBeginVolTopic");
            case xlsb_type::BrtEndVolTopic:
                return fmt::format_to(ctx.out(), "BrtEndVolTopic");
            case xlsb_type::BrtVolSubtopic:
                return fmt::format_to(ctx.out(), "BrtVolSubtopic");
            case xlsb_type::BrtVolRef:
                return fmt::format_to(ctx.out(), "BrtVolRef");
            case xlsb_type::BrtVolNum:
                return fmt::format_to(ctx.out(), "BrtVolNum");
            case xlsb_type::BrtVolErr:
                return fmt::format_to(ctx.out(), "BrtVolErr");
            case xlsb_type::BrtVolStr:
                return fmt::format_to(ctx.out(), "BrtVolStr");
            case xlsb_type::BrtVolBool:
                return fmt::format_to(ctx.out(), "BrtVolBool");
            case xlsb_type::BrtBeginSortState:
                return fmt::format_to(ctx.out(), "BrtBeginSortState");
            case xlsb_type::BrtEndSortState:
                return fmt::format_to(ctx.out(), "BrtEndSortState");
            case xlsb_type::BrtBeginSortCond:
                return fmt::format_to(ctx.out(), "BrtBeginSortCond");
            case xlsb_type::BrtEndSortCond:
                return fmt::format_to(ctx.out(), "BrtEndSortCond");
            case xlsb_type::BrtBookProtection:
                return fmt::format_to(ctx.out(), "BrtBookProtection");
            case xlsb_type::BrtSheetProtection:
                return fmt::format_to(ctx.out(), "BrtSheetProtection");
            case xlsb_type::BrtRangeProtection:
                return fmt::format_to(ctx.out(), "BrtRangeProtection");
            case xlsb_type::BrtPhoneticInfo:
                return fmt::format_to(ctx.out(), "BrtPhoneticInfo");
            case xlsb_type::BrtBeginECTxtWiz:
                return fmt::format_to(ctx.out(), "BrtBeginECTxtWiz");
            case xlsb_type::BrtEndECTxtWiz:
                return fmt::format_to(ctx.out(), "BrtEndECTxtWiz");
            case xlsb_type::BrtBeginECTWFldInfoLst:
                return fmt::format_to(ctx.out(), "BrtBeginECTWFldInfoLst");
            case xlsb_type::BrtEndECTWFldInfoLst:
                return fmt::format_to(ctx.out(), "BrtEndECTWFldInfoLst");
            case xlsb_type::BrtBeginECTwFldInfo:
                return fmt::format_to(ctx.out(), "BrtBeginECTwFldInfo");
            case xlsb_type::BrtFileSharing:
                return fmt::format_to(ctx.out(), "BrtFileSharing");
            case xlsb_type::BrtOleSize:
                return fmt::format_to(ctx.out(), "BrtOleSize");
            case xlsb_type::BrtDrawing:
                return fmt::format_to(ctx.out(), "BrtDrawing");
            case xlsb_type::BrtLegacyDrawing:
                return fmt::format_to(ctx.out(), "BrtLegacyDrawing");
            case xlsb_type::BrtLegacyDrawingHF:
                return fmt::format_to(ctx.out(), "BrtLegacyDrawingHF");
            case xlsb_type::BrtWebOpt:
                return fmt::format_to(ctx.out(), "BrtWebOpt");
            case xlsb_type::BrtBeginWebPubItems:
                return fmt::format_to(ctx.out(), "BrtBeginWebPubItems");
            case xlsb_type::BrtEndWebPubItems:
                return fmt::format_to(ctx.out(), "BrtEndWebPubItems");
            case xlsb_type::BrtBeginWebPubItem:
                return fmt::format_to(ctx.out(), "BrtBeginWebPubItem");
            case xlsb_type::BrtEndWebPubItem:
                return fmt::format_to(ctx.out(), "BrtEndWebPubItem");
            case xlsb_type::BrtBeginSXCondFmt:
                return fmt::format_to(ctx.out(), "BrtBeginSXCondFmt");
            case xlsb_type::BrtEndSXCondFmt:
                return fmt::format_to(ctx.out(), "BrtEndSXCondFmt");
            case xlsb_type::BrtBeginSXCondFmts:
                return fmt::format_to(ctx.out(), "BrtBeginSXCondFmts");
            case xlsb_type::BrtEndSXCondFmts:
                return fmt::format_to(ctx.out(), "BrtEndSXCondFmts");
            case xlsb_type::BrtBkHim:
                return fmt::format_to(ctx.out(), "BrtBkHim");
            case xlsb_type::BrtColor:
                return fmt::format_to(ctx.out(), "BrtColor");
            case xlsb_type::BrtBeginIndexedColors:
                return fmt::format_to(ctx.out(), "BrtBeginIndexedColors");
            case xlsb_type::BrtEndIndexedColors:
                return fmt::format_to(ctx.out(), "BrtEndIndexedColors");
            case xlsb_type::BrtBeginMRUColors:
                return fmt::format_to(ctx.out(), "BrtBeginMRUColors");
            case xlsb_type::BrtEndMRUColors:
                return fmt::format_to(ctx.out(), "BrtEndMRUColors");
            case xlsb_type::BrtMRUColor:
                return fmt::format_to(ctx.out(), "BrtMRUColor");
            case xlsb_type::BrtBeginDVals:
                return fmt::format_to(ctx.out(), "BrtBeginDVals");
            case xlsb_type::BrtEndDVals:
                return fmt::format_to(ctx.out(), "BrtEndDVals");
            case xlsb_type::BrtSupNameStart:
                return fmt::format_to(ctx.out(), "BrtSupNameStart");
            case xlsb_type::BrtSupNameValueStart:
                return fmt::format_to(ctx.out(), "BrtSupNameValueStart");
            case xlsb_type::BrtSupNameValueEnd:
                return fmt::format_to(ctx.out(), "BrtSupNameValueEnd");
            case xlsb_type::BrtSupNameNum:
                return fmt::format_to(ctx.out(), "BrtSupNameNum");
            case xlsb_type::BrtSupNameErr:
                return fmt::format_to(ctx.out(), "BrtSupNameErr");
            case xlsb_type::BrtSupNameSt:
                return fmt::format_to(ctx.out(), "BrtSupNameSt");
            case xlsb_type::BrtSupNameNil:
                return fmt::format_to(ctx.out(), "BrtSupNameNil");
            case xlsb_type::BrtSupNameBool:
                return fmt::format_to(ctx.out(), "BrtSupNameBool");
            case xlsb_type::BrtSupNameFmla:
                return fmt::format_to(ctx.out(), "BrtSupNameFmla");
            case xlsb_type::BrtSupNameBits:
                return fmt::format_to(ctx.out(), "BrtSupNameBits");
            case xlsb_type::BrtSupNameEnd:
                return fmt::format_to(ctx.out(), "BrtSupNameEnd");
            case xlsb_type::BrtEndSupBook:
                return fmt::format_to(ctx.out(), "BrtEndSupBook");
            case xlsb_type::BrtCellSmartTagProperty:
                return fmt::format_to(ctx.out(), "BrtCellSmartTagProperty");
            case xlsb_type::BrtBeginCellSmartTag:
                return fmt::format_to(ctx.out(), "BrtBeginCellSmartTag");
            case xlsb_type::BrtEndCellSmartTag:
                return fmt::format_to(ctx.out(), "BrtEndCellSmartTag");
            case xlsb_type::BrtBeginCellSmartTags:
                return fmt::format_to(ctx.out(), "BrtBeginCellSmartTags");
            case xlsb_type::BrtEndCellSmartTags:
                return fmt::format_to(ctx.out(), "BrtEndCellSmartTags");
            case xlsb_type::BrtBeginSmartTags:
                return fmt::format_to(ctx.out(), "BrtBeginSmartTags");
            case xlsb_type::BrtEndSmartTags:
                return fmt::format_to(ctx.out(), "BrtEndSmartTags");
            case xlsb_type::BrtSmartTagType:
                return fmt::format_to(ctx.out(), "BrtSmartTagType");
            case xlsb_type::BrtBeginSmartTagTypes:
                return fmt::format_to(ctx.out(), "BrtBeginSmartTagTypes");
            case xlsb_type::BrtEndSmartTagTypes:
                return fmt::format_to(ctx.out(), "BrtEndSmartTagTypes");
            case xlsb_type::BrtBeginSXFilters:
                return fmt::format_to(ctx.out(), "BrtBeginSXFilters");
            case xlsb_type::BrtEndSXFilters:
                return fmt::format_to(ctx.out(), "BrtEndSXFilters");
            case xlsb_type::BrtBeginSXFILTER:
                return fmt::format_to(ctx.out(), "BrtBeginSXFILTER");
            case xlsb_type::BrtEndSXFilter:
                return fmt::format_to(ctx.out(), "BrtEndSXFilter");
            case xlsb_type::BrtBeginFills:
                return fmt::format_to(ctx.out(), "BrtBeginFills");
            case xlsb_type::BrtEndFills:
                return fmt::format_to(ctx.out(), "BrtEndFills");
            case xlsb_type::BrtBeginCellWatches:
                return fmt::format_to(ctx.out(), "BrtBeginCellWatches");
            case xlsb_type::BrtEndCellWatches:
                return fmt::format_to(ctx.out(), "BrtEndCellWatches");
            case xlsb_type::BrtCellWatch:
                return fmt::format_to(ctx.out(), "BrtCellWatch");
            case xlsb_type::BrtBeginCRErrs:
                return fmt::format_to(ctx.out(), "BrtBeginCRErrs");
            case xlsb_type::BrtEndCRErrs:
                return fmt::format_to(ctx.out(), "BrtEndCRErrs");
            case xlsb_type::BrtCrashRecErr:
                return fmt::format_to(ctx.out(), "BrtCrashRecErr");
            case xlsb_type::BrtBeginFonts:
                return fmt::format_to(ctx.out(), "BrtBeginFonts");
            case xlsb_type::BrtEndFonts:
                return fmt::format_to(ctx.out(), "BrtEndFonts");
            case xlsb_type::BrtBeginBorders:
                return fmt::format_to(ctx.out(), "BrtBeginBorders");
            case xlsb_type::BrtEndBorders:
                return fmt::format_to(ctx.out(), "BrtEndBorders");
            case xlsb_type::BrtBeginFmts:
                return fmt::format_to(ctx.out(), "BrtBeginFmts");
            case xlsb_type::BrtEndFmts:
                return fmt::format_to(ctx.out(), "BrtEndFmts");
            case xlsb_type::BrtBeginCellXFs:
                return fmt::format_to(ctx.out(), "BrtBeginCellXFs");
            case xlsb_type::BrtEndCellXFs:
                return fmt::format_to(ctx.out(), "BrtEndCellXFs");
            case xlsb_type::BrtBeginStyles:
                return fmt::format_to(ctx.out(), "BrtBeginStyles");
            case xlsb_type::BrtEndStyles:
                return fmt::format_to(ctx.out(), "BrtEndStyles");
            case xlsb_type::BrtBigName:
                return fmt::format_to(ctx.out(), "BrtBigName");
            case xlsb_type::BrtBeginCellStyleXFs:
                return fmt::format_to(ctx.out(), "BrtBeginCellStyleXFs");
            case xlsb_type::BrtEndCellStyleXFs:
                return fmt::format_to(ctx.out(), "BrtEndCellStyleXFs");
            case xlsb_type::BrtBeginComments:
                return fmt::format_to(ctx.out(), "BrtBeginComments");
            case xlsb_type::BrtEndComments:
                return fmt::format_to(ctx.out(), "BrtEndComments");
            case xlsb_type::BrtBeginCommentAuthors:
                return fmt::format_to(ctx.out(), "BrtBeginCommentAuthors");
            case xlsb_type::BrtEndCommentAuthors:
                return fmt::format_to(ctx.out(), "BrtEndCommentAuthors");
            case xlsb_type::BrtCommentAuthor:
                return fmt::format_to(ctx.out(), "BrtCommentAuthor");
            case xlsb_type::BrtBeginCommentList:
                return fmt::format_to(ctx.out(), "BrtBeginCommentList");
            case xlsb_type::BrtEndCommentList:
                return fmt::format_to(ctx.out(), "BrtEndCommentList");
            case xlsb_type::BrtBeginComment:
                return fmt::format_to(ctx.out(), "BrtBeginComment");
            case xlsb_type::BrtEndComment:
                return fmt::format_to(ctx.out(), "BrtEndComment");
            case xlsb_type::BrtCommentText:
                return fmt::format_to(ctx.out(), "BrtCommentText");
            case xlsb_type::BrtBeginOleObjects:
                return fmt::format_to(ctx.out(), "BrtBeginOleObjects");
            case xlsb_type::BrtOleObject:
                return fmt::format_to(ctx.out(), "BrtOleObject");
            case xlsb_type::BrtEndOleObjects:
                return fmt::format_to(ctx.out(), "BrtEndOleObjects");
            case xlsb_type::BrtBeginSxrules:
                return fmt::format_to(ctx.out(), "BrtBeginSxrules");
            case xlsb_type::BrtEndSxRules:
                return fmt::format_to(ctx.out(), "BrtEndSxRules");
            case xlsb_type::BrtBeginActiveXControls:
                return fmt::format_to(ctx.out(), "BrtBeginActiveXControls");
            case xlsb_type::BrtActiveX:
                return fmt::format_to(ctx.out(), "BrtActiveX");
            case xlsb_type::BrtEndActiveXControls:
                return fmt::format_to(ctx.out(), "BrtEndActiveXControls");
            case xlsb_type::BrtBeginPCDSDTCEMembersSortBy:
                return fmt::format_to(ctx.out(), "BrtBeginPCDSDTCEMembersSortBy");
            case xlsb_type::BrtBeginCellIgnoreECs:
                return fmt::format_to(ctx.out(), "BrtBeginCellIgnoreECs");
            case xlsb_type::BrtCellIgnoreEC:
                return fmt::format_to(ctx.out(), "BrtCellIgnoreEC");
            case xlsb_type::BrtEndCellIgnoreECs:
                return fmt::format_to(ctx.out(), "BrtEndCellIgnoreECs");
            case xlsb_type::BrtCsProp:
                return fmt::format_to(ctx.out(), "BrtCsProp");
            case xlsb_type::BrtCsPageSetup:
                return fmt::format_to(ctx.out(), "BrtCsPageSetup");
            case xlsb_type::BrtBeginUserCsViews:
                return fmt::format_to(ctx.out(), "BrtBeginUserCsViews");
            case xlsb_type::BrtEndUserCsViews:
                return fmt::format_to(ctx.out(), "BrtEndUserCsViews");
            case xlsb_type::BrtBeginUserCsView:
                return fmt::format_to(ctx.out(), "BrtBeginUserCsView");
            case xlsb_type::BrtEndUserCsView:
                return fmt::format_to(ctx.out(), "BrtEndUserCsView");
            case xlsb_type::BrtBeginPcdSFCIEntries:
                return fmt::format_to(ctx.out(), "BrtBeginPcdSFCIEntries");
            case xlsb_type::BrtEndPCDSFCIEntries:
                return fmt::format_to(ctx.out(), "BrtEndPCDSFCIEntries");
            case xlsb_type::BrtPCDSFCIEntry:
                return fmt::format_to(ctx.out(), "BrtPCDSFCIEntry");
            case xlsb_type::BrtBeginListParts:
                return fmt::format_to(ctx.out(), "BrtBeginListParts");
            case xlsb_type::BrtListPart:
                return fmt::format_to(ctx.out(), "BrtListPart");
            case xlsb_type::BrtEndListParts:
                return fmt::format_to(ctx.out(), "BrtEndListParts");
            case xlsb_type::BrtSheetCalcProp:
                return fmt::format_to(ctx.out(), "BrtSheetCalcProp");
            case xlsb_type::BrtBeginFnGroup:
                return fmt::format_to(ctx.out(), "BrtBeginFnGroup");
            case xlsb_type::BrtFnGroup:
                return fmt::format_to(ctx.out(), "BrtFnGroup");
            case xlsb_type::BrtEndFnGroup:
                return fmt::format_to(ctx.out(), "BrtEndFnGroup");
            case xlsb_type::BrtSupAddin:
                return fmt::format_to(ctx.out(), "BrtSupAddin");
            case xlsb_type::BrtSXTDMPOrder:
                return fmt::format_to(ctx.out(), "BrtSXTDMPOrder");
            case xlsb_type::BrtCsProtection:
                return fmt::format_to(ctx.out(), "BrtCsProtection");
            case xlsb_type::BrtBeginWsSortMap:
                return fmt::format_to(ctx.out(), "BrtBeginWsSortMap");
            case xlsb_type::BrtEndWsSortMap:
                return fmt::format_to(ctx.out(), "BrtEndWsSortMap");
            case xlsb_type::BrtBeginRRSort:
                return fmt::format_to(ctx.out(), "BrtBeginRRSort");
            case xlsb_type::BrtEndRRSort:
                return fmt::format_to(ctx.out(), "BrtEndRRSort");
            case xlsb_type::BrtRRSortItem:
                return fmt::format_to(ctx.out(), "BrtRRSortItem");
            case xlsb_type::BrtFileSharingIso:
                return fmt::format_to(ctx.out(), "BrtFileSharingIso");
            case xlsb_type::BrtBookProtectionIso:
                return fmt::format_to(ctx.out(), "BrtBookProtectionIso");
            case xlsb_type::BrtSheetProtectionIso:
                return fmt::format_to(ctx.out(), "BrtSheetProtectionIso");
            case xlsb_type::BrtCsProtectionIso:
                return fmt::format_to(ctx.out(), "BrtCsProtectionIso");
            case xlsb_type::BrtRangeProtectionIso:
                return fmt::format_to(ctx.out(), "BrtRangeProtectionIso");
            case xlsb_type::BrtDValList:
                return fmt::format_to(ctx.out(), "BrtDValList");
            case xlsb_type::BrtRwDescent:
                return fmt::format_to(ctx.out(), "BrtRwDescent");
            case xlsb_type::BrtKnownFonts:
                return fmt::format_to(ctx.out(), "BrtKnownFonts");
            case xlsb_type::BrtBeginSXTupleSet:
                return fmt::format_to(ctx.out(), "BrtBeginSXTupleSet");
            case xlsb_type::BrtEndSXTupleSet:
                return fmt::format_to(ctx.out(), "BrtEndSXTupleSet");
            case xlsb_type::BrtBeginSXTupleSetHeader:
                return fmt::format_to(ctx.out(), "BrtBeginSXTupleSetHeader");
            case xlsb_type::BrtEndSXTupleSetHeader:
                return fmt::format_to(ctx.out(), "BrtEndSXTupleSetHeader");
            case xlsb_type::BrtSXTupleSetHeaderItem:
                return fmt::format_to(ctx.out(), "BrtSXTupleSetHeaderItem");
            case xlsb_type::BrtBeginSXTupleSetData:
                return fmt::format_to(ctx.out(), "BrtBeginSXTupleSetData");
            case xlsb_type::BrtEndSXTupleSetData:
                return fmt::format_to(ctx.out(), "BrtEndSXTupleSetData");
            case xlsb_type::BrtBeginSXTupleSetRow:
                return fmt::format_to(ctx.out(), "BrtBeginSXTupleSetRow");
            case xlsb_type::BrtEndSXTupleSetRow:
                return fmt::format_to(ctx.out(), "BrtEndSXTupleSetRow");
            case xlsb_type::BrtSXTupleSetRowItem:
                return fmt::format_to(ctx.out(), "BrtSXTupleSetRowItem");
            case xlsb_type::BrtNameExt:
                return fmt::format_to(ctx.out(), "BrtNameExt");
            case xlsb_type::BrtPCDH14:
                return fmt::format_to(ctx.out(), "BrtPCDH14");
            case xlsb_type::BrtBeginPCDCalcMem14:
                return fmt::format_to(ctx.out(), "BrtBeginPCDCalcMem14");
            case xlsb_type::BrtEndPCDCalcMem14:
                return fmt::format_to(ctx.out(), "BrtEndPCDCalcMem14");
            case xlsb_type::BrtSXTH14:
                return fmt::format_to(ctx.out(), "BrtSXTH14");
            case xlsb_type::BrtBeginSparklineGroup:
                return fmt::format_to(ctx.out(), "BrtBeginSparklineGroup");
            case xlsb_type::BrtEndSparklineGroup:
                return fmt::format_to(ctx.out(), "BrtEndSparklineGroup");
            case xlsb_type::BrtSparkline:
                return fmt::format_to(ctx.out(), "BrtSparkline");
            case xlsb_type::BrtSXDI14:
                return fmt::format_to(ctx.out(), "BrtSXDI14");
            case xlsb_type::BrtWsFmtInfoEx14:
                return fmt::format_to(ctx.out(), "BrtWsFmtInfoEx14");
            case xlsb_type::BrtBeginConditionalFormatting14:
                return fmt::format_to(ctx.out(), "BrtBeginConditionalFormatting14");
            case xlsb_type::BrtEndConditionalFormatting14:
                return fmt::format_to(ctx.out(), "BrtEndConditionalFormatting14");
            case xlsb_type::BrtBeginCFRule14:
                return fmt::format_to(ctx.out(), "BrtBeginCFRule14");
            case xlsb_type::BrtEndCFRule14:
                return fmt::format_to(ctx.out(), "BrtEndCFRule14");
            case xlsb_type::BrtCFVO14:
                return fmt::format_to(ctx.out(), "BrtCFVO14");
            case xlsb_type::BrtBeginDatabar14:
                return fmt::format_to(ctx.out(), "BrtBeginDatabar14");
            case xlsb_type::BrtBeginIconSet14:
                return fmt::format_to(ctx.out(), "BrtBeginIconSet14");
            case xlsb_type::BrtDVal14:
                return fmt::format_to(ctx.out(), "BrtDVal14");
            case xlsb_type::BrtBeginDVals14:
                return fmt::format_to(ctx.out(), "BrtBeginDVals14");
            case xlsb_type::BrtColor14:
                return fmt::format_to(ctx.out(), "BrtColor14");
            case xlsb_type::BrtBeginSparklines:
                return fmt::format_to(ctx.out(), "BrtBeginSparklines");
            case xlsb_type::BrtEndSparklines:
                return fmt::format_to(ctx.out(), "BrtEndSparklines");
            case xlsb_type::BrtBeginSparklineGroups:
                return fmt::format_to(ctx.out(), "BrtBeginSparklineGroups");
            case xlsb_type::BrtEndSparklineGroups:
                return fmt::format_to(ctx.out(), "BrtEndSparklineGroups");
            case xlsb_type::BrtSXVD14:
                return fmt::format_to(ctx.out(), "BrtSXVD14");
            case xlsb_type::BrtBeginSxView14:
                return fmt::format_to(ctx.out(), "BrtBeginSxView14");
            case xlsb_type::BrtEndSxView14:
                return fmt::format_to(ctx.out(), "BrtEndSxView14");
            case xlsb_type::BrtBeginSXView16:
                return fmt::format_to(ctx.out(), "BrtBeginSXView16");
            case xlsb_type::BrtEndSXView16:
                return fmt::format_to(ctx.out(), "BrtEndSXView16");
            case xlsb_type::BrtBeginPCD14:
                return fmt::format_to(ctx.out(), "BrtBeginPCD14");
            case xlsb_type::BrtEndPCD14:
                return fmt::format_to(ctx.out(), "BrtEndPCD14");
            case xlsb_type::BrtBeginExtConn14:
                return fmt::format_to(ctx.out(), "BrtBeginExtConn14");
            case xlsb_type::BrtEndExtConn14:
                return fmt::format_to(ctx.out(), "BrtEndExtConn14");
            case xlsb_type::BrtBeginSlicerCacheIDs:
                return fmt::format_to(ctx.out(), "BrtBeginSlicerCacheIDs");
            case xlsb_type::BrtEndSlicerCacheIDs:
                return fmt::format_to(ctx.out(), "BrtEndSlicerCacheIDs");
            case xlsb_type::BrtBeginSlicerCacheID:
                return fmt::format_to(ctx.out(), "BrtBeginSlicerCacheID");
            case xlsb_type::BrtEndSlicerCacheID:
                return fmt::format_to(ctx.out(), "BrtEndSlicerCacheID");
            case xlsb_type::BrtBeginSlicerCache:
                return fmt::format_to(ctx.out(), "BrtBeginSlicerCache");
            case xlsb_type::BrtEndSlicerCache:
                return fmt::format_to(ctx.out(), "BrtEndSlicerCache");
            case xlsb_type::BrtBeginSlicerCacheDef:
                return fmt::format_to(ctx.out(), "BrtBeginSlicerCacheDef");
            case xlsb_type::BrtEndSlicerCacheDef:
                return fmt::format_to(ctx.out(), "BrtEndSlicerCacheDef");
            case xlsb_type::BrtBeginSlicersEx:
                return fmt::format_to(ctx.out(), "BrtBeginSlicersEx");
            case xlsb_type::BrtEndSlicersEx:
                return fmt::format_to(ctx.out(), "BrtEndSlicersEx");
            case xlsb_type::BrtBeginSlicerEx:
                return fmt::format_to(ctx.out(), "BrtBeginSlicerEx");
            case xlsb_type::BrtEndSlicerEx:
                return fmt::format_to(ctx.out(), "BrtEndSlicerEx");
            case xlsb_type::BrtBeginSlicer:
                return fmt::format_to(ctx.out(), "BrtBeginSlicer");
            case xlsb_type::BrtEndSlicer:
                return fmt::format_to(ctx.out(), "BrtEndSlicer");
            case xlsb_type::BrtSlicerCachePivotTables:
                return fmt::format_to(ctx.out(), "BrtSlicerCachePivotTables");
            case xlsb_type::BrtBeginSlicerCacheOlapImpl:
                return fmt::format_to(ctx.out(), "BrtBeginSlicerCacheOlapImpl");
            case xlsb_type::BrtEndSlicerCacheOlapImpl:
                return fmt::format_to(ctx.out(), "BrtEndSlicerCacheOlapImpl");
            case xlsb_type::BrtBeginSlicerCacheLevelsData:
                return fmt::format_to(ctx.out(), "BrtBeginSlicerCacheLevelsData");
            case xlsb_type::BrtEndSlicerCacheLevelsData:
                return fmt::format_to(ctx.out(), "BrtEndSlicerCacheLevelsData");
            case xlsb_type::BrtBeginSlicerCacheLevelData:
                return fmt::format_to(ctx.out(), "BrtBeginSlicerCacheLevelData");
            case xlsb_type::BrtEndSlicerCacheLevelData:
                return fmt::format_to(ctx.out(), "BrtEndSlicerCacheLevelData");
            case xlsb_type::BrtBeginSlicerCacheSiRanges:
                return fmt::format_to(ctx.out(), "BrtBeginSlicerCacheSiRanges");
            case xlsb_type::BrtEndSlicerCacheSiRanges:
                return fmt::format_to(ctx.out(), "BrtEndSlicerCacheSiRanges");
            case xlsb_type::BrtBeginSlicerCacheSiRange:
                return fmt::format_to(ctx.out(), "BrtBeginSlicerCacheSiRange");
            case xlsb_type::BrtEndSlicerCacheSiRange:
                return fmt::format_to(ctx.out(), "BrtEndSlicerCacheSiRange");
            case xlsb_type::BrtSlicerCacheOlapItem:
                return fmt::format_to(ctx.out(), "BrtSlicerCacheOlapItem");
            case xlsb_type::BrtBeginSlicerCacheSelections:
                return fmt::format_to(ctx.out(), "BrtBeginSlicerCacheSelections");
            case xlsb_type::BrtSlicerCacheSelection:
                return fmt::format_to(ctx.out(), "BrtSlicerCacheSelection");
            case xlsb_type::BrtEndSlicerCacheSelections:
                return fmt::format_to(ctx.out(), "BrtEndSlicerCacheSelections");
            case xlsb_type::BrtBeginSlicerCacheNative:
                return fmt::format_to(ctx.out(), "BrtBeginSlicerCacheNative");
            case xlsb_type::BrtEndSlicerCacheNative:
                return fmt::format_to(ctx.out(), "BrtEndSlicerCacheNative");
            case xlsb_type::BrtSlicerCacheNativeItem:
                return fmt::format_to(ctx.out(), "BrtSlicerCacheNativeItem");
            case xlsb_type::BrtRangeProtection14:
                return fmt::format_to(ctx.out(), "BrtRangeProtection14");
            case xlsb_type::BrtRangeProtectionIso14:
                return fmt::format_to(ctx.out(), "BrtRangeProtectionIso14");
            case xlsb_type::BrtCellIgnoreEC14:
                return fmt::format_to(ctx.out(), "BrtCellIgnoreEC14");
            case xlsb_type::BrtList14:
                return fmt::format_to(ctx.out(), "BrtList14");
            case xlsb_type::BrtCFIcon:
                return fmt::format_to(ctx.out(), "BrtCFIcon");
            case xlsb_type::BrtBeginSlicerCachesPivotCacheIDs:
                return fmt::format_to(ctx.out(), "BrtBeginSlicerCachesPivotCacheIDs");
            case xlsb_type::BrtEndSlicerCachesPivotCacheIDs:
                return fmt::format_to(ctx.out(), "BrtEndSlicerCachesPivotCacheIDs");
            case xlsb_type::BrtBeginSlicers:
                return fmt::format_to(ctx.out(), "BrtBeginSlicers");
            case xlsb_type::BrtEndSlicers:
                return fmt::format_to(ctx.out(), "BrtEndSlicers");
            case xlsb_type::BrtWbProp14:
                return fmt::format_to(ctx.out(), "BrtWbProp14");
            case xlsb_type::BrtBeginSXEdit:
                return fmt::format_to(ctx.out(), "BrtBeginSXEdit");
            case xlsb_type::BrtEndSXEdit:
                return fmt::format_to(ctx.out(), "BrtEndSXEdit");
            case xlsb_type::BrtBeginSXEdits:
                return fmt::format_to(ctx.out(), "BrtBeginSXEdits");
            case xlsb_type::BrtEndSXEdits:
                return fmt::format_to(ctx.out(), "BrtEndSXEdits");
            case xlsb_type::BrtBeginSXChange:
                return fmt::format_to(ctx.out(), "BrtBeginSXChange");
            case xlsb_type::BrtEndSXChange:
                return fmt::format_to(ctx.out(), "BrtEndSXChange");
            case xlsb_type::BrtBeginSXChanges:
                return fmt::format_to(ctx.out(), "BrtBeginSXChanges");
            case xlsb_type::BrtEndSXChanges:
                return fmt::format_to(ctx.out(), "BrtEndSXChanges");
            case xlsb_type::BrtSXTupleItems:
                return fmt::format_to(ctx.out(), "BrtSXTupleItems");
            case xlsb_type::BrtBeginSlicerStyle:
                return fmt::format_to(ctx.out(), "BrtBeginSlicerStyle");
            case xlsb_type::BrtEndSlicerStyle:
                return fmt::format_to(ctx.out(), "BrtEndSlicerStyle");
            case xlsb_type::BrtSlicerStyleElement:
                return fmt::format_to(ctx.out(), "BrtSlicerStyleElement");
            case xlsb_type::BrtBeginStyleSheetExt14:
                return fmt::format_to(ctx.out(), "BrtBeginStyleSheetExt14");
            case xlsb_type::BrtEndStyleSheetExt14:
                return fmt::format_to(ctx.out(), "BrtEndStyleSheetExt14");
            case xlsb_type::BrtBeginSlicerCachesPivotCacheID:
                return fmt::format_to(ctx.out(), "BrtBeginSlicerCachesPivotCacheID");
            case xlsb_type::BrtEndSlicerCachesPivotCacheID:
                return fmt::format_to(ctx.out(), "BrtEndSlicerCachesPivotCacheID");
            case xlsb_type::BrtBeginConditionalFormattings:
                return fmt::format_to(ctx.out(), "BrtBeginConditionalFormattings");
            case xlsb_type::BrtEndConditionalFormattings:
                return fmt::format_to(ctx.out(), "BrtEndConditionalFormattings");
            case xlsb_type::BrtBeginPCDCalcMemExt:
                return fmt::format_to(ctx.out(), "BrtBeginPCDCalcMemExt");
            case xlsb_type::BrtEndPCDCalcMemExt:
                return fmt::format_to(ctx.out(), "BrtEndPCDCalcMemExt");
            case xlsb_type::BrtBeginPCDCalcMemsExt:
                return fmt::format_to(ctx.out(), "BrtBeginPCDCalcMemsExt");
            case xlsb_type::BrtEndPCDCalcMemsExt:
                return fmt::format_to(ctx.out(), "BrtEndPCDCalcMemsExt");
            case xlsb_type::BrtPCDField14:
                return fmt::format_to(ctx.out(), "BrtPCDField14");
            case xlsb_type::BrtBeginSlicerStyles:
                return fmt::format_to(ctx.out(), "BrtBeginSlicerStyles");
            case xlsb_type::BrtEndSlicerStyles:
                return fmt::format_to(ctx.out(), "BrtEndSlicerStyles");
            case xlsb_type::BrtBeginSlicerStyleElements:
                return fmt::format_to(ctx.out(), "BrtBeginSlicerStyleElements");
            case xlsb_type::BrtEndSlicerStyleElements:
                return fmt::format_to(ctx.out(), "BrtEndSlicerStyleElements");
            case xlsb_type::BrtCFRuleExt:
                return fmt::format_to(ctx.out(), "BrtCFRuleExt");
            case xlsb_type::BrtBeginSXCondFmt14:
                return fmt::format_to(ctx.out(), "BrtBeginSXCondFmt14");
            case xlsb_type::BrtEndSXCondFmt14:
                return fmt::format_to(ctx.out(), "BrtEndSXCondFmt14");
            case xlsb_type::BrtBeginSXCondFmts14:
                return fmt::format_to(ctx.out(), "BrtBeginSXCondFmts14");
            case xlsb_type::BrtEndSXCondFmts14:
                return fmt::format_to(ctx.out(), "BrtEndSXCondFmts14");
            case xlsb_type::BrtBeginSortCond14:
                return fmt::format_to(ctx.out(), "BrtBeginSortCond14");
            case xlsb_type::BrtEndSortCond14:
                return fmt::format_to(ctx.out(), "BrtEndSortCond14");
            case xlsb_type::BrtEndDVals14:
                return fmt::format_to(ctx.out(), "BrtEndDVals14");
            case xlsb_type::BrtEndIconSet14:
                return fmt::format_to(ctx.out(), "BrtEndIconSet14");
            case xlsb_type::BrtEndDatabar14:
                return fmt::format_to(ctx.out(), "BrtEndDatabar14");
            case xlsb_type::BrtBeginColorScale14:
                return fmt::format_to(ctx.out(), "BrtBeginColorScale14");
            case xlsb_type::BrtEndColorScale14:
                return fmt::format_to(ctx.out(), "BrtEndColorScale14");
            case xlsb_type::BrtBeginSxrules14:
                return fmt::format_to(ctx.out(), "BrtBeginSxrules14");
            case xlsb_type::BrtEndSxrules14:
                return fmt::format_to(ctx.out(), "BrtEndSxrules14");
            case xlsb_type::BrtBeginPRule14:
                return fmt::format_to(ctx.out(), "BrtBeginPRule14");
            case xlsb_type::BrtEndPRule14:
                return fmt::format_to(ctx.out(), "BrtEndPRule14");
            case xlsb_type::BrtBeginPRFilters14:
                return fmt::format_to(ctx.out(), "BrtBeginPRFilters14");
            case xlsb_type::BrtEndPRFilters14:
                return fmt::format_to(ctx.out(), "BrtEndPRFilters14");
            case xlsb_type::BrtBeginPRFilter14:
                return fmt::format_to(ctx.out(), "BrtBeginPRFilter14");
            case xlsb_type::BrtEndPRFilter14:
                return fmt::format_to(ctx.out(), "BrtEndPRFilter14");
            case xlsb_type::BrtBeginPRFItem14:
                return fmt::format_to(ctx.out(), "BrtBeginPRFItem14");
            case xlsb_type::BrtEndPRFItem14:
                return fmt::format_to(ctx.out(), "BrtEndPRFItem14");
            case xlsb_type::BrtBeginCellIgnoreECs14:
                return fmt::format_to(ctx.out(), "BrtBeginCellIgnoreECs14");
            case xlsb_type::BrtEndCellIgnoreECs14:
                return fmt::format_to(ctx.out(), "BrtEndCellIgnoreECs14");
            case xlsb_type::BrtDxf14:
                return fmt::format_to(ctx.out(), "BrtDxf14");
            case xlsb_type::BrtBeginDxF14s:
                return fmt::format_to(ctx.out(), "BrtBeginDxF14s");
            case xlsb_type::BrtEndDxf14s:
                return fmt::format_to(ctx.out(), "BrtEndDxf14s");
            case xlsb_type::BrtFilter14:
                return fmt::format_to(ctx.out(), "BrtFilter14");
            case xlsb_type::BrtBeginCustomFilters14:
                return fmt::format_to(ctx.out(), "BrtBeginCustomFilters14");
            case xlsb_type::BrtCustomFilter14:
                return fmt::format_to(ctx.out(), "BrtCustomFilter14");
            case xlsb_type::BrtIconFilter14:
                return fmt::format_to(ctx.out(), "BrtIconFilter14");
            case xlsb_type::BrtPivotCacheConnectionName:
                return fmt::format_to(ctx.out(), "BrtPivotCacheConnectionName");
            case xlsb_type::BrtBeginDecoupledPivotCacheIDs:
                return fmt::format_to(ctx.out(), "BrtBeginDecoupledPivotCacheIDs");
            case xlsb_type::BrtEndDecoupledPivotCacheIDs:
                return fmt::format_to(ctx.out(), "BrtEndDecoupledPivotCacheIDs");
            case xlsb_type::BrtDecoupledPivotCacheID:
                return fmt::format_to(ctx.out(), "BrtDecoupledPivotCacheID");
            case xlsb_type::BrtBeginPivotTableRefs:
                return fmt::format_to(ctx.out(), "BrtBeginPivotTableRefs");
            case xlsb_type::BrtEndPivotTableRefs:
                return fmt::format_to(ctx.out(), "BrtEndPivotTableRefs");
            case xlsb_type::BrtPivotTableRef:
                return fmt::format_to(ctx.out(), "BrtPivotTableRef");
            case xlsb_type::BrtSlicerCacheBookPivotTables:
                return fmt::format_to(ctx.out(), "BrtSlicerCacheBookPivotTables");
            case xlsb_type::BrtBeginSxvcells:
                return fmt::format_to(ctx.out(), "BrtBeginSxvcells");
            case xlsb_type::BrtEndSxvcells:
                return fmt::format_to(ctx.out(), "BrtEndSxvcells");
            case xlsb_type::BrtBeginSxRow:
                return fmt::format_to(ctx.out(), "BrtBeginSxRow");
            case xlsb_type::BrtEndSxRow:
                return fmt::format_to(ctx.out(), "BrtEndSxRow");
            case xlsb_type::BrtPcdCalcMem15:
                return fmt::format_to(ctx.out(), "BrtPcdCalcMem15");
            case xlsb_type::BrtQsi15:
                return fmt::format_to(ctx.out(), "BrtQsi15");
            case xlsb_type::BrtBeginWebExtensions:
                return fmt::format_to(ctx.out(), "BrtBeginWebExtensions");
            case xlsb_type::BrtEndWebExtensions:
                return fmt::format_to(ctx.out(), "BrtEndWebExtensions");
            case xlsb_type::BrtWebExtension:
                return fmt::format_to(ctx.out(), "BrtWebExtension");
            case xlsb_type::BrtAbsPath15:
                return fmt::format_to(ctx.out(), "BrtAbsPath15");
            case xlsb_type::BrtBeginPivotTableUISettings:
                return fmt::format_to(ctx.out(), "BrtBeginPivotTableUISettings");
            case xlsb_type::BrtEndPivotTableUISettings:
                return fmt::format_to(ctx.out(), "BrtEndPivotTableUISettings");
            case xlsb_type::BrtTableSlicerCacheIDs:
                return fmt::format_to(ctx.out(), "BrtTableSlicerCacheIDs");
            case xlsb_type::BrtTableSlicerCacheID:
                return fmt::format_to(ctx.out(), "BrtTableSlicerCacheID");
            case xlsb_type::BrtBeginTableSlicerCache:
                return fmt::format_to(ctx.out(), "BrtBeginTableSlicerCache");
            case xlsb_type::BrtEndTableSlicerCache:
                return fmt::format_to(ctx.out(), "BrtEndTableSlicerCache");
            case xlsb_type::BrtSxFilter15:
                return fmt::format_to(ctx.out(), "BrtSxFilter15");
            case xlsb_type::BrtBeginTimelineCachePivotCacheIDs:
                return fmt::format_to(ctx.out(), "BrtBeginTimelineCachePivotCacheIDs");
            case xlsb_type::BrtEndTimelineCachePivotCacheIDs:
                return fmt::format_to(ctx.out(), "BrtEndTimelineCachePivotCacheIDs");
            case xlsb_type::BrtTimelineCachePivotCacheID:
                return fmt::format_to(ctx.out(), "BrtTimelineCachePivotCacheID");
            case xlsb_type::BrtBeginTimelineCacheIDs:
                return fmt::format_to(ctx.out(), "BrtBeginTimelineCacheIDs");
            case xlsb_type::BrtEndTimelineCacheIDs:
                return fmt::format_to(ctx.out(), "BrtEndTimelineCacheIDs");
            case xlsb_type::BrtBeginTimelineCacheID:
                return fmt::format_to(ctx.out(), "BrtBeginTimelineCacheID");
            case xlsb_type::BrtEndTimelineCacheID:
                return fmt::format_to(ctx.out(), "BrtEndTimelineCacheID");
            case xlsb_type::BrtBeginTimelinesEx:
                return fmt::format_to(ctx.out(), "BrtBeginTimelinesEx");
            case xlsb_type::BrtEndTimelinesEx:
                return fmt::format_to(ctx.out(), "BrtEndTimelinesEx");
            case xlsb_type::BrtBeginTimelineEx:
                return fmt::format_to(ctx.out(), "BrtBeginTimelineEx");
            case xlsb_type::BrtEndTimelineEx:
                return fmt::format_to(ctx.out(), "BrtEndTimelineEx");
            case xlsb_type::BrtWorkBookPr15:
                return fmt::format_to(ctx.out(), "BrtWorkBookPr15");
            case xlsb_type::BrtPCDH15:
                return fmt::format_to(ctx.out(), "BrtPCDH15");
            case xlsb_type::BrtBeginTimelineStyle:
                return fmt::format_to(ctx.out(), "BrtBeginTimelineStyle");
            case xlsb_type::BrtEndTimelineStyle:
                return fmt::format_to(ctx.out(), "BrtEndTimelineStyle");
            case xlsb_type::BrtTimelineStyleElement:
                return fmt::format_to(ctx.out(), "BrtTimelineStyleElement");
            case xlsb_type::BrtBeginTimelineStylesheetExt15:
                return fmt::format_to(ctx.out(), "BrtBeginTimelineStylesheetExt15");
            case xlsb_type::BrtEndTimelineStylesheetExt15:
                return fmt::format_to(ctx.out(), "BrtEndTimelineStylesheetExt15");
            case xlsb_type::BrtBeginTimelineStyles:
                return fmt::format_to(ctx.out(), "BrtBeginTimelineStyles");
            case xlsb_type::BrtEndTimelineStyles:
                return fmt::format_to(ctx.out(), "BrtEndTimelineStyles");
            case xlsb_type::BrtBeginTimelineStyleElements:
                return fmt::format_to(ctx.out(), "BrtBeginTimelineStyleElements");
            case xlsb_type::BrtEndTimelineStyleElements:
                return fmt::format_to(ctx.out(), "BrtEndTimelineStyleElements");
            case xlsb_type::BrtDxf15:
                return fmt::format_to(ctx.out(), "BrtDxf15");
            case xlsb_type::BrtBeginDxfs15:
                return fmt::format_to(ctx.out(), "BrtBeginDxfs15");
            case xlsb_type::BrtEndDXFs15:
                return fmt::format_to(ctx.out(), "BrtEndDXFs15");
            case xlsb_type::BrtSlicerCacheHideItemsWithNoData:
                return fmt::format_to(ctx.out(), "BrtSlicerCacheHideItemsWithNoData");
            case xlsb_type::BrtBeginItemUniqueNames:
                return fmt::format_to(ctx.out(), "BrtBeginItemUniqueNames");
            case xlsb_type::BrtEndItemUniqueNames:
                return fmt::format_to(ctx.out(), "BrtEndItemUniqueNames");
            case xlsb_type::BrtItemUniqueName:
                return fmt::format_to(ctx.out(), "BrtItemUniqueName");
            case xlsb_type::BrtBeginExtConn15:
                return fmt::format_to(ctx.out(), "BrtBeginExtConn15");
            case xlsb_type::BrtEndExtConn15:
                return fmt::format_to(ctx.out(), "BrtEndExtConn15");
            case xlsb_type::BrtBeginOledbPr15:
                return fmt::format_to(ctx.out(), "BrtBeginOledbPr15");
            case xlsb_type::BrtEndOledbPr15:
                return fmt::format_to(ctx.out(), "BrtEndOledbPr15");
            case xlsb_type::BrtBeginDataFeedPr15:
                return fmt::format_to(ctx.out(), "BrtBeginDataFeedPr15");
            case xlsb_type::BrtEndDataFeedPr15:
                return fmt::format_to(ctx.out(), "BrtEndDataFeedPr15");
            case xlsb_type::BrtTextPr15:
                return fmt::format_to(ctx.out(), "BrtTextPr15");
            case xlsb_type::BrtRangePr15:
                return fmt::format_to(ctx.out(), "BrtRangePr15");
            case xlsb_type::BrtDbCommand15:
                return fmt::format_to(ctx.out(), "BrtDbCommand15");
            case xlsb_type::BrtBeginDbTables15:
                return fmt::format_to(ctx.out(), "BrtBeginDbTables15");
            case xlsb_type::BrtEndDbTables15:
                return fmt::format_to(ctx.out(), "BrtEndDbTables15");
            case xlsb_type::BrtDbTable15:
                return fmt::format_to(ctx.out(), "BrtDbTable15");
            case xlsb_type::BrtBeginDataModel:
                return fmt::format_to(ctx.out(), "BrtBeginDataModel");
            case xlsb_type::BrtEndDataModel:
                return fmt::format_to(ctx.out(), "BrtEndDataModel");
            case xlsb_type::BrtBeginModelTables:
                return fmt::format_to(ctx.out(), "BrtBeginModelTables");
            case xlsb_type::BrtEndModelTables:
                return fmt::format_to(ctx.out(), "BrtEndModelTables");
            case xlsb_type::BrtModelTable:
                return fmt::format_to(ctx.out(), "BrtModelTable");
            case xlsb_type::BrtBeginModelRelationships:
                return fmt::format_to(ctx.out(), "BrtBeginModelRelationships");
            case xlsb_type::BrtEndModelRelationships:
                return fmt::format_to(ctx.out(), "BrtEndModelRelationships");
            case xlsb_type::BrtModelRelationship:
                return fmt::format_to(ctx.out(), "BrtModelRelationship");
            case xlsb_type::BrtBeginECTxtWiz15:
                return fmt::format_to(ctx.out(), "BrtBeginECTxtWiz15");
            case xlsb_type::BrtEndECTxtWiz15:
                return fmt::format_to(ctx.out(), "BrtEndECTxtWiz15");
            case xlsb_type::BrtBeginECTWFldInfoLst15:
                return fmt::format_to(ctx.out(), "BrtBeginECTWFldInfoLst15");
            case xlsb_type::BrtEndECTWFldInfoLst15:
                return fmt::format_to(ctx.out(), "BrtEndECTWFldInfoLst15");
            case xlsb_type::BrtBeginECTWFldInfo15:
                return fmt::format_to(ctx.out(), "BrtBeginECTWFldInfo15");
            case xlsb_type::BrtFieldListActiveItem:
                return fmt::format_to(ctx.out(), "BrtFieldListActiveItem");
            case xlsb_type::BrtPivotCacheIdVersion:
                return fmt::format_to(ctx.out(), "BrtPivotCacheIdVersion");
            case xlsb_type::BrtSXDI15:
                return fmt::format_to(ctx.out(), "BrtSXDI15");
            case xlsb_type::brtBeginModelTimeGroupings:
                return fmt::format_to(ctx.out(), "brtBeginModelTimeGroupings");
            case xlsb_type::brtEndModelTimeGroupings:
                return fmt::format_to(ctx.out(), "brtEndModelTimeGroupings");
            case xlsb_type::brtBeginModelTimeGrouping:
                return fmt::format_to(ctx.out(), "brtBeginModelTimeGrouping");
            case xlsb_type::brtEndModelTimeGrouping:
                return fmt::format_to(ctx.out(), "brtEndModelTimeGrouping");
            case xlsb_type::brtModelTimeGroupingCalcCol:
                return fmt::format_to(ctx.out(), "brtModelTimeGroupingCalcCol");
            case xlsb_type::brtRevisionPtr:
                return fmt::format_to(ctx.out(), "brtRevisionPtr");
            case xlsb_type::BrtBeginDynamicArrayPr:
                return fmt::format_to(ctx.out(), "BrtBeginDynamicArrayPr");
            case xlsb_type::BrtEndDynamicArrayPr:
                return fmt::format_to(ctx.out(), "BrtEndDynamicArrayPr");
            case xlsb_type::BrtBeginRichValueBlock:
                return fmt::format_to(ctx.out(), "BrtBeginRichValueBlock");
            case xlsb_type::BrtEndRichValueBlock:
                return fmt::format_to(ctx.out(), "BrtEndRichValueBlock");
            case xlsb_type::BrtBeginRichFilters:
                return fmt::format_to(ctx.out(), "BrtBeginRichFilters");
            case xlsb_type::BrtEndRichFilters:
                return fmt::format_to(ctx.out(), "BrtEndRichFilters");
            case xlsb_type::BrtRichFilter:
                return fmt::format_to(ctx.out(), "BrtRichFilter");
            case xlsb_type::BrtBeginRichFilterColumn:
                return fmt::format_to(ctx.out(), "BrtBeginRichFilterColumn");
            case xlsb_type::BrtEndRichFilterColumn:
                return fmt::format_to(ctx.out(), "BrtEndRichFilterColumn");
            case xlsb_type::BrtBeginCustomRichFilters:
                return fmt::format_to(ctx.out(), "BrtBeginCustomRichFilters");
            case xlsb_type::BrtEndCustomRichFilters:
                return fmt::format_to(ctx.out(), "BrtEndCustomRichFilters");
            case xlsb_type::BRTCustomRichFilter:
                return fmt::format_to(ctx.out(), "BRTCustomRichFilter");
            case xlsb_type::BrtTop10RichFilter:
                return fmt::format_to(ctx.out(), "BrtTop10RichFilter");
            case xlsb_type::BrtDynamicRichFilter:
                return fmt::format_to(ctx.out(), "BrtDynamicRichFilter");
            case xlsb_type::BrtBeginRichSortCondition:
                return fmt::format_to(ctx.out(), "BrtBeginRichSortCondition");
            case xlsb_type::BrtEndRichSortCondition:
                return fmt::format_to(ctx.out(), "BrtEndRichSortCondition");
            case xlsb_type::BrtRichFilterDateGroupItem:
                return fmt::format_to(ctx.out(), "BrtRichFilterDateGroupItem");
            case xlsb_type::BrtBeginCalcFeatures:
                return fmt::format_to(ctx.out(), "BrtBeginCalcFeatures");
            case xlsb_type::BrtEndCalcFeatures:
                return fmt::format_to(ctx.out(), "BrtEndCalcFeatures");
            case xlsb_type::BrtCalcFeature:
                return fmt::format_to(ctx.out(), "BrtCalcFeature");
            case xlsb_type::BrtExternalLinksPr:
                return fmt::format_to(ctx.out(), "BrtExternalLinksPr");
            default:
                return fmt::format_to(ctx.out(), "{}", (unsigned int)t);
        }
    }
};

struct brt_bundle_sh {
    uint32_t hsState;
    uint32_t iTabID;
};

#pragma pack(push,1)

struct brt_row_hdr {
    uint32_t rw;
    uint32_t ixfe;
    uint16_t miyRw;
    // FIXME - make sure bitfields are in correct order
    uint16_t fExtraAsc : 1;
    uint16_t fExtraDsc : 1;
    uint16_t reserved1 : 6;
    uint16_t iOutLevel : 3;
    uint16_t fCollapsed : 1;
    uint16_t fDyZero : 1;
    uint16_t fUnsynced : 1;
    uint16_t fGhostDirty : 1;
    uint16_t fReserved : 1;
    uint8_t fPhShow : 1;
    uint8_t reserved2 : 7;
    uint32_t ccolspan;
};

struct xlsb_cell {
    uint32_t column;
    uint32_t iStyleRef : 24;
    uint32_t fPhShow : 1;
    uint32_t reserved : 7;
};

struct rk_number {
    uint32_t fx100 : 1;
    uint32_t fInt : 1;
    uint32_t num : 30;
};

#pragma pack(pop)

struct brt_cell_rk {
    xlsb_cell cell;
    rk_number value;
};

struct brt_cell_bool {
    xlsb_cell cell;
    uint8_t fBool;
};

struct brt_cell_real {
    xlsb_cell cell;
    double xnum;
};

struct brt_cell_st {
    xlsb_cell cell;
    uint32_t len;
    char16_t str[0];
};

#pragma pack(push,1)

struct rich_str {
    uint8_t fRichStr : 1;
    uint8_t fExtStr : 1;
    uint8_t unused1 : 6;
    uint32_t len;
    char16_t str[0];
};

#pragma pack(pop)

struct brt_sst_item {
    rich_str richStr;
};

struct brt_cell_rstring {
    xlsb_cell cell;
    rich_str value;
};

struct brt_cell_isst {
    xlsb_cell cell;
    uint32_t isst;
};

#pragma pack(push,1)

struct brt_xf {
    uint16_t ixfeParent;
    uint16_t iFmt;
    uint16_t iFont;
    uint16_t iFill;
    uint16_t ixBOrder;
    uint8_t trot;
    uint8_t indent;
    uint16_t f123Prefix : 1;
    uint16_t fSxButton : 1;
    uint16_t fHidden : 1;
    uint16_t fLocked : 1;
    uint16_t iReadingOrder : 2;
    uint16_t fMergeCell : 1;
    uint16_t fShrinkToFit : 1;
    uint16_t fJustLast : 1;
    uint16_t fWrap : 1;
    uint16_t alcv : 3;
    uint16_t alc : 3;
    uint16_t unused : 10;
    uint16_t xfGrbitAtr : 6;
};

#pragma pack(pop)

struct brt_wb_prop {
    uint32_t f1904 : 1;
    uint32_t reserved1 : 1;
    uint32_t fHideBorderUnselLists : 1;
    uint32_t fFilterPrivacy : 1;
    uint32_t fBuggedUserABoutSolution : 1;
    uint32_t fShowInkAnnotation : 1;
    uint32_t fBackup : 1;
    uint32_t fNoSaveSup : 1;
    uint32_t grbitUpdateLinks : 2;
    uint32_t fHidePivotTableFList : 1;
    uint32_t fPublishedBookItems : 1;
    uint32_t fCheckCompat : 1;
    uint32_t mdDspObj : 2;
    uint32_t fShowPivotChartFilter : 1;
    uint32_t fAutoCompressPictures : 1;
    uint32_t reserved2 : 1;
    uint32_t fRefreshAll : 1;
    uint32_t unused : 13;
    uint32_t dwThemeVersion;
    uint32_t strName_len;
};

#pragma pack(push,1)

struct brt_fmt {
    uint16_t ifmt;
    uint32_t stFmtCode_len;
    char16_t stFmtCode[0];
};

#pragma pack(pop)
