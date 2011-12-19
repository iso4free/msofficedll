{ This file contains some pascal declarations imported from a type library.}

{ Microsoft Word 10.0 Object Library }
{ Version 8.2 }
const
  cm2p = 28.35; //1 cm = 28.35 points

{ WdMailSystem }

  wdNoMailSystem = 0;
  wdMAPI = 1;
  wdPowerTalk = 2;
  wdMAPIandPowerTalk = 3;

{ WdTemplateType }

  wdNormalTemplate = 0;
  wdGlobalTemplate = 1;
  wdAttachedTemplate = 2;

{ WdContinue }

  wdContinueDisabled = 0;
  wdResetList = 1;
  wdContinueList = 2;

{ WdIMEMode }

  wdIMEModeNoControl = 0;
  wdIMEModeOn = 1;
  wdIMEModeOff = 2;
  wdIMEModeHiragana = 4;
  wdIMEModeKatakana = 5;
  wdIMEModeKatakanaHalf = 6;
  wdIMEModeAlphaFull = 7;
  wdIMEModeAlpha = 8;
  wdIMEModeHangulFull = 9;
  wdIMEModeHangul = 10;

{ WdBaselineAlignment }

  wdBaselineAlignTop = 0;
  wdBaselineAlignCenter = 1;
  wdBaselineAlignBaseline = 2;
  wdBaselineAlignFarEast50 = 3;
  wdBaselineAlignAuto = 4;

{ WdIndexFilter }

  wdIndexFilterNone = 0;
  wdIndexFilterAiueo = 1;
  wdIndexFilterAkasatana = 2;
  wdIndexFilterChosung = 3;
  wdIndexFilterLow = 4;
  wdIndexFilterMedium = 5;
  wdIndexFilterFull = 6;

{ WdIndexSortBy }

  wdIndexSortByStroke = 0;
  wdIndexSortBySyllable = 1;

{ WdJustificationMode }

  wdJustificationModeExpand = 0;
  wdJustificationModeCompress = 1;
  wdJustificationModeCompressKana = 2;

{ WdFarEastLineBreakLevel }

  wdFarEastLineBreakLevelNormal = 0;
  wdFarEastLineBreakLevelStrict = 1;
  wdFarEastLineBreakLevelCustom = 2;

{ WdMultipleWordConversionsMode }

  wdHangulToHanja = 0;
  wdHanjaToHangul = 1;

{ WdColorIndex }

  wdAuto = 0;
  wdBlack = 1;
  wdBlue = 2;
  wdTurquoise = 3;
  wdBrightGreen = 4;
  wdPink = 5;
  wdRed = 6;
  wdYellow = 7;
  wdWhite = 8;
  wdDarkBlue = 9;
  wdTeal = 10;
  wdGreen = 11;
  wdViolet = 12;
  wdDarkRed = 13;
  wdDarkYellow = 14;
  wdGray50 = 15;
  wdGray25 = 16;
  wdByAuthor = -1;
  wdNoHighlight = 0;

{ WdTextureIndex }

  wdTextureNone = 0;
  wdTexture2Pt5Percent = 25;
  wdTexture5Percent = 50;
  wdTexture7Pt5Percent = 75;
  wdTexture10Percent = 100;
  wdTexture12Pt5Percent = 125;
  wdTexture15Percent = 150;
  wdTexture17Pt5Percent = 175;
  wdTexture20Percent = 200;
  wdTexture22Pt5Percent = 225;
  wdTexture25Percent = 250;
  wdTexture27Pt5Percent = 275;
  wdTexture30Percent = 300;
  wdTexture32Pt5Percent = 325;
  wdTexture35Percent = 350;
  wdTexture37Pt5Percent = 375;
  wdTexture40Percent = 400;
  wdTexture42Pt5Percent = 425;
  wdTexture45Percent = 450;
  wdTexture47Pt5Percent = 475;
  wdTexture50Percent = 500;
  wdTexture52Pt5Percent = 525;
  wdTexture55Percent = 550;
  wdTexture57Pt5Percent = 575;
  wdTexture60Percent = 600;
  wdTexture62Pt5Percent = 625;
  wdTexture65Percent = 650;
  wdTexture67Pt5Percent = 675;
  wdTexture70Percent = 700;
  wdTexture72Pt5Percent = 725;
  wdTexture75Percent = 750;
  wdTexture77Pt5Percent = 775;
  wdTexture80Percent = 800;
  wdTexture82Pt5Percent = 825;
  wdTexture85Percent = 850;
  wdTexture87Pt5Percent = 875;
  wdTexture90Percent = 900;
  wdTexture92Pt5Percent = 925;
  wdTexture95Percent = 950;
  wdTexture97Pt5Percent = 975;
  wdTextureSolid = 1000;
  wdTextureDarkHorizontal = -1;
  wdTextureDarkVertical = -2;
  wdTextureDarkDiagonalDown = -3;
  wdTextureDarkDiagonalUp = -4;
  wdTextureDarkCross = -5;
  wdTextureDarkDiagonalCross = -6;
  wdTextureHorizontal = -7;
  wdTextureVertical = -8;
  wdTextureDiagonalDown = -9;
  wdTextureDiagonalUp = -10;
  wdTextureCross = -11;
  wdTextureDiagonalCross = -12;

{ WdUnderline }

  wdUnderlineNone = 0;
  wdUnderlineSingle = 1;
  wdUnderlineWords = 2;
  wdUnderlineDouble = 3;
  wdUnderlineDotted = 4;
  wdUnderlineThick = 6;
  wdUnderlineDash = 7;
  wdUnderlineDotDash = 9;
  wdUnderlineDotDotDash = 10;
  wdUnderlineWavy = 11;
  wdUnderlineWavyHeavy = 27;
  wdUnderlineDottedHeavy = 20;
  wdUnderlineDashHeavy = 23;
  wdUnderlineDotDashHeavy = 25;
  wdUnderlineDotDotDashHeavy = 26;
  wdUnderlineDashLong = 39;
  wdUnderlineDashLongHeavy = 55;
  wdUnderlineWavyDouble = 43;

{ WdEmphasisMark }

  wdEmphasisMarkNone = 0;
  wdEmphasisMarkOverSolidCircle = 1;
  wdEmphasisMarkOverComma = 2;
  wdEmphasisMarkOverWhiteCircle = 3;
  wdEmphasisMarkUnderSolidCircle = 4;

{ WdInternationalIndex }

  wdListSeparator = 17;
  wdDecimalSeparator = 18;
  wdThousandsSeparator = 19;
  wdCurrencyCode = 20;
  wd24HourClock = 21;
  wdInternationalAM = 22;
  wdInternationalPM = 23;
  wdTimeSeparator = 24;
  wdDateSeparator = 25;
  wdProductLanguageID = 26;

{ WdAutoMacros }

  wdAutoExec = 0;
  wdAutoNew = 1;
  wdAutoOpen = 2;
  wdAutoClose = 3;
  wdAutoExit = 4;

{ WdCaptionPosition }

  wdCaptionPositionAbove = 0;
  wdCaptionPositionBelow = 1;

{ WdCountry }

  wdUS = 1;
  wdCanada = 2;
  wdLatinAmerica = 3;
  wdNetherlands = 31;
  wdFrance = 33;
  wdSpain = 34;
  wdItaly = 39;
  wdUK = 44;
  wdDenmark = 45;
  wdSweden = 46;
  wdNorway = 47;
  wdGermany = 49;
  wdPeru = 51;
  wdMexico = 52;
  wdArgentina = 54;
  wdBrazil = 55;
  wdChile = 56;
  wdVenezuela = 58;
  wdJapan = 81;
  wdTaiwan = 886;
  wdChina = 86;
  wdKorea = 82;
  wdFinland = 358;
  wdIceland = 354;

{ WdHeadingSeparator }

  wdHeadingSeparatorNone = 0;
  wdHeadingSeparatorBlankLine = 1;
  wdHeadingSeparatorLetter = 2;
  wdHeadingSeparatorLetterLow = 3;
  wdHeadingSeparatorLetterFull = 4;

{ WdSeparatorType }

  wdSeparatorHyphen = 0;
  wdSeparatorPeriod = 1;
  wdSeparatorColon = 2;
  wdSeparatorEmDash = 3;
  wdSeparatorEnDash = 4;

{ WdPageNumberAlignment }

  wdAlignPageNumberLeft = 0;
  wdAlignPageNumberCenter = 1;
  wdAlignPageNumberRight = 2;
  wdAlignPageNumberInside = 3;
  wdAlignPageNumberOutside = 4;

{ WdBorderType }

  wdBorderTop = -1;
  wdBorderLeft = -2;
  wdBorderBottom = -3;
  wdBorderRight = -4;
  wdBorderHorizontal = -5;
  wdBorderVertical = -6;
  wdBorderDiagonalDown = -7;
  wdBorderDiagonalUp = -8;

{ WdBorderTypeHID }

  emptyenum = 0;

{ WdFramePosition }

  wdFrameTop = -999999;
  wdFrameLeft = -999998;
  wdFrameBottom = -999997;
  wdFrameRight = -999996;
  wdFrameCenter = -999995;
  wdFrameInside = -999994;
  wdFrameOutside = -999993;

{ WdAnimation }

  wdAnimationNone = 0;
  wdAnimationLasVegasLights = 1;
  wdAnimationBlinkingBackground = 2;
  wdAnimationSparkleText = 3;
  wdAnimationMarchingBlackAnts = 4;
  wdAnimationMarchingRedAnts = 5;
  wdAnimationShimmer = 6;

{ WdCharacterCase }

  wdNextCase = -1;
  wdLowerCase = 0;
  wdUpperCase = 1;
  wdTitleWord = 2;
  wdTitleSentence = 4;
  wdToggleCase = 5;
  wdHalfWidth = 6;
  wdFullWidth = 7;
  wdKatakana = 8;
  wdHiragana = 9;

{ WdCharacterCaseHID }

//emptyenum = 0;

{ WdSummaryMode }

  wdSummaryModeHighlight = 0;
  wdSummaryModeHideAllButSummary = 1;
  wdSummaryModeInsert = 2;
  wdSummaryModeCreateNew = 3;

{ WdSummaryLength }

  wd10Sentences = -2;
  wd20Sentences = -3;
  wd100Words = -4;
  wd500Words = -5;
  wd10Percent = -6;
  wd25Percent = -7;
  wd50Percent = -8;
  wd75Percent = -9;

{ WdStyleType }

  wdStyleTypeParagraph = 1;
  wdStyleTypeCharacter = 2;
  wdStyleTypeTable = 3;
  wdStyleTypeList = 4;

{ WdUnits }

  wdCharacter = 1;
  wdWord = 2;
  wdSentence = 3;
  wdParagraph = 4;
  wdLine = 5;
  wdStory = 6;
  wdScreen = 7;
  wdSection = 8;
  wdColumn = 9;
  wdRow = 10;
  wdWindow = 11;
  wdCell = 12;
  wdCharacterFormatting = 13;
  wdParagraphFormatting = 14;
  wdTable = 15;
  wdItem = 16;

{ WdGoToItem }

  wdGoToBookmark = -1;
  wdGoToSection = 0;
  wdGoToPage = 1;
  wdGoToTable = 2;
  wdGoToLine = 3;
  wdGoToFootnote = 4;
  wdGoToEndnote = 5;
  wdGoToComment = 6;
  wdGoToField = 7;
  wdGoToGraphic = 8;
  wdGoToObject = 9;
  wdGoToEquation = 10;
  wdGoToHeading = 11;
  wdGoToPercent = 12;
  wdGoToSpellingError = 13;
  wdGoToGrammaticalError = 14;
  wdGoToProofreadingError = 15;

{ WdGoToDirection }

  wdGoToFirst = 1;
  wdGoToLast = -1;
  wdGoToNext = 2;
  wdGoToRelative = 2;
  wdGoToPrevious = 3;
  wdGoToAbsolute = 1;

{ WdCollapseDirection }

  wdCollapseStart = 1;
  wdCollapseEnd = 0;

{ WdRowHeightRule }

  wdRowHeightAuto = 0;
  wdRowHeightAtLeast = 1;
  wdRowHeightExactly = 2;

{ WdFrameSizeRule }

  wdFrameAuto = 0;
  wdFrameAtLeast = 1;
  wdFrameExact = 2;

{ WdInsertCells }

  wdInsertCellsShiftRight = 0;
  wdInsertCellsShiftDown = 1;
  wdInsertCellsEntireRow = 2;
  wdInsertCellsEntireColumn = 3;

{ WdDeleteCells }

  wdDeleteCellsShiftLeft = 0;
  wdDeleteCellsShiftUp = 1;
  wdDeleteCellsEntireRow = 2;
  wdDeleteCellsEntireColumn = 3;

{ WdListApplyTo }

  wdListApplyToWholeList = 0;
  wdListApplyToThisPointForward = 1;
  wdListApplyToSelection = 2;

{ WdAlertLevel }

  wdAlertsNone = 0;
  wdAlertsMessageBox = -2;
  wdAlertsAll = -1;

{ WdCursorType }

  wdCursorWait = 0;
  wdCursorIBeam = 1;
  wdCursorNormal = 2;
  wdCursorNorthwestArrow = 3;

{ WdEnableCancelKey }

  wdCancelDisabled = 0;
  wdCancelInterrupt = 1;

{ WdRulerStyle }

  wdAdjustNone = 0;
  wdAdjustProportional = 1;
  wdAdjustFirstColumn = 2;
  wdAdjustSameWidth = 3;

{ WdParagraphAlignment }

  wdAlignParagraphLeft = 0;
  wdAlignParagraphCenter = 1;
  wdAlignParagraphRight = 2;
  wdAlignParagraphJustify = 3;
  wdAlignParagraphDistribute = 4;
  wdAlignParagraphJustifyMed = 5;
  wdAlignParagraphJustifyHi = 7;
  wdAlignParagraphJustifyLow = 8;
  wdAlignParagraphThaiJustify = 9;

{ WdParagraphAlignmentHID }

//emptyenum = 0;

{ WdListLevelAlignment }

  wdListLevelAlignLeft = 0;
  wdListLevelAlignCenter = 1;
  wdListLevelAlignRight = 2;

{ WdRowAlignment }

  wdAlignRowLeft = 0;
  wdAlignRowCenter = 1;
  wdAlignRowRight = 2;

{ WdTabAlignment }

  wdAlignTabLeft = 0;
  wdAlignTabCenter = 1;
  wdAlignTabRight = 2;
  wdAlignTabDecimal = 3;
  wdAlignTabBar = 4;
  wdAlignTabList = 6;

{ WdVerticalAlignment }

  wdAlignVerticalTop = 0;
  wdAlignVerticalCenter = 1;
  wdAlignVerticalJustify = 2;
  wdAlignVerticalBottom = 3;

{ WdCellVerticalAlignment }

  wdCellAlignVerticalTop = 0;
  wdCellAlignVerticalCenter = 1;
  wdCellAlignVerticalBottom = 3;

{ WdTrailingCharacter }

  wdTrailingTab = 0;
  wdTrailingSpace = 1;
  wdTrailingNone = 2;

{ WdListGalleryType }

  wdBulletGallery = 1;
  wdNumberGallery = 2;
  wdOutlineNumberGallery = 3;

{ WdListNumberStyle }

  wdListNumberStyleArabic = 0;
  wdListNumberStyleUppercaseRoman = 1;
  wdListNumberStyleLowercaseRoman = 2;
  wdListNumberStyleUppercaseLetter = 3;
  wdListNumberStyleLowercaseLetter = 4;
  wdListNumberStyleOrdinal = 5;
  wdListNumberStyleCardinalText = 6;
  wdListNumberStyleOrdinalText = 7;
  wdListNumberStyleKanji = 10;
  wdListNumberStyleKanjiDigit = 11;
  wdListNumberStyleAiueoHalfWidth = 12;
  wdListNumberStyleIrohaHalfWidth = 13;
  wdListNumberStyleArabicFullWidth = 14;
  wdListNumberStyleKanjiTraditional = 16;
  wdListNumberStyleKanjiTraditional2 = 17;
  wdListNumberStyleNumberInCircle = 18;
  wdListNumberStyleAiueo = 20;
  wdListNumberStyleIroha = 21;
  wdListNumberStyleArabicLZ = 22;
  wdListNumberStyleBullet = 23;
  wdListNumberStyleGanada = 24;
  wdListNumberStyleChosung = 25;
  wdListNumberStyleGBNum1 = 26;
  wdListNumberStyleGBNum2 = 27;
  wdListNumberStyleGBNum3 = 28;
  wdListNumberStyleGBNum4 = 29;
  wdListNumberStyleZodiac1 = 30;
  wdListNumberStyleZodiac2 = 31;
  wdListNumberStyleZodiac3 = 32;
  wdListNumberStyleTradChinNum1 = 33;
  wdListNumberStyleTradChinNum2 = 34;
  wdListNumberStyleTradChinNum3 = 35;
  wdListNumberStyleTradChinNum4 = 36;
  wdListNumberStyleSimpChinNum1 = 37;
  wdListNumberStyleSimpChinNum2 = 38;
  wdListNumberStyleSimpChinNum3 = 39;
  wdListNumberStyleSimpChinNum4 = 40;
  wdListNumberStyleHanjaRead = 41;
  wdListNumberStyleHanjaReadDigit = 42;
  wdListNumberStyleHangul = 43;
  wdListNumberStyleHanja = 44;
  wdListNumberStyleHebrew1 = 45;
  wdListNumberStyleArabic1 = 46;
  wdListNumberStyleHebrew2 = 47;
  wdListNumberStyleArabic2 = 48;
  wdListNumberStyleHindiLetter1 = 49;
  wdListNumberStyleHindiLetter2 = 50;
  wdListNumberStyleHindiArabic = 51;
  wdListNumberStyleHindiCardinalText = 52;
  wdListNumberStyleThaiLetter = 53;
  wdListNumberStyleThaiArabic = 54;
  wdListNumberStyleThaiCardinalText = 55;
  wdListNumberStyleVietCardinalText = 56;
  wdListNumberStyleLowercaseRussian = 58;
  wdListNumberStyleUppercaseRussian = 59;
  wdListNumberStylePictureBullet = 249;
  wdListNumberStyleLegal = 253;
  wdListNumberStyleLegalLZ = 254;
  wdListNumberStyleNone = 255;

{ WdListNumberStyleHID }

//emptyenum = 0;

{ WdNoteNumberStyle }

  wdNoteNumberStyleArabic = 0;
  wdNoteNumberStyleUppercaseRoman = 1;
  wdNoteNumberStyleLowercaseRoman = 2;
  wdNoteNumberStyleUppercaseLetter = 3;
  wdNoteNumberStyleLowercaseLetter = 4;
  wdNoteNumberStyleSymbol = 9;
  wdNoteNumberStyleArabicFullWidth = 14;
  wdNoteNumberStyleKanji = 10;
  wdNoteNumberStyleKanjiDigit = 11;
  wdNoteNumberStyleKanjiTraditional = 16;
  wdNoteNumberStyleNumberInCircle = 18;
  wdNoteNumberStyleHanjaRead = 41;
  wdNoteNumberStyleHanjaReadDigit = 42;
  wdNoteNumberStyleTradChinNum1 = 33;
  wdNoteNumberStyleTradChinNum2 = 34;
  wdNoteNumberStyleSimpChinNum1 = 37;
  wdNoteNumberStyleSimpChinNum2 = 38;
  wdNoteNumberStyleHebrewLetter1 = 45;
  wdNoteNumberStyleArabicLetter1 = 46;
  wdNoteNumberStyleHebrewLetter2 = 47;
  wdNoteNumberStyleArabicLetter2 = 48;
  wdNoteNumberStyleHindiLetter1 = 49;
  wdNoteNumberStyleHindiLetter2 = 50;
  wdNoteNumberStyleHindiArabic = 51;
  wdNoteNumberStyleHindiCardinalText = 52;
  wdNoteNumberStyleThaiLetter = 53;
  wdNoteNumberStyleThaiArabic = 54;
  wdNoteNumberStyleThaiCardinalText = 55;
  wdNoteNumberStyleVietCardinalText = 56;

{ WdNoteNumberStyleHID }

//emptyenum = 0;

{ WdCaptionNumberStyle }

  wdCaptionNumberStyleArabic = 0;
  wdCaptionNumberStyleUppercaseRoman = 1;
  wdCaptionNumberStyleLowercaseRoman = 2;
  wdCaptionNumberStyleUppercaseLetter = 3;
  wdCaptionNumberStyleLowercaseLetter = 4;
  wdCaptionNumberStyleArabicFullWidth = 14;
  wdCaptionNumberStyleKanji = 10;
  wdCaptionNumberStyleKanjiDigit = 11;
  wdCaptionNumberStyleKanjiTraditional = 16;
  wdCaptionNumberStyleNumberInCircle = 18;
  wdCaptionNumberStyleGanada = 24;
  wdCaptionNumberStyleChosung = 25;
  wdCaptionNumberStyleZodiac1 = 30;
  wdCaptionNumberStyleZodiac2 = 31;
  wdCaptionNumberStyleHanjaRead = 41;
  wdCaptionNumberStyleHanjaReadDigit = 42;
  wdCaptionNumberStyleTradChinNum2 = 34;
  wdCaptionNumberStyleTradChinNum3 = 35;
  wdCaptionNumberStyleSimpChinNum2 = 38;
  wdCaptionNumberStyleSimpChinNum3 = 39;
  wdCaptionNumberStyleHebrewLetter1 = 45;
  wdCaptionNumberStyleArabicLetter1 = 46;
  wdCaptionNumberStyleHebrewLetter2 = 47;
  wdCaptionNumberStyleArabicLetter2 = 48;
  wdCaptionNumberStyleHindiLetter1 = 49;
  wdCaptionNumberStyleHindiLetter2 = 50;
  wdCaptionNumberStyleHindiArabic = 51;
  wdCaptionNumberStyleHindiCardinalText = 52;
  wdCaptionNumberStyleThaiLetter = 53;
  wdCaptionNumberStyleThaiArabic = 54;
  wdCaptionNumberStyleThaiCardinalText = 55;
  wdCaptionNumberStyleVietCardinalText = 56;

{ WdCaptionNumberStyleHID }

//emptyenum = 0;

{ WdPageNumberStyle }

  wdPageNumberStyleArabic = 0;
  wdPageNumberStyleUppercaseRoman = 1;
  wdPageNumberStyleLowercaseRoman = 2;
  wdPageNumberStyleUppercaseLetter = 3;
  wdPageNumberStyleLowercaseLetter = 4;
  wdPageNumberStyleArabicFullWidth = 14;
  wdPageNumberStyleKanji = 10;
  wdPageNumberStyleKanjiDigit = 11;
  wdPageNumberStyleKanjiTraditional = 16;
  wdPageNumberStyleNumberInCircle = 18;
  wdPageNumberStyleHanjaRead = 41;
  wdPageNumberStyleHanjaReadDigit = 42;
  wdPageNumberStyleTradChinNum1 = 33;
  wdPageNumberStyleTradChinNum2 = 34;
  wdPageNumberStyleSimpChinNum1 = 37;
  wdPageNumberStyleSimpChinNum2 = 38;
  wdPageNumberStyleHebrewLetter1 = 45;
  wdPageNumberStyleArabicLetter1 = 46;
  wdPageNumberStyleHebrewLetter2 = 47;
  wdPageNumberStyleArabicLetter2 = 48;
  wdPageNumberStyleHindiLetter1 = 49;
  wdPageNumberStyleHindiLetter2 = 50;
  wdPageNumberStyleHindiArabic = 51;
  wdPageNumberStyleHindiCardinalText = 52;
  wdPageNumberStyleThaiLetter = 53;
  wdPageNumberStyleThaiArabic = 54;
  wdPageNumberStyleThaiCardinalText = 55;
  wdPageNumberStyleVietCardinalText = 56;
  wdPageNumberStyleNumberInDash = 57;

{ WdPageNumberStyleHID }

//emptyenum = 0;

{ WdStatistic }

  wdStatisticWords = 0;
  wdStatisticLines = 1;
  wdStatisticPages = 2;
  wdStatisticCharacters = 3;
  wdStatisticParagraphs = 4;
  wdStatisticCharactersWithSpaces = 5;
  wdStatisticFarEastCharacters = 6;

{ WdStatisticHID }

//emptyenum = 0;

{ WdBuiltInProperty }

  wdPropertyTitle = 1;
  wdPropertySubject = 2;
  wdPropertyAuthor = 3;
  wdPropertyKeywords = 4;
  wdPropertyComments = 5;
  wdPropertyTemplate = 6;
  wdPropertyLastAuthor = 7;
  wdPropertyRevision = 8;
  wdPropertyAppName = 9;
  wdPropertyTimeLastPrinted = 10;
  wdPropertyTimeCreated = 11;
  wdPropertyTimeLastSaved = 12;
  wdPropertyVBATotalEdit = 13;
  wdPropertyPages = 14;
  wdPropertyWords = 15;
  wdPropertyCharacters = 16;
  wdPropertySecurity = 17;
  wdPropertyCategory = 18;
  wdPropertyFormat = 19;
  wdPropertyManager = 20;
  wdPropertyCompany = 21;
  wdPropertyBytes = 22;
  wdPropertyLines = 23;
  wdPropertyParas = 24;
  wdPropertySlides = 25;
  wdPropertyNotes = 26;
  wdPropertyHiddenSlides = 27;
  wdPropertyMMClips = 28;
  wdPropertyHyperlinkBase = 29;
  wdPropertyCharsWSpaces = 30;

{ WdLineSpacing }

  wdLineSpaceSingle = 0;
  wdLineSpace1pt5 = 1;
  wdLineSpaceDouble = 2;
  wdLineSpaceAtLeast = 3;
  wdLineSpaceExactly = 4;
  wdLineSpaceMultiple = 5;

{ WdNumberType }

  wdNumberParagraph = 1;
  wdNumberListNum = 2;
  wdNumberAllNumbers = 3;

{ WdListType }

  wdListNoNumbering = 0;
  wdListListNumOnly = 1;
  wdListBullet = 2;
  wdListSimpleNumbering = 3;
  wdListOutlineNumbering = 4;
  wdListMixedNumbering = 5;
  wdListPictureBullet = 6;

{ WdStoryType }

  wdMainTextStory = 1;
  wdFootnotesStory = 2;
  wdEndnotesStory = 3;
  wdCommentsStory = 4;
  wdTextFrameStory = 5;
  wdEvenPagesHeaderStory = 6;
  wdPrimaryHeaderStory = 7;
  wdEvenPagesFooterStory = 8;
  wdPrimaryFooterStory = 9;
  wdFirstPageHeaderStory = 10;
  wdFirstPageFooterStory = 11;

{ WdSaveFormat }

  wdFormatDocument = 0;
  wdFormatTemplate = 1;
  wdFormatText = 2;
  wdFormatTextLineBreaks = 3;
  wdFormatDOSText = 4;
  wdFormatDOSTextLineBreaks = 5;
  wdFormatRTF = 6;
  wdFormatUnicodeText = 7;
  wdFormatEncodedText = 7;
  wdFormatHTML = 8;
  wdFormatWebArchive = 9;
  wdFormatFilteredHTML = 10;

{ WdOpenFormat }

  wdOpenFormatAuto = 0;
  wdOpenFormatDocument = 1;
  wdOpenFormatTemplate = 2;
  wdOpenFormatRTF = 3;
  wdOpenFormatText = 4;
  wdOpenFormatUnicodeText = 5;
  wdOpenFormatEncodedText = 5;
  wdOpenFormatAllWord = 6;
  wdOpenFormatWebPages = 7;

{ WdHeaderFooterIndex }

  wdHeaderFooterPrimary = 1;
  wdHeaderFooterFirstPage = 2;
  wdHeaderFooterEvenPages = 3;

{ WdTocFormat }

  wdTOCTemplate = 0;
  wdTOCClassic = 1;
  wdTOCDistinctive = 2;
  wdTOCFancy = 3;
  wdTOCModern = 4;
  wdTOCFormal = 5;
  wdTOCSimple = 6;

{ WdTofFormat }

  wdTOFTemplate = 0;
  wdTOFClassic = 1;
  wdTOFDistinctive = 2;
  wdTOFCentered = 3;
  wdTOFFormal = 4;
  wdTOFSimple = 5;

{ WdToaFormat }

  wdTOATemplate = 0;
  wdTOAClassic = 1;
  wdTOADistinctive = 2;
  wdTOAFormal = 3;
  wdTOASimple = 4;

{ WdLineStyle }

  wdLineStyleNone = 0;
  wdLineStyleSingle = 1;
  wdLineStyleDot = 2;
  wdLineStyleDashSmallGap = 3;
  wdLineStyleDashLargeGap = 4;
  wdLineStyleDashDot = 5;
  wdLineStyleDashDotDot = 6;
  wdLineStyleDouble = 7;
  wdLineStyleTriple = 8;
  wdLineStyleThinThickSmallGap = 9;
  wdLineStyleThickThinSmallGap = 10;
  wdLineStyleThinThickThinSmallGap = 11;
  wdLineStyleThinThickMedGap = 12;
  wdLineStyleThickThinMedGap = 13;
  wdLineStyleThinThickThinMedGap = 14;
  wdLineStyleThinThickLargeGap = 15;
  wdLineStyleThickThinLargeGap = 16;
  wdLineStyleThinThickThinLargeGap = 17;
  wdLineStyleSingleWavy = 18;
  wdLineStyleDoubleWavy = 19;
  wdLineStyleDashDotStroked = 20;
  wdLineStyleEmboss3D = 21;
  wdLineStyleEngrave3D = 22;
  wdLineStyleOutset = 23;
  wdLineStyleInset = 24;

{ WdLineWidth }

  wdLineWidth025pt = 2;
  wdLineWidth050pt = 4;
  wdLineWidth075pt = 6;
  wdLineWidth100pt = 8;
  wdLineWidth150pt = 12;
  wdLineWidth225pt = 18;
  wdLineWidth300pt = 24;
  wdLineWidth450pt = 36;
  wdLineWidth600pt = 48;

{ WdBreakType }

  wdSectionBreakNextPage = 2;
  wdSectionBreakContinuous = 3;
  wdSectionBreakEvenPage = 4;
  wdSectionBreakOddPage = 5;
  wdLineBreak = 6;
  wdPageBreak = 7;
  wdColumnBreak = 8;
  wdLineBreakClearLeft = 9;
  wdLineBreakClearRight = 10;
  wdTextWrappingBreak = 11;

{ WdTabLeader }

  wdTabLeaderSpaces = 0;
  wdTabLeaderDots = 1;
  wdTabLeaderDashes = 2;
  wdTabLeaderLines = 3;
  wdTabLeaderHeavy = 4;
  wdTabLeaderMiddleDot = 5;

{ WdTabLeaderHID }

//emptyenum = 0;

{ WdMeasurementUnits }

  wdInches = 0;
  wdCentimeters = 1;
  wdMillimeters = 2;
  wdPoints = 3;
  wdPicas = 4;

{ WdMeasurementUnitsHID }

//emptyenum = 0;

{ WdDropPosition }

  wdDropNone = 0;
  wdDropNormal = 1;
  wdDropMargin = 2;

{ WdNumberingRule }

  wdRestartContinuous = 0;
  wdRestartSection = 1;
  wdRestartPage = 2;

{ WdFootnoteLocation }

  wdBottomOfPage = 0;
  wdBeneathText = 1;

{ WdEndnoteLocation }

  wdEndOfSection = 0;
  wdEndOfDocument = 1;

{ WdSortSeparator }

  wdSortSeparateByTabs = 0;
  wdSortSeparateByCommas = 1;
  wdSortSeparateByDefaultTableSeparator = 2;

{ WdTableFieldSeparator }

  wdSeparateByParagraphs = 0;
  wdSeparateByTabs = 1;
  wdSeparateByCommas = 2;
  wdSeparateByDefaultListSeparator = 3;

{ WdSortFieldType }

  wdSortFieldAlphanumeric = 0;
  wdSortFieldNumeric = 1;
  wdSortFieldDate = 2;
  wdSortFieldSyllable = 3;
  wdSortFieldJapanJIS = 4;
  wdSortFieldStroke = 5;
  wdSortFieldKoreaKS = 6;

{ WdSortFieldTypeHID }

//emptyenum = 0;

{ WdSortOrder }

  wdSortOrderAscending = 0;
  wdSortOrderDescending = 1;

{ WdTableFormat }

  wdTableFormatNone = 0;
  wdTableFormatSimple1 = 1;
  wdTableFormatSimple2 = 2;
  wdTableFormatSimple3 = 3;
  wdTableFormatClassic1 = 4;
  wdTableFormatClassic2 = 5;
  wdTableFormatClassic3 = 6;
  wdTableFormatClassic4 = 7;
  wdTableFormatColorful1 = 8;
  wdTableFormatColorful2 = 9;
  wdTableFormatColorful3 = 10;
  wdTableFormatColumns1 = 11;
  wdTableFormatColumns2 = 12;
  wdTableFormatColumns3 = 13;
  wdTableFormatColumns4 = 14;
  wdTableFormatColumns5 = 15;
  wdTableFormatGrid1 = 16;
  wdTableFormatGrid2 = 17;
  wdTableFormatGrid3 = 18;
  wdTableFormatGrid4 = 19;
  wdTableFormatGrid5 = 20;
  wdTableFormatGrid6 = 21;
  wdTableFormatGrid7 = 22;
  wdTableFormatGrid8 = 23;
  wdTableFormatList1 = 24;
  wdTableFormatList2 = 25;
  wdTableFormatList3 = 26;
  wdTableFormatList4 = 27;
  wdTableFormatList5 = 28;
  wdTableFormatList6 = 29;
  wdTableFormatList7 = 30;
  wdTableFormatList8 = 31;
  wdTableFormat3DEffects1 = 32;
  wdTableFormat3DEffects2 = 33;
  wdTableFormat3DEffects3 = 34;
  wdTableFormatContemporary = 35;
  wdTableFormatElegant = 36;
  wdTableFormatProfessional = 37;
  wdTableFormatSubtle1 = 38;
  wdTableFormatSubtle2 = 39;
  wdTableFormatWeb1 = 40;
  wdTableFormatWeb2 = 41;
  wdTableFormatWeb3 = 42;

{ WdTableFormatApply }

  wdTableFormatApplyBorders = 1;
  wdTableFormatApplyShading = 2;
  wdTableFormatApplyFont = 4;
  wdTableFormatApplyColor = 8;
  wdTableFormatApplyAutoFit = 16;
  wdTableFormatApplyHeadingRows = 32;
  wdTableFormatApplyLastRow = 64;
  wdTableFormatApplyFirstColumn = 128;
  wdTableFormatApplyLastColumn = 256;

{ WdLanguageID }

  wdLanguageNone = 0;
  wdNoProofing = 1024;
  wdAfrikaans = 1078;
  wdAlbanian = 1052;
  wdAmharic = 1118;
  wdArabicAlgeria = 5121;
  wdArabicBahrain = 15361;
  wdArabicEgypt = 3073;
  wdArabicIraq = 2049;
  wdArabicJordan = 11265;
  wdArabicKuwait = 13313;
  wdArabicLebanon = 12289;
  wdArabicLibya = 4097;
  wdArabicMorocco = 6145;
  wdArabicOman = 8193;
  wdArabicQatar = 16385;
  wdArabic = 1025;
  wdArabicSyria = 10241;
  wdArabicTunisia = 7169;
  wdArabicUAE = 14337;
  wdArabicYemen = 9217;
  wdArmenian = 1067;
  wdAssamese = 1101;
  wdAzeriCyrillic = 2092;
  wdAzeriLatin = 1068;
  wdBasque = 1069;
  wdByelorussian = 1059;
  wdBengali = 1093;
  wdBulgarian = 1026;
  wdBurmese = 1109;
  wdCatalan = 1027;
  wdCherokee = 1116;
  wdChineseHongKongSAR = 3076;
  wdChineseMacaoSAR = 5124;
  wdSimplifiedChinese = 2052;
  wdChineseSingapore = 4100;
  wdTraditionalChinese = 1028;
  wdCroatian = 1050;
  wdCzech = 1029;
  wdDanish = 1030;
  wdDivehi = 1125;
  wdBelgianDutch = 2067;
  wdDutch = 1043;
  wdDzongkhaBhutan = 2129;
  wdEdo = 1126;
  wdEnglishAUS = 3081;
  wdEnglishBelize = 10249;
  wdEnglishCanadian = 4105;
  wdEnglishCaribbean = 9225;
  wdEnglishIreland = 6153;
  wdEnglishJamaica = 8201;
  wdEnglishNewZealand = 5129;
  wdEnglishPhilippines = 13321;
  wdEnglishSouthAfrica = 7177;
  wdEnglishTrinidadTobago = 11273;
  wdEnglishUK = 2057;
  wdEnglishUS = 1033;
  wdEnglishZimbabwe = 12297;
  wdEnglishIndonesia = 14345;
  wdEstonian = 1061;
  wdFaeroese = 1080;
  wdFarsi = 1065;
  wdFilipino = 1124;
  wdFinnish = 1035;
  wdFulfulde = 1127;
  wdBelgianFrench = 2060;
  wdFrenchCameroon = 11276;
  wdFrenchCanadian = 3084;
  wdFrenchCotedIvoire = 12300;
  wdFrench = 1036;
  wdFrenchLuxembourg = 5132;
  wdFrenchMali = 13324;
  wdFrenchMonaco = 6156;
  wdFrenchReunion = 8204;
  wdFrenchSenegal = 10252;
  wdFrenchMorocco = 14348;
  wdFrenchHaiti = 15372;
  wdSwissFrench = 4108;
  wdFrenchWestIndies = 7180;
  wdFrenchZaire = 9228;
  wdFrisianNetherlands = 1122;
  wdGaelicIreland = 2108;
  wdGaelicScotland = 1084;
  wdGalician = 1110;
  wdGeorgian = 1079;
  wdGermanAustria = 3079;
  wdGerman = 1031;
  wdGermanLiechtenstein = 5127;
  wdGermanLuxembourg = 4103;
  wdSwissGerman = 2055;
  wdGreek = 1032;
  wdGuarani = 1140;
  wdGujarati = 1095;
  wdHausa = 1128;
  wdHawaiian = 1141;
  wdHebrew = 1037;
  wdHindi = 1081;
  wdHungarian = 1038;
  wdIbibio = 1129;
  wdIcelandic = 1039;
  wdIgbo = 1136;
  wdIndonesian = 1057;
  wdInuktitut = 1117;
  wdItalian = 1040;
  wdSwissItalian = 2064;
  wdJapanese = 1041;
  wdKannada = 1099;
  wdKanuri = 1137;
  wdKashmiri = 1120;
  wdKazakh = 1087;
  wdKhmer = 1107;
  wdKirghiz = 1088;
  wdKonkani = 1111;
  wdKorean = 1042;
  wdKyrgyz = 1088;
  wdLao = 1108;
  wdLatin = 1142;
  wdLatvian = 1062;
  wdLithuanian = 1063;
  wdMacedonian = 1071;
  wdMalaysian = 1086;
  wdMalayBruneiDarussalam = 2110;
  wdMalayalam = 1100;
  wdMaltese = 1082;
  wdManipuri = 1112;
  wdMarathi = 1102;
  wdMongolian = 1104;
  wdNepali = 1121;
  wdNorwegianBokmol = 1044;
  wdNorwegianNynorsk = 2068;
  wdOriya = 1096;
  wdOromo = 1138;
  wdPashto = 1123;
  wdPolish = 1045;
  wdBrazilianPortuguese = 1046;
  wdPortuguese = 2070;
  wdPunjabi = 1094;
  wdRhaetoRomanic = 1047;
  wdRomanianMoldova = 2072;
  wdRomanian = 1048;
  wdRussianMoldova = 2073;
  wdRussian = 1049;
  wdSamiLappish = 1083;
  wdSanskrit = 1103;
  wdSerbianCyrillic = 3098;
  wdSerbianLatin = 2074;
  wdSinhalese = 1115;
  wdSindhi = 1113;
  wdSindhiPakistan = 2137;
  wdSlovak = 1051;
  wdSlovenian = 1060;
  wdSomali = 1143;
  wdSorbian = 1070;
  wdSpanishArgentina = 11274;
  wdSpanishBolivia = 16394;
  wdSpanishChile = 13322;
  wdSpanishColombia = 9226;
  wdSpanishCostaRica = 5130;
  wdSpanishDominicanRepublic = 7178;
  wdSpanishEcuador = 12298;
  wdSpanishElSalvador = 17418;
  wdSpanishGuatemala = 4106;
  wdSpanishHonduras = 18442;
  wdMexicanSpanish = 2058;
  wdSpanishNicaragua = 19466;
  wdSpanishPanama = 6154;
  wdSpanishParaguay = 15370;
  wdSpanishPeru = 10250;
  wdSpanishPuertoRico = 20490;
  wdSpanishModernSort = 3082;
  wdSpanish = 1034;
  wdSpanishUruguay = 14346;
  wdSpanishVenezuela = 8202;
  wdSesotho = 1072;
  wdSutu = 1072;
  wdSwahili = 1089;
  wdSwedishFinland = 2077;
  wdSwedish = 1053;
  wdSyriac = 1114;
  wdTajik = 1064;
  wdTamazight = 1119;
  wdTamazightLatin = 2143;
  wdTamil = 1097;
  wdTatar = 1092;
  wdTelugu = 1098;
  wdThai = 1054;
  wdTibetan = 1105;
  wdTigrignaEthiopic = 1139;
  wdTigrignaEritrea = 2163;
  wdTsonga = 1073;
  wdTswana = 1074;
  wdTurkish = 1055;
  wdTurkmen = 1090;
  wdUkrainian = 1058;
  wdUrdu = 1056;
  wdUzbekCyrillic = 2115;
  wdUzbekLatin = 1091;
  wdVenda = 1075;
  wdVietnamese = 1066;
  wdWelsh = 1106;
  wdXhosa = 1076;
  wdYi = 1144;
  wdYiddish = 1085;
  wdYoruba = 1130;
  wdZulu = 1077;

{ WdFieldType }

  wdFieldEmpty = -1;
  wdFieldRef = 3;
  wdFieldIndexEntry = 4;
  wdFieldFootnoteRef = 5;
  wdFieldSet = 6;
  wdFieldIf = 7;
  wdFieldIndex = 8;
  wdFieldTOCEntry = 9;
  wdFieldStyleRef = 10;
  wdFieldRefDoc = 11;
  wdFieldSequence = 12;
  wdFieldTOC = 13;
  wdFieldInfo = 14;
  wdFieldTitle = 15;
  wdFieldSubject = 16;
  wdFieldAuthor = 17;
  wdFieldKeyWord = 18;
  wdFieldComments = 19;
  wdFieldLastSavedBy = 20;
  wdFieldCreateDate = 21;
  wdFieldSaveDate = 22;
  wdFieldPrintDate = 23;
  wdFieldRevisionNum = 24;
  wdFieldEditTime = 25;
  wdFieldNumPages = 26;
  wdFieldNumWords = 27;
  wdFieldNumChars = 28;
  wdFieldFileName = 29;
  wdFieldTemplate = 30;
  wdFieldDate = 31;
  wdFieldTime = 32;
  wdFieldPage = 33;
  wdFieldExpression = 34;
  wdFieldQuote = 35;
  wdFieldInclude = 36;
  wdFieldPageRef = 37;
  wdFieldAsk = 38;
  wdFieldFillIn = 39;
  wdFieldData = 40;
  wdFieldNext = 41;
  wdFieldNextIf = 42;
  wdFieldSkipIf = 43;
  wdFieldMergeRec = 44;
  wdFieldDDE = 45;
  wdFieldDDEAuto = 46;
  wdFieldGlossary = 47;
  wdFieldPrint = 48;
  wdFieldFormula = 49;
  wdFieldGoToButton = 50;
  wdFieldMacroButton = 51;
  wdFieldAutoNumOutline = 52;
  wdFieldAutoNumLegal = 53;
  wdFieldAutoNum = 54;
  wdFieldImport = 55;
  wdFieldLink = 56;
  wdFieldSymbol = 57;
  wdFieldEmbed = 58;
  wdFieldMergeField = 59;
  wdFieldUserName = 60;
  wdFieldUserInitials = 61;
  wdFieldUserAddress = 62;
  wdFieldBarCode = 63;
  wdFieldDocVariable = 64;
  wdFieldSection = 65;
  wdFieldSectionPages = 66;
  wdFieldIncludePicture = 67;
  wdFieldIncludeText = 68;
  wdFieldFileSize = 69;
  wdFieldFormTextInput = 70;
  wdFieldFormCheckBox = 71;
  wdFieldNoteRef = 72;
  wdFieldTOA = 73;
  wdFieldTOAEntry = 74;
  wdFieldMergeSeq = 75;
  wdFieldPrivate = 77;
  wdFieldDatabase = 78;
  wdFieldAutoText = 79;
  wdFieldCompare = 80;
  wdFieldAddin = 81;
  wdFieldSubscriber = 82;
  wdFieldFormDropDown = 83;
  wdFieldAdvance = 84;
  wdFieldDocProperty = 85;
  wdFieldOCX = 87;
  wdFieldHyperlink = 88;
  wdFieldAutoTextList = 89;
  wdFieldListNum = 90;
  wdFieldHTMLActiveX = 91;
  wdFieldBidiOutline = 92;
  wdFieldAddressBlock = 93;
  wdFieldGreetingLine = 94;
  wdFieldShape = 95;

{ WdBuiltinStyle }

  wdStyleNormal = -1;
  wdStyleEnvelopeAddress = -37;
  wdStyleEnvelopeReturn = -38;
  wdStyleBodyText = -67;
  wdStyleHeading1 = -2;
  wdStyleHeading2 = -3;
  wdStyleHeading3 = -4;
  wdStyleHeading4 = -5;
  wdStyleHeading5 = -6;
  wdStyleHeading6 = -7;
  wdStyleHeading7 = -8;
  wdStyleHeading8 = -9;
  wdStyleHeading9 = -10;
  wdStyleIndex1 = -11;
  wdStyleIndex2 = -12;
  wdStyleIndex3 = -13;
  wdStyleIndex4 = -14;
  wdStyleIndex5 = -15;
  wdStyleIndex6 = -16;
  wdStyleIndex7 = -17;
  wdStyleIndex8 = -18;
  wdStyleIndex9 = -19;
  wdStyleTOC1 = -20;
  wdStyleTOC2 = -21;
  wdStyleTOC3 = -22;
  wdStyleTOC4 = -23;
  wdStyleTOC5 = -24;
  wdStyleTOC6 = -25;
  wdStyleTOC7 = -26;
  wdStyleTOC8 = -27;
  wdStyleTOC9 = -28;
  wdStyleNormalIndent = -29;
  wdStyleFootnoteText = -30;
  wdStyleCommentText = -31;
  wdStyleHeader = -32;
  wdStyleFooter = -33;
  wdStyleIndexHeading = -34;
  wdStyleCaption = -35;
  wdStyleTableOfFigures = -36;
  wdStyleFootnoteReference = -39;
  wdStyleCommentReference = -40;
  wdStyleLineNumber = -41;
  wdStylePageNumber = -42;
  wdStyleEndnoteReference = -43;
  wdStyleEndnoteText = -44;
  wdStyleTableOfAuthorities = -45;
  wdStyleMacroText = -46;
  wdStyleTOAHeading = -47;
  wdStyleList = -48;
  wdStyleListBullet = -49;
  wdStyleListNumber = -50;
  wdStyleList2 = -51;
  wdStyleList3 = -52;
  wdStyleList4 = -53;
  wdStyleList5 = -54;
  wdStyleListBullet2 = -55;
  wdStyleListBullet3 = -56;
  wdStyleListBullet4 = -57;
  wdStyleListBullet5 = -58;
  wdStyleListNumber2 = -59;
  wdStyleListNumber3 = -60;
  wdStyleListNumber4 = -61;
  wdStyleListNumber5 = -62;
  wdStyleTitle = -63;
  wdStyleClosing = -64;
  wdStyleSignature = -65;
  wdStyleDefaultParagraphFont = -66;
  wdStyleBodyTextIndent = -68;
  wdStyleListContinue = -69;
  wdStyleListContinue2 = -70;
  wdStyleListContinue3 = -71;
  wdStyleListContinue4 = -72;
  wdStyleListContinue5 = -73;
  wdStyleMessageHeader = -74;
  wdStyleSubtitle = -75;
  wdStyleSalutation = -76;
  wdStyleDate = -77;
  wdStyleBodyTextFirstIndent = -78;
  wdStyleBodyTextFirstIndent2 = -79;
  wdStyleNoteHeading = -80;
  wdStyleBodyText2 = -81;
  wdStyleBodyText3 = -82;
  wdStyleBodyTextIndent2 = -83;
  wdStyleBodyTextIndent3 = -84;
  wdStyleBlockQuotation = -85;
  wdStyleHyperlink = -86;
  wdStyleHyperlinkFollowed = -87;
  wdStyleStrong = -88;
  wdStyleEmphasis = -89;
  wdStyleNavPane = -90;
  wdStylePlainText = -91;
  wdStyleHtmlNormal = -95;
  wdStyleHtmlAcronym = -96;
  wdStyleHtmlAddress = -97;
  wdStyleHtmlCite = -98;
  wdStyleHtmlCode = -99;
  wdStyleHtmlDfn = -100;
  wdStyleHtmlKbd = -101;
  wdStyleHtmlPre = -102;
  wdStyleHtmlSamp = -103;
  wdStyleHtmlTt = -104;
  wdStyleHtmlVar = -105;
  wdStyleNormalTable = -106;

{ WdWordDialogTab }

  wdDialogToolsOptionsTabView = 204;
  wdDialogToolsOptionsTabGeneral = 203;
  wdDialogToolsOptionsTabEdit = 224;
  wdDialogToolsOptionsTabPrint = 208;
  wdDialogToolsOptionsTabSave = 209;
  wdDialogToolsOptionsTabProofread = 211;
  wdDialogToolsOptionsTabTrackChanges = 386;
  wdDialogToolsOptionsTabUserInfo = 213;
  wdDialogToolsOptionsTabCompatibility = 525;
  wdDialogToolsOptionsTabTypography = 739;
  wdDialogToolsOptionsTabFileLocations = 225;
  wdDialogToolsOptionsTabFuzzy = 790;
  wdDialogToolsOptionsTabHangulHanjaConversion = 786;
  wdDialogToolsOptionsTabBidi = 1029;
  wdDialogToolsOptionsTabSecurity = 1361;
  wdDialogFilePageSetupTabMargins = 150000;
  wdDialogFilePageSetupTabPaper = 150001;
  wdDialogFilePageSetupTabLayout = 150003;
  wdDialogFilePageSetupTabCharsLines = 150004;
  wdDialogInsertSymbolTabSymbols = 200000;
  wdDialogInsertSymbolTabSpecialCharacters = 200001;
  wdDialogNoteOptionsTabAllFootnotes = 300000;
  wdDialogNoteOptionsTabAllEndnotes = 300001;
  wdDialogInsertIndexAndTablesTabIndex = 400000;
  wdDialogInsertIndexAndTablesTabTableOfContents = 400001;
  wdDialogInsertIndexAndTablesTabTableOfFigures = 400002;
  wdDialogInsertIndexAndTablesTabTableOfAuthorities = 400003;
  wdDialogOrganizerTabStyles = 500000;
  wdDialogOrganizerTabAutoText = 500001;
  wdDialogOrganizerTabCommandBars = 500002;
  wdDialogOrganizerTabMacros = 500003;
  wdDialogFormatFontTabFont = 600000;
  wdDialogFormatFontTabCharacterSpacing = 600001;
  wdDialogFormatFontTabAnimation = 600002;
  wdDialogFormatBordersAndShadingTabBorders = 700000;
  wdDialogFormatBordersAndShadingTabPageBorder = 700001;
  wdDialogFormatBordersAndShadingTabShading = 700002;
  wdDialogToolsEnvelopesAndLabelsTabEnvelopes = 800000;
  wdDialogToolsEnvelopesAndLabelsTabLabels = 800001;
  wdDialogFormatParagraphTabIndentsAndSpacing = 1000000;
  wdDialogFormatParagraphTabTextFlow = 1000001;
  wdDialogFormatParagraphTabTeisai = 1000002;
  wdDialogFormatDrawingObjectTabColorsAndLines = 1200000;
  wdDialogFormatDrawingObjectTabSize = 1200001;
  wdDialogFormatDrawingObjectTabPosition = 1200002;
  wdDialogFormatDrawingObjectTabWrapping = 1200003;
  wdDialogFormatDrawingObjectTabPicture = 1200004;
  wdDialogFormatDrawingObjectTabTextbox = 1200005;
  wdDialogFormatDrawingObjectTabWeb = 1200006;
  wdDialogFormatDrawingObjectTabHR = 1200007;
  wdDialogToolsAutoCorrectExceptionsTabFirstLetter = 1400000;
  wdDialogToolsAutoCorrectExceptionsTabInitialCaps = 1400001;
  wdDialogToolsAutoCorrectExceptionsTabHangulAndAlphabet = 1400002;
  wdDialogToolsAutoCorrectExceptionsTabIac = 1400003;
  wdDialogFormatBulletsAndNumberingTabBulleted = 1500000;
  wdDialogFormatBulletsAndNumberingTabNumbered = 1500001;
  wdDialogFormatBulletsAndNumberingTabOutlineNumbered = 1500002;
  wdDialogLetterWizardTabLetterFormat = 1600000;
  wdDialogLetterWizardTabRecipientInfo = 1600001;
  wdDialogLetterWizardTabOtherElements = 1600002;
  wdDialogLetterWizardTabSenderInfo = 1600003;
  wdDialogToolsAutoManagerTabAutoCorrect = 1700000;
  wdDialogToolsAutoManagerTabAutoFormatAsYouType = 1700001;
  wdDialogToolsAutoManagerTabAutoText = 1700002;
  wdDialogToolsAutoManagerTabAutoFormat = 1700003;
  wdDialogToolsAutoManagerTabSmartTags = 1700004;
  wdDialogTablePropertiesTabTable = 1800000;
  wdDialogTablePropertiesTabRow = 1800001;
  wdDialogTablePropertiesTabColumn = 1800002;
  wdDialogTablePropertiesTabCell = 1800003;
  wdDialogEmailOptionsTabSignature = 1900000;
  wdDialogEmailOptionsTabStationary = 1900001;
  wdDialogEmailOptionsTabQuoting = 1900002;
  wdDialogWebOptionsBrowsers = 2000000;
  wdDialogWebOptionsGeneral = 2000000;
  wdDialogWebOptionsFiles = 2000001;
  wdDialogWebOptionsPictures = 2000002;
  wdDialogWebOptionsEncoding = 2000003;
  wdDialogWebOptionsFonts = 2000004;
  wdDialogToolsOptionsTabAcetate = 1266;

{ WdWordDialogTabHID }

  wdDialogFilePageSetupTabPaperSize = 150001;
  wdDialogFilePageSetupTabPaperSource = 150002;

{ WdWordDialog }

  wdDialogHelpAbout = 9;
  wdDialogHelpWordPerfectHelp = 10;
  wdDialogDocumentStatistics = 78;
  wdDialogFileNew = 79;
  wdDialogFileOpen = 80;
  wdDialogMailMergeOpenDataSource = 81;
  wdDialogMailMergeOpenHeaderSource = 82;
  wdDialogFileSaveAs = 84;
  wdDialogFileSummaryInfo = 86;
  wdDialogToolsTemplates = 87;
  wdDialogFilePrint = 88;
  wdDialogFilePrintSetup = 97;
  wdDialogFileFind = 99;
  wdDialogFormatAddrFonts = 103;
  wdDialogEditPasteSpecial = 111;
  wdDialogEditFind = 112;
  wdDialogEditReplace = 117;
  wdDialogEditStyle = 120;
  wdDialogEditLinks = 124;
  wdDialogEditObject = 125;
  wdDialogTableToText = 128;
  wdDialogTextToTable = 127;
  wdDialogTableInsertTable = 129;
  wdDialogTableInsertCells = 130;
  wdDialogTableInsertRow = 131;
  wdDialogTableDeleteCells = 133;
  wdDialogTableSplitCells = 137;
  wdDialogTableRowHeight = 142;
  wdDialogTableColumnWidth = 143;
  wdDialogToolsCustomize = 152;
  wdDialogInsertBreak = 159;
  wdDialogInsertSymbol = 162;
  wdDialogInsertPicture = 163;
  wdDialogInsertFile = 164;
  wdDialogInsertDateTime = 165;
  wdDialogInsertField = 166;
  wdDialogInsertMergeField = 167;
  wdDialogInsertBookmark = 168;
  wdDialogMarkIndexEntry = 169;
  wdDialogInsertIndex = 170;
  wdDialogInsertTableOfContents = 171;
  wdDialogInsertObject = 172;
  wdDialogToolsCreateEnvelope = 173;
  wdDialogFormatFont = 174;
  wdDialogFormatParagraph = 175;
  wdDialogFormatSectionLayout = 176;
  wdDialogFormatColumns = 177;
  wdDialogFileDocumentLayout = 178;
  wdDialogFilePageSetup = 178;
  wdDialogFormatTabs = 179;
  wdDialogFormatStyle = 180;
  wdDialogFormatDefineStyleFont = 181;
  wdDialogFormatDefineStylePara = 182;
  wdDialogFormatDefineStyleTabs = 183;
  wdDialogFormatDefineStyleFrame = 184;
  wdDialogFormatDefineStyleBorders = 185;
  wdDialogFormatDefineStyleLang = 186;
  wdDialogFormatPicture = 187;
  wdDialogToolsLanguage = 188;
  wdDialogFormatBordersAndShading = 189;
  wdDialogFormatFrame = 190;
  wdDialogToolsThesaurus = 194;
  wdDialogToolsHyphenation = 195;
  wdDialogToolsBulletsNumbers = 196;
  wdDialogToolsHighlightChanges = 197;
  wdDialogToolsRevisions = 197;
  wdDialogToolsCompareDocuments = 198;
  wdDialogTableSort = 199;
  wdDialogToolsOptionsGeneral = 203;
  wdDialogToolsOptionsView = 204;
  wdDialogToolsAdvancedSettings = 206;
  wdDialogToolsOptionsPrint = 208;
  wdDialogToolsOptionsSave = 209;
  wdDialogToolsOptionsSpellingAndGrammar = 211;
  wdDialogToolsOptionsUserInfo = 213;
  wdDialogToolsMacroRecord = 214;
  wdDialogToolsMacro = 215;
  wdDialogWindowActivate = 220;
  wdDialogFormatRetAddrFonts = 221;
  wdDialogOrganizer = 222;
  wdDialogToolsOptionsEdit = 224;
  wdDialogToolsOptionsFileLocations = 225;
  wdDialogToolsWordCount = 228;
  wdDialogControlRun = 235;
  wdDialogInsertPageNumbers = 294;
  wdDialogFormatPageNumber = 298;
  wdDialogCopyFile = 300;
  wdDialogFormatChangeCase = 322;
  wdDialogUpdateTOC = 331;
  wdDialogInsertDatabase = 341;
  wdDialogTableFormula = 348;
  wdDialogFormFieldOptions = 353;
  wdDialogInsertCaption = 357;
  wdDialogInsertCaptionNumbering = 358;
  wdDialogInsertAutoCaption = 359;
  wdDialogFormFieldHelp = 361;
  wdDialogInsertCrossReference = 367;
  wdDialogInsertFootnote = 370;
  wdDialogNoteOptions = 373;
  wdDialogToolsAutoCorrect = 378;
  wdDialogToolsOptionsTrackChanges = 386;
  wdDialogConvertObject = 392;
  wdDialogInsertAddCaption = 402;
  wdDialogConnect = 420;
  wdDialogToolsCustomizeKeyboard = 432;
  wdDialogToolsCustomizeMenus = 433;
  wdDialogToolsMergeDocuments = 435;
  wdDialogMarkTableOfContentsEntry = 442;
  wdDialogFileMacPageSetupGX = 444;
  wdDialogFilePrintOneCopy = 445;
  wdDialogEditFrame = 458;
  wdDialogMarkCitation = 463;
  wdDialogTableOfContentsOptions = 470;
  wdDialogInsertTableOfAuthorities = 471;
  wdDialogInsertTableOfFigures = 472;
  wdDialogInsertIndexAndTables = 473;
  wdDialogInsertFormField = 483;
  wdDialogFormatDropCap = 488;
  wdDialogToolsCreateLabels = 489;
  wdDialogToolsProtectDocument = 503;
  wdDialogFormatStyleGallery = 505;
  wdDialogToolsAcceptRejectChanges = 506;
  wdDialogHelpWordPerfectHelpOptions = 511;
  wdDialogToolsUnprotectDocument = 521;
  wdDialogToolsOptionsCompatibility = 525;
  wdDialogTableOfCaptionsOptions = 551;
  wdDialogTableAutoFormat = 563;
  wdDialogMailMergeFindRecord = 569;
  wdDialogReviewAfmtRevisions = 570;
  wdDialogViewZoom = 577;
  wdDialogToolsProtectSection = 578;
  wdDialogFontSubstitution = 581;
  wdDialogInsertSubdocument = 583;
  wdDialogNewToolbar = 586;
  wdDialogToolsEnvelopesAndLabels = 607;
  wdDialogFormatCallout = 610;
  wdDialogTableFormatCell = 612;
  wdDialogToolsCustomizeMenuBar = 615;
  wdDialogFileRoutingSlip = 624;
  wdDialogEditTOACategory = 625;
  wdDialogToolsManageFields = 631;
  wdDialogDrawSnapToGrid = 633;
  wdDialogDrawAlign = 634;
  wdDialogMailMergeCreateDataSource = 642;
  wdDialogMailMergeCreateHeaderSource = 643;
  wdDialogMailMerge = 676;
  wdDialogMailMergeCheck = 677;
  wdDialogMailMergeHelper = 680;
  wdDialogMailMergeQueryOptions = 681;
  wdDialogFileMacPageSetup = 685;
  wdDialogListCommands = 723;
  wdDialogEditCreatePublisher = 732;
  wdDialogEditSubscribeTo = 733;
  wdDialogEditPublishOptions = 735;
  wdDialogEditSubscribeOptions = 736;
  wdDialogFileMacCustomPageSetupGX = 737;
  wdDialogToolsOptionsTypography = 739;
  wdDialogToolsAutoCorrectExceptions = 762;
  wdDialogToolsOptionsAutoFormatAsYouType = 778;
  wdDialogMailMergeUseAddressBook = 779;
  wdDialogToolsHangulHanjaConversion = 784;
  wdDialogToolsOptionsFuzzy = 790;
  wdDialogEditGoToOld = 811;
  wdDialogInsertNumber = 812;
  wdDialogLetterWizard = 821;
  wdDialogFormatBulletsAndNumbering = 824;
  wdDialogToolsSpellingAndGrammar = 828;
  wdDialogToolsCreateDirectory = 833;
  wdDialogTableWrapping = 854;
  wdDialogFormatTheme = 855;
  wdDialogTableProperties = 861;
  wdDialogEmailOptions = 863;
  wdDialogCreateAutoText = 872;
  wdDialogToolsAutoSummarize = 874;
  wdDialogToolsGrammarSettings = 885;
  wdDialogEditGoTo = 896;
  wdDialogWebOptions = 898;
  wdDialogInsertHyperlink = 925;
  wdDialogToolsAutoManager = 915;
  wdDialogFileVersions = 945;
  wdDialogToolsOptionsAutoFormat = 959;
  wdDialogFormatDrawingObject = 960;
  wdDialogToolsOptions = 974;
  wdDialogFitText = 983;
  wdDialogEditAutoText = 985;
  wdDialogPhoneticGuide = 986;
  wdDialogToolsDictionary = 989;
  wdDialogFileSaveVersion = 1007;
  wdDialogToolsOptionsBidi = 1029;
  wdDialogFrameSetProperties = 1074;
  wdDialogTableTableOptions = 1080;
  wdDialogTableCellOptions = 1081;
  wdDialogIMESetDefault = 1094;
  wdDialogTCSCTranslator = 1156;
  wdDialogHorizontalInVertical = 1160;
  wdDialogTwoLinesInOne = 1161;
  wdDialogFormatEncloseCharacters = 1162;
  wdDialogConsistencyChecker = 1121;
  wdDialogToolsOptionsSmartTag = 1395;
  wdDialogFormatStylesCustom = 1248;
  wdDialogCSSLinks = 1261;
  wdDialogInsertWebComponent = 1324;
  wdDialogToolsOptionsEditCopyPaste = 1356;
  wdDialogToolsOptionsSecurity = 1361;
  wdDialogSearch = 1363;
  wdDialogShowRepairs = 1381;
  wdDialogMailMergeInsertAsk = 4047;
  wdDialogMailMergeInsertFillIn = 4048;
  wdDialogMailMergeInsertIf = 4049;
  wdDialogMailMergeInsertNextIf = 4053;
  wdDialogMailMergeInsertSet = 4054;
  wdDialogMailMergeInsertSkipIf = 4055;
  wdDialogMailMergeFieldMapping = 1304;
  wdDialogMailMergeInsertAddressBlock = 1305;
  wdDialogMailMergeInsertGreetingLine = 1306;
  wdDialogMailMergeInsertFields = 1307;
  wdDialogMailMergeRecipients = 1308;
  wdDialogMailMergeFindRecipient = 1326;
  wdDialogMailMergeSetDocumentType = 1339;

{ WdWordDialogHID }

//emptyenum = 0;

{ WdFieldKind }

  wdFieldKindNone = 0;
  wdFieldKindHot = 1;
  wdFieldKindWarm = 2;
  wdFieldKindCold = 3;

{ WdTextFormFieldType }

  wdRegularText = 0;
  wdNumberText = 1;
  wdDateText = 2;
  wdCurrentDateText = 3;
  wdCurrentTimeText = 4;
  wdCalculationText = 5;

{ WdChevronConvertRule }

  wdNeverConvert = 0;
  wdAlwaysConvert = 1;
  wdAskToNotConvert = 2;
  wdAskToConvert = 3;

{ WdMailMergeMainDocType }

  wdNotAMergeDocument = -1;
  wdFormLetters = 0;
  wdMailingLabels = 1;
  wdEnvelopes = 2;
  wdCatalog = 3;
  wdEMail = 4;
  wdFax = 5;
  wdDirectory = 3;

{ WdMailMergeState }

  wdNormalDocument = 0;
  wdMainDocumentOnly = 1;
  wdMainAndDataSource = 2;
  wdMainAndHeader = 3;
  wdMainAndSourceAndHeader = 4;
  wdDataSource = 5;

{ WdMailMergeDestination }

  wdSendToNewDocument = 0;
  wdSendToPrinter = 1;
  wdSendToEmail = 2;
  wdSendToFax = 3;

{ WdMailMergeActiveRecord }

  wdNoActiveRecord = -1;
  wdNextRecord = -2;
  wdPreviousRecord = -3;
  wdFirstRecord = -4;
  wdLastRecord = -5;
  wdFirstDataSourceRecord = -6;
  wdLastDataSourceRecord = -7;
  wdNextDataSourceRecord = -8;
  wdPreviousDataSourceRecord = -9;

{ WdMailMergeDefaultRecord }

  wdDefaultFirstRecord = 1;
  wdDefaultLastRecord = -16;

{ WdMailMergeDataSource }

  wdNoMergeInfo = -1;
  wdMergeInfoFromWord = 0;
  wdMergeInfoFromAccessDDE = 1;
  wdMergeInfoFromExcelDDE = 2;
  wdMergeInfoFromMSQueryDDE = 3;
  wdMergeInfoFromODBC = 4;
  wdMergeInfoFromODSO = 5;

{ WdMailMergeComparison }

  wdMergeIfEqual = 0;
  wdMergeIfNotEqual = 1;
  wdMergeIfLessThan = 2;
  wdMergeIfGreaterThan = 3;
  wdMergeIfLessThanOrEqual = 4;
  wdMergeIfGreaterThanOrEqual = 5;
  wdMergeIfIsBlank = 6;
  wdMergeIfIsNotBlank = 7;

{ WdBookmarkSortBy }

  wdSortByName = 0;
  wdSortByLocation = 1;

{ WdWindowState }

  wdWindowStateNormal = 0;
  wdWindowStateMaximize = 1;
  wdWindowStateMinimize = 2;

{ WdPictureLinkType }

  wdLinkNone = 0;
  wdLinkDataInDoc = 1;
  wdLinkDataOnDisk = 2;

{ WdLinkType }

  wdLinkTypeOLE = 0;
  wdLinkTypePicture = 1;
  wdLinkTypeText = 2;
  wdLinkTypeReference = 3;
  wdLinkTypeInclude = 4;
  wdLinkTypeImport = 5;
  wdLinkTypeDDE = 6;
  wdLinkTypeDDEAuto = 7;

{ WdWindowType }

  wdWindowDocument = 0;
  wdWindowTemplate = 1;

{ WdViewType }

  wdNormalView = 1;
  wdOutlineView = 2;
  wdPrintView = 3;
  wdPrintPreview = 4;
  wdMasterView = 5;
  wdWebView = 6;

{ WdSeekView }

  wdSeekMainDocument = 0;
  wdSeekPrimaryHeader = 1;
  wdSeekFirstPageHeader = 2;
  wdSeekEvenPagesHeader = 3;
  wdSeekPrimaryFooter = 4;
  wdSeekFirstPageFooter = 5;
  wdSeekEvenPagesFooter = 6;
  wdSeekFootnotes = 7;
  wdSeekEndnotes = 8;
  wdSeekCurrentPageHeader = 9;
  wdSeekCurrentPageFooter = 10;

{ WdSpecialPane }

  wdPaneNone = 0;
  wdPanePrimaryHeader = 1;
  wdPaneFirstPageHeader = 2;
  wdPaneEvenPagesHeader = 3;
  wdPanePrimaryFooter = 4;
  wdPaneFirstPageFooter = 5;
  wdPaneEvenPagesFooter = 6;
  wdPaneFootnotes = 7;
  wdPaneEndnotes = 8;
  wdPaneFootnoteContinuationNotice = 9;
  wdPaneFootnoteContinuationSeparator = 10;
  wdPaneFootnoteSeparator = 11;
  wdPaneEndnoteContinuationNotice = 12;
  wdPaneEndnoteContinuationSeparator = 13;
  wdPaneEndnoteSeparator = 14;
  wdPaneComments = 15;
  wdPaneCurrentPageHeader = 16;
  wdPaneCurrentPageFooter = 17;
  wdPaneRevisions = 18;

{ WdPageFit }

  wdPageFitNone = 0;
  wdPageFitFullPage = 1;
  wdPageFitBestFit = 2;
  wdPageFitTextFit = 3;

{ WdBrowseTarget }

  wdBrowsePage = 1;
  wdBrowseSection = 2;
  wdBrowseComment = 3;
  wdBrowseFootnote = 4;
  wdBrowseEndnote = 5;
  wdBrowseField = 6;
  wdBrowseTable = 7;
  wdBrowseGraphic = 8;
  wdBrowseHeading = 9;
  wdBrowseEdit = 10;
  wdBrowseFind = 11;
  wdBrowseGoTo = 12;

{ WdPaperTray }

  wdPrinterDefaultBin = 0;
  wdPrinterUpperBin = 1;
  wdPrinterOnlyBin = 1;
  wdPrinterLowerBin = 2;
  wdPrinterMiddleBin = 3;
  wdPrinterManualFeed = 4;
  wdPrinterEnvelopeFeed = 5;
  wdPrinterManualEnvelopeFeed = 6;
  wdPrinterAutomaticSheetFeed = 7;
  wdPrinterTractorFeed = 8;
  wdPrinterSmallFormatBin = 9;
  wdPrinterLargeFormatBin = 10;
  wdPrinterLargeCapacityBin = 11;
  wdPrinterPaperCassette = 14;
  wdPrinterFormSource = 15;

{ WdOrientation }

  wdOrientPortrait = 0;
  wdOrientLandscape = 1;

{ WdSelectionType }

  wdNoSelection = 0;
  wdSelectionIP = 1;
  wdSelectionNormal = 2;
  wdSelectionFrame = 3;
  wdSelectionColumn = 4;
  wdSelectionRow = 5;
  wdSelectionBlock = 6;
  wdSelectionInlineShape = 7;
  wdSelectionShape = 8;

{ WdCaptionLabelID }

  wdCaptionFigure = -1;
  wdCaptionTable = -2;
  wdCaptionEquation = -3;

{ WdReferenceType }

  wdRefTypeNumberedItem = 0;
  wdRefTypeHeading = 1;
  wdRefTypeBookmark = 2;
  wdRefTypeFootnote = 3;
  wdRefTypeEndnote = 4;

{ WdReferenceKind }

  wdContentText = -1;
  wdNumberRelativeContext = -2;
  wdNumberNoContext = -3;
  wdNumberFullContext = -4;
  wdEntireCaption = 2;
  wdOnlyLabelAndNumber = 3;
  wdOnlyCaptionText = 4;
  wdFootnoteNumber = 5;
  wdEndnoteNumber = 6;
  wdPageNumber = 7;
  wdPosition = 15;
  wdFootnoteNumberFormatted = 16;
  wdEndnoteNumberFormatted = 17;

{ WdIndexFormat }

  wdIndexTemplate = 0;
  wdIndexClassic = 1;
  wdIndexFancy = 2;
  wdIndexModern = 3;
  wdIndexBulleted = 4;
  wdIndexFormal = 5;
  wdIndexSimple = 6;

{ WdIndexType }

  wdIndexIndent = 0;
  wdIndexRunin = 1;

{ WdRevisionsWrap }

  wdWrapNever = 0;
  wdWrapAlways = 1;
  wdWrapAsk = 2;

{ WdRevisionType }

  wdNoRevision = 0;
  wdRevisionInsert = 1;
  wdRevisionDelete = 2;
  wdRevisionProperty = 3;
  wdRevisionParagraphNumber = 4;
  wdRevisionDisplayField = 5;
  wdRevisionReconcile = 6;
  wdRevisionConflict = 7;
  wdRevisionStyle = 8;
  wdRevisionReplace = 9;
  wdRevisionParagraphProperty = 10;
  wdRevisionTableProperty = 11;
  wdRevisionSectionProperty = 12;
  wdRevisionStyleDefinition = 13;

{ WdRoutingSlipDelivery }

  wdOneAfterAnother = 0;
  wdAllAtOnce = 1;

{ WdRoutingSlipStatus }

  wdNotYetRouted = 0;
  wdRouteInProgress = 1;
  wdRouteComplete = 2;

{ WdSectionStart }

  wdSectionContinuous = 0;
  wdSectionNewColumn = 1;
  wdSectionNewPage = 2;
  wdSectionEvenPage = 3;
  wdSectionOddPage = 4;

{ WdSaveOptions }

  wdDoNotSaveChanges = 0;
  wdSaveChanges = -1;
  wdPromptToSaveChanges = -2;

{ WdDocumentKind }

  wdDocumentNotSpecified = 0;
  wdDocumentLetter = 1;
  wdDocumentEmail = 2;

{ WdDocumentType }

  wdTypeDocument = 0;
  wdTypeTemplate = 1;
  wdTypeFrameset = 2;

{ WdOriginalFormat }

  wdWordDocument = 0;
  wdOriginalDocumentFormat = 1;
  wdPromptUser = 2;

{ WdRelocate }

  wdRelocateUp = 0;
  wdRelocateDown = 1;

{ WdInsertedTextMark }

  wdInsertedTextMarkNone = 0;
  wdInsertedTextMarkBold = 1;
  wdInsertedTextMarkItalic = 2;
  wdInsertedTextMarkUnderline = 3;
  wdInsertedTextMarkDoubleUnderline = 4;
  wdInsertedTextMarkColorOnly = 5;
  wdInsertedTextMarkStrikeThrough = 6;

{ WdRevisedLinesMark }

  wdRevisedLinesMarkNone = 0;
  wdRevisedLinesMarkLeftBorder = 1;
  wdRevisedLinesMarkRightBorder = 2;
  wdRevisedLinesMarkOutsideBorder = 3;

{ WdDeletedTextMark }

  wdDeletedTextMarkHidden = 0;
  wdDeletedTextMarkStrikeThrough = 1;
  wdDeletedTextMarkCaret = 2;
  wdDeletedTextMarkPound = 3;
  wdDeletedTextMarkNone = 4;
  wdDeletedTextMarkBold = 5;
  wdDeletedTextMarkItalic = 6;
  wdDeletedTextMarkUnderline = 7;
  wdDeletedTextMarkDoubleUnderline = 8;
  wdDeletedTextMarkColorOnly = 9;

{ WdRevisedPropertiesMark }

  wdRevisedPropertiesMarkNone = 0;
  wdRevisedPropertiesMarkBold = 1;
  wdRevisedPropertiesMarkItalic = 2;
  wdRevisedPropertiesMarkUnderline = 3;
  wdRevisedPropertiesMarkDoubleUnderline = 4;
  wdRevisedPropertiesMarkColorOnly = 5;
  wdRevisedPropertiesMarkStrikeThrough = 6;

{ WdFieldShading }

  wdFieldShadingNever = 0;
  wdFieldShadingAlways = 1;
  wdFieldShadingWhenSelected = 2;

{ WdDefaultFilePath }

  wdDocumentsPath = 0;
  wdPicturesPath = 1;
  wdUserTemplatesPath = 2;
  wdWorkgroupTemplatesPath = 3;
  wdUserOptionsPath = 4;
  wdAutoRecoverPath = 5;
  wdToolsPath = 6;
  wdTutorialPath = 7;
  wdStartupPath = 8;
  wdProgramPath = 9;
  wdGraphicsFiltersPath = 10;
  wdTextConvertersPath = 11;
  wdProofingToolsPath = 12;
  wdTempFilePath = 13;
  wdCurrentFolderPath = 14;
  wdStyleGalleryPath = 15;
  wdBorderArtPath = 19;

{ WdCompatibility }

  wdNoTabHangIndent = 1;
  wdNoSpaceRaiseLower = 2;
  wdPrintColBlack = 3;
  wdWrapTrailSpaces = 4;
  wdNoColumnBalance = 5;
  wdConvMailMergeEsc = 6;
  wdSuppressSpBfAfterPgBrk = 7;
  wdSuppressTopSpacing = 8;
  wdOrigWordTableRules = 9;
  wdTransparentMetafiles = 10;
  wdShowBreaksInFrames = 11;
  wdSwapBordersFacingPages = 12;
  wdLeaveBackslashAlone = 13;
  wdExpandShiftReturn = 14;
  wdDontULTrailSpace = 15;
  wdDontBalanceSingleByteDoubleByteWidth = 16;
  wdSuppressTopSpacingMac5 = 17;
  wdSpacingInWholePoints = 18;
  wdPrintBodyTextBeforeHeader = 19;
  wdNoLeading = 20;
  wdNoSpaceForUL = 21;
  wdMWSmallCaps = 22;
  wdNoExtraLineSpacing = 23;
  wdTruncateFontHeight = 24;
  wdSubFontBySize = 25;
  wdUsePrinterMetrics = 26;
  wdWW6BorderRules = 27;
  wdExactOnTop = 28;
  wdSuppressBottomSpacing = 29;
  wdWPSpaceWidth = 30;
  wdWPJustification = 31;
  wdLineWrapLikeWord6 = 32;
  wdShapeLayoutLikeWW8 = 33;
  wdFootnoteLayoutLikeWW8 = 34;
  wdDontUseHTMLParagraphAutoSpacing = 35;
  wdDontAdjustLineHeightInTable = 36;
  wdForgetLastTabAlignment = 37;
  wdAutospaceLikeWW7 = 38;
  wdAlignTablesRowByRow = 39;
  wdLayoutRawTableWidth = 40;
  wdLayoutTableRowsApart = 41;
  wdUseWord97LineBreakingRules = 42;
  wdDontBreakWrappedTables = 43;
  wdDontSnapTextToGridInTableWithObjects = 44;
  wdSelectFieldWithFirstOrLastCharacter = 45;
  wdApplyBreakingRules = 46;
  wdDontWrapTextWithPunctuation = 47;
  wdDontUseAsianBreakRulesInGrid = 48;

{ WdPaperSize }

  wdPaper10x14 = 0;
  wdPaper11x17 = 1;
  wdPaperLetter = 2;
  wdPaperLetterSmall = 3;
  wdPaperLegal = 4;
  wdPaperExecutive = 5;
  wdPaperA3 = 6;
  wdPaperA4 = 7;
  wdPaperA4Small = 8;
  wdPaperA5 = 9;
  wdPaperB4 = 10;
  wdPaperB5 = 11;
  wdPaperCSheet = 12;
  wdPaperDSheet = 13;
  wdPaperESheet = 14;
  wdPaperFanfoldLegalGerman = 15;
  wdPaperFanfoldStdGerman = 16;
  wdPaperFanfoldUS = 17;
  wdPaperFolio = 18;
  wdPaperLedger = 19;
  wdPaperNote = 20;
  wdPaperQuarto = 21;
  wdPaperStatement = 22;
  wdPaperTabloid = 23;
  wdPaperEnvelope9 = 24;
  wdPaperEnvelope10 = 25;
  wdPaperEnvelope11 = 26;
  wdPaperEnvelope12 = 27;
  wdPaperEnvelope14 = 28;
  wdPaperEnvelopeB4 = 29;
  wdPaperEnvelopeB5 = 30;
  wdPaperEnvelopeB6 = 31;
  wdPaperEnvelopeC3 = 32;
  wdPaperEnvelopeC4 = 33;
  wdPaperEnvelopeC5 = 34;
  wdPaperEnvelopeC6 = 35;
  wdPaperEnvelopeC65 = 36;
  wdPaperEnvelopeDL = 37;
  wdPaperEnvelopeItaly = 38;
  wdPaperEnvelopeMonarch = 39;
  wdPaperEnvelopePersonal = 40;
  wdPaperCustom = 41;

{ WdCustomLabelPageSize }

  wdCustomLabelLetter = 0;
  wdCustomLabelLetterLS = 1;
  wdCustomLabelA4 = 2;
  wdCustomLabelA4LS = 3;
  wdCustomLabelA5 = 4;
  wdCustomLabelA5LS = 5;
  wdCustomLabelB5 = 6;
  wdCustomLabelMini = 7;
  wdCustomLabelFanfold = 8;
  wdCustomLabelVertHalfSheet = 9;
  wdCustomLabelVertHalfSheetLS = 10;
  wdCustomLabelHigaki = 11;
  wdCustomLabelHigakiLS = 12;
  wdCustomLabelB4JIS = 13;

{ WdProtectionType }

  wdNoProtection = -1;
  wdAllowOnlyRevisions = 0;
  wdAllowOnlyComments = 1;
  wdAllowOnlyFormFields = 2;

{ WdPartOfSpeech }

  wdAdjective = 0;
  wdNoun = 1;
  wdAdverb = 2;
  wdVerb = 3;
  wdPronoun = 4;
  wdConjunction = 5;
  wdPreposition = 6;
  wdInterjection = 7;
  wdIdiom = 8;
  wdOther = 9;

{ WdSubscriberFormats }

  wdSubscriberBestFormat = 0;
  wdSubscriberRTF = 1;
  wdSubscriberText = 2;
  wdSubscriberPict = 4;

{ WdEditionType }

  wdPublisher = 0;
  wdSubscriber = 1;

{ WdEditionOption }

  wdCancelPublisher = 0;
  wdSendPublisher = 1;
  wdSelectPublisher = 2;
  wdAutomaticUpdate = 3;
  wdManualUpdate = 4;
  wdChangeAttributes = 5;
  wdUpdateSubscriber = 6;
  wdOpenSource = 7;

{ WdRelativeHorizontalPosition }

  wdRelativeHorizontalPositionMargin = 0;
  wdRelativeHorizontalPositionPage = 1;
  wdRelativeHorizontalPositionColumn = 2;
  wdRelativeHorizontalPositionCharacter = 3;

{ WdRelativeVerticalPosition }

  wdRelativeVerticalPositionMargin = 0;
  wdRelativeVerticalPositionPage = 1;
  wdRelativeVerticalPositionParagraph = 2;
  wdRelativeVerticalPositionLine = 3;

{ WdHelpType }

  wdHelp = 0;
  wdHelpAbout = 1;
  wdHelpActiveWindow = 2;
  wdHelpContents = 3;
  wdHelpExamplesAndDemos = 4;
  wdHelpIndex = 5;
  wdHelpKeyboard = 6;
  wdHelpPSSHelp = 7;
  wdHelpQuickPreview = 8;
  wdHelpSearch = 9;
  wdHelpUsingHelp = 10;
  wdHelpIchitaro = 11;
  wdHelpPE2 = 12;
  wdHelpHWP = 13;

{ WdHelpTypeHID }

//emptyenum = 0;

{ WdKeyCategory }

  wdKeyCategoryNil = -1;
  wdKeyCategoryDisable = 0;
  wdKeyCategoryCommand = 1;
  wdKeyCategoryMacro = 2;
  wdKeyCategoryFont = 3;
  wdKeyCategoryAutoText = 4;
  wdKeyCategoryStyle = 5;
  wdKeyCategorySymbol = 6;
  wdKeyCategoryPrefix = 7;

{ WdKey }

  wdNoKey = 255;
  wdKeyShift = 256;
  wdKeyControl = 512;
  wdKeyCommand = 512;
  wdKeyAlt = 1024;
  wdKeyOption = 1024;
  wdKeyA = 65;
  wdKeyB = 66;
  wdKeyC = 67;
  wdKeyD = 68;
  wdKeyE = 69;
  wdKeyF = 70;
  wdKeyG = 71;
  wdKeyH = 72;
  wdKeyI = 73;
  wdKeyJ = 74;
  wdKeyK = 75;
  wdKeyL = 76;
  wdKeyM = 77;
  wdKeyN = 78;
  wdKeyO = 79;
  wdKeyP = 80;
  wdKeyQ = 81;
  wdKeyR = 82;
  wdKeyS = 83;
  wdKeyT = 84;
  wdKeyU = 85;
  wdKeyV = 86;
  wdKeyW = 87;
  wdKeyX = 88;
  wdKeyY = 89;
  wdKeyZ = 90;
  wdKey0 = 48;
  wdKey1 = 49;
  wdKey2 = 50;
  wdKey3 = 51;
  wdKey4 = 52;
  wdKey5 = 53;
  wdKey6 = 54;
  wdKey7 = 55;
  wdKey8 = 56;
  wdKey9 = 57;
  wdKeyBackspace = 8;
  wdKeyTab = 9;
  wdKeyNumeric5Special = 12;
  wdKeyReturn = 13;
  wdKeyPause = 19;
  wdKeyEsc = 27;
  wdKeySpacebar = 32;
  wdKeyPageUp = 33;
  wdKeyPageDown = 34;
  wdKeyEnd = 35;
  wdKeyHome = 36;
  wdKeyInsert = 45;
  wdKeyDelete = 46;
  wdKeyNumeric0 = 96;
  wdKeyNumeric1 = 97;
  wdKeyNumeric2 = 98;
  wdKeyNumeric3 = 99;
  wdKeyNumeric4 = 100;
  wdKeyNumeric5 = 101;
  wdKeyNumeric6 = 102;
  wdKeyNumeric7 = 103;
  wdKeyNumeric8 = 104;
  wdKeyNumeric9 = 105;
  wdKeyNumericMultiply = 106;
  wdKeyNumericAdd = 107;
  wdKeyNumericSubtract = 109;
  wdKeyNumericDecimal = 110;
  wdKeyNumericDivide = 111;
  wdKeyF1 = 112;
  wdKeyF2 = 113;
  wdKeyF3 = 114;
  wdKeyF4 = 115;
  wdKeyF5 = 116;
  wdKeyF6 = 117;
  wdKeyF7 = 118;
  wdKeyF8 = 119;
  wdKeyF9 = 120;
  wdKeyF10 = 121;
  wdKeyF11 = 122;
  wdKeyF12 = 123;
  wdKeyF13 = 124;
  wdKeyF14 = 125;
  wdKeyF15 = 126;
  wdKeyF16 = 127;
  wdKeyScrollLock = 145;
  wdKeySemiColon = 186;
  wdKeyEquals = 187;
  wdKeyComma = 188;
  wdKeyHyphen = 189;
  wdKeyPeriod = 190;
  wdKeySlash = 191;
  wdKeyBackSingleQuote = 192;
  wdKeyOpenSquareBrace = 219;
  wdKeyBackSlash = 220;
  wdKeyCloseSquareBrace = 221;
  wdKeySingleQuote = 222;

{ WdOLEType }

  wdOLELink = 0;
  wdOLEEmbed = 1;
  wdOLEControl = 2;

{ WdOLEVerb }

  wdOLEVerbPrimary = 0;
  wdOLEVerbShow = -1;
  wdOLEVerbOpen = -2;
  wdOLEVerbHide = -3;
  wdOLEVerbUIActivate = -4;
  wdOLEVerbInPlaceActivate = -5;
  wdOLEVerbDiscardUndoState = -6;

{ WdOLEPlacement }

  wdInLine = 0;
  wdFloatOverText = 1;

{ WdEnvelopeOrientation }

  wdLeftPortrait = 0;
  wdCenterPortrait = 1;
  wdRightPortrait = 2;
  wdLeftLandscape = 3;
  wdCenterLandscape = 4;
  wdRightLandscape = 5;
  wdLeftClockwise = 6;
  wdCenterClockwise = 7;
  wdRightClockwise = 8;

{ WdLetterStyle }

  wdFullBlock = 0;
  wdModifiedBlock = 1;
  wdSemiBlock = 2;

{ WdLetterheadLocation }

  wdLetterTop = 0;
  wdLetterBottom = 1;
  wdLetterLeft = 2;
  wdLetterRight = 3;

{ WdSalutationType }

  wdSalutationInformal = 0;
  wdSalutationFormal = 1;
  wdSalutationBusiness = 2;
  wdSalutationOther = 3;

{ WdSalutationGender }

  wdGenderFemale = 0;
  wdGenderMale = 1;
  wdGenderNeutral = 2;
  wdGenderUnknown = 3;

{ WdMovementType }

  wdMove = 0;
  wdExtend = 1;

{ WdConstants }

  wdUndefined = 9999999;
  wdToggle = 9999998;
  wdForward = 1073741823;
  wdBackward = -1073741823;
  wdAutoPosition = 0;
  wdFirst = 1;
  wdCreatorCode = 1297307460;

{ WdPasteDataType }

  wdPasteOLEObject = 0;
  wdPasteRTF = 1;
  wdPasteText = 2;
  wdPasteMetafilePicture = 3;
  wdPasteBitmap = 4;
  wdPasteDeviceIndependentBitmap = 5;
  wdPasteHyperlink = 7;
  wdPasteShape = 8;
  wdPasteEnhancedMetafile = 9;
  wdPasteHTML = 10;

{ WdPrintOutItem }

  wdPrintDocumentContent = 0;
  wdPrintProperties = 1;
  wdPrintComments = 2;
  wdPrintMarkup = 2;
  wdPrintStyles = 3;
  wdPrintAutoTextEntries = 4;
  wdPrintKeyAssignments = 5;
  wdPrintEnvelope = 6;
  wdPrintDocumentWithMarkup = 7;

{ WdPrintOutPages }

  wdPrintAllPages = 0;
  wdPrintOddPagesOnly = 1;
  wdPrintEvenPagesOnly = 2;

{ WdPrintOutRange }

  wdPrintAllDocument = 0;
  wdPrintSelection = 1;
  wdPrintCurrentPage = 2;
  wdPrintFromTo = 3;
  wdPrintRangeOfPages = 4;

{ WdDictionaryType }

  wdSpelling = 0;
  wdGrammar = 1;
  wdThesaurus = 2;
  wdHyphenation = 3;
  wdSpellingComplete = 4;
  wdSpellingCustom = 5;
  wdSpellingLegal = 6;
  wdSpellingMedical = 7;
  wdHangulHanjaConversion = 8;
  wdHangulHanjaConversionCustom = 9;

{ WdDictionaryTypeHID }

//emptyenum = 0;

{ WdSpellingWordType }

  wdSpellword = 0;
  wdWildcard = 1;
  wdAnagram = 2;

{ WdSpellingErrorType }

  wdSpellingCorrect = 0;
  wdSpellingNotInDictionary = 1;
  wdSpellingCapitalization = 2;

{ WdProofreadingErrorType }

  wdSpellingError = 0;
  wdGrammaticalError = 1;

{ WdInlineShapeType }

  wdInlineShapeEmbeddedOLEObject = 1;
  wdInlineShapeLinkedOLEObject = 2;
  wdInlineShapePicture = 3;
  wdInlineShapeLinkedPicture = 4;
  wdInlineShapeOLEControlObject = 5;
  wdInlineShapeHorizontalLine = 6;
  wdInlineShapePictureHorizontalLine = 7;
  wdInlineShapeLinkedPictureHorizontalLine = 8;
  wdInlineShapePictureBullet = 9;
  wdInlineShapeScriptAnchor = 10;
  wdInlineShapeOWSAnchor = 11;

{ WdArrangeStyle }

  wdTiled = 0;
  wdIcons = 1;

{ WdSelectionFlags }

  wdSelStartActive = 1;
  wdSelAtEOL = 2;
  wdSelOvertype = 4;
  wdSelActive = 8;
  wdSelReplace = 16;

{ WdAutoVersions }

  wdAutoVersionOff = 0;
  wdAutoVersionOnClose = 1;

{ WdOrganizerObject }

  wdOrganizerObjectStyles = 0;
  wdOrganizerObjectAutoText = 1;
  wdOrganizerObjectCommandBars = 2;
  wdOrganizerObjectProjectItems = 3;

{ WdFindMatch }

  wdMatchParagraphMark = 65551;
  wdMatchTabCharacter = 9;
  wdMatchCommentMark = 5;
  wdMatchAnyCharacter = 65599;
  wdMatchAnyDigit = 65567;
  wdMatchAnyLetter = 65583;
  wdMatchCaretCharacter = 11;
  wdMatchColumnBreak = 14;
  wdMatchEmDash = 8212;
  wdMatchEnDash = 8211;
  wdMatchEndnoteMark = 65555;
  wdMatchField = 19;
  wdMatchFootnoteMark = 65554;
  wdMatchGraphic = 1;
  wdMatchManualLineBreak = 65551;
  wdMatchManualPageBreak = 65564;
  wdMatchNonbreakingHyphen = 30;
  wdMatchNonbreakingSpace = 160;
  wdMatchOptionalHyphen = 31;
  wdMatchSectionBreak = 65580;
  wdMatchWhiteSpace = 65655;

{ WdFindWrap }

  wdFindStop = 0;
  wdFindContinue = 1;
  wdFindAsk = 2;

{ WdInformation }

  wdActiveEndAdjustedPageNumber = 1;
  wdActiveEndSectionNumber = 2;
  wdActiveEndPageNumber = 3;
  wdNumberOfPagesInDocument = 4;
  wdHorizontalPositionRelativeToPage = 5;
  wdVerticalPositionRelativeToPage = 6;
  wdHorizontalPositionRelativeToTextBoundary = 7;
  wdVerticalPositionRelativeToTextBoundary = 8;
  wdFirstCharacterColumnNumber = 9;
  wdFirstCharacterLineNumber = 10;
  wdFrameIsSelected = 11;
  wdWithInTable = 12;
  wdStartOfRangeRowNumber = 13;
  wdEndOfRangeRowNumber = 14;
  wdMaximumNumberOfRows = 15;
  wdStartOfRangeColumnNumber = 16;
  wdEndOfRangeColumnNumber = 17;
  wdMaximumNumberOfColumns = 18;
  wdZoomPercentage = 19;
  wdSelectionMode = 20;
  wdCapsLock = 21;
  wdNumLock = 22;
  wdOverType = 23;
  wdRevisionMarking = 24;
  wdInFootnoteEndnotePane = 25;
  wdInCommentPane = 26;
  wdInHeaderFooter = 28;
  wdAtEndOfRowMarker = 31;
  wdReferenceOfType = 32;
  wdHeaderFooterType = 33;
  wdInMasterDocument = 34;
  wdInFootnote = 35;
  wdInEndnote = 36;
  wdInWordMail = 37;
  wdInClipboard = 38;

{ WdWrapType }

  wdWrapSquare = 0;
  wdWrapTight = 1;
  wdWrapThrough = 2;
  wdWrapNone = 3;
  wdWrapTopBottom = 4;
  wdWrapInline = 7;

{ WdWrapSideType }

  wdWrapBoth = 0;
  wdWrapLeft = 1;
  wdWrapRight = 2;
  wdWrapLargest = 3;

{ WdOutlineLevel }

  wdOutlineLevel1 = 1;
  wdOutlineLevel2 = 2;
  wdOutlineLevel3 = 3;
  wdOutlineLevel4 = 4;
  wdOutlineLevel5 = 5;
  wdOutlineLevel6 = 6;
  wdOutlineLevel7 = 7;
  wdOutlineLevel8 = 8;
  wdOutlineLevel9 = 9;
  wdOutlineLevelBodyText = 10;

{ WdTextOrientation }

  wdTextOrientationHorizontal = 0;
  wdTextOrientationUpward = 2;
  wdTextOrientationDownward = 3;
  wdTextOrientationVerticalFarEast = 1;
  wdTextOrientationHorizontalRotatedFarEast = 4;

{ WdTextOrientationHID }

//emptyenum = 0;

{ WdPageBorderArt }

  wdArtApples = 1;
  wdArtMapleMuffins = 2;
  wdArtCakeSlice = 3;
  wdArtCandyCorn = 4;
  wdArtIceCreamCones = 5;
  wdArtChampagneBottle = 6;
  wdArtPartyGlass = 7;
  wdArtChristmasTree = 8;
  wdArtTrees = 9;
  wdArtPalmsColor = 10;
  wdArtBalloons3Colors = 11;
  wdArtBalloonsHotAir = 12;
  wdArtPartyFavor = 13;
  wdArtConfettiStreamers = 14;
  wdArtHearts = 15;
  wdArtHeartBalloon = 16;
  wdArtStars3D = 17;
  wdArtStarsShadowed = 18;
  wdArtStars = 19;
  wdArtSun = 20;
  wdArtEarth2 = 21;
  wdArtEarth1 = 22;
  wdArtPeopleHats = 23;
  wdArtSombrero = 24;
  wdArtPencils = 25;
  wdArtPackages = 26;
  wdArtClocks = 27;
  wdArtFirecrackers = 28;
  wdArtRings = 29;
  wdArtMapPins = 30;
  wdArtConfetti = 31;
  wdArtCreaturesButterfly = 32;
  wdArtCreaturesLadyBug = 33;
  wdArtCreaturesFish = 34;
  wdArtBirdsFlight = 35;
  wdArtScaredCat = 36;
  wdArtBats = 37;
  wdArtFlowersRoses = 38;
  wdArtFlowersRedRose = 39;
  wdArtPoinsettias = 40;
  wdArtHolly = 41;
  wdArtFlowersTiny = 42;
  wdArtFlowersPansy = 43;
  wdArtFlowersModern2 = 44;
  wdArtFlowersModern1 = 45;
  wdArtWhiteFlowers = 46;
  wdArtVine = 47;
  wdArtFlowersDaisies = 48;
  wdArtFlowersBlockPrint = 49;
  wdArtDecoArchColor = 50;
  wdArtFans = 51;
  wdArtFilm = 52;
  wdArtLightning1 = 53;
  wdArtCompass = 54;
  wdArtDoubleD = 55;
  wdArtClassicalWave = 56;
  wdArtShadowedSquares = 57;
  wdArtTwistedLines1 = 58;
  wdArtWaveline = 59;
  wdArtQuadrants = 60;
  wdArtCheckedBarColor = 61;
  wdArtSwirligig = 62;
  wdArtPushPinNote1 = 63;
  wdArtPushPinNote2 = 64;
  wdArtPumpkin1 = 65;
  wdArtEggsBlack = 66;
  wdArtCup = 67;
  wdArtHeartGray = 68;
  wdArtGingerbreadMan = 69;
  wdArtBabyPacifier = 70;
  wdArtBabyRattle = 71;
  wdArtCabins = 72;
  wdArtHouseFunky = 73;
  wdArtStarsBlack = 74;
  wdArtSnowflakes = 75;
  wdArtSnowflakeFancy = 76;
  wdArtSkyrocket = 77;
  wdArtSeattle = 78;
  wdArtMusicNotes = 79;
  wdArtPalmsBlack = 80;
  wdArtMapleLeaf = 81;
  wdArtPaperClips = 82;
  wdArtShorebirdTracks = 83;
  wdArtPeople = 84;
  wdArtPeopleWaving = 85;
  wdArtEclipsingSquares2 = 86;
  wdArtHypnotic = 87;
  wdArtDiamondsGray = 88;
  wdArtDecoArch = 89;
  wdArtDecoBlocks = 90;
  wdArtCirclesLines = 91;
  wdArtPapyrus = 92;
  wdArtWoodwork = 93;
  wdArtWeavingBraid = 94;
  wdArtWeavingRibbon = 95;
  wdArtWeavingAngles = 96;
  wdArtArchedScallops = 97;
  wdArtSafari = 98;
  wdArtCelticKnotwork = 99;
  wdArtCrazyMaze = 100;
  wdArtEclipsingSquares1 = 101;
  wdArtBirds = 102;
  wdArtFlowersTeacup = 103;
  wdArtNorthwest = 104;
  wdArtSouthwest = 105;
  wdArtTribal6 = 106;
  wdArtTribal4 = 107;
  wdArtTribal3 = 108;
  wdArtTribal2 = 109;
  wdArtTribal5 = 110;
  wdArtXIllusions = 111;
  wdArtZanyTriangles = 112;
  wdArtPyramids = 113;
  wdArtPyramidsAbove = 114;
  wdArtConfettiGrays = 115;
  wdArtConfettiOutline = 116;
  wdArtConfettiWhite = 117;
  wdArtMosaic = 118;
  wdArtLightning2 = 119;
  wdArtHeebieJeebies = 120;
  wdArtLightBulb = 121;
  wdArtGradient = 122;
  wdArtTriangleParty = 123;
  wdArtTwistedLines2 = 124;
  wdArtMoons = 125;
  wdArtOvals = 126;
  wdArtDoubleDiamonds = 127;
  wdArtChainLink = 128;
  wdArtTriangles = 129;
  wdArtTribal1 = 130;
  wdArtMarqueeToothed = 131;
  wdArtSharksTeeth = 132;
  wdArtSawtooth = 133;
  wdArtSawtoothGray = 134;
  wdArtPostageStamp = 135;
  wdArtWeavingStrips = 136;
  wdArtZigZag = 137;
  wdArtCrossStitch = 138;
  wdArtGems = 139;
  wdArtCirclesRectangles = 140;
  wdArtCornerTriangles = 141;
  wdArtCreaturesInsects = 142;
  wdArtZigZagStitch = 143;
  wdArtCheckered = 144;
  wdArtCheckedBarBlack = 145;
  wdArtMarquee = 146;
  wdArtBasicWhiteDots = 147;
  wdArtBasicWideMidline = 148;
  wdArtBasicWideOutline = 149;
  wdArtBasicWideInline = 150;
  wdArtBasicThinLines = 151;
  wdArtBasicWhiteDashes = 152;
  wdArtBasicWhiteSquares = 153;
  wdArtBasicBlackSquares = 154;
  wdArtBasicBlackDashes = 155;
  wdArtBasicBlackDots = 156;
  wdArtStarsTop = 157;
  wdArtCertificateBanner = 158;
  wdArtHandmade1 = 159;
  wdArtHandmade2 = 160;
  wdArtTornPaper = 161;
  wdArtTornPaperBlack = 162;
  wdArtCouponCutoutDashes = 163;
  wdArtCouponCutoutDots = 164;

{ WdBorderDistanceFrom }

  wdBorderDistanceFromText = 0;
  wdBorderDistanceFromPageEdge = 1;

{ WdReplace }

  wdReplaceNone = 0;
  wdReplaceOne = 1;
  wdReplaceAll = 2;

{ WdFontBias }

  wdFontBiasDontCare = 255;
  wdFontBiasDefault = 0;
  wdFontBiasFareast = 1;

{ WdBrowserLevel }

  wdBrowserLevelV4 = 0;
  wdBrowserLevelMicrosoftInternetExplorer5 = 1;
  wdBrowserLevelMicrosoftInternetExplorer6 = 2;

{ WdEnclosureType }

  wdEnclosureCircle = 0;
  wdEnclosureSquare = 1;
  wdEnclosureTriangle = 2;
  wdEnclosureDiamond = 3;

{ WdEncloseStyle }

  wdEncloseStyleNone = 0;
  wdEncloseStyleSmall = 1;
  wdEncloseStyleLarge = 2;

{ WdHighAnsiText }

  wdHighAnsiIsFarEast = 0;
  wdHighAnsiIsHighAnsi = 1;
  wdAutoDetectHighAnsiFarEast = 2;

{ WdLayoutMode }

  wdLayoutModeDefault = 0;
  wdLayoutModeGrid = 1;
  wdLayoutModeLineGrid = 2;
  wdLayoutModeGenko = 3;

{ WdDocumentMedium }

  wdEmailMessage = 0;
  wdDocument = 1;
  wdWebPage = 2;

{ WdMailerPriority }

  wdPriorityNormal = 1;
  wdPriorityLow = 2;
  wdPriorityHigh = 3;

{ WdDocumentViewDirection }

  wdDocumentViewRtl = 0;
  wdDocumentViewLtr = 1;

{ WdArabicNumeral }

  wdNumeralArabic = 0;
  wdNumeralHindi = 1;
  wdNumeralContext = 2;
  wdNumeralSystem = 3;

{ WdMonthNames }

  wdMonthNamesArabic = 0;
  wdMonthNamesEnglish = 1;
  wdMonthNamesFrench = 2;

{ WdCursorMovement }

  wdCursorMovementLogical = 0;
  wdCursorMovementVisual = 1;

{ WdVisualSelection }

  wdVisualSelectionBlock = 0;
  wdVisualSelectionContinuous = 1;

{ WdTableDirection }

  wdTableDirectionRtl = 0;
  wdTableDirectionLtr = 1;

{ WdFlowDirection }

  wdFlowLtr = 0;
  wdFlowRtl = 1;

{ WdDiacriticColor }

  wdDiacriticColorBidi = 0;
  wdDiacriticColorLatin = 1;

{ WdGutterStyle }

  wdGutterPosLeft = 0;
  wdGutterPosTop = 1;
  wdGutterPosRight = 2;

{ WdGutterStyleOld }

  wdGutterStyleLatin = -10;
  wdGutterStyleBidi = 2;

{ WdSectionDirection }

  wdSectionDirectionRtl = 0;
  wdSectionDirectionLtr = 1;

{ WdDateLanguage }

  wdDateLanguageBidi = 10;
  wdDateLanguageLatin = 1033;

{ WdCalendarTypeBi }

  wdCalendarTypeBidi = 99;
  wdCalendarTypeGregorian = 100;

{ WdCalendarType }

  wdCalendarWestern = 0;
  wdCalendarArabic = 1;
  wdCalendarHebrew = 2;
  wdCalendarChina = 3;
  wdCalendarJapan = 4;
  wdCalendarThai = 5;
  wdCalendarKorean = 6;
  wdCalendarSakaEra = 7;

{ WdReadingOrder }

  wdReadingOrderRtl = 0;
  wdReadingOrderLtr = 1;

{ WdHebSpellStart }

  wdFullScript = 0;
  wdPartialScript = 1;
  wdMixedScript = 2;
  wdMixedAuthorizedScript = 3;

{ WdAraSpeller }

  wdNone = 0;
  wdInitialAlef = 1;
  wdFinalYaa = 2;
  wdBoth = 3;

{ WdColor }

  wdColorAutomatic = -16777216;
  wdColorBlack = 0;
  wdColorBlue = 16711680;
  wdColorTurquoise = 16776960;
  wdColorBrightGreen = 65280;
  wdColorPink = 16711935;
  wdColorRed = 255;
  wdColorYellow = 65535;
  wdColorWhite = 16777215;
  wdColorDarkBlue = 8388608;
  wdColorTeal = 8421376;
  wdColorGreen = 32768;
  wdColorViolet = 8388736;
  wdColorDarkRed = 128;
  wdColorDarkYellow = 32896;
  wdColorBrown = 13209;
  wdColorOliveGreen = 13107;
  wdColorDarkGreen = 13056;
  wdColorDarkTeal = 6697728;
  wdColorIndigo = 10040115;
  wdColorOrange = 26367;
  wdColorBlueGray = 10053222;
  wdColorLightOrange = 39423;
  wdColorLime = 52377;
  wdColorSeaGreen = 6723891;
  wdColorAqua = 13421619;
  wdColorLightBlue = 16737843;
  wdColorGold = 52479;
  wdColorSkyBlue = 16763904;
  wdColorPlum = 6697881;
  wdColorRose = 13408767;
  wdColorTan = 10079487;
  wdColorLightYellow = 10092543;
  wdColorLightGreen = 13434828;
  wdColorLightTurquoise = 16777164;
  wdColorPaleBlue = 16764057;
  wdColorLavender = 16751052;
  wdColorGray05 = 15987699;
  wdColorGray10 = 15132390;
  wdColorGray125 = 14737632;
  wdColorGray15 = 14277081;
  wdColorGray20 = 13421772;
  wdColorGray25 = 12632256;
  wdColorGray30 = 11776947;
  wdColorGray35 = 10921638;
  wdColorGray375 = 10526880;
  wdColorGray40 = 10066329;
  wdColorGray45 = 9211020;
  wdColorGray50 = 8421504;
  wdColorGray55 = 7566195;
  wdColorGray60 = 6710886;
  wdColorGray625 = 6316128;
  wdColorGray65 = 5855577;
  wdColorGray70 = 5000268;
  wdColorGray75 = 4210752;
  wdColorGray80 = 3355443;
  wdColorGray85 = 2500134;
  wdColorGray875 = 2105376;
  wdColorGray90 = 1644825;
  wdColorGray95 = 789516;

{ WdShapePosition }

  wdShapeTop = -999999;
  wdShapeLeft = -999998;
  wdShapeBottom = -999997;
  wdShapeRight = -999996;
  wdShapeCenter = -999995;
  wdShapeInside = -999994;
  wdShapeOutside = -999993;

{ WdTablePosition }

  wdTableTop = -999999;
  wdTableLeft = -999998;
  wdTableBottom = -999997;
  wdTableRight = -999996;
  wdTableCenter = -999995;
  wdTableInside = -999994;
  wdTableOutside = -999993;

{ WdDefaultListBehavior }

  wdWord8ListBehavior = 0;
  wdWord9ListBehavior = 1;
  wdWord10ListBehavior = 2;

{ WdDefaultTableBehavior }

  wdWord8TableBehavior = 0;
  wdWord9TableBehavior = 1;

{ WdAutoFitBehavior }

  wdAutoFitFixed = 0;
  wdAutoFitContent = 1;
  wdAutoFitWindow = 2;

{ WdPreferredWidthType }

  wdPreferredWidthAuto = 1;
  wdPreferredWidthPercent = 2;
  wdPreferredWidthPoints = 3;

{ WdFarEastLineBreakLanguageID }

  wdLineBreakJapanese = 1041;
  wdLineBreakKorean = 1042;
  wdLineBreakSimplifiedChinese = 2052;
  wdLineBreakTraditionalChinese = 1028;

{ WdViewTypeOld }

  wdPageView = 3;
  wdOnlineView = 6;

{ WdFramesetType }

  wdFramesetTypeFrameset = 0;
  wdFramesetTypeFrame = 1;

{ WdFramesetSizeType }

  wdFramesetSizeTypePercent = 0;
  wdFramesetSizeTypeFixed = 1;
  wdFramesetSizeTypeRelative = 2;

{ WdFramesetNewFrameLocation }

  wdFramesetNewFrameAbove = 0;
  wdFramesetNewFrameBelow = 1;
  wdFramesetNewFrameRight = 2;
  wdFramesetNewFrameLeft = 3;

{ WdScrollbarType }

  wdScrollbarTypeAuto = 0;
  wdScrollbarTypeYes = 1;
  wdScrollbarTypeNo = 2;

{ WdTwoLinesInOneType }

  wdTwoLinesInOneNone = 0;
  wdTwoLinesInOneNoBrackets = 1;
  wdTwoLinesInOneParentheses = 2;
  wdTwoLinesInOneSquareBrackets = 3;
  wdTwoLinesInOneAngleBrackets = 4;
  wdTwoLinesInOneCurlyBrackets = 5;

{ WdHorizontalInVerticalType }

  wdHorizontalInVerticalNone = 0;
  wdHorizontalInVerticalFitInLine = 1;
  wdHorizontalInVerticalResizeLine = 2;

{ WdHorizontalLineAlignment }

  wdHorizontalLineAlignLeft = 0;
  wdHorizontalLineAlignCenter = 1;
  wdHorizontalLineAlignRight = 2;

{ WdHorizontalLineWidthType }

  wdHorizontalLinePercentWidth = -1;
  wdHorizontalLineFixedWidth = -2;

{ WdPhoneticGuideAlignmentType }

  wdPhoneticGuideAlignmentCenter = 0;
  wdPhoneticGuideAlignmentZeroOneZero = 1;
  wdPhoneticGuideAlignmentOneTwoOne = 2;
  wdPhoneticGuideAlignmentLeft = 3;
  wdPhoneticGuideAlignmentRight = 4;
  wdPhoneticGuideAlignmentRightVertical = 5;

{ WdNewDocumentType }

  wdNewBlankDocument = 0;
  wdNewWebPage = 1;
  wdNewEmailMessage = 2;
  wdNewFrameset = 3;

{ WdKana }

  wdKanaKatakana = 8;
  wdKanaHiragana = 9;

{ WdCharacterWidth }

  wdWidthHalfWidth = 6;
  wdWidthFullWidth = 7;

{ WdNumberStyleWordBasicBiDi }

  wdListNumberStyleBidi1 = 49;
  wdListNumberStyleBidi2 = 50;
  wdCaptionNumberStyleBidiLetter1 = 49;
  wdCaptionNumberStyleBidiLetter2 = 50;
  wdNoteNumberStyleBidiLetter1 = 49;
  wdNoteNumberStyleBidiLetter2 = 50;
  wdPageNumberStyleBidiLetter1 = 49;
  wdPageNumberStyleBidiLetter2 = 50;

{ WdTCSCConverterDirection }

  wdTCSCConverterDirectionSCTC = 0;
  wdTCSCConverterDirectionTCSC = 1;
  wdTCSCConverterDirectionAuto = 2;

{ WdDisableFeaturesIntroducedAfter }

  wd70 = 0;
  wd70FE = 1;
  wd80 = 2;

{ WdWrapTypeMerged }

  wdWrapMergeInline = 0;
  wdWrapMergeSquare = 1;
  wdWrapMergeTight = 2;
  wdWrapMergeBehind = 3;
  wdWrapMergeFront = 4;
  wdWrapMergeThrough = 5;
  wdWrapMergeTopBottom = 6;

{ WdRecoveryType }

  wdPasteDefault = 0;
  wdSingleCellText = 5;
  wdSingleCellTable = 6;
  wdListContinueNumbering = 7;
  wdListRestartNumbering = 8;
  wdTableInsertAsRows = 11;
  wdTableAppendTable = 10;
  wdTableOriginalFormatting = 12;
  wdChartPicture = 13;
  wdChart = 14;
  wdChartLinked = 15;
  wdFormatOriginalFormatting = 16;
  wdFormatSurroundingFormattingWithEmphasis = 20;
  wdFormatPlainText = 22;
  wdTableOverwriteCells = 23;
  wdListCombineWithExistingList = 24;
  wdListDontMerge = 25;

{ WdLineEndingType }

  wdCRLF = 0;
  wdCROnly = 1;
  wdLFOnly = 2;
  wdLFCR = 3;
  wdLSPS = 4;

{ WdStyleSheetLinkType }

  wdStyleSheetLinkTypeLinked = 0;
  wdStyleSheetLinkTypeImported = 1;

{ WdStyleSheetPrecedence }

  wdStyleSheetPrecedenceHigher = -1;
  wdStyleSheetPrecedenceLower = -2;
  wdStyleSheetPrecedenceHighest = 1;
  wdStyleSheetPrecedenceLowest = 0;

{ WdEmailHTMLFidelity }

  wdEmailHTMLFidelityLow = 1;
  wdEmailHTMLFidelityMedium = 2;
  wdEmailHTMLFidelityHigh = 3;

{ WdMailMergeMailFormat }

  wdMailFormatPlainText = 0;
  wdMailFormatHTML = 1;

{ WdMappedDataFields }

  wdUniqueIdentifier = 1;
  wdCourtesyTitle = 2;
  wdFirstName = 3;
  wdMiddleName = 4;
  wdLastName = 5;
  wdSuffix = 6;
  wdNickname = 7;
  wdJobTitle = 8;
  wdCompany = 9;
  wdAddress1 = 10;
  wdAddress2 = 11;
  wdCity = 12;
  wdState = 13;
  wdPostalCode = 14;
  wdCountryRegion = 15;
  wdBusinessPhone = 16;
  wdBusinessFax = 17;
  wdHomePhone = 18;
  wdHomeFax = 19;
  wdEmailAddress = 20;
  wdWebPageURL = 21;
  wdSpouseCourtesyTitle = 22;
  wdSpouseFirstName = 23;
  wdSpouseMiddleName = 24;
  wdSpouseLastName = 25;
  wdSpouseNickname = 26;
  wdRubyFirstName = 27;
  wdRubyLastName = 28;
  wdAddress3 = 29;
  wdDepartment = 30;

{ WdConditionCode }

  wdFirstRow = 0;
  wdLastRow = 1;
  wdOddRowBanding = 2;
  wdEvenRowBanding = 3;
  wdFirstColumn = 4;
  wdLastColumn = 5;
  wdOddColumnBanding = 6;
  wdEvenColumnBanding = 7;
  wdNECell = 8;
  wdNWCell = 9;
  wdSECell = 10;
  wdSWCell = 11;

{ WdCompareTarget }

  wdCompareTargetSelected = 0;
  wdCompareTargetCurrent = 1;
  wdCompareTargetNew = 2;

{ WdMergeTarget }

  wdMergeTargetSelected = 0;
  wdMergeTargetCurrent = 1;
  wdMergeTargetNew = 2;

{ WdUseFormattingFrom }

  wdFormattingFromCurrent = 0;
  wdFormattingFromSelected = 1;
  wdFormattingFromPrompt = 2;

{ WdRevisionsView }

  wdRevisionsViewFinal = 0;
  wdRevisionsViewOriginal = 1;

{ WdRevisionsMode }

  wdBalloonRevisions = 0;
  wdInLineRevisions = 1;

{ WdRevisionsBalloonWidthType }

  wdBalloonWidthPercent = 0;
  wdBalloonWidthPoints = 1;

{ WdRevisionsBalloonPrintOrientation }

  wdBalloonPrintOrientationAuto = 0;
  wdBalloonPrintOrientationPreserve = 1;
  wdBalloonPrintOrientationForceLandscape = 2;

{ WdRevisionsBalloonMargin }

  wdLeftMargin = 0;
  wdRightMargin = 1;

{ WdTaskPanes }

  wdTaskPaneFormatting = 0;
  wdTaskPaneRevealFormatting = 1;
  wdTaskPaneMailMerge = 2;
  wdTaskPaneTranslate = 3;
  wdTaskPaneSearch = 4;

{ WdShowFilter }

  wdShowFilterStylesAvailable = 0;
  wdShowFilterStylesInUse = 1;
  wdShowFilterStylesAll = 2;
  wdShowFilterFormattingInUse = 3;
  wdShowFilterFormattingAvailable = 4;

{ WdMergeSubType }

  wdMergeSubTypeOther = 0;
  wdMergeSubTypeAccess = 1;
  wdMergeSubTypeOAL = 2;
  wdMergeSubTypeOLEDBWord = 3;
  wdMergeSubTypeWorks = 4;
  wdMergeSubTypeOLEDBText = 5;
  wdMergeSubTypeOutlook = 6;
  wdMergeSubTypeWord = 7;
  wdMergeSubTypeWord2000 = 8;

{ WdDocumentDirection }

  wdLeftToRight = 0;
  wdRightToLeft = 1;

{ WdLanguageID2000 }

  wdChineseHongKong = 3076;
  wdChineseMacao = 5124;
  wdEnglishTrinidad = 11273;

{ Для размещения страниц}
  wdNormalPage = 0;   //нормальне
  wdTwoOnOne = 1;     //2 стор. на 1
  wdBookFold = 2;     //брошура

{ Константы для M$ Excel}
const

{ Constants }

  xlAll = -4104;
  xlAutomatic = -4105;
  xlBoth = 1;
  xlCenter = -4108;
  xlChecker = 9;
  xlCircle = 8;
  xlCorner = 2;
  xlCrissCross = 16;
  xlCross = 4;
  xlDiamond = 2;
  xlDistributed = -4117;
  xlDoubleAccounting = 5;
  xlFixedValue = 1;
  xlFormats = -4122;
  xlGray16 = 17;
  xlGray8 = 18;
  xlGrid = 15;
  xlHigh = -4127;
  xlInside = 2;
  xlJustify = -4130;
  xlLightDown = 13;
  xlLightHorizontal = 11;
  xlLightUp = 14;
  xlLightVertical = 12;
  xlLow = -4134;
  xlManual = -4135;
  xlMinusValues = 3;
  xlModule = -4141;
  xlNextToAxis = 4;
  xlNone = -4142;
  xlNotes = -4144;
  xlOff = -4146;
  xlOn = 1;
  xlPercent = 2;
  xlPlus = 9;
  xlPlusValues = 2;
  xlSemiGray75 = 10;
  xlShowLabel = 4;
  xlShowLabelAndPercent = 5;
  xlShowPercent = 3;
  xlShowValue = 2;
  xlSimple = -4154;
  xlSingle = 2;
  xlSingleAccounting = 4;
  xlSolid = 1;
  xlSquare = 1;
  xlStar = 5;
  xlStError = 4;
  xlToolbarButton = 2;
  xlTriangle = 3;
  xlGray25 = -4124;
  xlGray50 = -4125;
  xlGray75 = -4126;
  xlBottom = -4107;
  xlLeft = -4131;
  xlRight = -4152;
  xlTop = -4160;
  xl3DBar = -4099;
  xl3DSurface = -4103;
  xlBar = 2;
  xlColumn = 3;
  xlCombination = -4111;
  xlCustom = -4114;
  xlDefaultAutoFormat = -1;
  xlMaximum = 2;
  xlMinimum = 4;
  xlOpaque = 3;
  xlTransparent = 2;
  xlBidi = -5000;
  xlLatin = -5001;
  xlContext = -5002;
  xlLTR = -5003;
  xlRTL = -5004;
  xlFullScript = 1;
  xlPartialScript = 2;
  xlMixedScript = 3;
  xlMixedAuthorizedScript = 4;
  xlVisualCursor = 2;
  xlLogicalCursor = 1;
  xlSystem = 1;
  xlPartial = 3;
  xlHindiNumerals = 3;
  xlBidiCalendar = 3;
  xlGregorian = 2;
  xlComplete = 4;
  xlScale = 3;
  xlClosed = 3;
  xlColor1 = 7;
  xlColor2 = 8;
  xlColor3 = 9;
  xlConstants = 2;
  xlContents = 2;
  xlBelow = 1;
  xlCascade = 7;
  xlCenterAcrossSelection = 7;
  xlChart4 = 2;
  xlChartSeries = 17;
  xlChartShort = 6;
  xlChartTitles = 18;
  xlClassic1 = 1;
  xlClassic2 = 2;
  xlClassic3 = 3;
  xl3DEffects1 = 13;
  xl3DEffects2 = 14;
  xlAbove = 0;
  xlAccounting1 = 4;
  xlAccounting2 = 5;
  xlAccounting3 = 6;
  xlAccounting4 = 17;
  xlAdd = 2;
  xlDebugCodePane = 13;
  xlDesktop = 9;
  xlDirect = 1;
  xlDivide = 5;
  xlDoubleClosed = 5;
  xlDoubleOpen = 4;
  xlDoubleQuote = 1;
  xlEntireChart = 20;
  xlExcelMenus = 1;
  xlExtended = 3;
  xlFill = 5;
  xlFirst = 0;
  xlFloating = 5;
  xlFormula = 5;
  xlGeneral = 1;
  xlGridline = 22;
  xlIcons = 1;
  xlImmediatePane = 12;
  xlInteger = 2;
  xlLast = 1;
  xlLastCell = 11;
  xlList1 = 10;
  xlList2 = 11;
  xlList3 = 12;
  xlLocalFormat1 = 15;
  xlLocalFormat2 = 16;
  xlLong = 3;
  xlLotusHelp = 2;
  xlMacrosheetCell = 7;
  xlMixed = 2;
  xlMultiply = 4;
  xlNarrow = 1;
  xlNoDocuments = 3;
  xlOpen = 2;
  xlOutside = 3;
  xlReference = 4;
  xlSemiautomatic = 2;
  xlShort = 1;
  xlSingleQuote = 2;
  xlStrict = 2;
  xlSubtract = 3;
  xlTextBox = 16;
  xlTiled = 1;
  xlTitleBar = 8;
  xlToolbar = 1;
  xlVisible = 12;
  xlWatchPane = 11;
  xlWide = 3;
  xlWorkbookTab = 6;
  xlWorksheet4 = 1;
  xlWorksheetCell = 3;
  xlWorksheetShort = 5;
  xlAllExceptBorders = 7;
  xlLeftToRight = 2;
  xlTopToBottom = 1;
  xlVeryHidden = 2;
  xlDrawingObject = 14;

{ XlCreator }

  xlCreatorCode = 1480803660;

{ XlChartGallery }

  xlBuiltIn = 21;
  xlUserDefined = 22;
  xlAnyGallery = 23;

{ XlColorIndex }

  xlColorIndexAutomatic = -4105;
  xlColorIndexNone = -4142;

{ XlEndStyleCap }

  xlCap = 1;
  xlNoCap = 2;

{ XlRowCol }

  xlColumns = 2;
  xlRows = 1;

{ XlScaleType }

  xlScaleLinear = -4132;
  xlScaleLogarithmic = -4133;

{ XlDataSeriesType }

  xlAutoFill = 4;
  xlChronological = 3;
  xlGrowth = 2;
  xlDataSeriesLinear = -4132;

{ XlAxisCrosses }

  xlAxisCrossesAutomatic = -4105;
  xlAxisCrossesCustom = -4114;
  xlAxisCrossesMaximum = 2;
  xlAxisCrossesMinimum = 4;

{ XlAxisGroup }

  xlPrimary = 1;
  xlSecondary = 2;

{ XlBackground }

  xlBackgroundAutomatic = -4105;
  xlBackgroundOpaque = 3;
  xlBackgroundTransparent = 2;

{ XlWindowState }

  xlMaximized = -4137;
  xlMinimized = -4140;
  xlNormal = -4143;

{ XlAxisType }

  xlCategory = 1;
  xlSeriesAxis = 3;
  xlValue = 2;

{ XlArrowHeadLength }

  xlArrowHeadLengthLong = 3;
  xlArrowHeadLengthMedium = -4138;
  xlArrowHeadLengthShort = 1;

{ XlVAlign }

  xlVAlignBottom = -4107;
  xlVAlignCenter = -4108;
  xlVAlignDistributed = -4117;
  xlVAlignJustify = -4130;
  xlVAlignTop = -4160;

{ XlTickMark }

  xlTickMarkCross = 4;
  xlTickMarkInside = 2;
  xlTickMarkNone = -4142;
  xlTickMarkOutside = 3;

{ XlErrorBarDirection }

  xlX = -4168;
  xlY = 1;

{ XlErrorBarInclude }

  xlErrorBarIncludeBoth = 1;
  xlErrorBarIncludeMinusValues = 3;
  xlErrorBarIncludeNone = -4142;
  xlErrorBarIncludePlusValues = 2;

{ XlDisplayBlanksAs }

  xlInterpolated = 3;
  xlNotPlotted = 1;
  xlZero = 2;

{ XlArrowHeadStyle }

  xlArrowHeadStyleClosed = 3;
  xlArrowHeadStyleDoubleClosed = 5;
  xlArrowHeadStyleDoubleOpen = 4;
  xlArrowHeadStyleNone = -4142;
  xlArrowHeadStyleOpen = 2;

{ XlArrowHeadWidth }

  xlArrowHeadWidthMedium = -4138;
  xlArrowHeadWidthNarrow = 1;
  xlArrowHeadWidthWide = 3;

{ XlHAlign }

  xlHAlignCenter = -4108;
  xlHAlignCenterAcrossSelection = 7;
  xlHAlignDistributed = -4117;
  xlHAlignFill = 5;
  xlHAlignGeneral = 1;
  xlHAlignJustify = -4130;
  xlHAlignLeft = -4131;
  xlHAlignRight = -4152;

{ XlTickLabelPosition }

  xlTickLabelPositionHigh = -4127;
  xlTickLabelPositionLow = -4134;
  xlTickLabelPositionNextToAxis = 4;
  xlTickLabelPositionNone = -4142;

{ XlLegendPosition }

  xlLegendPositionBottom = -4107;
  xlLegendPositionCorner = 2;
  xlLegendPositionLeft = -4131;
  xlLegendPositionRight = -4152;
  xlLegendPositionTop = -4160;

{ XlChartPictureType }

  xlStackScale = 3;
  xlStack = 2;
  xlStretch = 1;

{ XlChartPicturePlacement }

  xlSides = 1;
  xlEnd = 2;
  xlEndSides = 3;
  xlFront = 4;
  xlFrontSides = 5;
  xlFrontEnd = 6;
  xlAllFaces = 7;

{ XlOrientation }

  xlDownward = -4170;
  xlHorizontal = -4128;
  xlUpward = -4171;
  xlVertical = -4166;

{ XlTickLabelOrientation }

  xlTickLabelOrientationAutomatic = -4105;
  xlTickLabelOrientationDownward = -4170;
  xlTickLabelOrientationHorizontal = -4128;
  xlTickLabelOrientationUpward = -4171;
  xlTickLabelOrientationVertical = -4166;

{ XlBorderWeight }

  xlHairline = 1;
  xlMedium = -4138;
  xlThick = 4;
  xlThin = 2;

{ XlDataSeriesDate }

  xlDay = 1;
  xlMonth = 3;
  xlWeekday = 2;
  xlYear = 4;

{ XlUnderlineStyle }

  xlUnderlineStyleDouble = -4119;
  xlUnderlineStyleDoubleAccounting = 5;
  xlUnderlineStyleNone = -4142;
  xlUnderlineStyleSingle = 2;
  xlUnderlineStyleSingleAccounting = 4;

{ XlErrorBarType }

  xlErrorBarTypeCustom = -4114;
  xlErrorBarTypeFixedValue = 1;
  xlErrorBarTypePercent = 2;
  xlErrorBarTypeStDev = -4155;
  xlErrorBarTypeStError = 4;

{ XlTrendlineType }

  xlExponential = 5;
  xlLinear = -4132;
  xlLogarithmic = -4133;
  xlMovingAvg = 6;
  xlPolynomial = 3;
  xlPower = 4;

{ XlLineStyle }

  xlContinuous = 1;
  xlDash = -4115;
  xlDashDot = 4;
  xlDashDotDot = 5;
  xlDot = -4118;
  xlDouble = -4119;
  xlSlantDashDot = 13;
  xlLineStyleNone = -4142;

{ XlDataLabelsType }

  xlDataLabelsShowNone = -4142;
  xlDataLabelsShowValue = 2;
  xlDataLabelsShowPercent = 3;
  xlDataLabelsShowLabel = 4;
  xlDataLabelsShowLabelAndPercent = 5;
  xlDataLabelsShowBubbleSizes = 6;

{ XlMarkerStyle }

  xlMarkerStyleAutomatic = -4105;
  xlMarkerStyleCircle = 8;
  xlMarkerStyleDash = -4115;
  xlMarkerStyleDiamond = 2;
  xlMarkerStyleDot = -4118;
  xlMarkerStyleNone = -4142;
  xlMarkerStylePicture = -4147;
  xlMarkerStylePlus = 9;
  xlMarkerStyleSquare = 1;
  xlMarkerStyleStar = 5;
  xlMarkerStyleTriangle = 3;
  xlMarkerStyleX = -4168;

{ XlPictureConvertorType }

  xlBMP = 1;
  xlCGM = 7;
  xlDRW = 4;
  xlDXF = 5;
  xlEPS = 8;
  xlHGL = 6;
  xlPCT = 13;
  xlPCX = 10;
  xlPIC = 11;
  xlPLT = 12;
  xlTIF = 9;
  xlWMF = 2;
  xlWPG = 3;

{ XlPattern }

  xlPatternAutomatic = -4105;
  xlPatternChecker = 9;
  xlPatternCrissCross = 16;
  xlPatternDown = -4121;
  xlPatternGray16 = 17;
  xlPatternGray25 = -4124;
  xlPatternGray50 = -4125;
  xlPatternGray75 = -4126;
  xlPatternGray8 = 18;
  xlPatternGrid = 15;
  xlPatternHorizontal = -4128;
  xlPatternLightDown = 13;
  xlPatternLightHorizontal = 11;
  xlPatternLightUp = 14;
  xlPatternLightVertical = 12;
  xlPatternNone = -4142;
  xlPatternSemiGray75 = 10;
  xlPatternSolid = 1;
  xlPatternUp = -4162;
  xlPatternVertical = -4166;

{ XlChartSplitType }

  xlSplitByPosition = 1;
  xlSplitByPercentValue = 3;
  xlSplitByCustomSplit = 4;
  xlSplitByValue = 2;

{ XlDisplayUnit }

  xlHundreds = -2;
  xlThousands = -3;
  xlTenThousands = -4;
  xlHundredThousands = -5;
  xlMillions = -6;
  xlTenMillions = -7;
  xlHundredMillions = -8;
  xlThousandMillions = -9;
  xlMillionMillions = -10;

{ XlDataLabelPosition }

  xlLabelPositionCenter = -4108;
  xlLabelPositionAbove = 0;
  xlLabelPositionBelow = 1;
  xlLabelPositionLeft = -4131;
  xlLabelPositionRight = -4152;
  xlLabelPositionOutsideEnd = 2;
  xlLabelPositionInsideEnd = 3;
  xlLabelPositionInsideBase = 4;
  xlLabelPositionBestFit = 5;
  xlLabelPositionMixed = 6;
  xlLabelPositionCustom = 7;

{ XlTimeUnit }

  xlDays = 0;
  xlMonths = 1;
  xlYears = 2;

{ XlCategoryType }

  xlCategoryScale = 2;
  xlTimeScale = 3;
  xlAutomaticScale = -4105;

{ XlBarShape }

  xlBox = 0;
  xlPyramidToPoint = 1;
  xlPyramidToMax = 2;
  xlCylinder = 3;
  xlConeToPoint = 4;
  xlConeToMax = 5;

{ XlChartType }

  xlColumnClustered = 51;
  xlColumnStacked = 52;
  xlColumnStacked100 = 53;
  xl3DColumnClustered = 54;
  xl3DColumnStacked = 55;
  xl3DColumnStacked100 = 56;
  xlBarClustered = 57;
  xlBarStacked = 58;
  xlBarStacked100 = 59;
  xl3DBarClustered = 60;
  xl3DBarStacked = 61;
  xl3DBarStacked100 = 62;
  xlLineStacked = 63;
  xlLineStacked100 = 64;
  xlLineMarkers = 65;
  xlLineMarkersStacked = 66;
  xlLineMarkersStacked100 = 67;
  xlPieOfPie = 68;
  xlPieExploded = 69;
  xl3DPieExploded = 70;
  xlBarOfPie = 71;
  xlXYScatterSmooth = 72;
  xlXYScatterSmoothNoMarkers = 73;
  xlXYScatterLines = 74;
  xlXYScatterLinesNoMarkers = 75;
  xlAreaStacked = 76;
  xlAreaStacked100 = 77;
  xl3DAreaStacked = 78;
  xl3DAreaStacked100 = 79;
  xlDoughnutExploded = 80;
  xlRadarMarkers = 81;
  xlRadarFilled = 82;
  xlSurface = 83;
  xlSurfaceWireframe = 84;
  xlSurfaceTopView = 85;
  xlSurfaceTopViewWireframe = 86;
  xlBubble = 15;
  xlBubble3DEffect = 87;
  xlStockHLC = 88;
  xlStockOHLC = 89;
  xlStockVHLC = 90;
  xlStockVOHLC = 91;
  xlCylinderColClustered = 92;
  xlCylinderColStacked = 93;
  xlCylinderColStacked100 = 94;
  xlCylinderBarClustered = 95;
  xlCylinderBarStacked = 96;
  xlCylinderBarStacked100 = 97;
  xlCylinderCol = 98;
  xlConeColClustered = 99;
  xlConeColStacked = 100;
  xlConeColStacked100 = 101;
  xlConeBarClustered = 102;
  xlConeBarStacked = 103;
  xlConeBarStacked100 = 104;
  xlConeCol = 105;
  xlPyramidColClustered = 106;
  xlPyramidColStacked = 107;
  xlPyramidColStacked100 = 108;
  xlPyramidBarClustered = 109;
  xlPyramidBarStacked = 110;
  xlPyramidBarStacked100 = 111;
  xlPyramidCol = 112;
  xl3DColumn = -4100;
  xlLine = 4;
  xl3DLine = -4101;
  xl3DPie = -4102;
  xlPie = 5;
  xlXYScatter = -4169;
  xl3DArea = -4098;
  xlArea = 1;
  xlDoughnut = -4120;
  xlRadar = -4151;

{ XlChartItem }

  xlDataLabel = 0;
  xlChartArea = 2;
  xlSeries = 3;
  xlChartTitle = 4;
  xlWalls = 5;
  xlCorners = 6;
  xlDataTable = 7;
  xlTrendline = 8;
  xlErrorBars = 9;
  xlXErrorBars = 10;
  xlYErrorBars = 11;
  xlLegendEntry = 12;
  xlLegendKey = 13;
  xlShape = 14;
  xlMajorGridlines = 15;
  xlMinorGridlines = 16;
  xlAxisTitle = 17;
  xlUpBars = 18;
  xlPlotArea = 19;
  xlDownBars = 20;
  xlAxis = 21;
  xlSeriesLines = 22;
  xlFloor = 23;
  xlLegend = 24;
  xlHiLoLines = 25;
  xlDropLines = 26;
  xlRadarAxisLabels = 27;
  xlNothing = 28;
  xlLeaderLines = 29;
  xlDisplayUnitLabel = 30;
  xlPivotChartFieldButton = 31;
  xlPivotChartDropZone = 32;

{ XlSizeRepresents }

  xlSizeIsWidth = 2;
  xlSizeIsArea = 1;

{ XlInsertShiftDirection }

  xlShiftDown = -4121;
  xlShiftToRight = -4161;

{ XlDeleteShiftDirection }

  xlShiftToLeft = -4159;
  xlShiftUp = -4162;

{ XlDirection }

  xlDown = -4121;
  xlToLeft = -4159;
  xlToRight = -4161;
  xlUp = -4162;

{ XlConsolidationFunction }

  xlAverage = -4106;
  xlCount = -4112;
  xlCountNums = -4113;
  xlMax = -4136;
  xlMin = -4139;
  xlProduct = -4149;
  xlStDev = -4155;
  xlStDevP = -4156;
  xlSum = -4157;
  xlVar = -4164;
  xlVarP = -4165;
  xlUnknown = 1000;

{ XlSheetType }

  xlChart = -4109;
  xlDialogSheet = -4116;
  xlExcel4IntlMacroSheet = 4;
  xlExcel4MacroSheet = 3;
  xlWorksheet = -4167;

{ XlLocationInTable }

  xlColumnHeader = -4110;
  xlColumnItem = 5;
  xlDataHeader = 3;
  xlDataItem = 7;
  xlPageHeader = 2;
  xlPageItem = 6;
  xlRowHeader = -4153;
  xlRowItem = 4;
  xlTableBody = 8;

{ XlFindLookIn }

  xlFormulas = -4123;
  xlComments = -4144;
  xlValues = -4163;

{ XlWindowType }

  xlChartAsWindow = 5;
  xlChartInPlace = 4;
  xlClipboard = 3;
  xlInfo = -4129;
  xlWorkbook = 1;

{ XlPivotFieldDataType }

  xlDate = 2;
  xlNumber = -4145;
  xlText = -4158;

{ XlCopyPictureFormat }

  xlBitmap = 2;
  xlPicture = -4147;

{ XlPivotTableSourceType }

  xlScenario = 4;
  xlConsolidation = 3;
  xlDatabase = 1;
  xlExternal = 2;
  xlPivotTable = -4148;

{ XlReferenceStyle }

  xlA1 = 1;
  xlR1C1 = -4150;

{ XlMSApplication }

  xlMicrosoftAccess = 4;
  xlMicrosoftFoxPro = 5;
  xlMicrosoftMail = 3;
  xlMicrosoftPowerPoint = 2;
  xlMicrosoftProject = 6;
  xlMicrosoftSchedulePlus = 7;
  xlMicrosoftWord = 1;

{ XlMouseButton }

  xlNoButton = 0;
  xlPrimaryButton = 1;
  xlSecondaryButton = 2;

{ XlCutCopyMode }

  xlCopy = 1;
  xlCut = 2;

{ XlFillWith }

  xlFillWithAll = -4104;
  xlFillWithContents = 2;
  xlFillWithFormats = -4122;

{ XlFilterAction }

  xlFilterCopy = 2;
  xlFilterInPlace = 1;

{ XlOrder }

  xlDownThenOver = 1;
  xlOverThenDown = 2;

{ XlLinkType }

  xlLinkTypeExcelLinks = 1;
  xlLinkTypeOLELinks = 2;

{ XlApplyNamesOrder }

  xlColumnThenRow = 2;
  xlRowThenColumn = 1;

{ XlEnableCancelKey }

  xlDisabled = 0;
  xlErrorHandler = 2;
  xlInterrupt = 1;

{ XlPageBreak }

  xlPageBreakAutomatic = -4105;
  xlPageBreakManual = -4135;
  xlPageBreakNone = -4142;

{ XlOLEType }

  xlOLEControl = 2;
  xlOLEEmbed = 1;
  xlOLELink = 0;

{ XlPageOrientation }

  xlLandscape = 2;
  xlPortrait = 1;

{ XlLinkInfo }

  xlEditionDate = 2;
  xlUpdateState = 1;
  xlLinkInfoStatus = 3;

{ XlCommandUnderlines }

  xlCommandUnderlinesAutomatic = -4105;
  xlCommandUnderlinesOff = -4146;
  xlCommandUnderlinesOn = 1;

{ XlOLEVerb }

  xlVerbOpen = 2;
  xlVerbPrimary = 1;

{ XlCalculation }

  xlCalculationAutomatic = -4105;
  xlCalculationManual = -4135;
  xlCalculationSemiautomatic = 2;

{ XlFileAccess }

  xlReadOnly = 3;
  xlReadWrite = 2;

{ XlEditionType }

  xlPublisher = 1;
  xlSubscriber = 2;

{ XlObjectSize }

  xlFitToPage = 2;
  xlFullPage = 3;
  xlScreenSize = 1;

{ XlLookAt }

  xlPart = 2;
  xlWhole = 1;

{ XlMailSystem }

  xlMAPI = 1;
  xlNoMailSystem = 0;
  xlPowerTalk = 2;

{ XlLinkInfoType }

  xlLinkInfoOLELinks = 2;
  xlLinkInfoPublishers = 5;
  xlLinkInfoSubscribers = 6;

{ XlCVError }

  xlErrDiv0 = 2007;
  xlErrNA = 2042;
  xlErrName = 2029;
  xlErrNull = 2000;
  xlErrNum = 2036;
  xlErrRef = 2023;
  xlErrValue = 2015;

{ XlEditionFormat }

  xlBIFF = 2;
  xlPICT = 1;
  xlRTF = 4;
  xlVALU = 8;

{ XlLink }

  xlExcelLinks = 1;
  xlOLELinks = 2;
  xlPublishers = 5;
  xlSubscribers = 6;

{ XlCellType }

  xlCellTypeBlanks = 4;
  xlCellTypeConstants = 2;
  xlCellTypeFormulas = -4123;
  xlCellTypeLastCell = 11;
  xlCellTypeComments = -4144;
  xlCellTypeVisible = 12;
  xlCellTypeAllFormatConditions = -4172;
  xlCellTypeSameFormatConditions = -4173;
  xlCellTypeAllValidation = -4174;
  xlCellTypeSameValidation = -4175;

{ XlArrangeStyle }

  xlArrangeStyleCascade = 7;
  xlArrangeStyleHorizontal = -4128;
  xlArrangeStyleTiled = 1;
  xlArrangeStyleVertical = -4166;

{ XlMousePointer }

  xlIBeam = 3;
  xlDefault = -4143;
  xlNorthwestArrow = 1;
  xlWait = 2;

{ XlEditionOptionsOption }

  xlAutomaticUpdate = 4;
  xlCancel = 1;
  xlChangeAttributes = 6;
  xlManualUpdate = 5;
  xlOpenSource = 3;
  xlSelect = 3;
  xlSendPublisher = 2;
  xlUpdateSubscriber = 2;

{ XlAutoFillType }

  xlFillCopy = 1;
  xlFillDays = 5;
  xlFillDefault = 0;
  xlFillFormats = 3;
  xlFillMonths = 7;
  xlFillSeries = 2;
  xlFillValues = 4;
  xlFillWeekdays = 6;
  xlFillYears = 8;
  xlGrowthTrend = 10;
  xlLinearTrend = 9;

{ XlAutoFilterOperator }

  xlAnd = 1;
  xlBottom10Items = 4;
  xlBottom10Percent = 6;
  xlOr = 2;
  xlTop10Items = 3;
  xlTop10Percent = 5;

{ XlClipboardFormat }

  xlClipboardFormatBIFF = 8;
  xlClipboardFormatBIFF2 = 18;
  xlClipboardFormatBIFF3 = 20;
  xlClipboardFormatBIFF4 = 30;
  xlClipboardFormatBinary = 15;
  xlClipboardFormatBitmap = 9;
  xlClipboardFormatCGM = 13;
  xlClipboardFormatCSV = 5;
  xlClipboardFormatDIF = 4;
  xlClipboardFormatDspText = 12;
  xlClipboardFormatEmbeddedObject = 21;
  xlClipboardFormatEmbedSource = 22;
  xlClipboardFormatLink = 11;
  xlClipboardFormatLinkSource = 23;
  xlClipboardFormatLinkSourceDesc = 32;
  xlClipboardFormatMovie = 24;
  xlClipboardFormatNative = 14;
  xlClipboardFormatObjectDesc = 31;
  xlClipboardFormatObjectLink = 19;
  xlClipboardFormatOwnerLink = 17;
  xlClipboardFormatPICT = 2;
  xlClipboardFormatPrintPICT = 3;
  xlClipboardFormatRTF = 7;
  xlClipboardFormatScreenPICT = 29;
  xlClipboardFormatStandardFont = 28;
  xlClipboardFormatStandardScale = 27;
  xlClipboardFormatSYLK = 6;
  xlClipboardFormatTable = 16;
  xlClipboardFormatText = 0;
  xlClipboardFormatToolFace = 25;
  xlClipboardFormatToolFacePICT = 26;
  xlClipboardFormatVALU = 1;
  xlClipboardFormatWK1 = 10;

{ XlFileFormat }

  xlAddIn = 18;
  xlCSV = 6;
  xlCSVMac = 22;
  xlCSVMSDOS = 24;
  xlCSVWindows = 23;
  xlDBF2 = 7;
  xlDBF3 = 8;
  xlDBF4 = 11;
  xlDIF = 9;
  xlExcel2 = 16;
  xlExcel2FarEast = 27;
  xlExcel3 = 29;
  xlExcel4 = 33;
  xlExcel5 = 39;
  xlExcel7 = 39;
  xlExcel9795 = 43;
  xlExcel4Workbook = 35;
  xlIntlAddIn = 26;
  xlIntlMacro = 25;
  xlWorkbookNormal = -4143;
  xlSYLK = 2;
  xlTemplate = 17;
  xlCurrentPlatformText = -4158;
  xlTextMac = 19;
  xlTextMSDOS = 21;
  xlTextPrinter = 36;
  xlTextWindows = 20;
  xlWJ2WD1 = 14;
  xlWK1 = 5;
  xlWK1ALL = 31;
  xlWK1FMT = 30;
  xlWK3 = 15;
  xlWK4 = 38;
  xlWK3FM3 = 32;
  xlWKS = 4;
  xlWorks2FarEast = 28;
  xlWQ1 = 34;
  xlWJ3 = 40;
  xlWJ3FJ3 = 41;
  xlUnicodeText = 42;
  xlHtml = 44;
  xlWebArchive = 45;
  xlXMLSpreadsheet = 46;

{ XlApplicationInternational }

  xl24HourClock = 33;
  xl4DigitYears = 43;
  xlAlternateArraySeparator = 16;
  xlColumnSeparator = 14;
  xlCountryCode = 1;
  xlCountrySetting = 2;
  xlCurrencyBefore = 37;
  xlCurrencyCode = 25;
  xlCurrencyDigits = 27;
  xlCurrencyLeadingZeros = 40;
  xlCurrencyMinusSign = 38;
  xlCurrencyNegative = 28;
  xlCurrencySpaceBefore = 36;
  xlCurrencyTrailingZeros = 39;
  xlDateOrder = 32;
  xlDateSeparator = 17;
  xlDayCode = 21;
  xlDayLeadingZero = 42;
  xlDecimalSeparator = 3;
  xlGeneralFormatName = 26;
  xlHourCode = 22;
  xlLeftBrace = 12;
  xlLeftBracket = 10;
  xlListSeparator = 5;
  xlLowerCaseColumnLetter = 9;
  xlLowerCaseRowLetter = 8;
  xlMDY = 44;
  xlMetric = 35;
  xlMinuteCode = 23;
  xlMonthCode = 20;
  xlMonthLeadingZero = 41;
  xlMonthNameChars = 30;
  xlNoncurrencyDigits = 29;
  xlNonEnglishFunctions = 34;
  xlRightBrace = 13;
  xlRightBracket = 11;
  xlRowSeparator = 15;
  xlSecondCode = 24;
  xlThousandsSeparator = 4;
  xlTimeLeadingZero = 45;
  xlTimeSeparator = 18;
  xlUpperCaseColumnLetter = 7;
  xlUpperCaseRowLetter = 6;
  xlWeekdayNameChars = 31;
  xlYearCode = 19;

{ XlPageBreakExtent }

  xlPageBreakFull = 1;
  xlPageBreakPartial = 2;

{ XlCellInsertionMode }

  xlOverwriteCells = 0;
  xlInsertDeleteCells = 1;
  xlInsertEntireRows = 2;

{ XlFormulaLabel }

  xlNoLabels = -4142;
  xlRowLabels = 1;
  xlColumnLabels = 2;
  xlMixedLabels = 3;

{ XlHighlightChangesTime }

  xlSinceMyLastSave = 1;
  xlAllChanges = 2;
  xlNotYetReviewed = 3;

{ XlCommentDisplayMode }

  xlNoIndicator = 0;
  xlCommentIndicatorOnly = -1;
  xlCommentAndIndicator = 1;

{ XlFormatConditionType }

  xlCellValue = 1;
  xlExpression = 2;

{ XlFormatConditionOperator }

  xlBetween = 1;
  xlNotBetween = 2;
  xlEqual = 3;
  xlNotEqual = 4;
  xlGreater = 5;
  xlLess = 6;
  xlGreaterEqual = 7;
  xlLessEqual = 8;

{ XlEnableSelection }

  xlNoRestrictions = 0;
  xlUnlockedCells = 1;
  xlNoSelection = -4142;

{ XlDVType }

  xlValidateInputOnly = 0;
  xlValidateWholeNumber = 1;
  xlValidateDecimal = 2;
  xlValidateList = 3;
  xlValidateDate = 4;
  xlValidateTime = 5;
  xlValidateTextLength = 6;
  xlValidateCustom = 7;

{ XlIMEMode }

  xlIMEModeNoControl = 0;
  xlIMEModeOn = 1;
  xlIMEModeOff = 2;
  xlIMEModeDisable = 3;
  xlIMEModeHiragana = 4;
  xlIMEModeKatakana = 5;
  xlIMEModeKatakanaHalf = 6;
  xlIMEModeAlphaFull = 7;
  xlIMEModeAlpha = 8;
  xlIMEModeHangulFull = 9;
  xlIMEModeHangul = 10;

{ XlDVAlertStyle }

  xlValidAlertStop = 1;
  xlValidAlertWarning = 2;
  xlValidAlertInformation = 3;

{ XlChartLocation }

  xlLocationAsNewSheet = 1;
  xlLocationAsObject = 2;
  xlLocationAutomatic = 3;

{ XlPaperSize }

  xlPaper10x14 = 16;
  xlPaper11x17 = 17;
  xlPaperA3 = 8;
  xlPaperA4 = 9;
  xlPaperA4Small = 10;
  xlPaperA5 = 11;
  xlPaperB4 = 12;
  xlPaperB5 = 13;
  xlPaperCsheet = 24;
  xlPaperDsheet = 25;
  xlPaperEnvelope10 = 20;
  xlPaperEnvelope11 = 21;
  xlPaperEnvelope12 = 22;
  xlPaperEnvelope14 = 23;
  xlPaperEnvelope9 = 19;
  xlPaperEnvelopeB4 = 33;
  xlPaperEnvelopeB5 = 34;
  xlPaperEnvelopeB6 = 35;
  xlPaperEnvelopeC3 = 29;
  xlPaperEnvelopeC4 = 30;
  xlPaperEnvelopeC5 = 28;
  xlPaperEnvelopeC6 = 31;
  xlPaperEnvelopeC65 = 32;
  xlPaperEnvelopeDL = 27;
  xlPaperEnvelopeItaly = 36;
  xlPaperEnvelopeMonarch = 37;
  xlPaperEnvelopePersonal = 38;
  xlPaperEsheet = 26;
  xlPaperExecutive = 7;
  xlPaperFanfoldLegalGerman = 41;
  xlPaperFanfoldStdGerman = 40;
  xlPaperFanfoldUS = 39;
  xlPaperFolio = 14;
  xlPaperLedger = 4;
  xlPaperLegal = 5;
  xlPaperLetter = 1;
  xlPaperLetterSmall = 2;
  xlPaperNote = 18;
  xlPaperQuarto = 15;
  xlPaperStatement = 6;
  xlPaperTabloid = 3;
  xlPaperUser = 256;

{ XlPasteSpecialOperation }

  xlPasteSpecialOperationAdd = 2;
  xlPasteSpecialOperationDivide = 5;
  xlPasteSpecialOperationMultiply = 4;
  xlPasteSpecialOperationNone = -4142;
  xlPasteSpecialOperationSubtract = 3;

{ XlPasteType }

  xlPasteAll = -4104;
  xlPasteAllExceptBorders = 7;
  xlPasteFormats = -4122;
  xlPasteFormulas = -4123;
  xlPasteComments = -4144;
  xlPasteValues = -4163;
  xlPasteColumnWidths = 8;
  xlPasteValidation = 6;
  xlPasteFormulasAndNumberFormats = 11;
  xlPasteValuesAndNumberFormats = 12;

{ XlPhoneticCharacterType }

  xlKatakanaHalf = 0;
  xlKatakana = 1;
  xlHiragana = 2;
  xlNoConversion = 3;

{ XlPhoneticAlignment }

  xlPhoneticAlignNoControl = 0;
  xlPhoneticAlignLeft = 1;
  xlPhoneticAlignCenter = 2;
  xlPhoneticAlignDistributed = 3;

{ XlPictureAppearance }

  xlPrinter = 2;
  xlScreen = 1;

{ XlPivotFieldOrientation }

  xlColumnField = 2;
  xlDataField = 4;
  xlHidden = 0;
  xlPageField = 3;
  xlRowField = 1;

{ XlPivotFieldCalculation }

  xlDifferenceFrom = 2;
  xlIndex = 9;
  xlNoAdditionalCalculation = -4143;
  xlPercentDifferenceFrom = 4;
  xlPercentOf = 3;
  xlPercentOfColumn = 7;
  xlPercentOfRow = 6;
  xlPercentOfTotal = 8;
  xlRunningTotal = 5;

{ XlPlacement }

  xlFreeFloating = 3;
  xlMove = 2;
  xlMoveAndSize = 1;

{ XlPlatform }

  xlMacintosh = 1;
  xlMSDOS = 3;
  xlWindows = 2;

{ XlPrintLocation }

  xlPrintSheetEnd = 1;
  xlPrintInPlace = 16;
  xlPrintNoComments = -4142;

{ XlPriority }

  xlPriorityHigh = -4127;
  xlPriorityLow = -4134;
  xlPriorityNormal = -4143;

{ XlPTSelectionMode }

  xlLabelOnly = 1;
  xlDataAndLabel = 0;
  xlDataOnly = 2;
  xlOrigin = 3;
  xlButton = 15;
  xlBlanks = 4;
  xlFirstRow = 256;

{ XlRangeAutoFormat }

  xlRangeAutoFormat3DEffects1 = 13;
  xlRangeAutoFormat3DEffects2 = 14;
  xlRangeAutoFormatAccounting1 = 4;
  xlRangeAutoFormatAccounting2 = 5;
  xlRangeAutoFormatAccounting3 = 6;
  xlRangeAutoFormatAccounting4 = 17;
  xlRangeAutoFormatClassic1 = 1;
  xlRangeAutoFormatClassic2 = 2;
  xlRangeAutoFormatClassic3 = 3;
  xlRangeAutoFormatColor1 = 7;
  xlRangeAutoFormatColor2 = 8;
  xlRangeAutoFormatColor3 = 9;
  xlRangeAutoFormatList1 = 10;
  xlRangeAutoFormatList2 = 11;
  xlRangeAutoFormatList3 = 12;
  xlRangeAutoFormatLocalFormat1 = 15;
  xlRangeAutoFormatLocalFormat2 = 16;
  xlRangeAutoFormatLocalFormat3 = 19;
  xlRangeAutoFormatLocalFormat4 = 20;
  xlRangeAutoFormatReport1 = 21;
  xlRangeAutoFormatReport2 = 22;
  xlRangeAutoFormatReport3 = 23;
  xlRangeAutoFormatReport4 = 24;
  xlRangeAutoFormatReport5 = 25;
  xlRangeAutoFormatReport6 = 26;
  xlRangeAutoFormatReport7 = 27;
  xlRangeAutoFormatReport8 = 28;
  xlRangeAutoFormatReport9 = 29;
  xlRangeAutoFormatReport10 = 30;
  xlRangeAutoFormatClassicPivotTable = 31;
  xlRangeAutoFormatTable1 = 32;
  xlRangeAutoFormatTable2 = 33;
  xlRangeAutoFormatTable3 = 34;
  xlRangeAutoFormatTable4 = 35;
  xlRangeAutoFormatTable5 = 36;
  xlRangeAutoFormatTable6 = 37;
  xlRangeAutoFormatTable7 = 38;
  xlRangeAutoFormatTable8 = 39;
  xlRangeAutoFormatTable9 = 40;
  xlRangeAutoFormatTable10 = 41;
  xlRangeAutoFormatPTNone = 42;
  xlRangeAutoFormatNone = -4142;
  xlRangeAutoFormatSimple = -4154;

{ XlReferenceType }

  xlAbsolute = 1;
  xlAbsRowRelColumn = 2;
  xlRelative = 4;
  xlRelRowAbsColumn = 3;

{ XlLayoutFormType }

  xlTabular = 0;
  xlOutline = 1;

{ XlRoutingSlipDelivery }

  xlAllAtOnce = 2;
  xlOneAfterAnother = 1;

{ XlRoutingSlipStatus }

  xlNotYetRouted = 0;
  xlRoutingComplete = 2;
  xlRoutingInProgress = 1;

{ XlRunAutoMacro }

  xlAutoActivate = 3;
  xlAutoClose = 2;
  xlAutoDeactivate = 4;
  xlAutoOpen = 1;

{ XlSaveAction }

  xlDoNotSaveChanges = 2;
  xlSaveChanges = 1;

{ XlSaveAsAccessMode }

  xlExclusive = 3;
  xlNoChange = 1;
  xlShared = 2;

{ XlSaveConflictResolution }

  xlLocalSessionChanges = 2;
  xlOtherSessionChanges = 3;
  xlUserResolution = 1;

{ XlSearchDirection }

  xlNext = 1;
  xlPrevious = 2;

{ XlSearchOrder }

  xlByColumns = 2;
  xlByRows = 1;

{ XlSheetVisibility }

  xlSheetVisible = -1;
  xlSheetHidden = 0;
  xlSheetVeryHidden = 2;

{ XlSortMethod }

  xlPinYin = 1;
  xlStroke = 2;

{ XlSortMethodOld }

  xlCodePage = 2;
  xlSyllabary = 1;

{ XlSortOrder }

  xlAscending = 1;
  xlDescending = 2;

{ XlSortOrientation }

  xlSortRows = 2;
  xlSortColumns = 1;

{ XlSortType }

  xlSortLabels = 2;
  xlSortValues = 1;

{ XlSpecialCellsValue }

  xlErrors = 16;
  xlLogical = 4;
  xlNumbers = 1;
  xlTextValues = 2;

{ XlSubscribeToFormat }

  xlSubscribeToPicture = -4147;
  xlSubscribeToText = -4158;

{ XlSummaryRow }

  xlSummaryAbove = 0;
  xlSummaryBelow = 1;

{ XlSummaryColumn }

  xlSummaryOnLeft = -4131;
  xlSummaryOnRight = -4152;

{ XlSummaryReportType }

  xlSummaryPivotTable = -4148;
  xlStandardSummary = 1;

{ XlTabPosition }

  xlTabPositionFirst = 0;
  xlTabPositionLast = 1;

{ XlTextParsingType }

  xlDelimited = 1;
  xlFixedWidth = 2;

{ XlTextQualifier }

  xlTextQualifierDoubleQuote = 1;
  xlTextQualifierNone = -4142;
  xlTextQualifierSingleQuote = 2;

{ XlWBATemplate }

  xlWBATChart = -4109;
  xlWBATExcel4IntlMacroSheet = 4;
  xlWBATExcel4MacroSheet = 3;
  xlWBATWorksheet = -4167;

{ XlWindowView }

  xlNormalView = 1;
  xlPageBreakPreview = 2;

{ XlXLMMacroType }

  xlCommand = 2;
  xlFunction = 1;
  xlNotXLM = 3;

{ XlYesNoGuess }

  xlGuess = 0;
  xlNo = 2;
  xlYes = 1;

{ XlBordersIndex }

  xlInsideHorizontal = 12;
  xlInsideVertical = 11;
  xlDiagonalDown = 5;
  xlDiagonalUp = 6;
  xlEdgeBottom = 9;
  xlEdgeLeft = 7;
  xlEdgeRight = 10;
  xlEdgeTop = 8;

{ XlToolbarProtection }

  xlNoButtonChanges = 1;
  xlNoChanges = 4;
  xlNoDockingChanges = 3;
  xlToolbarProtectionNone = -4143;
  xlNoShapeChanges = 2;

{ XlBuiltInDialog }

  xlDialogOpen = 1;
  xlDialogOpenLinks = 2;
  xlDialogSaveAs = 5;
  xlDialogFileDelete = 6;
  xlDialogPageSetup = 7;
  xlDialogPrint = 8;
  xlDialogPrinterSetup = 9;
  xlDialogArrangeAll = 12;
  xlDialogWindowSize = 13;
  xlDialogWindowMove = 14;
  xlDialogRun = 17;
  xlDialogSetPrintTitles = 23;
  xlDialogFont = 26;
  xlDialogDisplay = 27;
  xlDialogProtectDocument = 28;
  xlDialogCalculation = 32;
  xlDialogExtract = 35;
  xlDialogDataDelete = 36;
  xlDialogSort = 39;
  xlDialogDataSeries = 40;
  xlDialogTable = 41;
  xlDialogFormatNumber = 42;
  xlDialogAlignment = 43;
  xlDialogStyle = 44;
  xlDialogBorder = 45;
  xlDialogCellProtection = 46;
  xlDialogColumnWidth = 47;
  xlDialogClear = 52;
  xlDialogPasteSpecial = 53;
  xlDialogEditDelete = 54;
  xlDialogInsert = 55;
  xlDialogPasteNames = 58;
  xlDialogDefineName = 61;
  xlDialogCreateNames = 62;
  xlDialogFormulaGoto = 63;
  xlDialogFormulaFind = 64;
  xlDialogGalleryArea = 67;
  xlDialogGalleryBar = 68;
  xlDialogGalleryColumn = 69;
  xlDialogGalleryLine = 70;
  xlDialogGalleryPie = 71;
  xlDialogGalleryScatter = 72;
  xlDialogCombination = 73;
  xlDialogGridlines = 76;
  xlDialogAxes = 78;
  xlDialogAttachText = 80;
  xlDialogPatterns = 84;
  xlDialogMainChart = 85;
  xlDialogOverlay = 86;
  xlDialogScale = 87;
  xlDialogFormatLegend = 88;
  xlDialogFormatText = 89;
  xlDialogParse = 91;
  xlDialogUnhide = 94;
  xlDialogWorkspace = 95;
  xlDialogActivate = 103;
  xlDialogCopyPicture = 108;
  xlDialogDeleteName = 110;
  xlDialogDeleteFormat = 111;
  xlDialogNew = 119;
  xlDialogRowHeight = 127;
  xlDialogFormatMove = 128;
  xlDialogFormatSize = 129;
  xlDialogFormulaReplace = 130;
  xlDialogSelectSpecial = 132;
  xlDialogApplyNames = 133;
  xlDialogReplaceFont = 134;
  xlDialogSplit = 137;
  xlDialogOutline = 142;
  xlDialogSaveWorkbook = 145;
  xlDialogCopyChart = 147;
  xlDialogFormatFont = 150;
  xlDialogNote = 154;
  xlDialogSetUpdateStatus = 159;
  xlDialogColorPalette = 161;
  xlDialogChangeLink = 166;
  xlDialogAppMove = 170;
  xlDialogAppSize = 171;
  xlDialogMainChartType = 185;
  xlDialogOverlayChartType = 186;
  xlDialogOpenMail = 188;
  xlDialogSendMail = 189;
  xlDialogStandardFont = 190;
  xlDialogConsolidate = 191;
  xlDialogSortSpecial = 192;
  xlDialogGallery3dArea = 193;
  xlDialogGallery3dColumn = 194;
  xlDialogGallery3dLine = 195;
  xlDialogGallery3dPie = 196;
  xlDialogView3d = 197;
  xlDialogGoalSeek = 198;
  xlDialogWorkgroup = 199;
  xlDialogFillGroup = 200;
  xlDialogUpdateLink = 201;
  xlDialogPromote = 202;
  xlDialogDemote = 203;
  xlDialogShowDetail = 204;
  xlDialogObjectProperties = 207;
  xlDialogSaveNewObject = 208;
  xlDialogApplyStyle = 212;
  xlDialogAssignToObject = 213;
  xlDialogObjectProtection = 214;
  xlDialogCreatePublisher = 217;
  xlDialogSubscribeTo = 218;
  xlDialogShowToolbar = 220;
  xlDialogPrintPreview = 222;
  xlDialogEditColor = 223;
  xlDialogFormatMain = 225;
  xlDialogFormatOverlay = 226;
  xlDialogEditSeries = 228;
  xlDialogDefineStyle = 229;
  xlDialogGalleryRadar = 249;
  xlDialogEditionOptions = 251;
  xlDialogZoom = 256;
  xlDialogInsertObject = 259;
  xlDialogSize = 261;
  xlDialogMove = 262;
  xlDialogFormatAuto = 269;
  xlDialogGallery3dBar = 272;
  xlDialogGallery3dSurface = 273;
  xlDialogCustomizeToolbar = 276;
  xlDialogWorkbookAdd = 281;
  xlDialogWorkbookMove = 282;
  xlDialogWorkbookCopy = 283;
  xlDialogWorkbookOptions = 284;
  xlDialogSaveWorkspace = 285;
  xlDialogChartWizard = 288;
  xlDialogAssignToTool = 293;
  xlDialogPlacement = 300;
  xlDialogFillWorkgroup = 301;
  xlDialogWorkbookNew = 302;
  xlDialogScenarioCells = 305;
  xlDialogScenarioAdd = 307;
  xlDialogScenarioEdit = 308;
  xlDialogScenarioSummary = 311;
  xlDialogPivotTableWizard = 312;
  xlDialogPivotFieldProperties = 313;
  xlDialogOptionsCalculation = 318;
  xlDialogOptionsEdit = 319;
  xlDialogOptionsView = 320;
  xlDialogAddinManager = 321;
  xlDialogMenuEditor = 322;
  xlDialogAttachToolbars = 323;
  xlDialogOptionsChart = 325;
  xlDialogVbaInsertFile = 328;
  xlDialogVbaProcedureDefinition = 330;
  xlDialogRoutingSlip = 336;
  xlDialogMailLogon = 339;
  xlDialogInsertPicture = 342;
  xlDialogGalleryDoughnut = 344;
  xlDialogChartTrend = 350;
  xlDialogWorkbookInsert = 354;
  xlDialogOptionsTransition = 355;
  xlDialogOptionsGeneral = 356;
  xlDialogFilterAdvanced = 370;
  xlDialogMailNextLetter = 378;
  xlDialogDataLabel = 379;
  xlDialogInsertTitle = 380;
  xlDialogFontProperties = 381;
  xlDialogMacroOptions = 382;
  xlDialogWorkbookUnhide = 384;
  xlDialogWorkbookName = 386;
  xlDialogGalleryCustom = 388;
  xlDialogAddChartAutoformat = 390;
  xlDialogChartAddData = 392;
  xlDialogTabOrder = 394;
  xlDialogSubtotalCreate = 398;
  xlDialogWorkbookTabSplit = 415;
  xlDialogWorkbookProtect = 417;
  xlDialogScrollbarProperties = 420;
  xlDialogPivotShowPages = 421;
  xlDialogTextToColumns = 422;
  xlDialogFormatCharttype = 423;
  xlDialogPivotFieldGroup = 433;
  xlDialogPivotFieldUngroup = 434;
  xlDialogCheckboxProperties = 435;
  xlDialogLabelProperties = 436;
  xlDialogListboxProperties = 437;
  xlDialogEditboxProperties = 438;
  xlDialogOpenText = 441;
  xlDialogPushbuttonProperties = 445;
  xlDialogFilter = 447;
  xlDialogFunctionWizard = 450;
  xlDialogSaveCopyAs = 456;
  xlDialogOptionsListsAdd = 458;
  xlDialogSeriesAxes = 460;
  xlDialogSeriesX = 461;
  xlDialogSeriesY = 462;
  xlDialogErrorbarX = 463;
  xlDialogErrorbarY = 464;
  xlDialogFormatChart = 465;
  xlDialogSeriesOrder = 466;
  xlDialogMailEditMailer = 470;
  xlDialogStandardWidth = 472;
  xlDialogScenarioMerge = 473;
  xlDialogProperties = 474;
  xlDialogSummaryInfo = 474;
  xlDialogFindFile = 475;
  xlDialogActiveCellFont = 476;
  xlDialogVbaMakeAddin = 478;
  xlDialogFileSharing = 481;
  xlDialogAutoCorrect = 485;
  xlDialogCustomViews = 493;
  xlDialogInsertNameLabel = 496;
  xlDialogSeriesShape = 504;
  xlDialogChartOptionsDataLabels = 505;
  xlDialogChartOptionsDataTable = 506;
  xlDialogSetBackgroundPicture = 509;
  xlDialogDataValidation = 525;
  xlDialogChartType = 526;
  xlDialogChartLocation = 527;
  _xlDialogPhonetic = 538;
  xlDialogChartSourceData = 540;
  _xlDialogChartSourceData = 541;
  xlDialogSeriesOptions = 557;
  xlDialogPivotTableOptions = 567;
  xlDialogPivotSolveOrder = 568;
  xlDialogPivotCalculatedField = 570;
  xlDialogPivotCalculatedItem = 572;
  xlDialogConditionalFormatting = 583;
  xlDialogInsertHyperlink = 596;
  xlDialogProtectSharing = 620;
  xlDialogOptionsME = 647;
  xlDialogPublishAsWebPage = 653;
  xlDialogPhonetic = 656;
  xlDialogNewWebQuery = 667;
  xlDialogImportTextFile = 666;
  xlDialogExternalDataProperties = 530;
  xlDialogWebOptionsGeneral = 683;
  xlDialogWebOptionsFiles = 684;
  xlDialogWebOptionsPictures = 685;
  xlDialogWebOptionsEncoding = 686;
  xlDialogWebOptionsFonts = 687;
  xlDialogPivotClientServerSet = 689;
  xlDialogPropertyFields = 754;
  xlDialogSearch = 731;
  xlDialogEvaluateFormula = 709;
  xlDialogDataLabelMultiple = 723;
  xlDialogChartOptionsDataLabelMultiple = 724;
  xlDialogErrorChecking = 732;
  xlDialogWebOptionsBrowsers = 773;
  xlDialogCreateList = 796;
  xlDialogPermission = 832;
  xlDialogMyPermission = 834;

{ XlParameterType }

  xlPrompt = 0;
  xlConstant = 1;
  xlRange = 2;

{ XlParameterDataType }

  xlParamTypeUnknown = 0;
  xlParamTypeChar = 1;
  xlParamTypeNumeric = 2;
  xlParamTypeDecimal = 3;
  xlParamTypeInteger = 4;
  xlParamTypeSmallInt = 5;
  xlParamTypeFloat = 6;
  xlParamTypeReal = 7;
  xlParamTypeDouble = 8;
  xlParamTypeVarChar = 12;
  xlParamTypeDate = 9;
  xlParamTypeTime = 10;
  xlParamTypeTimestamp = 11;
  xlParamTypeLongVarChar = -1;
  xlParamTypeBinary = -2;
  xlParamTypeVarBinary = -3;
  xlParamTypeLongVarBinary = -4;
  xlParamTypeBigInt = -5;
  xlParamTypeTinyInt = -6;
  xlParamTypeBit = -7;
  xlParamTypeWChar = -8;

{ XlFormControl }

  xlButtonControl = 0;
  xlCheckBox = 1;
  xlDropDown = 2;
  xlEditBox = 3;
  xlGroupBox = 4;
  xlLabel = 5;
  xlListBox = 6;
  xlOptionButton = 7;
  xlScrollBar = 8;
  xlSpinner = 9;

{ XlSourceType }

  xlSourceWorkbook = 0;
  xlSourceSheet = 1;
  xlSourcePrintArea = 2;
  xlSourceAutoFilter = 3;
  xlSourceRange = 4;
  xlSourceChart = 5;
  xlSourcePivotTable = 6;
  xlSourceQuery = 7;

{ XlHtmlType }

  xlHtmlStatic = 0;
  xlHtmlCalc = 1;
  xlHtmlList = 2;
  xlHtmlChart = 3;

{ XlPivotFormatType }

  xlReport1 = 0;
  xlReport2 = 1;
  xlReport3 = 2;
  xlReport4 = 3;
  xlReport5 = 4;
  xlReport6 = 5;
  xlReport7 = 6;
  xlReport8 = 7;
  xlReport9 = 8;
  xlReport10 = 9;
  xlTable1 = 10;
  xlTable2 = 11;
  xlTable3 = 12;
  xlTable4 = 13;
  xlTable5 = 14;
  xlTable6 = 15;
  xlTable7 = 16;
  xlTable8 = 17;
  xlTable9 = 18;
  xlTable10 = 19;
  xlPTClassic = 20;
  xlPTNone = 21;

{ XlCmdType }

  xlCmdCube = 1;
  xlCmdSql = 2;
  xlCmdTable = 3;
  xlCmdDefault = 4;
  xlCmdList = 5;

{ XlColumnDataType }

  xlGeneralFormat = 1;
  xlTextFormat = 2;
  xlMDYFormat = 3;
  xlDMYFormat = 4;
  xlYMDFormat = 5;
  xlMYDFormat = 6;
  xlDYMFormat = 7;
  xlYDMFormat = 8;
  xlSkipColumn = 9;
  xlEMDFormat = 10;

{ XlQueryType }

  xlODBCQuery = 1;
  xlDAORecordset = 2;
  xlWebQuery = 4;
  xlOLEDBQuery = 5;
  xlTextImport = 6;
  xlADORecordset = 7;

{ XlWebSelectionType }

  xlEntirePage = 1;
  xlAllTables = 2;
  xlSpecifiedTables = 3;

{ XlCubeFieldType }

  xlHierarchy = 1;
  xlMeasure = 2;
  xlSet = 3;

{ XlWebFormatting }

  xlWebFormattingAll = 1;
  xlWebFormattingRTF = 2;
  xlWebFormattingNone = 3;

{ XlDisplayDrawingObjects }

  xlDisplayShapes = -4104;
  xlHide = 3;
  xlPlaceholders = 2;

{ XlSubtototalLocationType }

  xlAtTop = 1;
  xlAtBottom = 2;

{ XlPivotTableVersionList }

  xlPivotTableVersion2000 = 0;
  xlPivotTableVersion10 = 1;
  xlPivotTableVersionCurrent = -1;

{ XlPrintErrors }

  xlPrintErrorsDisplayed = 0;
  xlPrintErrorsBlank = 1;
  xlPrintErrorsDash = 2;
  xlPrintErrorsNA = 3;

{ XlPivotCellType }

  xlPivotCellValue = 0;
  xlPivotCellPivotItem = 1;
  xlPivotCellSubtotal = 2;
  xlPivotCellGrandTotal = 3;
  xlPivotCellDataField = 4;
  xlPivotCellPivotField = 5;
  xlPivotCellPageFieldItem = 6;
  xlPivotCellCustomSubtotal = 7;
  xlPivotCellDataPivotField = 8;
  xlPivotCellBlankCell = 9;

{ XlPivotTableMissingItems }

  xlMissingItemsDefault = -1;
  xlMissingItemsNone = 0;
  xlMissingItemsMax = 32500;

{ XlCalculationState }

  xlDone = 0;
  xlCalculating = 1;
  xlPending = 2;

{ XlCalculationInterruptKey }

  xlNoKey = 0;
  xlEscKey = 1;
  xlAnyKey = 2;

{ XlSortDataOption }

  xlSortNormal = 0;
  xlSortTextAsNumbers = 1;

{ XlUpdateLinks }

  xlUpdateLinksUserSetting = 1;
  xlUpdateLinksNever = 2;
  xlUpdateLinksAlways = 3;

{ XlLinkStatus }

  xlLinkStatusOK = 0;
  xlLinkStatusMissingFile = 1;
  xlLinkStatusMissingSheet = 2;
  xlLinkStatusOld = 3;
  xlLinkStatusSourceNotCalculated = 4;
  xlLinkStatusIndeterminate = 5;
  xlLinkStatusNotStarted = 6;
  xlLinkStatusInvalidName = 7;
  xlLinkStatusSourceNotOpen = 8;
  xlLinkStatusSourceOpen = 9;
  xlLinkStatusCopiedValues = 10;

{ XlSearchWithin }

  xlWithinSheet = 1;
  xlWithinWorkbook = 2;

{ XlCorruptLoad }

  xlNormalLoad = 0;
  xlRepairFile = 1;
  xlExtractData = 2;

{ XlRobustConnect }

  xlAsRequired = 0;
  xlAlways = 1;
  xlNever = 2;

{ XlErrorChecks }

  xlEvaluateToError = 1;
  xlTextDate = 2;
  xlNumberAsText = 3;
  xlInconsistentFormula = 4;
  xlOmittedCells = 5;
  xlUnlockedFormulaCells = 6;
  xlEmptyCellReferences = 7;
  xlListDataValidation = 8;

{ XlDataLabelSeparator }

  xlDataLabelSeparatorDefault = 1;

{ XlSmartTagDisplayMode }

  xlIndicatorAndButton = 0;
  xlDisplayNone = 1;
  xlButtonOnly = 2;

{ XlRangeValueDataType }

  xlRangeValueDefault = 10;
  xlRangeValueXMLSpreadsheet = 11;
  xlRangeValueMSPersistXML = 12;

{ XlSpeakDirection }

  xlSpeakByRows = 0;
  xlSpeakByColumns = 1;

{ XlInsertFormatOrigin }

  xlFormatFromLeftOrAbove = 0;
  xlFormatFromRightOrBelow = 1;

{ XlArabicModes }

  xlArabicNone = 0;
  xlArabicStrictAlefHamza = 1;
  xlArabicStrictFinalYaa = 2;
  xlArabicBothStrict = 3;

{ XlImportDataAs }

  xlQueryTable = 0;
  xlPivotTableReport = 1;

{ XlCalculatedMemberType }

  xlCalculatedMember = 0;
  xlCalculatedSet = 1;

{ XlHebrewModes }

  xlHebrewFullScript = 0;
  xlHebrewPartialScript = 1;
  xlHebrewMixedScript = 2;
  xlHebrewMixedAuthorizedScript = 3;

{ XlListObjectSourceType }

  xlSrcExternal = 0;
  xlSrcRange = 1;
  xlSrcXml = 2;

{ XlTextVisualLayoutType }

  xlTextVisualLTR = 1;
  xlTextVisualRTL = 2;

{ XlListDataType }

  xlListDataTypeNone = 0;
  xlListDataTypeText = 1;
  xlListDataTypeMultiLineText = 2;
  xlListDataTypeNumber = 3;
  xlListDataTypeCurrency = 4;
  xlListDataTypeDateTime = 5;
  xlListDataTypeChoice = 6;
  xlListDataTypeChoiceMulti = 7;
  xlListDataTypeListLookup = 8;
  xlListDataTypeCheckbox = 9;
  xlListDataTypeHyperLink = 10;
  xlListDataTypeCounter = 11;
  xlListDataTypeMultiLineRichText = 12;

{ XlTotalsCalculation }

  xlTotalsCalculationNone = 0;
  xlTotalsCalculationSum = 1;
  xlTotalsCalculationAverage = 2;
  xlTotalsCalculationCount = 3;
  xlTotalsCalculationCountNums = 4;
  xlTotalsCalculationMin = 5;
  xlTotalsCalculationMax = 6;
  xlTotalsCalculationStdDev = 7;
  xlTotalsCalculationVar = 8;

{ XlXmlLoadOption }

  xlXmlLoadPromptUser = 0;
  xlXmlLoadOpenXml = 1;
  xlXmlLoadImportToList = 2;
  xlXmlLoadMapXml = 3;

{ XlSmartTagControlType }

  xlSmartTagControlSmartTag = 1;
  xlSmartTagControlLink = 2;
  xlSmartTagControlHelp = 3;
  xlSmartTagControlHelpURL = 4;
  xlSmartTagControlSeparator = 5;
  xlSmartTagControlButton = 6;
  xlSmartTagControlLabel = 7;
  xlSmartTagControlImage = 8;
  xlSmartTagControlCheckbox = 9;
  xlSmartTagControlTextbox = 10;
  xlSmartTagControlListbox = 11;
  xlSmartTagControlCombo = 12;
  xlSmartTagControlActiveX = 13;
  xlSmartTagControlRadioGroup = 14;

{ XlListConflict }

  xlListConflictDialog = 0;
  xlListConflictRetryAllConflicts = 1;
  xlListConflictDiscardAllConflicts = 2;
  xlListConflictError = 3;

{ XlXmlExportResult }

  xlXmlExportSuccess = 0;
  xlXmlExportValidationFailed = 1;

{ XlXmlImportResult }

  xlXmlImportSuccess = 0;
  xlXmlImportElementsTruncated = 1;
  xlXmlImportValidationFailed = 2;


{ Константы для M$ Outlook}  
{ OlDefaultFolders }

  olFolderDeletedItems = $00000003;
  olFolderOutbox = $00000004;
  olFolderSentMail = $00000005;
  olFolderInbox = $00000006;
  olFolderCalendar = $00000009;
  olFolderContacts = $0000000A;
  olFolderJournal = $0000000B;
  olFolderNotes = $0000000C;
  olFolderTasks = $0000000D;
  olFolderDrafts = $00000010;
  olPublicFoldersAllPublicFolders = $00000012;

