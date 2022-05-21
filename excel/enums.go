package excel

// enum Constants
var Constants = struct {
	XlAll int32
	XlAutomatic int32
	XlBoth int32
	XlCenter int32
	XlChecker int32
	XlCircle int32
	XlCorner int32
	XlCrissCross int32
	XlCross int32
	XlDiamond int32
	XlDistributed int32
	XlDoubleAccounting int32
	XlFixedValue int32
	XlFormats int32
	XlGray16 int32
	XlGray8 int32
	XlGrid int32
	XlHigh int32
	XlInside int32
	XlJustify int32
	XlLightDown int32
	XlLightHorizontal int32
	XlLightUp int32
	XlLightVertical int32
	XlLow int32
	XlManual int32
	XlMinusValues int32
	XlModule int32
	XlNextToAxis int32
	XlNone int32
	XlNotes int32
	XlOff int32
	XlOn int32
	XlPercent int32
	XlPlus int32
	XlPlusValues int32
	XlSemiGray75 int32
	XlShowLabel int32
	XlShowLabelAndPercent int32
	XlShowPercent int32
	XlShowValue int32
	XlSimple int32
	XlSingle int32
	XlSingleAccounting int32
	XlSolid int32
	XlSquare int32
	XlStar int32
	XlStError int32
	XlToolbarButton int32
	XlTriangle int32
	XlGray25 int32
	XlGray50 int32
	XlGray75 int32
	XlBottom int32
	XlLeft int32
	XlRight int32
	XlTop int32
	Xl3DBar int32
	Xl3DSurface int32
	XlBar int32
	XlColumn int32
	XlCombination int32
	XlCustom int32
	XlDefaultAutoFormat int32
	XlMaximum int32
	XlMinimum int32
	XlOpaque int32
	XlTransparent int32
	XlBidi int32
	XlLatin int32
	XlContext int32
	XlLTR int32
	XlRTL int32
	XlFullScript int32
	XlPartialScript int32
	XlMixedScript int32
	XlMixedAuthorizedScript int32
	XlVisualCursor int32
	XlLogicalCursor int32
	XlSystem int32
	XlPartial int32
	XlHindiNumerals int32
	XlBidiCalendar int32
	XlGregorian int32
	XlComplete int32
	XlScale int32
	XlClosed int32
	XlColor1 int32
	XlColor2 int32
	XlColor3 int32
	XlConstants int32
	XlContents int32
	XlBelow int32
	XlCascade int32
	XlCenterAcrossSelection int32
	XlChart4 int32
	XlChartSeries int32
	XlChartShort int32
	XlChartTitles int32
	XlClassic1 int32
	XlClassic2 int32
	XlClassic3 int32
	Xl3DEffects1 int32
	Xl3DEffects2 int32
	XlAbove int32
	XlAccounting1 int32
	XlAccounting2 int32
	XlAccounting3 int32
	XlAccounting4 int32
	XlAdd int32
	XlDebugCodePane int32
	XlDesktop int32
	XlDirect int32
	XlDivide int32
	XlDoubleClosed int32
	XlDoubleOpen int32
	XlDoubleQuote int32
	XlEntireChart int32
	XlExcelMenus int32
	XlExtended int32
	XlFill int32
	XlFirst int32
	XlFloating int32
	XlFormula int32
	XlGeneral int32
	XlGridline int32
	XlIcons int32
	XlImmediatePane int32
	XlInteger int32
	XlLast int32
	XlLastCell int32
	XlList1 int32
	XlList2 int32
	XlList3 int32
	XlLocalFormat1 int32
	XlLocalFormat2 int32
	XlLong int32
	XlLotusHelp int32
	XlMacrosheetCell int32
	XlMixed int32
	XlMultiply int32
	XlNarrow int32
	XlNoDocuments int32
	XlOpen int32
	XlOutside int32
	XlReference int32
	XlSemiautomatic int32
	XlShort int32
	XlSingleQuote int32
	XlStrict int32
	XlSubtract int32
	XlTextBox int32
	XlTiled int32
	XlTitleBar int32
	XlToolbar int32
	XlVisible int32
	XlWatchPane int32
	XlWide int32
	XlWorkbookTab int32
	XlWorksheet4 int32
	XlWorksheetCell int32
	XlWorksheetShort int32
	XlAllExceptBorders int32
	XlLeftToRight int32
	XlTopToBottom int32
	XlVeryHidden int32
	XlDrawingObject int32
}{
	XlAll: -4104,
	XlAutomatic: -4105,
	XlBoth: 1,
	XlCenter: -4108,
	XlChecker: 9,
	XlCircle: 8,
	XlCorner: 2,
	XlCrissCross: 16,
	XlCross: 4,
	XlDiamond: 2,
	XlDistributed: -4117,
	XlDoubleAccounting: 5,
	XlFixedValue: 1,
	XlFormats: -4122,
	XlGray16: 17,
	XlGray8: 18,
	XlGrid: 15,
	XlHigh: -4127,
	XlInside: 2,
	XlJustify: -4130,
	XlLightDown: 13,
	XlLightHorizontal: 11,
	XlLightUp: 14,
	XlLightVertical: 12,
	XlLow: -4134,
	XlManual: -4135,
	XlMinusValues: 3,
	XlModule: -4141,
	XlNextToAxis: 4,
	XlNone: -4142,
	XlNotes: -4144,
	XlOff: -4146,
	XlOn: 1,
	XlPercent: 2,
	XlPlus: 9,
	XlPlusValues: 2,
	XlSemiGray75: 10,
	XlShowLabel: 4,
	XlShowLabelAndPercent: 5,
	XlShowPercent: 3,
	XlShowValue: 2,
	XlSimple: -4154,
	XlSingle: 2,
	XlSingleAccounting: 4,
	XlSolid: 1,
	XlSquare: 1,
	XlStar: 5,
	XlStError: 4,
	XlToolbarButton: 2,
	XlTriangle: 3,
	XlGray25: -4124,
	XlGray50: -4125,
	XlGray75: -4126,
	XlBottom: -4107,
	XlLeft: -4131,
	XlRight: -4152,
	XlTop: -4160,
	Xl3DBar: -4099,
	Xl3DSurface: -4103,
	XlBar: 2,
	XlColumn: 3,
	XlCombination: -4111,
	XlCustom: -4114,
	XlDefaultAutoFormat: -1,
	XlMaximum: 2,
	XlMinimum: 4,
	XlOpaque: 3,
	XlTransparent: 2,
	XlBidi: -5000,
	XlLatin: -5001,
	XlContext: -5002,
	XlLTR: -5003,
	XlRTL: -5004,
	XlFullScript: 1,
	XlPartialScript: 2,
	XlMixedScript: 3,
	XlMixedAuthorizedScript: 4,
	XlVisualCursor: 2,
	XlLogicalCursor: 1,
	XlSystem: 1,
	XlPartial: 3,
	XlHindiNumerals: 3,
	XlBidiCalendar: 3,
	XlGregorian: 2,
	XlComplete: 4,
	XlScale: 3,
	XlClosed: 3,
	XlColor1: 7,
	XlColor2: 8,
	XlColor3: 9,
	XlConstants: 2,
	XlContents: 2,
	XlBelow: 1,
	XlCascade: 7,
	XlCenterAcrossSelection: 7,
	XlChart4: 2,
	XlChartSeries: 17,
	XlChartShort: 6,
	XlChartTitles: 18,
	XlClassic1: 1,
	XlClassic2: 2,
	XlClassic3: 3,
	Xl3DEffects1: 13,
	Xl3DEffects2: 14,
	XlAbove: 0,
	XlAccounting1: 4,
	XlAccounting2: 5,
	XlAccounting3: 6,
	XlAccounting4: 17,
	XlAdd: 2,
	XlDebugCodePane: 13,
	XlDesktop: 9,
	XlDirect: 1,
	XlDivide: 5,
	XlDoubleClosed: 5,
	XlDoubleOpen: 4,
	XlDoubleQuote: 1,
	XlEntireChart: 20,
	XlExcelMenus: 1,
	XlExtended: 3,
	XlFill: 5,
	XlFirst: 0,
	XlFloating: 5,
	XlFormula: 5,
	XlGeneral: 1,
	XlGridline: 22,
	XlIcons: 1,
	XlImmediatePane: 12,
	XlInteger: 2,
	XlLast: 1,
	XlLastCell: 11,
	XlList1: 10,
	XlList2: 11,
	XlList3: 12,
	XlLocalFormat1: 15,
	XlLocalFormat2: 16,
	XlLong: 3,
	XlLotusHelp: 2,
	XlMacrosheetCell: 7,
	XlMixed: 2,
	XlMultiply: 4,
	XlNarrow: 1,
	XlNoDocuments: 3,
	XlOpen: 2,
	XlOutside: 3,
	XlReference: 4,
	XlSemiautomatic: 2,
	XlShort: 1,
	XlSingleQuote: 2,
	XlStrict: 2,
	XlSubtract: 3,
	XlTextBox: 16,
	XlTiled: 1,
	XlTitleBar: 8,
	XlToolbar: 1,
	XlVisible: 12,
	XlWatchPane: 11,
	XlWide: 3,
	XlWorkbookTab: 6,
	XlWorksheet4: 1,
	XlWorksheetCell: 3,
	XlWorksheetShort: 5,
	XlAllExceptBorders: 7,
	XlLeftToRight: 2,
	XlTopToBottom: 1,
	XlVeryHidden: 2,
	XlDrawingObject: 14,
}

// enum XlCreator
var XlCreator = struct {
	XlCreatorCode int32
}{
	XlCreatorCode: 1480803660,
}

// enum XlChartGallery
var XlChartGallery = struct {
	XlBuiltIn int32
	XlUserDefined int32
	XlAnyGallery int32
}{
	XlBuiltIn: 21,
	XlUserDefined: 22,
	XlAnyGallery: 23,
}

// enum XlColorIndex
var XlColorIndex = struct {
	XlColorIndexAutomatic int32
	XlColorIndexNone int32
}{
	XlColorIndexAutomatic: -4105,
	XlColorIndexNone: -4142,
}

// enum XlEndStyleCap
var XlEndStyleCap = struct {
	XlCap int32
	XlNoCap int32
}{
	XlCap: 1,
	XlNoCap: 2,
}

// enum XlRowCol
var XlRowCol = struct {
	XlColumns int32
	XlRows int32
}{
	XlColumns: 2,
	XlRows: 1,
}

// enum XlScaleType
var XlScaleType = struct {
	XlScaleLinear int32
	XlScaleLogarithmic int32
}{
	XlScaleLinear: -4132,
	XlScaleLogarithmic: -4133,
}

// enum XlDataSeriesType
var XlDataSeriesType = struct {
	XlAutoFill int32
	XlChronological int32
	XlGrowth int32
	XlDataSeriesLinear int32
}{
	XlAutoFill: 4,
	XlChronological: 3,
	XlGrowth: 2,
	XlDataSeriesLinear: -4132,
}

// enum XlAxisCrosses
var XlAxisCrosses = struct {
	XlAxisCrossesAutomatic int32
	XlAxisCrossesCustom int32
	XlAxisCrossesMaximum int32
	XlAxisCrossesMinimum int32
}{
	XlAxisCrossesAutomatic: -4105,
	XlAxisCrossesCustom: -4114,
	XlAxisCrossesMaximum: 2,
	XlAxisCrossesMinimum: 4,
}

// enum XlAxisGroup
var XlAxisGroup = struct {
	XlPrimary int32
	XlSecondary int32
}{
	XlPrimary: 1,
	XlSecondary: 2,
}

// enum XlBackground
var XlBackground = struct {
	XlBackgroundAutomatic int32
	XlBackgroundOpaque int32
	XlBackgroundTransparent int32
}{
	XlBackgroundAutomatic: -4105,
	XlBackgroundOpaque: 3,
	XlBackgroundTransparent: 2,
}

// enum XlWindowState
var XlWindowState = struct {
	XlMaximized int32
	XlMinimized int32
	XlNormal int32
}{
	XlMaximized: -4137,
	XlMinimized: -4140,
	XlNormal: -4143,
}

// enum XlAxisType
var XlAxisType = struct {
	XlCategory int32
	XlSeriesAxis int32
	XlValue int32
}{
	XlCategory: 1,
	XlSeriesAxis: 3,
	XlValue: 2,
}

// enum XlArrowHeadLength
var XlArrowHeadLength = struct {
	XlArrowHeadLengthLong int32
	XlArrowHeadLengthMedium int32
	XlArrowHeadLengthShort int32
}{
	XlArrowHeadLengthLong: 3,
	XlArrowHeadLengthMedium: -4138,
	XlArrowHeadLengthShort: 1,
}

// enum XlVAlign
var XlVAlign = struct {
	XlVAlignBottom int32
	XlVAlignCenter int32
	XlVAlignDistributed int32
	XlVAlignJustify int32
	XlVAlignTop int32
}{
	XlVAlignBottom: -4107,
	XlVAlignCenter: -4108,
	XlVAlignDistributed: -4117,
	XlVAlignJustify: -4130,
	XlVAlignTop: -4160,
}

// enum XlTickMark
var XlTickMark = struct {
	XlTickMarkCross int32
	XlTickMarkInside int32
	XlTickMarkNone int32
	XlTickMarkOutside int32
}{
	XlTickMarkCross: 4,
	XlTickMarkInside: 2,
	XlTickMarkNone: -4142,
	XlTickMarkOutside: 3,
}

// enum XlErrorBarDirection
var XlErrorBarDirection = struct {
	XlX int32
	XlY int32
}{
	XlX: -4168,
	XlY: 1,
}

// enum XlErrorBarInclude
var XlErrorBarInclude = struct {
	XlErrorBarIncludeBoth int32
	XlErrorBarIncludeMinusValues int32
	XlErrorBarIncludeNone int32
	XlErrorBarIncludePlusValues int32
}{
	XlErrorBarIncludeBoth: 1,
	XlErrorBarIncludeMinusValues: 3,
	XlErrorBarIncludeNone: -4142,
	XlErrorBarIncludePlusValues: 2,
}

// enum XlDisplayBlanksAs
var XlDisplayBlanksAs = struct {
	XlInterpolated int32
	XlNotPlotted int32
	XlZero int32
}{
	XlInterpolated: 3,
	XlNotPlotted: 1,
	XlZero: 2,
}

// enum XlArrowHeadStyle
var XlArrowHeadStyle = struct {
	XlArrowHeadStyleClosed int32
	XlArrowHeadStyleDoubleClosed int32
	XlArrowHeadStyleDoubleOpen int32
	XlArrowHeadStyleNone int32
	XlArrowHeadStyleOpen int32
}{
	XlArrowHeadStyleClosed: 3,
	XlArrowHeadStyleDoubleClosed: 5,
	XlArrowHeadStyleDoubleOpen: 4,
	XlArrowHeadStyleNone: -4142,
	XlArrowHeadStyleOpen: 2,
}

// enum XlArrowHeadWidth
var XlArrowHeadWidth = struct {
	XlArrowHeadWidthMedium int32
	XlArrowHeadWidthNarrow int32
	XlArrowHeadWidthWide int32
}{
	XlArrowHeadWidthMedium: -4138,
	XlArrowHeadWidthNarrow: 1,
	XlArrowHeadWidthWide: 3,
}

// enum XlHAlign
var XlHAlign = struct {
	XlHAlignCenter int32
	XlHAlignCenterAcrossSelection int32
	XlHAlignDistributed int32
	XlHAlignFill int32
	XlHAlignGeneral int32
	XlHAlignJustify int32
	XlHAlignLeft int32
	XlHAlignRight int32
}{
	XlHAlignCenter: -4108,
	XlHAlignCenterAcrossSelection: 7,
	XlHAlignDistributed: -4117,
	XlHAlignFill: 5,
	XlHAlignGeneral: 1,
	XlHAlignJustify: -4130,
	XlHAlignLeft: -4131,
	XlHAlignRight: -4152,
}

// enum XlTickLabelPosition
var XlTickLabelPosition = struct {
	XlTickLabelPositionHigh int32
	XlTickLabelPositionLow int32
	XlTickLabelPositionNextToAxis int32
	XlTickLabelPositionNone int32
}{
	XlTickLabelPositionHigh: -4127,
	XlTickLabelPositionLow: -4134,
	XlTickLabelPositionNextToAxis: 4,
	XlTickLabelPositionNone: -4142,
}

// enum XlLegendPosition
var XlLegendPosition = struct {
	XlLegendPositionBottom int32
	XlLegendPositionCorner int32
	XlLegendPositionLeft int32
	XlLegendPositionRight int32
	XlLegendPositionTop int32
	XlLegendPositionCustom int32
}{
	XlLegendPositionBottom: -4107,
	XlLegendPositionCorner: 2,
	XlLegendPositionLeft: -4131,
	XlLegendPositionRight: -4152,
	XlLegendPositionTop: -4160,
	XlLegendPositionCustom: -4161,
}

// enum XlChartPictureType
var XlChartPictureType = struct {
	XlStackScale int32
	XlStack int32
	XlStretch int32
}{
	XlStackScale: 3,
	XlStack: 2,
	XlStretch: 1,
}

// enum XlChartPicturePlacement
var XlChartPicturePlacement = struct {
	XlSides int32
	XlEnd int32
	XlEndSides int32
	XlFront int32
	XlFrontSides int32
	XlFrontEnd int32
	XlAllFaces int32
}{
	XlSides: 1,
	XlEnd: 2,
	XlEndSides: 3,
	XlFront: 4,
	XlFrontSides: 5,
	XlFrontEnd: 6,
	XlAllFaces: 7,
}

// enum XlOrientation
var XlOrientation = struct {
	XlDownward int32
	XlHorizontal int32
	XlUpward int32
	XlVertical int32
}{
	XlDownward: -4170,
	XlHorizontal: -4128,
	XlUpward: -4171,
	XlVertical: -4166,
}

// enum XlTickLabelOrientation
var XlTickLabelOrientation = struct {
	XlTickLabelOrientationAutomatic int32
	XlTickLabelOrientationDownward int32
	XlTickLabelOrientationHorizontal int32
	XlTickLabelOrientationUpward int32
	XlTickLabelOrientationVertical int32
}{
	XlTickLabelOrientationAutomatic: -4105,
	XlTickLabelOrientationDownward: -4170,
	XlTickLabelOrientationHorizontal: -4128,
	XlTickLabelOrientationUpward: -4171,
	XlTickLabelOrientationVertical: -4166,
}

// enum XlBorderWeight
var XlBorderWeight = struct {
	XlHairline int32
	XlMedium int32
	XlThick int32
	XlThin int32
}{
	XlHairline: 1,
	XlMedium: -4138,
	XlThick: 4,
	XlThin: 2,
}

// enum XlDataSeriesDate
var XlDataSeriesDate = struct {
	XlDay int32
	XlMonth int32
	XlWeekday int32
	XlYear int32
}{
	XlDay: 1,
	XlMonth: 3,
	XlWeekday: 2,
	XlYear: 4,
}

// enum XlUnderlineStyle
var XlUnderlineStyle = struct {
	XlUnderlineStyleDouble int32
	XlUnderlineStyleDoubleAccounting int32
	XlUnderlineStyleNone int32
	XlUnderlineStyleSingle int32
	XlUnderlineStyleSingleAccounting int32
}{
	XlUnderlineStyleDouble: -4119,
	XlUnderlineStyleDoubleAccounting: 5,
	XlUnderlineStyleNone: -4142,
	XlUnderlineStyleSingle: 2,
	XlUnderlineStyleSingleAccounting: 4,
}

// enum XlErrorBarType
var XlErrorBarType = struct {
	XlErrorBarTypeCustom int32
	XlErrorBarTypeFixedValue int32
	XlErrorBarTypePercent int32
	XlErrorBarTypeStDev int32
	XlErrorBarTypeStError int32
}{
	XlErrorBarTypeCustom: -4114,
	XlErrorBarTypeFixedValue: 1,
	XlErrorBarTypePercent: 2,
	XlErrorBarTypeStDev: -4155,
	XlErrorBarTypeStError: 4,
}

// enum XlTrendlineType
var XlTrendlineType = struct {
	XlExponential int32
	XlLinear int32
	XlLogarithmic int32
	XlMovingAvg int32
	XlPolynomial int32
	XlPower int32
}{
	XlExponential: 5,
	XlLinear: -4132,
	XlLogarithmic: -4133,
	XlMovingAvg: 6,
	XlPolynomial: 3,
	XlPower: 4,
}

// enum XlLineStyle
var XlLineStyle = struct {
	XlContinuous int32
	XlDash int32
	XlDashDot int32
	XlDashDotDot int32
	XlDot int32
	XlDouble int32
	XlSlantDashDot int32
	XlLineStyleNone int32
}{
	XlContinuous: 1,
	XlDash: -4115,
	XlDashDot: 4,
	XlDashDotDot: 5,
	XlDot: -4118,
	XlDouble: -4119,
	XlSlantDashDot: 13,
	XlLineStyleNone: -4142,
}

// enum XlDataLabelsType
var XlDataLabelsType = struct {
	XlDataLabelsShowNone int32
	XlDataLabelsShowValue int32
	XlDataLabelsShowPercent int32
	XlDataLabelsShowLabel int32
	XlDataLabelsShowLabelAndPercent int32
	XlDataLabelsShowBubbleSizes int32
}{
	XlDataLabelsShowNone: -4142,
	XlDataLabelsShowValue: 2,
	XlDataLabelsShowPercent: 3,
	XlDataLabelsShowLabel: 4,
	XlDataLabelsShowLabelAndPercent: 5,
	XlDataLabelsShowBubbleSizes: 6,
}

// enum XlMarkerStyle
var XlMarkerStyle = struct {
	XlMarkerStyleAutomatic int32
	XlMarkerStyleCircle int32
	XlMarkerStyleDash int32
	XlMarkerStyleDiamond int32
	XlMarkerStyleDot int32
	XlMarkerStyleNone int32
	XlMarkerStylePicture int32
	XlMarkerStylePlus int32
	XlMarkerStyleSquare int32
	XlMarkerStyleStar int32
	XlMarkerStyleTriangle int32
	XlMarkerStyleX int32
}{
	XlMarkerStyleAutomatic: -4105,
	XlMarkerStyleCircle: 8,
	XlMarkerStyleDash: -4115,
	XlMarkerStyleDiamond: 2,
	XlMarkerStyleDot: -4118,
	XlMarkerStyleNone: -4142,
	XlMarkerStylePicture: -4147,
	XlMarkerStylePlus: 9,
	XlMarkerStyleSquare: 1,
	XlMarkerStyleStar: 5,
	XlMarkerStyleTriangle: 3,
	XlMarkerStyleX: -4168,
}

// enum XlPictureConvertorType
var XlPictureConvertorType = struct {
	XlBMP int32
	XlCGM int32
	XlDRW int32
	XlDXF int32
	XlEPS int32
	XlHGL int32
	XlPCT int32
	XlPCX int32
	XlPIC int32
	XlPLT int32
	XlTIF int32
	XlWMF int32
	XlWPG int32
}{
	XlBMP: 1,
	XlCGM: 7,
	XlDRW: 4,
	XlDXF: 5,
	XlEPS: 8,
	XlHGL: 6,
	XlPCT: 13,
	XlPCX: 10,
	XlPIC: 11,
	XlPLT: 12,
	XlTIF: 9,
	XlWMF: 2,
	XlWPG: 3,
}

// enum XlPattern
var XlPattern = struct {
	XlPatternAutomatic int32
	XlPatternChecker int32
	XlPatternCrissCross int32
	XlPatternDown int32
	XlPatternGray16 int32
	XlPatternGray25 int32
	XlPatternGray50 int32
	XlPatternGray75 int32
	XlPatternGray8 int32
	XlPatternGrid int32
	XlPatternHorizontal int32
	XlPatternLightDown int32
	XlPatternLightHorizontal int32
	XlPatternLightUp int32
	XlPatternLightVertical int32
	XlPatternNone int32
	XlPatternSemiGray75 int32
	XlPatternSolid int32
	XlPatternUp int32
	XlPatternVertical int32
	XlPatternLinearGradient int32
	XlPatternRectangularGradient int32
}{
	XlPatternAutomatic: -4105,
	XlPatternChecker: 9,
	XlPatternCrissCross: 16,
	XlPatternDown: -4121,
	XlPatternGray16: 17,
	XlPatternGray25: -4124,
	XlPatternGray50: -4125,
	XlPatternGray75: -4126,
	XlPatternGray8: 18,
	XlPatternGrid: 15,
	XlPatternHorizontal: -4128,
	XlPatternLightDown: 13,
	XlPatternLightHorizontal: 11,
	XlPatternLightUp: 14,
	XlPatternLightVertical: 12,
	XlPatternNone: -4142,
	XlPatternSemiGray75: 10,
	XlPatternSolid: 1,
	XlPatternUp: -4162,
	XlPatternVertical: -4166,
	XlPatternLinearGradient: 4000,
	XlPatternRectangularGradient: 4001,
}

// enum XlChartSplitType
var XlChartSplitType = struct {
	XlSplitByPosition int32
	XlSplitByPercentValue int32
	XlSplitByCustomSplit int32
	XlSplitByValue int32
}{
	XlSplitByPosition: 1,
	XlSplitByPercentValue: 3,
	XlSplitByCustomSplit: 4,
	XlSplitByValue: 2,
}

// enum XlDisplayUnit
var XlDisplayUnit = struct {
	XlHundreds int32
	XlThousands int32
	XlTenThousands int32
	XlHundredThousands int32
	XlMillions int32
	XlTenMillions int32
	XlHundredMillions int32
	XlThousandMillions int32
	XlMillionMillions int32
}{
	XlHundreds: -2,
	XlThousands: -3,
	XlTenThousands: -4,
	XlHundredThousands: -5,
	XlMillions: -6,
	XlTenMillions: -7,
	XlHundredMillions: -8,
	XlThousandMillions: -9,
	XlMillionMillions: -10,
}

// enum XlDataLabelPosition
var XlDataLabelPosition = struct {
	XlLabelPositionCenter int32
	XlLabelPositionAbove int32
	XlLabelPositionBelow int32
	XlLabelPositionLeft int32
	XlLabelPositionRight int32
	XlLabelPositionOutsideEnd int32
	XlLabelPositionInsideEnd int32
	XlLabelPositionInsideBase int32
	XlLabelPositionBestFit int32
	XlLabelPositionMixed int32
	XlLabelPositionCustom int32
}{
	XlLabelPositionCenter: -4108,
	XlLabelPositionAbove: 0,
	XlLabelPositionBelow: 1,
	XlLabelPositionLeft: -4131,
	XlLabelPositionRight: -4152,
	XlLabelPositionOutsideEnd: 2,
	XlLabelPositionInsideEnd: 3,
	XlLabelPositionInsideBase: 4,
	XlLabelPositionBestFit: 5,
	XlLabelPositionMixed: 6,
	XlLabelPositionCustom: 7,
}

// enum XlTimeUnit
var XlTimeUnit = struct {
	XlDays int32
	XlMonths int32
	XlYears int32
}{
	XlDays: 0,
	XlMonths: 1,
	XlYears: 2,
}

// enum XlCategoryType
var XlCategoryType = struct {
	XlCategoryScale int32
	XlTimeScale int32
	XlAutomaticScale int32
}{
	XlCategoryScale: 2,
	XlTimeScale: 3,
	XlAutomaticScale: -4105,
}

// enum XlBarShape
var XlBarShape = struct {
	XlBox int32
	XlPyramidToPoint int32
	XlPyramidToMax int32
	XlCylinder int32
	XlConeToPoint int32
	XlConeToMax int32
}{
	XlBox: 0,
	XlPyramidToPoint: 1,
	XlPyramidToMax: 2,
	XlCylinder: 3,
	XlConeToPoint: 4,
	XlConeToMax: 5,
}

// enum XlChartType
var XlChartType = struct {
	XlColumnClustered int32
	XlColumnStacked int32
	XlColumnStacked100 int32
	Xl3DColumnClustered int32
	Xl3DColumnStacked int32
	Xl3DColumnStacked100 int32
	XlBarClustered int32
	XlBarStacked int32
	XlBarStacked100 int32
	Xl3DBarClustered int32
	Xl3DBarStacked int32
	Xl3DBarStacked100 int32
	XlLineStacked int32
	XlLineStacked100 int32
	XlLineMarkers int32
	XlLineMarkersStacked int32
	XlLineMarkersStacked100 int32
	XlPieOfPie int32
	XlPieExploded int32
	Xl3DPieExploded int32
	XlBarOfPie int32
	XlXYScatterSmooth int32
	XlXYScatterSmoothNoMarkers int32
	XlXYScatterLines int32
	XlXYScatterLinesNoMarkers int32
	XlAreaStacked int32
	XlAreaStacked100 int32
	Xl3DAreaStacked int32
	Xl3DAreaStacked100 int32
	XlDoughnutExploded int32
	XlRadarMarkers int32
	XlRadarFilled int32
	XlSurface int32
	XlSurfaceWireframe int32
	XlSurfaceTopView int32
	XlSurfaceTopViewWireframe int32
	XlBubble int32
	XlBubble3DEffect int32
	XlStockHLC int32
	XlStockOHLC int32
	XlStockVHLC int32
	XlStockVOHLC int32
	XlCylinderColClustered int32
	XlCylinderColStacked int32
	XlCylinderColStacked100 int32
	XlCylinderBarClustered int32
	XlCylinderBarStacked int32
	XlCylinderBarStacked100 int32
	XlCylinderCol int32
	XlConeColClustered int32
	XlConeColStacked int32
	XlConeColStacked100 int32
	XlConeBarClustered int32
	XlConeBarStacked int32
	XlConeBarStacked100 int32
	XlConeCol int32
	XlPyramidColClustered int32
	XlPyramidColStacked int32
	XlPyramidColStacked100 int32
	XlPyramidBarClustered int32
	XlPyramidBarStacked int32
	XlPyramidBarStacked100 int32
	XlPyramidCol int32
	Xl3DColumn int32
	XlLine int32
	Xl3DLine int32
	Xl3DPie int32
	XlPie int32
	XlXYScatter int32
	Xl3DArea int32
	XlArea int32
	XlDoughnut int32
	XlRadar int32
}{
	XlColumnClustered: 51,
	XlColumnStacked: 52,
	XlColumnStacked100: 53,
	Xl3DColumnClustered: 54,
	Xl3DColumnStacked: 55,
	Xl3DColumnStacked100: 56,
	XlBarClustered: 57,
	XlBarStacked: 58,
	XlBarStacked100: 59,
	Xl3DBarClustered: 60,
	Xl3DBarStacked: 61,
	Xl3DBarStacked100: 62,
	XlLineStacked: 63,
	XlLineStacked100: 64,
	XlLineMarkers: 65,
	XlLineMarkersStacked: 66,
	XlLineMarkersStacked100: 67,
	XlPieOfPie: 68,
	XlPieExploded: 69,
	Xl3DPieExploded: 70,
	XlBarOfPie: 71,
	XlXYScatterSmooth: 72,
	XlXYScatterSmoothNoMarkers: 73,
	XlXYScatterLines: 74,
	XlXYScatterLinesNoMarkers: 75,
	XlAreaStacked: 76,
	XlAreaStacked100: 77,
	Xl3DAreaStacked: 78,
	Xl3DAreaStacked100: 79,
	XlDoughnutExploded: 80,
	XlRadarMarkers: 81,
	XlRadarFilled: 82,
	XlSurface: 83,
	XlSurfaceWireframe: 84,
	XlSurfaceTopView: 85,
	XlSurfaceTopViewWireframe: 86,
	XlBubble: 15,
	XlBubble3DEffect: 87,
	XlStockHLC: 88,
	XlStockOHLC: 89,
	XlStockVHLC: 90,
	XlStockVOHLC: 91,
	XlCylinderColClustered: 92,
	XlCylinderColStacked: 93,
	XlCylinderColStacked100: 94,
	XlCylinderBarClustered: 95,
	XlCylinderBarStacked: 96,
	XlCylinderBarStacked100: 97,
	XlCylinderCol: 98,
	XlConeColClustered: 99,
	XlConeColStacked: 100,
	XlConeColStacked100: 101,
	XlConeBarClustered: 102,
	XlConeBarStacked: 103,
	XlConeBarStacked100: 104,
	XlConeCol: 105,
	XlPyramidColClustered: 106,
	XlPyramidColStacked: 107,
	XlPyramidColStacked100: 108,
	XlPyramidBarClustered: 109,
	XlPyramidBarStacked: 110,
	XlPyramidBarStacked100: 111,
	XlPyramidCol: 112,
	Xl3DColumn: -4100,
	XlLine: 4,
	Xl3DLine: -4101,
	Xl3DPie: -4102,
	XlPie: 5,
	XlXYScatter: -4169,
	Xl3DArea: -4098,
	XlArea: 1,
	XlDoughnut: -4120,
	XlRadar: -4151,
}

// enum XlChartItem
var XlChartItem = struct {
	XlDataLabel int32
	XlChartArea int32
	XlSeries int32
	XlChartTitle int32
	XlWalls int32
	XlCorners int32
	XlDataTable int32
	XlTrendline int32
	XlErrorBars int32
	XlXErrorBars int32
	XlYErrorBars int32
	XlLegendEntry int32
	XlLegendKey int32
	XlShape int32
	XlMajorGridlines int32
	XlMinorGridlines int32
	XlAxisTitle int32
	XlUpBars int32
	XlPlotArea int32
	XlDownBars int32
	XlAxis int32
	XlSeriesLines int32
	XlFloor int32
	XlLegend int32
	XlHiLoLines int32
	XlDropLines int32
	XlRadarAxisLabels int32
	XlNothing int32
	XlLeaderLines int32
	XlDisplayUnitLabel int32
	XlPivotChartFieldButton int32
	XlPivotChartDropZone int32
}{
	XlDataLabel: 0,
	XlChartArea: 2,
	XlSeries: 3,
	XlChartTitle: 4,
	XlWalls: 5,
	XlCorners: 6,
	XlDataTable: 7,
	XlTrendline: 8,
	XlErrorBars: 9,
	XlXErrorBars: 10,
	XlYErrorBars: 11,
	XlLegendEntry: 12,
	XlLegendKey: 13,
	XlShape: 14,
	XlMajorGridlines: 15,
	XlMinorGridlines: 16,
	XlAxisTitle: 17,
	XlUpBars: 18,
	XlPlotArea: 19,
	XlDownBars: 20,
	XlAxis: 21,
	XlSeriesLines: 22,
	XlFloor: 23,
	XlLegend: 24,
	XlHiLoLines: 25,
	XlDropLines: 26,
	XlRadarAxisLabels: 27,
	XlNothing: 28,
	XlLeaderLines: 29,
	XlDisplayUnitLabel: 30,
	XlPivotChartFieldButton: 31,
	XlPivotChartDropZone: 32,
}

// enum XlSizeRepresents
var XlSizeRepresents = struct {
	XlSizeIsWidth int32
	XlSizeIsArea int32
}{
	XlSizeIsWidth: 2,
	XlSizeIsArea: 1,
}

// enum XlInsertShiftDirection
var XlInsertShiftDirection = struct {
	XlShiftDown int32
	XlShiftToRight int32
}{
	XlShiftDown: -4121,
	XlShiftToRight: -4161,
}

// enum XlDeleteShiftDirection
var XlDeleteShiftDirection = struct {
	XlShiftToLeft int32
	XlShiftUp int32
}{
	XlShiftToLeft: -4159,
	XlShiftUp: -4162,
}

// enum XlDirection
var XlDirection = struct {
	XlDown int32
	XlToLeft int32
	XlToRight int32
	XlUp int32
}{
	XlDown: -4121,
	XlToLeft: -4159,
	XlToRight: -4161,
	XlUp: -4162,
}

// enum XlConsolidationFunction
var XlConsolidationFunction = struct {
	XlAverage int32
	XlCount int32
	XlCountNums int32
	XlMax int32
	XlMin int32
	XlProduct int32
	XlStDev int32
	XlStDevP int32
	XlSum int32
	XlVar int32
	XlVarP int32
	XlUnknown int32
}{
	XlAverage: -4106,
	XlCount: -4112,
	XlCountNums: -4113,
	XlMax: -4136,
	XlMin: -4139,
	XlProduct: -4149,
	XlStDev: -4155,
	XlStDevP: -4156,
	XlSum: -4157,
	XlVar: -4164,
	XlVarP: -4165,
	XlUnknown: 1000,
}

// enum XlSheetType
var XlSheetType = struct {
	XlChart int32
	XlDialogSheet int32
	XlExcel4IntlMacroSheet int32
	XlExcel4MacroSheet int32
	XlWorksheet int32
}{
	XlChart: -4109,
	XlDialogSheet: -4116,
	XlExcel4IntlMacroSheet: 4,
	XlExcel4MacroSheet: 3,
	XlWorksheet: -4167,
}

// enum XlLocationInTable
var XlLocationInTable = struct {
	XlColumnHeader int32
	XlColumnItem int32
	XlDataHeader int32
	XlDataItem int32
	XlPageHeader int32
	XlPageItem int32
	XlRowHeader int32
	XlRowItem int32
	XlTableBody int32
}{
	XlColumnHeader: -4110,
	XlColumnItem: 5,
	XlDataHeader: 3,
	XlDataItem: 7,
	XlPageHeader: 2,
	XlPageItem: 6,
	XlRowHeader: -4153,
	XlRowItem: 4,
	XlTableBody: 8,
}

// enum XlFindLookIn
var XlFindLookIn = struct {
	XlFormulas int32
	XlComments int32
	XlValues int32
}{
	XlFormulas: -4123,
	XlComments: -4144,
	XlValues: -4163,
}

// enum XlWindowType
var XlWindowType = struct {
	XlChartAsWindow int32
	XlChartInPlace int32
	XlClipboard int32
	XlInfo int32
	XlWorkbook int32
}{
	XlChartAsWindow: 5,
	XlChartInPlace: 4,
	XlClipboard: 3,
	XlInfo: -4129,
	XlWorkbook: 1,
}

// enum XlPivotFieldDataType
var XlPivotFieldDataType = struct {
	XlDate int32
	XlNumber int32
	XlText int32
}{
	XlDate: 2,
	XlNumber: -4145,
	XlText: -4158,
}

// enum XlCopyPictureFormat
var XlCopyPictureFormat = struct {
	XlBitmap int32
	XlPicture int32
}{
	XlBitmap: 2,
	XlPicture: -4147,
}

// enum XlPivotTableSourceType
var XlPivotTableSourceType = struct {
	XlScenario int32
	XlConsolidation int32
	XlDatabase int32
	XlExternal int32
	XlPivotTable int32
}{
	XlScenario: 4,
	XlConsolidation: 3,
	XlDatabase: 1,
	XlExternal: 2,
	XlPivotTable: -4148,
}

// enum XlReferenceStyle
var XlReferenceStyle = struct {
	XlA1 int32
	XlR1C1 int32
}{
	XlA1: 1,
	XlR1C1: -4150,
}

// enum XlMSApplication
var XlMSApplication = struct {
	XlMicrosoftAccess int32
	XlMicrosoftFoxPro int32
	XlMicrosoftMail int32
	XlMicrosoftPowerPoint int32
	XlMicrosoftProject int32
	XlMicrosoftSchedulePlus int32
	XlMicrosoftWord int32
}{
	XlMicrosoftAccess: 4,
	XlMicrosoftFoxPro: 5,
	XlMicrosoftMail: 3,
	XlMicrosoftPowerPoint: 2,
	XlMicrosoftProject: 6,
	XlMicrosoftSchedulePlus: 7,
	XlMicrosoftWord: 1,
}

// enum XlMouseButton
var XlMouseButton = struct {
	XlNoButton int32
	XlPrimaryButton int32
	XlSecondaryButton int32
}{
	XlNoButton: 0,
	XlPrimaryButton: 1,
	XlSecondaryButton: 2,
}

// enum XlCutCopyMode
var XlCutCopyMode = struct {
	XlCopy int32
	XlCut int32
}{
	XlCopy: 1,
	XlCut: 2,
}

// enum XlFillWith
var XlFillWith = struct {
	XlFillWithAll int32
	XlFillWithContents int32
	XlFillWithFormats int32
}{
	XlFillWithAll: -4104,
	XlFillWithContents: 2,
	XlFillWithFormats: -4122,
}

// enum XlFilterAction
var XlFilterAction = struct {
	XlFilterCopy int32
	XlFilterInPlace int32
}{
	XlFilterCopy: 2,
	XlFilterInPlace: 1,
}

// enum XlOrder
var XlOrder = struct {
	XlDownThenOver int32
	XlOverThenDown int32
}{
	XlDownThenOver: 1,
	XlOverThenDown: 2,
}

// enum XlLinkType
var XlLinkType = struct {
	XlLinkTypeExcelLinks int32
	XlLinkTypeOLELinks int32
}{
	XlLinkTypeExcelLinks: 1,
	XlLinkTypeOLELinks: 2,
}

// enum XlApplyNamesOrder
var XlApplyNamesOrder = struct {
	XlColumnThenRow int32
	XlRowThenColumn int32
}{
	XlColumnThenRow: 2,
	XlRowThenColumn: 1,
}

// enum XlEnableCancelKey
var XlEnableCancelKey = struct {
	XlDisabled int32
	XlErrorHandler int32
	XlInterrupt int32
}{
	XlDisabled: 0,
	XlErrorHandler: 2,
	XlInterrupt: 1,
}

// enum XlPageBreak
var XlPageBreak = struct {
	XlPageBreakAutomatic int32
	XlPageBreakManual int32
	XlPageBreakNone int32
}{
	XlPageBreakAutomatic: -4105,
	XlPageBreakManual: -4135,
	XlPageBreakNone: -4142,
}

// enum XlOLEType
var XlOLEType = struct {
	XlOLEControl int32
	XlOLEEmbed int32
	XlOLELink int32
}{
	XlOLEControl: 2,
	XlOLEEmbed: 1,
	XlOLELink: 0,
}

// enum XlPageOrientation
var XlPageOrientation = struct {
	XlLandscape int32
	XlPortrait int32
}{
	XlLandscape: 2,
	XlPortrait: 1,
}

// enum XlLinkInfo
var XlLinkInfo = struct {
	XlEditionDate int32
	XlUpdateState int32
	XlLinkInfoStatus int32
}{
	XlEditionDate: 2,
	XlUpdateState: 1,
	XlLinkInfoStatus: 3,
}

// enum XlCommandUnderlines
var XlCommandUnderlines = struct {
	XlCommandUnderlinesAutomatic int32
	XlCommandUnderlinesOff int32
	XlCommandUnderlinesOn int32
}{
	XlCommandUnderlinesAutomatic: -4105,
	XlCommandUnderlinesOff: -4146,
	XlCommandUnderlinesOn: 1,
}

// enum XlOLEVerb
var XlOLEVerb = struct {
	XlVerbOpen int32
	XlVerbPrimary int32
}{
	XlVerbOpen: 2,
	XlVerbPrimary: 1,
}

// enum XlCalculation
var XlCalculation = struct {
	XlCalculationAutomatic int32
	XlCalculationManual int32
	XlCalculationSemiautomatic int32
}{
	XlCalculationAutomatic: -4105,
	XlCalculationManual: -4135,
	XlCalculationSemiautomatic: 2,
}

// enum XlFileAccess
var XlFileAccess = struct {
	XlReadOnly int32
	XlReadWrite int32
}{
	XlReadOnly: 3,
	XlReadWrite: 2,
}

// enum XlEditionType
var XlEditionType = struct {
	XlPublisher int32
	XlSubscriber int32
}{
	XlPublisher: 1,
	XlSubscriber: 2,
}

// enum XlObjectSize
var XlObjectSize = struct {
	XlFitToPage int32
	XlFullPage int32
	XlScreenSize int32
}{
	XlFitToPage: 2,
	XlFullPage: 3,
	XlScreenSize: 1,
}

// enum XlLookAt
var XlLookAt = struct {
	XlPart int32
	XlWhole int32
}{
	XlPart: 2,
	XlWhole: 1,
}

// enum XlMailSystem
var XlMailSystem = struct {
	XlMAPI int32
	XlNoMailSystem int32
	XlPowerTalk int32
}{
	XlMAPI: 1,
	XlNoMailSystem: 0,
	XlPowerTalk: 2,
}

// enum XlLinkInfoType
var XlLinkInfoType = struct {
	XlLinkInfoOLELinks int32
	XlLinkInfoPublishers int32
	XlLinkInfoSubscribers int32
}{
	XlLinkInfoOLELinks: 2,
	XlLinkInfoPublishers: 5,
	XlLinkInfoSubscribers: 6,
}

// enum XlCVError
var XlCVError = struct {
	XlErrDiv0 int32
	XlErrNA int32
	XlErrName int32
	XlErrNull int32
	XlErrNum int32
	XlErrRef int32
	XlErrValue int32
}{
	XlErrDiv0: 2007,
	XlErrNA: 2042,
	XlErrName: 2029,
	XlErrNull: 2000,
	XlErrNum: 2036,
	XlErrRef: 2023,
	XlErrValue: 2015,
}

// enum XlEditionFormat
var XlEditionFormat = struct {
	XlBIFF int32
	XlPICT int32
	XlRTF int32
	XlVALU int32
}{
	XlBIFF: 2,
	XlPICT: 1,
	XlRTF: 4,
	XlVALU: 8,
}

// enum XlLink
var XlLink = struct {
	XlExcelLinks int32
	XlOLELinks int32
	XlPublishers int32
	XlSubscribers int32
}{
	XlExcelLinks: 1,
	XlOLELinks: 2,
	XlPublishers: 5,
	XlSubscribers: 6,
}

// enum XlCellType
var XlCellType = struct {
	XlCellTypeBlanks int32
	XlCellTypeConstants int32
	XlCellTypeFormulas int32
	XlCellTypeLastCell int32
	XlCellTypeComments int32
	XlCellTypeVisible int32
	XlCellTypeAllFormatConditions int32
	XlCellTypeSameFormatConditions int32
	XlCellTypeAllValidation int32
	XlCellTypeSameValidation int32
}{
	XlCellTypeBlanks: 4,
	XlCellTypeConstants: 2,
	XlCellTypeFormulas: -4123,
	XlCellTypeLastCell: 11,
	XlCellTypeComments: -4144,
	XlCellTypeVisible: 12,
	XlCellTypeAllFormatConditions: -4172,
	XlCellTypeSameFormatConditions: -4173,
	XlCellTypeAllValidation: -4174,
	XlCellTypeSameValidation: -4175,
}

// enum XlArrangeStyle
var XlArrangeStyle = struct {
	XlArrangeStyleCascade int32
	XlArrangeStyleHorizontal int32
	XlArrangeStyleTiled int32
	XlArrangeStyleVertical int32
}{
	XlArrangeStyleCascade: 7,
	XlArrangeStyleHorizontal: -4128,
	XlArrangeStyleTiled: 1,
	XlArrangeStyleVertical: -4166,
}

// enum XlMousePointer
var XlMousePointer = struct {
	XlIBeam int32
	XlDefault int32
	XlNorthwestArrow int32
	XlWait int32
}{
	XlIBeam: 3,
	XlDefault: -4143,
	XlNorthwestArrow: 1,
	XlWait: 2,
}

// enum XlEditionOptionsOption
var XlEditionOptionsOption = struct {
	XlAutomaticUpdate int32
	XlCancel int32
	XlChangeAttributes int32
	XlManualUpdate int32
	XlOpenSource int32
	XlSelect int32
	XlSendPublisher int32
	XlUpdateSubscriber int32
}{
	XlAutomaticUpdate: 4,
	XlCancel: 1,
	XlChangeAttributes: 6,
	XlManualUpdate: 5,
	XlOpenSource: 3,
	XlSelect: 3,
	XlSendPublisher: 2,
	XlUpdateSubscriber: 2,
}

// enum XlAutoFillType
var XlAutoFillType = struct {
	XlFillCopy int32
	XlFillDays int32
	XlFillDefault int32
	XlFillFormats int32
	XlFillMonths int32
	XlFillSeries int32
	XlFillValues int32
	XlFillWeekdays int32
	XlFillYears int32
	XlGrowthTrend int32
	XlLinearTrend int32
}{
	XlFillCopy: 1,
	XlFillDays: 5,
	XlFillDefault: 0,
	XlFillFormats: 3,
	XlFillMonths: 7,
	XlFillSeries: 2,
	XlFillValues: 4,
	XlFillWeekdays: 6,
	XlFillYears: 8,
	XlGrowthTrend: 10,
	XlLinearTrend: 9,
}

// enum XlAutoFilterOperator
var XlAutoFilterOperator = struct {
	XlAnd int32
	XlBottom10Items int32
	XlBottom10Percent int32
	XlOr int32
	XlTop10Items int32
	XlTop10Percent int32
	XlFilterValues int32
	XlFilterCellColor int32
	XlFilterFontColor int32
	XlFilterIcon int32
	XlFilterDynamic int32
	XlFilterNoFill int32
	XlFilterAutomaticFontColor int32
	XlFilterNoIcon int32
}{
	XlAnd: 1,
	XlBottom10Items: 4,
	XlBottom10Percent: 6,
	XlOr: 2,
	XlTop10Items: 3,
	XlTop10Percent: 5,
	XlFilterValues: 7,
	XlFilterCellColor: 8,
	XlFilterFontColor: 9,
	XlFilterIcon: 10,
	XlFilterDynamic: 11,
	XlFilterNoFill: 12,
	XlFilterAutomaticFontColor: 13,
	XlFilterNoIcon: 14,
}

// enum XlClipboardFormat
var XlClipboardFormat = struct {
	XlClipboardFormatBIFF12 int32
	XlClipboardFormatBIFF int32
	XlClipboardFormatBIFF2 int32
	XlClipboardFormatBIFF3 int32
	XlClipboardFormatBIFF4 int32
	XlClipboardFormatBinary int32
	XlClipboardFormatBitmap int32
	XlClipboardFormatCGM int32
	XlClipboardFormatCSV int32
	XlClipboardFormatDIF int32
	XlClipboardFormatDspText int32
	XlClipboardFormatEmbeddedObject int32
	XlClipboardFormatEmbedSource int32
	XlClipboardFormatLink int32
	XlClipboardFormatLinkSource int32
	XlClipboardFormatLinkSourceDesc int32
	XlClipboardFormatMovie int32
	XlClipboardFormatNative int32
	XlClipboardFormatObjectDesc int32
	XlClipboardFormatObjectLink int32
	XlClipboardFormatOwnerLink int32
	XlClipboardFormatPICT int32
	XlClipboardFormatPrintPICT int32
	XlClipboardFormatRTF int32
	XlClipboardFormatScreenPICT int32
	XlClipboardFormatStandardFont int32
	XlClipboardFormatStandardScale int32
	XlClipboardFormatSYLK int32
	XlClipboardFormatTable int32
	XlClipboardFormatText int32
	XlClipboardFormatToolFace int32
	XlClipboardFormatToolFacePICT int32
	XlClipboardFormatVALU int32
	XlClipboardFormatWK1 int32
}{
	XlClipboardFormatBIFF12: 63,
	XlClipboardFormatBIFF: 8,
	XlClipboardFormatBIFF2: 18,
	XlClipboardFormatBIFF3: 20,
	XlClipboardFormatBIFF4: 30,
	XlClipboardFormatBinary: 15,
	XlClipboardFormatBitmap: 9,
	XlClipboardFormatCGM: 13,
	XlClipboardFormatCSV: 5,
	XlClipboardFormatDIF: 4,
	XlClipboardFormatDspText: 12,
	XlClipboardFormatEmbeddedObject: 21,
	XlClipboardFormatEmbedSource: 22,
	XlClipboardFormatLink: 11,
	XlClipboardFormatLinkSource: 23,
	XlClipboardFormatLinkSourceDesc: 32,
	XlClipboardFormatMovie: 24,
	XlClipboardFormatNative: 14,
	XlClipboardFormatObjectDesc: 31,
	XlClipboardFormatObjectLink: 19,
	XlClipboardFormatOwnerLink: 17,
	XlClipboardFormatPICT: 2,
	XlClipboardFormatPrintPICT: 3,
	XlClipboardFormatRTF: 7,
	XlClipboardFormatScreenPICT: 29,
	XlClipboardFormatStandardFont: 28,
	XlClipboardFormatStandardScale: 27,
	XlClipboardFormatSYLK: 6,
	XlClipboardFormatTable: 16,
	XlClipboardFormatText: 0,
	XlClipboardFormatToolFace: 25,
	XlClipboardFormatToolFacePICT: 26,
	XlClipboardFormatVALU: 1,
	XlClipboardFormatWK1: 10,
}

// enum XlFileFormat
var XlFileFormat = struct {
	XlAddIn int32
	XlCSV int32
	XlCSVMac int32
	XlCSVMSDOS int32
	XlCSVWindows int32
	XlDBF2 int32
	XlDBF3 int32
	XlDBF4 int32
	XlDIF int32
	XlExcel2 int32
	XlExcel2FarEast int32
	XlExcel3 int32
	XlExcel4 int32
	XlExcel5 int32
	XlExcel7 int32
	XlExcel9795 int32
	XlExcel4Workbook int32
	XlIntlAddIn int32
	XlIntlMacro int32
	XlWorkbookNormal int32
	XlSYLK int32
	XlTemplate int32
	XlCurrentPlatformText int32
	XlTextMac int32
	XlTextMSDOS int32
	XlTextPrinter int32
	XlTextWindows int32
	XlWJ2WD1 int32
	XlWK1 int32
	XlWK1ALL int32
	XlWK1FMT int32
	XlWK3 int32
	XlWK4 int32
	XlWK3FM3 int32
	XlWKS int32
	XlWorks2FarEast int32
	XlWQ1 int32
	XlWJ3 int32
	XlWJ3FJ3 int32
	XlUnicodeText int32
	XlHtml int32
	XlWebArchive int32
	XlXMLSpreadsheet int32
	XlExcel12 int32
	XlOpenXMLWorkbook int32
	XlOpenXMLWorkbookMacroEnabled int32
	XlOpenXMLTemplateMacroEnabled int32
	XlTemplate8 int32
	XlOpenXMLTemplate int32
	XlAddIn8 int32
	XlOpenXMLAddIn int32
	XlExcel8 int32
	XlOpenDocumentSpreadsheet int32
	XlWorkbookDefault int32
}{
	XlAddIn: 18,
	XlCSV: 6,
	XlCSVMac: 22,
	XlCSVMSDOS: 24,
	XlCSVWindows: 23,
	XlDBF2: 7,
	XlDBF3: 8,
	XlDBF4: 11,
	XlDIF: 9,
	XlExcel2: 16,
	XlExcel2FarEast: 27,
	XlExcel3: 29,
	XlExcel4: 33,
	XlExcel5: 39,
	XlExcel7: 39,
	XlExcel9795: 43,
	XlExcel4Workbook: 35,
	XlIntlAddIn: 26,
	XlIntlMacro: 25,
	XlWorkbookNormal: -4143,
	XlSYLK: 2,
	XlTemplate: 17,
	XlCurrentPlatformText: -4158,
	XlTextMac: 19,
	XlTextMSDOS: 21,
	XlTextPrinter: 36,
	XlTextWindows: 20,
	XlWJ2WD1: 14,
	XlWK1: 5,
	XlWK1ALL: 31,
	XlWK1FMT: 30,
	XlWK3: 15,
	XlWK4: 38,
	XlWK3FM3: 32,
	XlWKS: 4,
	XlWorks2FarEast: 28,
	XlWQ1: 34,
	XlWJ3: 40,
	XlWJ3FJ3: 41,
	XlUnicodeText: 42,
	XlHtml: 44,
	XlWebArchive: 45,
	XlXMLSpreadsheet: 46,
	XlExcel12: 50,
	XlOpenXMLWorkbook: 51,
	XlOpenXMLWorkbookMacroEnabled: 52,
	XlOpenXMLTemplateMacroEnabled: 53,
	XlTemplate8: 17,
	XlOpenXMLTemplate: 54,
	XlAddIn8: 18,
	XlOpenXMLAddIn: 55,
	XlExcel8: 56,
	XlOpenDocumentSpreadsheet: 60,
	XlWorkbookDefault: 51,
}

// enum XlApplicationInternational
var XlApplicationInternational = struct {
	Xl24HourClock int32
	Xl4DigitYears int32
	XlAlternateArraySeparator int32
	XlColumnSeparator int32
	XlCountryCode int32
	XlCountrySetting int32
	XlCurrencyBefore int32
	XlCurrencyCode int32
	XlCurrencyDigits int32
	XlCurrencyLeadingZeros int32
	XlCurrencyMinusSign int32
	XlCurrencyNegative int32
	XlCurrencySpaceBefore int32
	XlCurrencyTrailingZeros int32
	XlDateOrder int32
	XlDateSeparator int32
	XlDayCode int32
	XlDayLeadingZero int32
	XlDecimalSeparator int32
	XlGeneralFormatName int32
	XlHourCode int32
	XlLeftBrace int32
	XlLeftBracket int32
	XlListSeparator int32
	XlLowerCaseColumnLetter int32
	XlLowerCaseRowLetter int32
	XlMDY int32
	XlMetric int32
	XlMinuteCode int32
	XlMonthCode int32
	XlMonthLeadingZero int32
	XlMonthNameChars int32
	XlNoncurrencyDigits int32
	XlNonEnglishFunctions int32
	XlRightBrace int32
	XlRightBracket int32
	XlRowSeparator int32
	XlSecondCode int32
	XlThousandsSeparator int32
	XlTimeLeadingZero int32
	XlTimeSeparator int32
	XlUpperCaseColumnLetter int32
	XlUpperCaseRowLetter int32
	XlWeekdayNameChars int32
	XlYearCode int32
}{
	Xl24HourClock: 33,
	Xl4DigitYears: 43,
	XlAlternateArraySeparator: 16,
	XlColumnSeparator: 14,
	XlCountryCode: 1,
	XlCountrySetting: 2,
	XlCurrencyBefore: 37,
	XlCurrencyCode: 25,
	XlCurrencyDigits: 27,
	XlCurrencyLeadingZeros: 40,
	XlCurrencyMinusSign: 38,
	XlCurrencyNegative: 28,
	XlCurrencySpaceBefore: 36,
	XlCurrencyTrailingZeros: 39,
	XlDateOrder: 32,
	XlDateSeparator: 17,
	XlDayCode: 21,
	XlDayLeadingZero: 42,
	XlDecimalSeparator: 3,
	XlGeneralFormatName: 26,
	XlHourCode: 22,
	XlLeftBrace: 12,
	XlLeftBracket: 10,
	XlListSeparator: 5,
	XlLowerCaseColumnLetter: 9,
	XlLowerCaseRowLetter: 8,
	XlMDY: 44,
	XlMetric: 35,
	XlMinuteCode: 23,
	XlMonthCode: 20,
	XlMonthLeadingZero: 41,
	XlMonthNameChars: 30,
	XlNoncurrencyDigits: 29,
	XlNonEnglishFunctions: 34,
	XlRightBrace: 13,
	XlRightBracket: 11,
	XlRowSeparator: 15,
	XlSecondCode: 24,
	XlThousandsSeparator: 4,
	XlTimeLeadingZero: 45,
	XlTimeSeparator: 18,
	XlUpperCaseColumnLetter: 7,
	XlUpperCaseRowLetter: 6,
	XlWeekdayNameChars: 31,
	XlYearCode: 19,
}

// enum XlPageBreakExtent
var XlPageBreakExtent = struct {
	XlPageBreakFull int32
	XlPageBreakPartial int32
}{
	XlPageBreakFull: 1,
	XlPageBreakPartial: 2,
}

// enum XlCellInsertionMode
var XlCellInsertionMode = struct {
	XlOverwriteCells int32
	XlInsertDeleteCells int32
	XlInsertEntireRows int32
}{
	XlOverwriteCells: 0,
	XlInsertDeleteCells: 1,
	XlInsertEntireRows: 2,
}

// enum XlFormulaLabel
var XlFormulaLabel = struct {
	XlNoLabels int32
	XlRowLabels int32
	XlColumnLabels int32
	XlMixedLabels int32
}{
	XlNoLabels: -4142,
	XlRowLabels: 1,
	XlColumnLabels: 2,
	XlMixedLabels: 3,
}

// enum XlHighlightChangesTime
var XlHighlightChangesTime = struct {
	XlSinceMyLastSave int32
	XlAllChanges int32
	XlNotYetReviewed int32
}{
	XlSinceMyLastSave: 1,
	XlAllChanges: 2,
	XlNotYetReviewed: 3,
}

// enum XlCommentDisplayMode
var XlCommentDisplayMode = struct {
	XlNoIndicator int32
	XlCommentIndicatorOnly int32
	XlCommentAndIndicator int32
}{
	XlNoIndicator: 0,
	XlCommentIndicatorOnly: -1,
	XlCommentAndIndicator: 1,
}

// enum XlFormatConditionType
var XlFormatConditionType = struct {
	XlCellValue int32
	XlExpression int32
	XlColorScale int32
	XlDatabar int32
	XlTop10 int32
	XlIconSets int32
	XlUniqueValues int32
	XlTextString int32
	XlBlanksCondition int32
	XlTimePeriod int32
	XlAboveAverageCondition int32
	XlNoBlanksCondition int32
	XlErrorsCondition int32
	XlNoErrorsCondition int32
}{
	XlCellValue: 1,
	XlExpression: 2,
	XlColorScale: 3,
	XlDatabar: 4,
	XlTop10: 5,
	XlIconSets: 6,
	XlUniqueValues: 8,
	XlTextString: 9,
	XlBlanksCondition: 10,
	XlTimePeriod: 11,
	XlAboveAverageCondition: 12,
	XlNoBlanksCondition: 13,
	XlErrorsCondition: 16,
	XlNoErrorsCondition: 17,
}

// enum XlFormatConditionOperator
var XlFormatConditionOperator = struct {
	XlBetween int32
	XlNotBetween int32
	XlEqual int32
	XlNotEqual int32
	XlGreater int32
	XlLess int32
	XlGreaterEqual int32
	XlLessEqual int32
}{
	XlBetween: 1,
	XlNotBetween: 2,
	XlEqual: 3,
	XlNotEqual: 4,
	XlGreater: 5,
	XlLess: 6,
	XlGreaterEqual: 7,
	XlLessEqual: 8,
}

// enum XlEnableSelection
var XlEnableSelection = struct {
	XlNoRestrictions int32
	XlUnlockedCells int32
	XlNoSelection int32
}{
	XlNoRestrictions: 0,
	XlUnlockedCells: 1,
	XlNoSelection: -4142,
}

// enum XlDVType
var XlDVType = struct {
	XlValidateInputOnly int32
	XlValidateWholeNumber int32
	XlValidateDecimal int32
	XlValidateList int32
	XlValidateDate int32
	XlValidateTime int32
	XlValidateTextLength int32
	XlValidateCustom int32
}{
	XlValidateInputOnly: 0,
	XlValidateWholeNumber: 1,
	XlValidateDecimal: 2,
	XlValidateList: 3,
	XlValidateDate: 4,
	XlValidateTime: 5,
	XlValidateTextLength: 6,
	XlValidateCustom: 7,
}

// enum XlIMEMode
var XlIMEMode = struct {
	XlIMEModeNoControl int32
	XlIMEModeOn int32
	XlIMEModeOff int32
	XlIMEModeDisable int32
	XlIMEModeHiragana int32
	XlIMEModeKatakana int32
	XlIMEModeKatakanaHalf int32
	XlIMEModeAlphaFull int32
	XlIMEModeAlpha int32
	XlIMEModeHangulFull int32
	XlIMEModeHangul int32
}{
	XlIMEModeNoControl: 0,
	XlIMEModeOn: 1,
	XlIMEModeOff: 2,
	XlIMEModeDisable: 3,
	XlIMEModeHiragana: 4,
	XlIMEModeKatakana: 5,
	XlIMEModeKatakanaHalf: 6,
	XlIMEModeAlphaFull: 7,
	XlIMEModeAlpha: 8,
	XlIMEModeHangulFull: 9,
	XlIMEModeHangul: 10,
}

// enum XlDVAlertStyle
var XlDVAlertStyle = struct {
	XlValidAlertStop int32
	XlValidAlertWarning int32
	XlValidAlertInformation int32
}{
	XlValidAlertStop: 1,
	XlValidAlertWarning: 2,
	XlValidAlertInformation: 3,
}

// enum XlChartLocation
var XlChartLocation = struct {
	XlLocationAsNewSheet int32
	XlLocationAsObject int32
	XlLocationAutomatic int32
}{
	XlLocationAsNewSheet: 1,
	XlLocationAsObject: 2,
	XlLocationAutomatic: 3,
}

// enum XlPaperSize
var XlPaperSize = struct {
	XlPaper10x14 int32
	XlPaper11x17 int32
	XlPaperA3 int32
	XlPaperA4 int32
	XlPaperA4Small int32
	XlPaperA5 int32
	XlPaperB4 int32
	XlPaperB5 int32
	XlPaperCsheet int32
	XlPaperDsheet int32
	XlPaperEnvelope10 int32
	XlPaperEnvelope11 int32
	XlPaperEnvelope12 int32
	XlPaperEnvelope14 int32
	XlPaperEnvelope9 int32
	XlPaperEnvelopeB4 int32
	XlPaperEnvelopeB5 int32
	XlPaperEnvelopeB6 int32
	XlPaperEnvelopeC3 int32
	XlPaperEnvelopeC4 int32
	XlPaperEnvelopeC5 int32
	XlPaperEnvelopeC6 int32
	XlPaperEnvelopeC65 int32
	XlPaperEnvelopeDL int32
	XlPaperEnvelopeItaly int32
	XlPaperEnvelopeMonarch int32
	XlPaperEnvelopePersonal int32
	XlPaperEsheet int32
	XlPaperExecutive int32
	XlPaperFanfoldLegalGerman int32
	XlPaperFanfoldStdGerman int32
	XlPaperFanfoldUS int32
	XlPaperFolio int32
	XlPaperLedger int32
	XlPaperLegal int32
	XlPaperLetter int32
	XlPaperLetterSmall int32
	XlPaperNote int32
	XlPaperQuarto int32
	XlPaperStatement int32
	XlPaperTabloid int32
	XlPaperUser int32
}{
	XlPaper10x14: 16,
	XlPaper11x17: 17,
	XlPaperA3: 8,
	XlPaperA4: 9,
	XlPaperA4Small: 10,
	XlPaperA5: 11,
	XlPaperB4: 12,
	XlPaperB5: 13,
	XlPaperCsheet: 24,
	XlPaperDsheet: 25,
	XlPaperEnvelope10: 20,
	XlPaperEnvelope11: 21,
	XlPaperEnvelope12: 22,
	XlPaperEnvelope14: 23,
	XlPaperEnvelope9: 19,
	XlPaperEnvelopeB4: 33,
	XlPaperEnvelopeB5: 34,
	XlPaperEnvelopeB6: 35,
	XlPaperEnvelopeC3: 29,
	XlPaperEnvelopeC4: 30,
	XlPaperEnvelopeC5: 28,
	XlPaperEnvelopeC6: 31,
	XlPaperEnvelopeC65: 32,
	XlPaperEnvelopeDL: 27,
	XlPaperEnvelopeItaly: 36,
	XlPaperEnvelopeMonarch: 37,
	XlPaperEnvelopePersonal: 38,
	XlPaperEsheet: 26,
	XlPaperExecutive: 7,
	XlPaperFanfoldLegalGerman: 41,
	XlPaperFanfoldStdGerman: 40,
	XlPaperFanfoldUS: 39,
	XlPaperFolio: 14,
	XlPaperLedger: 4,
	XlPaperLegal: 5,
	XlPaperLetter: 1,
	XlPaperLetterSmall: 2,
	XlPaperNote: 18,
	XlPaperQuarto: 15,
	XlPaperStatement: 6,
	XlPaperTabloid: 3,
	XlPaperUser: 256,
}

// enum XlPasteSpecialOperation
var XlPasteSpecialOperation = struct {
	XlPasteSpecialOperationAdd int32
	XlPasteSpecialOperationDivide int32
	XlPasteSpecialOperationMultiply int32
	XlPasteSpecialOperationNone int32
	XlPasteSpecialOperationSubtract int32
}{
	XlPasteSpecialOperationAdd: 2,
	XlPasteSpecialOperationDivide: 5,
	XlPasteSpecialOperationMultiply: 4,
	XlPasteSpecialOperationNone: -4142,
	XlPasteSpecialOperationSubtract: 3,
}

// enum XlPasteType
var XlPasteType = struct {
	XlPasteAll int32
	XlPasteAllUsingSourceTheme int32
	XlPasteAllMergingConditionalFormats int32
	XlPasteAllExceptBorders int32
	XlPasteFormats int32
	XlPasteFormulas int32
	XlPasteComments int32
	XlPasteValues int32
	XlPasteColumnWidths int32
	XlPasteValidation int32
	XlPasteFormulasAndNumberFormats int32
	XlPasteValuesAndNumberFormats int32
}{
	XlPasteAll: -4104,
	XlPasteAllUsingSourceTheme: 13,
	XlPasteAllMergingConditionalFormats: 14,
	XlPasteAllExceptBorders: 7,
	XlPasteFormats: -4122,
	XlPasteFormulas: -4123,
	XlPasteComments: -4144,
	XlPasteValues: -4163,
	XlPasteColumnWidths: 8,
	XlPasteValidation: 6,
	XlPasteFormulasAndNumberFormats: 11,
	XlPasteValuesAndNumberFormats: 12,
}

// enum XlPhoneticCharacterType
var XlPhoneticCharacterType = struct {
	XlKatakanaHalf int32
	XlKatakana int32
	XlHiragana int32
	XlNoConversion int32
}{
	XlKatakanaHalf: 0,
	XlKatakana: 1,
	XlHiragana: 2,
	XlNoConversion: 3,
}

// enum XlPhoneticAlignment
var XlPhoneticAlignment = struct {
	XlPhoneticAlignNoControl int32
	XlPhoneticAlignLeft int32
	XlPhoneticAlignCenter int32
	XlPhoneticAlignDistributed int32
}{
	XlPhoneticAlignNoControl: 0,
	XlPhoneticAlignLeft: 1,
	XlPhoneticAlignCenter: 2,
	XlPhoneticAlignDistributed: 3,
}

// enum XlPictureAppearance
var XlPictureAppearance = struct {
	XlPrinter int32
	XlScreen int32
}{
	XlPrinter: 2,
	XlScreen: 1,
}

// enum XlPivotFieldOrientation
var XlPivotFieldOrientation = struct {
	XlColumnField int32
	XlDataField int32
	XlHidden int32
	XlPageField int32
	XlRowField int32
}{
	XlColumnField: 2,
	XlDataField: 4,
	XlHidden: 0,
	XlPageField: 3,
	XlRowField: 1,
}

// enum XlPivotFieldCalculation
var XlPivotFieldCalculation = struct {
	XlDifferenceFrom int32
	XlIndex int32
	XlNoAdditionalCalculation int32
	XlPercentDifferenceFrom int32
	XlPercentOf int32
	XlPercentOfColumn int32
	XlPercentOfRow int32
	XlPercentOfTotal int32
	XlRunningTotal int32
	XlPercentOfParentRow int32
	XlPercentOfParentColumn int32
	XlPercentOfParent int32
	XlPercentRunningTotal int32
	XlRankAscending int32
	XlRankDecending int32
}{
	XlDifferenceFrom: 2,
	XlIndex: 9,
	XlNoAdditionalCalculation: -4143,
	XlPercentDifferenceFrom: 4,
	XlPercentOf: 3,
	XlPercentOfColumn: 7,
	XlPercentOfRow: 6,
	XlPercentOfTotal: 8,
	XlRunningTotal: 5,
	XlPercentOfParentRow: 10,
	XlPercentOfParentColumn: 11,
	XlPercentOfParent: 12,
	XlPercentRunningTotal: 13,
	XlRankAscending: 14,
	XlRankDecending: 15,
}

// enum XlPlacement
var XlPlacement = struct {
	XlFreeFloating int32
	XlMove int32
	XlMoveAndSize int32
}{
	XlFreeFloating: 3,
	XlMove: 2,
	XlMoveAndSize: 1,
}

// enum XlPlatform
var XlPlatform = struct {
	XlMacintosh int32
	XlMSDOS int32
	XlWindows int32
}{
	XlMacintosh: 1,
	XlMSDOS: 3,
	XlWindows: 2,
}

// enum XlPrintLocation
var XlPrintLocation = struct {
	XlPrintSheetEnd int32
	XlPrintInPlace int32
	XlPrintNoComments int32
}{
	XlPrintSheetEnd: 1,
	XlPrintInPlace: 16,
	XlPrintNoComments: -4142,
}

// enum XlPriority
var XlPriority = struct {
	XlPriorityHigh int32
	XlPriorityLow int32
	XlPriorityNormal int32
}{
	XlPriorityHigh: -4127,
	XlPriorityLow: -4134,
	XlPriorityNormal: -4143,
}

// enum XlPTSelectionMode
var XlPTSelectionMode = struct {
	XlLabelOnly int32
	XlDataAndLabel int32
	XlDataOnly int32
	XlOrigin int32
	XlButton int32
	XlBlanks int32
	XlFirstRow int32
}{
	XlLabelOnly: 1,
	XlDataAndLabel: 0,
	XlDataOnly: 2,
	XlOrigin: 3,
	XlButton: 15,
	XlBlanks: 4,
	XlFirstRow: 256,
}

// enum XlRangeAutoFormat
var XlRangeAutoFormat = struct {
	XlRangeAutoFormat3DEffects1 int32
	XlRangeAutoFormat3DEffects2 int32
	XlRangeAutoFormatAccounting1 int32
	XlRangeAutoFormatAccounting2 int32
	XlRangeAutoFormatAccounting3 int32
	XlRangeAutoFormatAccounting4 int32
	XlRangeAutoFormatClassic1 int32
	XlRangeAutoFormatClassic2 int32
	XlRangeAutoFormatClassic3 int32
	XlRangeAutoFormatColor1 int32
	XlRangeAutoFormatColor2 int32
	XlRangeAutoFormatColor3 int32
	XlRangeAutoFormatList1 int32
	XlRangeAutoFormatList2 int32
	XlRangeAutoFormatList3 int32
	XlRangeAutoFormatLocalFormat1 int32
	XlRangeAutoFormatLocalFormat2 int32
	XlRangeAutoFormatLocalFormat3 int32
	XlRangeAutoFormatLocalFormat4 int32
	XlRangeAutoFormatReport1 int32
	XlRangeAutoFormatReport2 int32
	XlRangeAutoFormatReport3 int32
	XlRangeAutoFormatReport4 int32
	XlRangeAutoFormatReport5 int32
	XlRangeAutoFormatReport6 int32
	XlRangeAutoFormatReport7 int32
	XlRangeAutoFormatReport8 int32
	XlRangeAutoFormatReport9 int32
	XlRangeAutoFormatReport10 int32
	XlRangeAutoFormatClassicPivotTable int32
	XlRangeAutoFormatTable1 int32
	XlRangeAutoFormatTable2 int32
	XlRangeAutoFormatTable3 int32
	XlRangeAutoFormatTable4 int32
	XlRangeAutoFormatTable5 int32
	XlRangeAutoFormatTable6 int32
	XlRangeAutoFormatTable7 int32
	XlRangeAutoFormatTable8 int32
	XlRangeAutoFormatTable9 int32
	XlRangeAutoFormatTable10 int32
	XlRangeAutoFormatPTNone int32
	XlRangeAutoFormatNone int32
	XlRangeAutoFormatSimple int32
}{
	XlRangeAutoFormat3DEffects1: 13,
	XlRangeAutoFormat3DEffects2: 14,
	XlRangeAutoFormatAccounting1: 4,
	XlRangeAutoFormatAccounting2: 5,
	XlRangeAutoFormatAccounting3: 6,
	XlRangeAutoFormatAccounting4: 17,
	XlRangeAutoFormatClassic1: 1,
	XlRangeAutoFormatClassic2: 2,
	XlRangeAutoFormatClassic3: 3,
	XlRangeAutoFormatColor1: 7,
	XlRangeAutoFormatColor2: 8,
	XlRangeAutoFormatColor3: 9,
	XlRangeAutoFormatList1: 10,
	XlRangeAutoFormatList2: 11,
	XlRangeAutoFormatList3: 12,
	XlRangeAutoFormatLocalFormat1: 15,
	XlRangeAutoFormatLocalFormat2: 16,
	XlRangeAutoFormatLocalFormat3: 19,
	XlRangeAutoFormatLocalFormat4: 20,
	XlRangeAutoFormatReport1: 21,
	XlRangeAutoFormatReport2: 22,
	XlRangeAutoFormatReport3: 23,
	XlRangeAutoFormatReport4: 24,
	XlRangeAutoFormatReport5: 25,
	XlRangeAutoFormatReport6: 26,
	XlRangeAutoFormatReport7: 27,
	XlRangeAutoFormatReport8: 28,
	XlRangeAutoFormatReport9: 29,
	XlRangeAutoFormatReport10: 30,
	XlRangeAutoFormatClassicPivotTable: 31,
	XlRangeAutoFormatTable1: 32,
	XlRangeAutoFormatTable2: 33,
	XlRangeAutoFormatTable3: 34,
	XlRangeAutoFormatTable4: 35,
	XlRangeAutoFormatTable5: 36,
	XlRangeAutoFormatTable6: 37,
	XlRangeAutoFormatTable7: 38,
	XlRangeAutoFormatTable8: 39,
	XlRangeAutoFormatTable9: 40,
	XlRangeAutoFormatTable10: 41,
	XlRangeAutoFormatPTNone: 42,
	XlRangeAutoFormatNone: -4142,
	XlRangeAutoFormatSimple: -4154,
}

// enum XlReferenceType
var XlReferenceType = struct {
	XlAbsolute int32
	XlAbsRowRelColumn int32
	XlRelative int32
	XlRelRowAbsColumn int32
}{
	XlAbsolute: 1,
	XlAbsRowRelColumn: 2,
	XlRelative: 4,
	XlRelRowAbsColumn: 3,
}

// enum XlLayoutFormType
var XlLayoutFormType = struct {
	XlTabular int32
	XlOutline int32
}{
	XlTabular: 0,
	XlOutline: 1,
}

// enum XlRoutingSlipDelivery
var XlRoutingSlipDelivery = struct {
	XlAllAtOnce int32
	XlOneAfterAnother int32
}{
	XlAllAtOnce: 2,
	XlOneAfterAnother: 1,
}

// enum XlRoutingSlipStatus
var XlRoutingSlipStatus = struct {
	XlNotYetRouted int32
	XlRoutingComplete int32
	XlRoutingInProgress int32
}{
	XlNotYetRouted: 0,
	XlRoutingComplete: 2,
	XlRoutingInProgress: 1,
}

// enum XlRunAutoMacro
var XlRunAutoMacro = struct {
	XlAutoActivate int32
	XlAutoClose int32
	XlAutoDeactivate int32
	XlAutoOpen int32
}{
	XlAutoActivate: 3,
	XlAutoClose: 2,
	XlAutoDeactivate: 4,
	XlAutoOpen: 1,
}

// enum XlSaveAction
var XlSaveAction = struct {
	XlDoNotSaveChanges int32
	XlSaveChanges int32
}{
	XlDoNotSaveChanges: 2,
	XlSaveChanges: 1,
}

// enum XlSaveAsAccessMode
var XlSaveAsAccessMode = struct {
	XlExclusive int32
	XlNoChange int32
	XlShared int32
}{
	XlExclusive: 3,
	XlNoChange: 1,
	XlShared: 2,
}

// enum XlSaveConflictResolution
var XlSaveConflictResolution = struct {
	XlLocalSessionChanges int32
	XlOtherSessionChanges int32
	XlUserResolution int32
}{
	XlLocalSessionChanges: 2,
	XlOtherSessionChanges: 3,
	XlUserResolution: 1,
}

// enum XlSearchDirection
var XlSearchDirection = struct {
	XlNext int32
	XlPrevious int32
}{
	XlNext: 1,
	XlPrevious: 2,
}

// enum XlSearchOrder
var XlSearchOrder = struct {
	XlByColumns int32
	XlByRows int32
}{
	XlByColumns: 2,
	XlByRows: 1,
}

// enum XlSheetVisibility
var XlSheetVisibility = struct {
	XlSheetVisible int32
	XlSheetHidden int32
	XlSheetVeryHidden int32
}{
	XlSheetVisible: -1,
	XlSheetHidden: 0,
	XlSheetVeryHidden: 2,
}

// enum XlSortMethod
var XlSortMethod = struct {
	XlPinYin int32
	XlStroke int32
}{
	XlPinYin: 1,
	XlStroke: 2,
}

// enum XlSortMethodOld
var XlSortMethodOld = struct {
	XlCodePage int32
	XlSyllabary int32
}{
	XlCodePage: 2,
	XlSyllabary: 1,
}

// enum XlSortOrder
var XlSortOrder = struct {
	XlAscending int32
	XlDescending int32
}{
	XlAscending: 1,
	XlDescending: 2,
}

// enum XlSortOrientation
var XlSortOrientation = struct {
	XlSortRows int32
	XlSortColumns int32
}{
	XlSortRows: 2,
	XlSortColumns: 1,
}

// enum XlSortType
var XlSortType = struct {
	XlSortLabels int32
	XlSortValues int32
}{
	XlSortLabels: 2,
	XlSortValues: 1,
}

// enum XlSpecialCellsValue
var XlSpecialCellsValue = struct {
	XlErrors int32
	XlLogical int32
	XlNumbers int32
	XlTextValues int32
}{
	XlErrors: 16,
	XlLogical: 4,
	XlNumbers: 1,
	XlTextValues: 2,
}

// enum XlSubscribeToFormat
var XlSubscribeToFormat = struct {
	XlSubscribeToPicture int32
	XlSubscribeToText int32
}{
	XlSubscribeToPicture: -4147,
	XlSubscribeToText: -4158,
}

// enum XlSummaryRow
var XlSummaryRow = struct {
	XlSummaryAbove int32
	XlSummaryBelow int32
}{
	XlSummaryAbove: 0,
	XlSummaryBelow: 1,
}

// enum XlSummaryColumn
var XlSummaryColumn = struct {
	XlSummaryOnLeft int32
	XlSummaryOnRight int32
}{
	XlSummaryOnLeft: -4131,
	XlSummaryOnRight: -4152,
}

// enum XlSummaryReportType
var XlSummaryReportType = struct {
	XlSummaryPivotTable int32
	XlStandardSummary int32
}{
	XlSummaryPivotTable: -4148,
	XlStandardSummary: 1,
}

// enum XlTabPosition
var XlTabPosition = struct {
	XlTabPositionFirst int32
	XlTabPositionLast int32
}{
	XlTabPositionFirst: 0,
	XlTabPositionLast: 1,
}

// enum XlTextParsingType
var XlTextParsingType = struct {
	XlDelimited int32
	XlFixedWidth int32
}{
	XlDelimited: 1,
	XlFixedWidth: 2,
}

// enum XlTextQualifier
var XlTextQualifier = struct {
	XlTextQualifierDoubleQuote int32
	XlTextQualifierNone int32
	XlTextQualifierSingleQuote int32
}{
	XlTextQualifierDoubleQuote: 1,
	XlTextQualifierNone: -4142,
	XlTextQualifierSingleQuote: 2,
}

// enum XlWBATemplate
var XlWBATemplate = struct {
	XlWBATChart int32
	XlWBATExcel4IntlMacroSheet int32
	XlWBATExcel4MacroSheet int32
	XlWBATWorksheet int32
}{
	XlWBATChart: -4109,
	XlWBATExcel4IntlMacroSheet: 4,
	XlWBATExcel4MacroSheet: 3,
	XlWBATWorksheet: -4167,
}

// enum XlWindowView
var XlWindowView = struct {
	XlNormalView int32
	XlPageBreakPreview int32
	XlPageLayoutView int32
}{
	XlNormalView: 1,
	XlPageBreakPreview: 2,
	XlPageLayoutView: 3,
}

// enum XlXLMMacroType
var XlXLMMacroType = struct {
	XlCommand int32
	XlFunction int32
	XlNotXLM int32
}{
	XlCommand: 2,
	XlFunction: 1,
	XlNotXLM: 3,
}

// enum XlYesNoGuess
var XlYesNoGuess = struct {
	XlGuess int32
	XlNo int32
	XlYes int32
}{
	XlGuess: 0,
	XlNo: 2,
	XlYes: 1,
}

// enum XlBordersIndex
var XlBordersIndex = struct {
	XlInsideHorizontal int32
	XlInsideVertical int32
	XlDiagonalDown int32
	XlDiagonalUp int32
	XlEdgeBottom int32
	XlEdgeLeft int32
	XlEdgeRight int32
	XlEdgeTop int32
}{
	XlInsideHorizontal: 12,
	XlInsideVertical: 11,
	XlDiagonalDown: 5,
	XlDiagonalUp: 6,
	XlEdgeBottom: 9,
	XlEdgeLeft: 7,
	XlEdgeRight: 10,
	XlEdgeTop: 8,
}

// enum XlToolbarProtection
var XlToolbarProtection = struct {
	XlNoButtonChanges int32
	XlNoChanges int32
	XlNoDockingChanges int32
	XlToolbarProtectionNone int32
	XlNoShapeChanges int32
}{
	XlNoButtonChanges: 1,
	XlNoChanges: 4,
	XlNoDockingChanges: 3,
	XlToolbarProtectionNone: -4143,
	XlNoShapeChanges: 2,
}

// enum XlBuiltInDialog
var XlBuiltInDialog = struct {
	XlDialogOpen int32
	XlDialogOpenLinks int32
	XlDialogSaveAs int32
	XlDialogFileDelete int32
	XlDialogPageSetup int32
	XlDialogPrint int32
	XlDialogPrinterSetup int32
	XlDialogArrangeAll int32
	XlDialogWindowSize int32
	XlDialogWindowMove int32
	XlDialogRun int32
	XlDialogSetPrintTitles int32
	XlDialogFont int32
	XlDialogDisplay int32
	XlDialogProtectDocument int32
	XlDialogCalculation int32
	XlDialogExtract int32
	XlDialogDataDelete int32
	XlDialogSort int32
	XlDialogDataSeries int32
	XlDialogTable int32
	XlDialogFormatNumber int32
	XlDialogAlignment int32
	XlDialogStyle int32
	XlDialogBorder int32
	XlDialogCellProtection int32
	XlDialogColumnWidth int32
	XlDialogClear int32
	XlDialogPasteSpecial int32
	XlDialogEditDelete int32
	XlDialogInsert int32
	XlDialogPasteNames int32
	XlDialogDefineName int32
	XlDialogCreateNames int32
	XlDialogFormulaGoto int32
	XlDialogFormulaFind int32
	XlDialogGalleryArea int32
	XlDialogGalleryBar int32
	XlDialogGalleryColumn int32
	XlDialogGalleryLine int32
	XlDialogGalleryPie int32
	XlDialogGalleryScatter int32
	XlDialogCombination int32
	XlDialogGridlines int32
	XlDialogAxes int32
	XlDialogAttachText int32
	XlDialogPatterns int32
	XlDialogMainChart int32
	XlDialogOverlay int32
	XlDialogScale int32
	XlDialogFormatLegend int32
	XlDialogFormatText int32
	XlDialogParse int32
	XlDialogUnhide int32
	XlDialogWorkspace int32
	XlDialogActivate int32
	XlDialogCopyPicture int32
	XlDialogDeleteName int32
	XlDialogDeleteFormat int32
	XlDialogNew int32
	XlDialogRowHeight int32
	XlDialogFormatMove int32
	XlDialogFormatSize int32
	XlDialogFormulaReplace int32
	XlDialogSelectSpecial int32
	XlDialogApplyNames int32
	XlDialogReplaceFont int32
	XlDialogSplit int32
	XlDialogOutline int32
	XlDialogSaveWorkbook int32
	XlDialogCopyChart int32
	XlDialogFormatFont int32
	XlDialogNote int32
	XlDialogSetUpdateStatus int32
	XlDialogColorPalette int32
	XlDialogChangeLink int32
	XlDialogAppMove int32
	XlDialogAppSize int32
	XlDialogMainChartType int32
	XlDialogOverlayChartType int32
	XlDialogOpenMail int32
	XlDialogSendMail int32
	XlDialogStandardFont int32
	XlDialogConsolidate int32
	XlDialogSortSpecial int32
	XlDialogGallery3dArea int32
	XlDialogGallery3dColumn int32
	XlDialogGallery3dLine int32
	XlDialogGallery3dPie int32
	XlDialogView3d int32
	XlDialogGoalSeek int32
	XlDialogWorkgroup int32
	XlDialogFillGroup int32
	XlDialogUpdateLink int32
	XlDialogPromote int32
	XlDialogDemote int32
	XlDialogShowDetail int32
	XlDialogObjectProperties int32
	XlDialogSaveNewObject int32
	XlDialogApplyStyle int32
	XlDialogAssignToObject int32
	XlDialogObjectProtection int32
	XlDialogCreatePublisher int32
	XlDialogSubscribeTo int32
	XlDialogShowToolbar int32
	XlDialogPrintPreview int32
	XlDialogEditColor int32
	XlDialogFormatMain int32
	XlDialogFormatOverlay int32
	XlDialogEditSeries int32
	XlDialogDefineStyle int32
	XlDialogGalleryRadar int32
	XlDialogEditionOptions int32
	XlDialogZoom int32
	XlDialogInsertObject int32
	XlDialogSize int32
	XlDialogMove int32
	XlDialogFormatAuto int32
	XlDialogGallery3dBar int32
	XlDialogGallery3dSurface int32
	XlDialogCustomizeToolbar int32
	XlDialogWorkbookAdd int32
	XlDialogWorkbookMove int32
	XlDialogWorkbookCopy int32
	XlDialogWorkbookOptions int32
	XlDialogSaveWorkspace int32
	XlDialogChartWizard int32
	XlDialogAssignToTool int32
	XlDialogPlacement int32
	XlDialogFillWorkgroup int32
	XlDialogWorkbookNew int32
	XlDialogScenarioCells int32
	XlDialogScenarioAdd int32
	XlDialogScenarioEdit int32
	XlDialogScenarioSummary int32
	XlDialogPivotTableWizard int32
	XlDialogPivotFieldProperties int32
	XlDialogOptionsCalculation int32
	XlDialogOptionsEdit int32
	XlDialogOptionsView int32
	XlDialogAddinManager int32
	XlDialogMenuEditor int32
	XlDialogAttachToolbars int32
	XlDialogOptionsChart int32
	XlDialogVbaInsertFile int32
	XlDialogVbaProcedureDefinition int32
	XlDialogRoutingSlip int32
	XlDialogMailLogon int32
	XlDialogInsertPicture int32
	XlDialogGalleryDoughnut int32
	XlDialogChartTrend int32
	XlDialogWorkbookInsert int32
	XlDialogOptionsTransition int32
	XlDialogOptionsGeneral int32
	XlDialogFilterAdvanced int32
	XlDialogMailNextLetter int32
	XlDialogDataLabel int32
	XlDialogInsertTitle int32
	XlDialogFontProperties int32
	XlDialogMacroOptions int32
	XlDialogWorkbookUnhide int32
	XlDialogWorkbookName int32
	XlDialogGalleryCustom int32
	XlDialogAddChartAutoformat int32
	XlDialogChartAddData int32
	XlDialogTabOrder int32
	XlDialogSubtotalCreate int32
	XlDialogWorkbookTabSplit int32
	XlDialogWorkbookProtect int32
	XlDialogScrollbarProperties int32
	XlDialogPivotShowPages int32
	XlDialogTextToColumns int32
	XlDialogFormatCharttype int32
	XlDialogPivotFieldGroup int32
	XlDialogPivotFieldUngroup int32
	XlDialogCheckboxProperties int32
	XlDialogLabelProperties int32
	XlDialogListboxProperties int32
	XlDialogEditboxProperties int32
	XlDialogOpenText int32
	XlDialogPushbuttonProperties int32
	XlDialogFilter int32
	XlDialogFunctionWizard int32
	XlDialogSaveCopyAs int32
	XlDialogOptionsListsAdd int32
	XlDialogSeriesAxes int32
	XlDialogSeriesX int32
	XlDialogSeriesY int32
	XlDialogErrorbarX int32
	XlDialogErrorbarY int32
	XlDialogFormatChart int32
	XlDialogSeriesOrder int32
	XlDialogMailEditMailer int32
	XlDialogStandardWidth int32
	XlDialogScenarioMerge int32
	XlDialogProperties int32
	XlDialogSummaryInfo int32
	XlDialogFindFile int32
	XlDialogActiveCellFont int32
	XlDialogVbaMakeAddin int32
	XlDialogFileSharing int32
	XlDialogAutoCorrect int32
	XlDialogCustomViews int32
	XlDialogInsertNameLabel int32
	XlDialogSeriesShape int32
	XlDialogChartOptionsDataLabels int32
	XlDialogChartOptionsDataTable int32
	XlDialogSetBackgroundPicture int32
	XlDialogDataValidation int32
	XlDialogChartType int32
	XlDialogChartLocation int32
	XlDialogPhonetic_ int32
	XlDialogChartSourceData int32
	XlDialogChartSourceData_ int32
	XlDialogSeriesOptions int32
	XlDialogPivotTableOptions int32
	XlDialogPivotSolveOrder int32
	XlDialogPivotCalculatedField int32
	XlDialogPivotCalculatedItem int32
	XlDialogConditionalFormatting int32
	XlDialogInsertHyperlink int32
	XlDialogProtectSharing int32
	XlDialogOptionsME int32
	XlDialogPublishAsWebPage int32
	XlDialogPhonetic int32
	XlDialogNewWebQuery int32
	XlDialogImportTextFile int32
	XlDialogExternalDataProperties int32
	XlDialogWebOptionsGeneral int32
	XlDialogWebOptionsFiles int32
	XlDialogWebOptionsPictures int32
	XlDialogWebOptionsEncoding int32
	XlDialogWebOptionsFonts int32
	XlDialogPivotClientServerSet int32
	XlDialogPropertyFields int32
	XlDialogSearch int32
	XlDialogEvaluateFormula int32
	XlDialogDataLabelMultiple int32
	XlDialogChartOptionsDataLabelMultiple int32
	XlDialogErrorChecking int32
	XlDialogWebOptionsBrowsers int32
	XlDialogCreateList int32
	XlDialogPermission int32
	XlDialogMyPermission int32
	XlDialogDocumentInspector int32
	XlDialogNameManager int32
	XlDialogNewName int32
	XlDialogSparklineInsertLine int32
	XlDialogSparklineInsertColumn int32
	XlDialogSparklineInsertWinLoss int32
	XlDialogSlicerSettings int32
	XlDialogSlicerCreation int32
	XlDialogSlicerPivotTableConnections int32
	XlDialogPivotTableSlicerConnections int32
	XlDialogPivotTableWhatIfAnalysisSettings int32
	XlDialogSetManager int32
	XlDialogSetMDXEditor int32
	XlDialogSetTupleEditorOnRows int32
	XlDialogSetTupleEditorOnColumns int32
}{
	XlDialogOpen: 1,
	XlDialogOpenLinks: 2,
	XlDialogSaveAs: 5,
	XlDialogFileDelete: 6,
	XlDialogPageSetup: 7,
	XlDialogPrint: 8,
	XlDialogPrinterSetup: 9,
	XlDialogArrangeAll: 12,
	XlDialogWindowSize: 13,
	XlDialogWindowMove: 14,
	XlDialogRun: 17,
	XlDialogSetPrintTitles: 23,
	XlDialogFont: 26,
	XlDialogDisplay: 27,
	XlDialogProtectDocument: 28,
	XlDialogCalculation: 32,
	XlDialogExtract: 35,
	XlDialogDataDelete: 36,
	XlDialogSort: 39,
	XlDialogDataSeries: 40,
	XlDialogTable: 41,
	XlDialogFormatNumber: 42,
	XlDialogAlignment: 43,
	XlDialogStyle: 44,
	XlDialogBorder: 45,
	XlDialogCellProtection: 46,
	XlDialogColumnWidth: 47,
	XlDialogClear: 52,
	XlDialogPasteSpecial: 53,
	XlDialogEditDelete: 54,
	XlDialogInsert: 55,
	XlDialogPasteNames: 58,
	XlDialogDefineName: 61,
	XlDialogCreateNames: 62,
	XlDialogFormulaGoto: 63,
	XlDialogFormulaFind: 64,
	XlDialogGalleryArea: 67,
	XlDialogGalleryBar: 68,
	XlDialogGalleryColumn: 69,
	XlDialogGalleryLine: 70,
	XlDialogGalleryPie: 71,
	XlDialogGalleryScatter: 72,
	XlDialogCombination: 73,
	XlDialogGridlines: 76,
	XlDialogAxes: 78,
	XlDialogAttachText: 80,
	XlDialogPatterns: 84,
	XlDialogMainChart: 85,
	XlDialogOverlay: 86,
	XlDialogScale: 87,
	XlDialogFormatLegend: 88,
	XlDialogFormatText: 89,
	XlDialogParse: 91,
	XlDialogUnhide: 94,
	XlDialogWorkspace: 95,
	XlDialogActivate: 103,
	XlDialogCopyPicture: 108,
	XlDialogDeleteName: 110,
	XlDialogDeleteFormat: 111,
	XlDialogNew: 119,
	XlDialogRowHeight: 127,
	XlDialogFormatMove: 128,
	XlDialogFormatSize: 129,
	XlDialogFormulaReplace: 130,
	XlDialogSelectSpecial: 132,
	XlDialogApplyNames: 133,
	XlDialogReplaceFont: 134,
	XlDialogSplit: 137,
	XlDialogOutline: 142,
	XlDialogSaveWorkbook: 145,
	XlDialogCopyChart: 147,
	XlDialogFormatFont: 150,
	XlDialogNote: 154,
	XlDialogSetUpdateStatus: 159,
	XlDialogColorPalette: 161,
	XlDialogChangeLink: 166,
	XlDialogAppMove: 170,
	XlDialogAppSize: 171,
	XlDialogMainChartType: 185,
	XlDialogOverlayChartType: 186,
	XlDialogOpenMail: 188,
	XlDialogSendMail: 189,
	XlDialogStandardFont: 190,
	XlDialogConsolidate: 191,
	XlDialogSortSpecial: 192,
	XlDialogGallery3dArea: 193,
	XlDialogGallery3dColumn: 194,
	XlDialogGallery3dLine: 195,
	XlDialogGallery3dPie: 196,
	XlDialogView3d: 197,
	XlDialogGoalSeek: 198,
	XlDialogWorkgroup: 199,
	XlDialogFillGroup: 200,
	XlDialogUpdateLink: 201,
	XlDialogPromote: 202,
	XlDialogDemote: 203,
	XlDialogShowDetail: 204,
	XlDialogObjectProperties: 207,
	XlDialogSaveNewObject: 208,
	XlDialogApplyStyle: 212,
	XlDialogAssignToObject: 213,
	XlDialogObjectProtection: 214,
	XlDialogCreatePublisher: 217,
	XlDialogSubscribeTo: 218,
	XlDialogShowToolbar: 220,
	XlDialogPrintPreview: 222,
	XlDialogEditColor: 223,
	XlDialogFormatMain: 225,
	XlDialogFormatOverlay: 226,
	XlDialogEditSeries: 228,
	XlDialogDefineStyle: 229,
	XlDialogGalleryRadar: 249,
	XlDialogEditionOptions: 251,
	XlDialogZoom: 256,
	XlDialogInsertObject: 259,
	XlDialogSize: 261,
	XlDialogMove: 262,
	XlDialogFormatAuto: 269,
	XlDialogGallery3dBar: 272,
	XlDialogGallery3dSurface: 273,
	XlDialogCustomizeToolbar: 276,
	XlDialogWorkbookAdd: 281,
	XlDialogWorkbookMove: 282,
	XlDialogWorkbookCopy: 283,
	XlDialogWorkbookOptions: 284,
	XlDialogSaveWorkspace: 285,
	XlDialogChartWizard: 288,
	XlDialogAssignToTool: 293,
	XlDialogPlacement: 300,
	XlDialogFillWorkgroup: 301,
	XlDialogWorkbookNew: 302,
	XlDialogScenarioCells: 305,
	XlDialogScenarioAdd: 307,
	XlDialogScenarioEdit: 308,
	XlDialogScenarioSummary: 311,
	XlDialogPivotTableWizard: 312,
	XlDialogPivotFieldProperties: 313,
	XlDialogOptionsCalculation: 318,
	XlDialogOptionsEdit: 319,
	XlDialogOptionsView: 320,
	XlDialogAddinManager: 321,
	XlDialogMenuEditor: 322,
	XlDialogAttachToolbars: 323,
	XlDialogOptionsChart: 325,
	XlDialogVbaInsertFile: 328,
	XlDialogVbaProcedureDefinition: 330,
	XlDialogRoutingSlip: 336,
	XlDialogMailLogon: 339,
	XlDialogInsertPicture: 342,
	XlDialogGalleryDoughnut: 344,
	XlDialogChartTrend: 350,
	XlDialogWorkbookInsert: 354,
	XlDialogOptionsTransition: 355,
	XlDialogOptionsGeneral: 356,
	XlDialogFilterAdvanced: 370,
	XlDialogMailNextLetter: 378,
	XlDialogDataLabel: 379,
	XlDialogInsertTitle: 380,
	XlDialogFontProperties: 381,
	XlDialogMacroOptions: 382,
	XlDialogWorkbookUnhide: 384,
	XlDialogWorkbookName: 386,
	XlDialogGalleryCustom: 388,
	XlDialogAddChartAutoformat: 390,
	XlDialogChartAddData: 392,
	XlDialogTabOrder: 394,
	XlDialogSubtotalCreate: 398,
	XlDialogWorkbookTabSplit: 415,
	XlDialogWorkbookProtect: 417,
	XlDialogScrollbarProperties: 420,
	XlDialogPivotShowPages: 421,
	XlDialogTextToColumns: 422,
	XlDialogFormatCharttype: 423,
	XlDialogPivotFieldGroup: 433,
	XlDialogPivotFieldUngroup: 434,
	XlDialogCheckboxProperties: 435,
	XlDialogLabelProperties: 436,
	XlDialogListboxProperties: 437,
	XlDialogEditboxProperties: 438,
	XlDialogOpenText: 441,
	XlDialogPushbuttonProperties: 445,
	XlDialogFilter: 447,
	XlDialogFunctionWizard: 450,
	XlDialogSaveCopyAs: 456,
	XlDialogOptionsListsAdd: 458,
	XlDialogSeriesAxes: 460,
	XlDialogSeriesX: 461,
	XlDialogSeriesY: 462,
	XlDialogErrorbarX: 463,
	XlDialogErrorbarY: 464,
	XlDialogFormatChart: 465,
	XlDialogSeriesOrder: 466,
	XlDialogMailEditMailer: 470,
	XlDialogStandardWidth: 472,
	XlDialogScenarioMerge: 473,
	XlDialogProperties: 474,
	XlDialogSummaryInfo: 474,
	XlDialogFindFile: 475,
	XlDialogActiveCellFont: 476,
	XlDialogVbaMakeAddin: 478,
	XlDialogFileSharing: 481,
	XlDialogAutoCorrect: 485,
	XlDialogCustomViews: 493,
	XlDialogInsertNameLabel: 496,
	XlDialogSeriesShape: 504,
	XlDialogChartOptionsDataLabels: 505,
	XlDialogChartOptionsDataTable: 506,
	XlDialogSetBackgroundPicture: 509,
	XlDialogDataValidation: 525,
	XlDialogChartType: 526,
	XlDialogChartLocation: 527,
	XlDialogPhonetic_: 538,
	XlDialogChartSourceData: 540,
	XlDialogChartSourceData_: 541,
	XlDialogSeriesOptions: 557,
	XlDialogPivotTableOptions: 567,
	XlDialogPivotSolveOrder: 568,
	XlDialogPivotCalculatedField: 570,
	XlDialogPivotCalculatedItem: 572,
	XlDialogConditionalFormatting: 583,
	XlDialogInsertHyperlink: 596,
	XlDialogProtectSharing: 620,
	XlDialogOptionsME: 647,
	XlDialogPublishAsWebPage: 653,
	XlDialogPhonetic: 656,
	XlDialogNewWebQuery: 667,
	XlDialogImportTextFile: 666,
	XlDialogExternalDataProperties: 530,
	XlDialogWebOptionsGeneral: 683,
	XlDialogWebOptionsFiles: 684,
	XlDialogWebOptionsPictures: 685,
	XlDialogWebOptionsEncoding: 686,
	XlDialogWebOptionsFonts: 687,
	XlDialogPivotClientServerSet: 689,
	XlDialogPropertyFields: 754,
	XlDialogSearch: 731,
	XlDialogEvaluateFormula: 709,
	XlDialogDataLabelMultiple: 723,
	XlDialogChartOptionsDataLabelMultiple: 724,
	XlDialogErrorChecking: 732,
	XlDialogWebOptionsBrowsers: 773,
	XlDialogCreateList: 796,
	XlDialogPermission: 832,
	XlDialogMyPermission: 834,
	XlDialogDocumentInspector: 862,
	XlDialogNameManager: 977,
	XlDialogNewName: 978,
	XlDialogSparklineInsertLine: 1133,
	XlDialogSparklineInsertColumn: 1134,
	XlDialogSparklineInsertWinLoss: 1135,
	XlDialogSlicerSettings: 1179,
	XlDialogSlicerCreation: 1182,
	XlDialogSlicerPivotTableConnections: 1184,
	XlDialogPivotTableSlicerConnections: 1183,
	XlDialogPivotTableWhatIfAnalysisSettings: 1153,
	XlDialogSetManager: 1109,
	XlDialogSetMDXEditor: 1208,
	XlDialogSetTupleEditorOnRows: 1107,
	XlDialogSetTupleEditorOnColumns: 1108,
}

// enum XlParameterType
var XlParameterType = struct {
	XlPrompt int32
	XlConstant int32
	XlRange int32
}{
	XlPrompt: 0,
	XlConstant: 1,
	XlRange: 2,
}

// enum XlParameterDataType
var XlParameterDataType = struct {
	XlParamTypeUnknown int32
	XlParamTypeChar int32
	XlParamTypeNumeric int32
	XlParamTypeDecimal int32
	XlParamTypeInteger int32
	XlParamTypeSmallInt int32
	XlParamTypeFloat int32
	XlParamTypeReal int32
	XlParamTypeDouble int32
	XlParamTypeVarChar int32
	XlParamTypeDate int32
	XlParamTypeTime int32
	XlParamTypeTimestamp int32
	XlParamTypeLongVarChar int32
	XlParamTypeBinary int32
	XlParamTypeVarBinary int32
	XlParamTypeLongVarBinary int32
	XlParamTypeBigInt int32
	XlParamTypeTinyInt int32
	XlParamTypeBit int32
	XlParamTypeWChar int32
}{
	XlParamTypeUnknown: 0,
	XlParamTypeChar: 1,
	XlParamTypeNumeric: 2,
	XlParamTypeDecimal: 3,
	XlParamTypeInteger: 4,
	XlParamTypeSmallInt: 5,
	XlParamTypeFloat: 6,
	XlParamTypeReal: 7,
	XlParamTypeDouble: 8,
	XlParamTypeVarChar: 12,
	XlParamTypeDate: 9,
	XlParamTypeTime: 10,
	XlParamTypeTimestamp: 11,
	XlParamTypeLongVarChar: -1,
	XlParamTypeBinary: -2,
	XlParamTypeVarBinary: -3,
	XlParamTypeLongVarBinary: -4,
	XlParamTypeBigInt: -5,
	XlParamTypeTinyInt: -6,
	XlParamTypeBit: -7,
	XlParamTypeWChar: -8,
}

// enum XlFormControl
var XlFormControl = struct {
	XlButtonControl int32
	XlCheckBox int32
	XlDropDown int32
	XlEditBox int32
	XlGroupBox int32
	XlLabel int32
	XlListBox int32
	XlOptionButton int32
	XlScrollBar int32
	XlSpinner int32
}{
	XlButtonControl: 0,
	XlCheckBox: 1,
	XlDropDown: 2,
	XlEditBox: 3,
	XlGroupBox: 4,
	XlLabel: 5,
	XlListBox: 6,
	XlOptionButton: 7,
	XlScrollBar: 8,
	XlSpinner: 9,
}

// enum XlSourceType
var XlSourceType = struct {
	XlSourceWorkbook int32
	XlSourceSheet int32
	XlSourcePrintArea int32
	XlSourceAutoFilter int32
	XlSourceRange int32
	XlSourceChart int32
	XlSourcePivotTable int32
	XlSourceQuery int32
}{
	XlSourceWorkbook: 0,
	XlSourceSheet: 1,
	XlSourcePrintArea: 2,
	XlSourceAutoFilter: 3,
	XlSourceRange: 4,
	XlSourceChart: 5,
	XlSourcePivotTable: 6,
	XlSourceQuery: 7,
}

// enum XlHtmlType
var XlHtmlType = struct {
	XlHtmlStatic int32
	XlHtmlCalc int32
	XlHtmlList int32
	XlHtmlChart int32
}{
	XlHtmlStatic: 0,
	XlHtmlCalc: 1,
	XlHtmlList: 2,
	XlHtmlChart: 3,
}

// enum XlPivotFormatType
var XlPivotFormatType = struct {
	XlReport1 int32
	XlReport2 int32
	XlReport3 int32
	XlReport4 int32
	XlReport5 int32
	XlReport6 int32
	XlReport7 int32
	XlReport8 int32
	XlReport9 int32
	XlReport10 int32
	XlTable1 int32
	XlTable2 int32
	XlTable3 int32
	XlTable4 int32
	XlTable5 int32
	XlTable6 int32
	XlTable7 int32
	XlTable8 int32
	XlTable9 int32
	XlTable10 int32
	XlPTClassic int32
	XlPTNone int32
}{
	XlReport1: 0,
	XlReport2: 1,
	XlReport3: 2,
	XlReport4: 3,
	XlReport5: 4,
	XlReport6: 5,
	XlReport7: 6,
	XlReport8: 7,
	XlReport9: 8,
	XlReport10: 9,
	XlTable1: 10,
	XlTable2: 11,
	XlTable3: 12,
	XlTable4: 13,
	XlTable5: 14,
	XlTable6: 15,
	XlTable7: 16,
	XlTable8: 17,
	XlTable9: 18,
	XlTable10: 19,
	XlPTClassic: 20,
	XlPTNone: 21,
}

// enum XlCmdType
var XlCmdType = struct {
	XlCmdCube int32
	XlCmdSql int32
	XlCmdTable int32
	XlCmdDefault int32
	XlCmdList int32
}{
	XlCmdCube: 1,
	XlCmdSql: 2,
	XlCmdTable: 3,
	XlCmdDefault: 4,
	XlCmdList: 5,
}

// enum XlColumnDataType
var XlColumnDataType = struct {
	XlGeneralFormat int32
	XlTextFormat int32
	XlMDYFormat int32
	XlDMYFormat int32
	XlYMDFormat int32
	XlMYDFormat int32
	XlDYMFormat int32
	XlYDMFormat int32
	XlSkipColumn int32
	XlEMDFormat int32
}{
	XlGeneralFormat: 1,
	XlTextFormat: 2,
	XlMDYFormat: 3,
	XlDMYFormat: 4,
	XlYMDFormat: 5,
	XlMYDFormat: 6,
	XlDYMFormat: 7,
	XlYDMFormat: 8,
	XlSkipColumn: 9,
	XlEMDFormat: 10,
}

// enum XlQueryType
var XlQueryType = struct {
	XlODBCQuery int32
	XlDAORecordset int32
	XlWebQuery int32
	XlOLEDBQuery int32
	XlTextImport int32
	XlADORecordset int32
}{
	XlODBCQuery: 1,
	XlDAORecordset: 2,
	XlWebQuery: 4,
	XlOLEDBQuery: 5,
	XlTextImport: 6,
	XlADORecordset: 7,
}

// enum XlWebSelectionType
var XlWebSelectionType = struct {
	XlEntirePage int32
	XlAllTables int32
	XlSpecifiedTables int32
}{
	XlEntirePage: 1,
	XlAllTables: 2,
	XlSpecifiedTables: 3,
}

// enum XlCubeFieldType
var XlCubeFieldType = struct {
	XlHierarchy int32
	XlMeasure int32
	XlSet int32
}{
	XlHierarchy: 1,
	XlMeasure: 2,
	XlSet: 3,
}

// enum XlWebFormatting
var XlWebFormatting = struct {
	XlWebFormattingAll int32
	XlWebFormattingRTF int32
	XlWebFormattingNone int32
}{
	XlWebFormattingAll: 1,
	XlWebFormattingRTF: 2,
	XlWebFormattingNone: 3,
}

// enum XlDisplayDrawingObjects
var XlDisplayDrawingObjects = struct {
	XlDisplayShapes int32
	XlHide int32
	XlPlaceholders int32
}{
	XlDisplayShapes: -4104,
	XlHide: 3,
	XlPlaceholders: 2,
}

// enum XlSubtototalLocationType
var XlSubtototalLocationType = struct {
	XlAtTop int32
	XlAtBottom int32
}{
	XlAtTop: 1,
	XlAtBottom: 2,
}

// enum XlPivotTableVersionList
var XlPivotTableVersionList = struct {
	XlPivotTableVersion2000 int32
	XlPivotTableVersion10 int32
	XlPivotTableVersion11 int32
	XlPivotTableVersion12 int32
	XlPivotTableVersion14 int32
	XlPivotTableVersionCurrent int32
}{
	XlPivotTableVersion2000: 0,
	XlPivotTableVersion10: 1,
	XlPivotTableVersion11: 2,
	XlPivotTableVersion12: 3,
	XlPivotTableVersion14: 4,
	XlPivotTableVersionCurrent: -1,
}

// enum XlPrintErrors
var XlPrintErrors = struct {
	XlPrintErrorsDisplayed int32
	XlPrintErrorsBlank int32
	XlPrintErrorsDash int32
	XlPrintErrorsNA int32
}{
	XlPrintErrorsDisplayed: 0,
	XlPrintErrorsBlank: 1,
	XlPrintErrorsDash: 2,
	XlPrintErrorsNA: 3,
}

// enum XlPivotCellType
var XlPivotCellType = struct {
	XlPivotCellValue int32
	XlPivotCellPivotItem int32
	XlPivotCellSubtotal int32
	XlPivotCellGrandTotal int32
	XlPivotCellDataField int32
	XlPivotCellPivotField int32
	XlPivotCellPageFieldItem int32
	XlPivotCellCustomSubtotal int32
	XlPivotCellDataPivotField int32
	XlPivotCellBlankCell int32
}{
	XlPivotCellValue: 0,
	XlPivotCellPivotItem: 1,
	XlPivotCellSubtotal: 2,
	XlPivotCellGrandTotal: 3,
	XlPivotCellDataField: 4,
	XlPivotCellPivotField: 5,
	XlPivotCellPageFieldItem: 6,
	XlPivotCellCustomSubtotal: 7,
	XlPivotCellDataPivotField: 8,
	XlPivotCellBlankCell: 9,
}

// enum XlPivotTableMissingItems
var XlPivotTableMissingItems = struct {
	XlMissingItemsDefault int32
	XlMissingItemsNone int32
	XlMissingItemsMax int32
	XlMissingItemsMax2 int32
}{
	XlMissingItemsDefault: -1,
	XlMissingItemsNone: 0,
	XlMissingItemsMax: 32500,
	XlMissingItemsMax2: 1048576,
}

// enum XlCalculationState
var XlCalculationState = struct {
	XlDone int32
	XlCalculating int32
	XlPending int32
}{
	XlDone: 0,
	XlCalculating: 1,
	XlPending: 2,
}

// enum XlCalculationInterruptKey
var XlCalculationInterruptKey = struct {
	XlNoKey int32
	XlEscKey int32
	XlAnyKey int32
}{
	XlNoKey: 0,
	XlEscKey: 1,
	XlAnyKey: 2,
}

// enum XlSortDataOption
var XlSortDataOption = struct {
	XlSortNormal int32
	XlSortTextAsNumbers int32
}{
	XlSortNormal: 0,
	XlSortTextAsNumbers: 1,
}

// enum XlUpdateLinks
var XlUpdateLinks = struct {
	XlUpdateLinksUserSetting int32
	XlUpdateLinksNever int32
	XlUpdateLinksAlways int32
}{
	XlUpdateLinksUserSetting: 1,
	XlUpdateLinksNever: 2,
	XlUpdateLinksAlways: 3,
}

// enum XlLinkStatus
var XlLinkStatus = struct {
	XlLinkStatusOK int32
	XlLinkStatusMissingFile int32
	XlLinkStatusMissingSheet int32
	XlLinkStatusOld int32
	XlLinkStatusSourceNotCalculated int32
	XlLinkStatusIndeterminate int32
	XlLinkStatusNotStarted int32
	XlLinkStatusInvalidName int32
	XlLinkStatusSourceNotOpen int32
	XlLinkStatusSourceOpen int32
	XlLinkStatusCopiedValues int32
}{
	XlLinkStatusOK: 0,
	XlLinkStatusMissingFile: 1,
	XlLinkStatusMissingSheet: 2,
	XlLinkStatusOld: 3,
	XlLinkStatusSourceNotCalculated: 4,
	XlLinkStatusIndeterminate: 5,
	XlLinkStatusNotStarted: 6,
	XlLinkStatusInvalidName: 7,
	XlLinkStatusSourceNotOpen: 8,
	XlLinkStatusSourceOpen: 9,
	XlLinkStatusCopiedValues: 10,
}

// enum XlSearchWithin
var XlSearchWithin = struct {
	XlWithinSheet int32
	XlWithinWorkbook int32
}{
	XlWithinSheet: 1,
	XlWithinWorkbook: 2,
}

// enum XlCorruptLoad
var XlCorruptLoad = struct {
	XlNormalLoad int32
	XlRepairFile int32
	XlExtractData int32
}{
	XlNormalLoad: 0,
	XlRepairFile: 1,
	XlExtractData: 2,
}

// enum XlRobustConnect
var XlRobustConnect = struct {
	XlAsRequired int32
	XlAlways int32
	XlNever int32
}{
	XlAsRequired: 0,
	XlAlways: 1,
	XlNever: 2,
}

// enum XlErrorChecks
var XlErrorChecks = struct {
	XlEvaluateToError int32
	XlTextDate int32
	XlNumberAsText int32
	XlInconsistentFormula int32
	XlOmittedCells int32
	XlUnlockedFormulaCells int32
	XlEmptyCellReferences int32
	XlListDataValidation int32
	XlInconsistentListFormula int32
}{
	XlEvaluateToError: 1,
	XlTextDate: 2,
	XlNumberAsText: 3,
	XlInconsistentFormula: 4,
	XlOmittedCells: 5,
	XlUnlockedFormulaCells: 6,
	XlEmptyCellReferences: 7,
	XlListDataValidation: 8,
	XlInconsistentListFormula: 9,
}

// enum XlDataLabelSeparator
var XlDataLabelSeparator = struct {
	XlDataLabelSeparatorDefault int32
}{
	XlDataLabelSeparatorDefault: 1,
}

// enum XlSmartTagDisplayMode
var XlSmartTagDisplayMode = struct {
	XlIndicatorAndButton int32
	XlDisplayNone int32
	XlButtonOnly int32
}{
	XlIndicatorAndButton: 0,
	XlDisplayNone: 1,
	XlButtonOnly: 2,
}

// enum XlRangeValueDataType
var XlRangeValueDataType = struct {
	XlRangeValueDefault int32
	XlRangeValueXMLSpreadsheet int32
	XlRangeValueMSPersistXML int32
}{
	XlRangeValueDefault: 10,
	XlRangeValueXMLSpreadsheet: 11,
	XlRangeValueMSPersistXML: 12,
}

// enum XlSpeakDirection
var XlSpeakDirection = struct {
	XlSpeakByRows int32
	XlSpeakByColumns int32
}{
	XlSpeakByRows: 0,
	XlSpeakByColumns: 1,
}

// enum XlInsertFormatOrigin
var XlInsertFormatOrigin = struct {
	XlFormatFromLeftOrAbove int32
	XlFormatFromRightOrBelow int32
}{
	XlFormatFromLeftOrAbove: 0,
	XlFormatFromRightOrBelow: 1,
}

// enum XlArabicModes
var XlArabicModes = struct {
	XlArabicNone int32
	XlArabicStrictAlefHamza int32
	XlArabicStrictFinalYaa int32
	XlArabicBothStrict int32
}{
	XlArabicNone: 0,
	XlArabicStrictAlefHamza: 1,
	XlArabicStrictFinalYaa: 2,
	XlArabicBothStrict: 3,
}

// enum XlImportDataAs
var XlImportDataAs = struct {
	XlQueryTable int32
	XlPivotTableReport int32
	XlTable int32
}{
	XlQueryTable: 0,
	XlPivotTableReport: 1,
	XlTable: 2,
}

// enum XlCalculatedMemberType
var XlCalculatedMemberType = struct {
	XlCalculatedMember int32
	XlCalculatedSet int32
}{
	XlCalculatedMember: 0,
	XlCalculatedSet: 1,
}

// enum XlHebrewModes
var XlHebrewModes = struct {
	XlHebrewFullScript int32
	XlHebrewPartialScript int32
	XlHebrewMixedScript int32
	XlHebrewMixedAuthorizedScript int32
}{
	XlHebrewFullScript: 0,
	XlHebrewPartialScript: 1,
	XlHebrewMixedScript: 2,
	XlHebrewMixedAuthorizedScript: 3,
}

// enum XlListObjectSourceType
var XlListObjectSourceType = struct {
	XlSrcExternal int32
	XlSrcRange int32
	XlSrcXml int32
	XlSrcQuery int32
}{
	XlSrcExternal: 0,
	XlSrcRange: 1,
	XlSrcXml: 2,
	XlSrcQuery: 3,
}

// enum XlTextVisualLayoutType
var XlTextVisualLayoutType = struct {
	XlTextVisualLTR int32
	XlTextVisualRTL int32
}{
	XlTextVisualLTR: 1,
	XlTextVisualRTL: 2,
}

// enum XlListDataType
var XlListDataType = struct {
	XlListDataTypeNone int32
	XlListDataTypeText int32
	XlListDataTypeMultiLineText int32
	XlListDataTypeNumber int32
	XlListDataTypeCurrency int32
	XlListDataTypeDateTime int32
	XlListDataTypeChoice int32
	XlListDataTypeChoiceMulti int32
	XlListDataTypeListLookup int32
	XlListDataTypeCheckbox int32
	XlListDataTypeHyperLink int32
	XlListDataTypeCounter int32
	XlListDataTypeMultiLineRichText int32
}{
	XlListDataTypeNone: 0,
	XlListDataTypeText: 1,
	XlListDataTypeMultiLineText: 2,
	XlListDataTypeNumber: 3,
	XlListDataTypeCurrency: 4,
	XlListDataTypeDateTime: 5,
	XlListDataTypeChoice: 6,
	XlListDataTypeChoiceMulti: 7,
	XlListDataTypeListLookup: 8,
	XlListDataTypeCheckbox: 9,
	XlListDataTypeHyperLink: 10,
	XlListDataTypeCounter: 11,
	XlListDataTypeMultiLineRichText: 12,
}

// enum XlTotalsCalculation
var XlTotalsCalculation = struct {
	XlTotalsCalculationNone int32
	XlTotalsCalculationSum int32
	XlTotalsCalculationAverage int32
	XlTotalsCalculationCount int32
	XlTotalsCalculationCountNums int32
	XlTotalsCalculationMin int32
	XlTotalsCalculationMax int32
	XlTotalsCalculationStdDev int32
	XlTotalsCalculationVar int32
	XlTotalsCalculationCustom int32
}{
	XlTotalsCalculationNone: 0,
	XlTotalsCalculationSum: 1,
	XlTotalsCalculationAverage: 2,
	XlTotalsCalculationCount: 3,
	XlTotalsCalculationCountNums: 4,
	XlTotalsCalculationMin: 5,
	XlTotalsCalculationMax: 6,
	XlTotalsCalculationStdDev: 7,
	XlTotalsCalculationVar: 8,
	XlTotalsCalculationCustom: 9,
}

// enum XlXmlLoadOption
var XlXmlLoadOption = struct {
	XlXmlLoadPromptUser int32
	XlXmlLoadOpenXml int32
	XlXmlLoadImportToList int32
	XlXmlLoadMapXml int32
}{
	XlXmlLoadPromptUser: 0,
	XlXmlLoadOpenXml: 1,
	XlXmlLoadImportToList: 2,
	XlXmlLoadMapXml: 3,
}

// enum XlSmartTagControlType
var XlSmartTagControlType = struct {
	XlSmartTagControlSmartTag int32
	XlSmartTagControlLink int32
	XlSmartTagControlHelp int32
	XlSmartTagControlHelpURL int32
	XlSmartTagControlSeparator int32
	XlSmartTagControlButton int32
	XlSmartTagControlLabel int32
	XlSmartTagControlImage int32
	XlSmartTagControlCheckbox int32
	XlSmartTagControlTextbox int32
	XlSmartTagControlListbox int32
	XlSmartTagControlCombo int32
	XlSmartTagControlActiveX int32
	XlSmartTagControlRadioGroup int32
}{
	XlSmartTagControlSmartTag: 1,
	XlSmartTagControlLink: 2,
	XlSmartTagControlHelp: 3,
	XlSmartTagControlHelpURL: 4,
	XlSmartTagControlSeparator: 5,
	XlSmartTagControlButton: 6,
	XlSmartTagControlLabel: 7,
	XlSmartTagControlImage: 8,
	XlSmartTagControlCheckbox: 9,
	XlSmartTagControlTextbox: 10,
	XlSmartTagControlListbox: 11,
	XlSmartTagControlCombo: 12,
	XlSmartTagControlActiveX: 13,
	XlSmartTagControlRadioGroup: 14,
}

// enum XlListConflict
var XlListConflict = struct {
	XlListConflictDialog int32
	XlListConflictRetryAllConflicts int32
	XlListConflictDiscardAllConflicts int32
	XlListConflictError int32
}{
	XlListConflictDialog: 0,
	XlListConflictRetryAllConflicts: 1,
	XlListConflictDiscardAllConflicts: 2,
	XlListConflictError: 3,
}

// enum XlXmlExportResult
var XlXmlExportResult = struct {
	XlXmlExportSuccess int32
	XlXmlExportValidationFailed int32
}{
	XlXmlExportSuccess: 0,
	XlXmlExportValidationFailed: 1,
}

// enum XlXmlImportResult
var XlXmlImportResult = struct {
	XlXmlImportSuccess int32
	XlXmlImportElementsTruncated int32
	XlXmlImportValidationFailed int32
}{
	XlXmlImportSuccess: 0,
	XlXmlImportElementsTruncated: 1,
	XlXmlImportValidationFailed: 2,
}

// enum XlRemoveDocInfoType
var XlRemoveDocInfoType = struct {
	XlRDIComments int32
	XlRDIRemovePersonalInformation int32
	XlRDIEmailHeader int32
	XlRDIRoutingSlip int32
	XlRDISendForReview int32
	XlRDIDocumentProperties int32
	XlRDIDocumentWorkspace int32
	XlRDIInkAnnotations int32
	XlRDIScenarioComments int32
	XlRDIPublishInfo int32
	XlRDIDocumentServerProperties int32
	XlRDIDocumentManagementPolicy int32
	XlRDIContentType int32
	XlRDIDefinedNameComments int32
	XlRDIInactiveDataConnections int32
	XlRDIPrinterPath int32
	XlRDIAll int32
}{
	XlRDIComments: 1,
	XlRDIRemovePersonalInformation: 4,
	XlRDIEmailHeader: 5,
	XlRDIRoutingSlip: 6,
	XlRDISendForReview: 7,
	XlRDIDocumentProperties: 8,
	XlRDIDocumentWorkspace: 10,
	XlRDIInkAnnotations: 11,
	XlRDIScenarioComments: 12,
	XlRDIPublishInfo: 13,
	XlRDIDocumentServerProperties: 14,
	XlRDIDocumentManagementPolicy: 15,
	XlRDIContentType: 16,
	XlRDIDefinedNameComments: 18,
	XlRDIInactiveDataConnections: 19,
	XlRDIPrinterPath: 20,
	XlRDIAll: 99,
}

// enum XlRgbColor
var XlRgbColor = struct {
	RgbAliceBlue int32
	RgbAntiqueWhite int32
	RgbAqua int32
	RgbAquamarine int32
	RgbAzure int32
	RgbBeige int32
	RgbBisque int32
	RgbBlack int32
	RgbBlanchedAlmond int32
	RgbBlue int32
	RgbBlueViolet int32
	RgbBrown int32
	RgbBurlyWood int32
	RgbCadetBlue int32
	RgbChartreuse int32
	RgbCoral int32
	RgbCornflowerBlue int32
	RgbCornsilk int32
	RgbCrimson int32
	RgbDarkBlue int32
	RgbDarkCyan int32
	RgbDarkGoldenrod int32
	RgbDarkGreen int32
	RgbDarkGray int32
	RgbDarkGrey int32
	RgbDarkKhaki int32
	RgbDarkMagenta int32
	RgbDarkOliveGreen int32
	RgbDarkOrange int32
	RgbDarkOrchid int32
	RgbDarkRed int32
	RgbDarkSalmon int32
	RgbDarkSeaGreen int32
	RgbDarkSlateBlue int32
	RgbDarkSlateGray int32
	RgbDarkSlateGrey int32
	RgbDarkTurquoise int32
	RgbDarkViolet int32
	RgbDeepPink int32
	RgbDeepSkyBlue int32
	RgbDimGray int32
	RgbDimGrey int32
	RgbDodgerBlue int32
	RgbFireBrick int32
	RgbFloralWhite int32
	RgbForestGreen int32
	RgbFuchsia int32
	RgbGainsboro int32
	RgbGhostWhite int32
	RgbGold int32
	RgbGoldenrod int32
	RgbGray int32
	RgbGreen int32
	RgbGrey int32
	RgbGreenYellow int32
	RgbHoneydew int32
	RgbHotPink int32
	RgbIndianRed int32
	RgbIndigo int32
	RgbIvory int32
	RgbKhaki int32
	RgbLavender int32
	RgbLavenderBlush int32
	RgbLawnGreen int32
	RgbLemonChiffon int32
	RgbLightBlue int32
	RgbLightCoral int32
	RgbLightCyan int32
	RgbLightGoldenrodYellow int32
	RgbLightGray int32
	RgbLightGreen int32
	RgbLightGrey int32
	RgbLightPink int32
	RgbLightSalmon int32
	RgbLightSeaGreen int32
	RgbLightSkyBlue int32
	RgbLightSlateGray int32
	RgbLightSlateGrey int32
	RgbLightSteelBlue int32
	RgbLightYellow int32
	RgbLime int32
	RgbLimeGreen int32
	RgbLinen int32
	RgbMaroon int32
	RgbMediumAquamarine int32
	RgbMediumBlue int32
	RgbMediumOrchid int32
	RgbMediumPurple int32
	RgbMediumSeaGreen int32
	RgbMediumSlateBlue int32
	RgbMediumSpringGreen int32
	RgbMediumTurquoise int32
	RgbMediumVioletRed int32
	RgbMidnightBlue int32
	RgbMintCream int32
	RgbMistyRose int32
	RgbMoccasin int32
	RgbNavajoWhite int32
	RgbNavy int32
	RgbNavyBlue int32
	RgbOldLace int32
	RgbOlive int32
	RgbOliveDrab int32
	RgbOrange int32
	RgbOrangeRed int32
	RgbOrchid int32
	RgbPaleGoldenrod int32
	RgbPaleGreen int32
	RgbPaleTurquoise int32
	RgbPaleVioletRed int32
	RgbPapayaWhip int32
	RgbPeachPuff int32
	RgbPeru int32
	RgbPink int32
	RgbPlum int32
	RgbPowderBlue int32
	RgbPurple int32
	RgbRed int32
	RgbRosyBrown int32
	RgbRoyalBlue int32
	RgbSalmon int32
	RgbSandyBrown int32
	RgbSeaGreen int32
	RgbSeashell int32
	RgbSienna int32
	RgbSilver int32
	RgbSkyBlue int32
	RgbSlateBlue int32
	RgbSlateGray int32
	RgbSlateGrey int32
	RgbSnow int32
	RgbSpringGreen int32
	RgbSteelBlue int32
	RgbTan int32
	RgbTeal int32
	RgbThistle int32
	RgbTomato int32
	RgbTurquoise int32
	RgbYellow int32
	RgbYellowGreen int32
	RgbViolet int32
	RgbWheat int32
	RgbWhite int32
	RgbWhiteSmoke int32
}{
	RgbAliceBlue: 16775408,
	RgbAntiqueWhite: 14150650,
	RgbAqua: 16776960,
	RgbAquamarine: 13959039,
	RgbAzure: 16777200,
	RgbBeige: 14480885,
	RgbBisque: 12903679,
	RgbBlack: 0,
	RgbBlanchedAlmond: 13495295,
	RgbBlue: 16711680,
	RgbBlueViolet: 14822282,
	RgbBrown: 2763429,
	RgbBurlyWood: 8894686,
	RgbCadetBlue: 10526303,
	RgbChartreuse: 65407,
	RgbCoral: 5275647,
	RgbCornflowerBlue: 15570276,
	RgbCornsilk: 14481663,
	RgbCrimson: 3937500,
	RgbDarkBlue: 9109504,
	RgbDarkCyan: 9145088,
	RgbDarkGoldenrod: 755384,
	RgbDarkGreen: 25600,
	RgbDarkGray: 11119017,
	RgbDarkGrey: 11119017,
	RgbDarkKhaki: 7059389,
	RgbDarkMagenta: 9109643,
	RgbDarkOliveGreen: 3107669,
	RgbDarkOrange: 36095,
	RgbDarkOrchid: 13382297,
	RgbDarkRed: 139,
	RgbDarkSalmon: 8034025,
	RgbDarkSeaGreen: 9419919,
	RgbDarkSlateBlue: 9125192,
	RgbDarkSlateGray: 5197615,
	RgbDarkSlateGrey: 5197615,
	RgbDarkTurquoise: 13749760,
	RgbDarkViolet: 13828244,
	RgbDeepPink: 9639167,
	RgbDeepSkyBlue: 16760576,
	RgbDimGray: 6908265,
	RgbDimGrey: 6908265,
	RgbDodgerBlue: 16748574,
	RgbFireBrick: 2237106,
	RgbFloralWhite: 15792895,
	RgbForestGreen: 2263842,
	RgbFuchsia: 16711935,
	RgbGainsboro: 14474460,
	RgbGhostWhite: 16775416,
	RgbGold: 55295,
	RgbGoldenrod: 2139610,
	RgbGray: 8421504,
	RgbGreen: 32768,
	RgbGrey: 8421504,
	RgbGreenYellow: 3145645,
	RgbHoneydew: 15794160,
	RgbHotPink: 11823615,
	RgbIndianRed: 6053069,
	RgbIndigo: 8519755,
	RgbIvory: 15794175,
	RgbKhaki: 9234160,
	RgbLavender: 16443110,
	RgbLavenderBlush: 16118015,
	RgbLawnGreen: 64636,
	RgbLemonChiffon: 13499135,
	RgbLightBlue: 15128749,
	RgbLightCoral: 8421616,
	RgbLightCyan: 9145088,
	RgbLightGoldenrodYellow: 13826810,
	RgbLightGray: 13882323,
	RgbLightGreen: 9498256,
	RgbLightGrey: 13882323,
	RgbLightPink: 12695295,
	RgbLightSalmon: 8036607,
	RgbLightSeaGreen: 11186720,
	RgbLightSkyBlue: 16436871,
	RgbLightSlateGray: 10061943,
	RgbLightSlateGrey: 10061943,
	RgbLightSteelBlue: 14599344,
	RgbLightYellow: 14745599,
	RgbLime: 65280,
	RgbLimeGreen: 3329330,
	RgbLinen: 15134970,
	RgbMaroon: 128,
	RgbMediumAquamarine: 11206502,
	RgbMediumBlue: 13434880,
	RgbMediumOrchid: 13850042,
	RgbMediumPurple: 14381203,
	RgbMediumSeaGreen: 7451452,
	RgbMediumSlateBlue: 15624315,
	RgbMediumSpringGreen: 10156544,
	RgbMediumTurquoise: 13422920,
	RgbMediumVioletRed: 8721863,
	RgbMidnightBlue: 7346457,
	RgbMintCream: 16449525,
	RgbMistyRose: 14804223,
	RgbMoccasin: 11920639,
	RgbNavajoWhite: 11394815,
	RgbNavy: 8388608,
	RgbNavyBlue: 8388608,
	RgbOldLace: 15136253,
	RgbOlive: 32896,
	RgbOliveDrab: 2330219,
	RgbOrange: 42495,
	RgbOrangeRed: 17919,
	RgbOrchid: 14053594,
	RgbPaleGoldenrod: 7071982,
	RgbPaleGreen: 10025880,
	RgbPaleTurquoise: 15658671,
	RgbPaleVioletRed: 9662683,
	RgbPapayaWhip: 14020607,
	RgbPeachPuff: 12180223,
	RgbPeru: 4163021,
	RgbPink: 13353215,
	RgbPlum: 14524637,
	RgbPowderBlue: 15130800,
	RgbPurple: 8388736,
	RgbRed: 255,
	RgbRosyBrown: 9408444,
	RgbRoyalBlue: 14772545,
	RgbSalmon: 7504122,
	RgbSandyBrown: 6333684,
	RgbSeaGreen: 5737262,
	RgbSeashell: 15660543,
	RgbSienna: 2970272,
	RgbSilver: 12632256,
	RgbSkyBlue: 15453831,
	RgbSlateBlue: 13458026,
	RgbSlateGray: 9470064,
	RgbSlateGrey: 9470064,
	RgbSnow: 16448255,
	RgbSpringGreen: 8388352,
	RgbSteelBlue: 11829830,
	RgbTan: 9221330,
	RgbTeal: 8421376,
	RgbThistle: 14204888,
	RgbTomato: 4678655,
	RgbTurquoise: 13688896,
	RgbYellow: 65535,
	RgbYellowGreen: 3329434,
	RgbViolet: 15631086,
	RgbWheat: 11788021,
	RgbWhite: 16777215,
	RgbWhiteSmoke: 16119285,
}

// enum XlStdColorScale
var XlStdColorScale = struct {
	XlColorScaleRYG int32
	XlColorScaleGYR int32
	XlColorScaleBlackWhite int32
	XlColorScaleWhiteBlack int32
}{
	XlColorScaleRYG: 1,
	XlColorScaleGYR: 2,
	XlColorScaleBlackWhite: 3,
	XlColorScaleWhiteBlack: 4,
}

// enum XlConditionValueTypes
var XlConditionValueTypes = struct {
	XlConditionValueNone int32
	XlConditionValueNumber int32
	XlConditionValueLowestValue int32
	XlConditionValueHighestValue int32
	XlConditionValuePercent int32
	XlConditionValueFormula int32
	XlConditionValuePercentile int32
	XlConditionValueAutomaticMin int32
	XlConditionValueAutomaticMax int32
}{
	XlConditionValueNone: -1,
	XlConditionValueNumber: 0,
	XlConditionValueLowestValue: 1,
	XlConditionValueHighestValue: 2,
	XlConditionValuePercent: 3,
	XlConditionValueFormula: 4,
	XlConditionValuePercentile: 5,
	XlConditionValueAutomaticMin: 6,
	XlConditionValueAutomaticMax: 7,
}

// enum XlFormatFilterTypes
var XlFormatFilterTypes = struct {
	XlFilterBottom int32
	XlFilterTop int32
	XlFilterBottomPercent int32
	XlFilterTopPercent int32
}{
	XlFilterBottom: 0,
	XlFilterTop: 1,
	XlFilterBottomPercent: 2,
	XlFilterTopPercent: 3,
}

// enum XlContainsOperator
var XlContainsOperator = struct {
	XlContains int32
	XlDoesNotContain int32
	XlBeginsWith int32
	XlEndsWith int32
}{
	XlContains: 0,
	XlDoesNotContain: 1,
	XlBeginsWith: 2,
	XlEndsWith: 3,
}

// enum XlAboveBelow
var XlAboveBelow = struct {
	XlAboveAverage int32
	XlBelowAverage int32
	XlEqualAboveAverage int32
	XlEqualBelowAverage int32
	XlAboveStdDev int32
	XlBelowStdDev int32
}{
	XlAboveAverage: 0,
	XlBelowAverage: 1,
	XlEqualAboveAverage: 2,
	XlEqualBelowAverage: 3,
	XlAboveStdDev: 4,
	XlBelowStdDev: 5,
}

// enum XlLookFor
var XlLookFor = struct {
	XlLookForBlanks int32
	XlLookForErrors int32
	XlLookForFormulas int32
}{
	XlLookForBlanks: 0,
	XlLookForErrors: 1,
	XlLookForFormulas: 2,
}

// enum XlTimePeriods
var XlTimePeriods = struct {
	XlToday int32
	XlYesterday int32
	XlLast7Days int32
	XlThisWeek int32
	XlLastWeek int32
	XlLastMonth int32
	XlTomorrow int32
	XlNextWeek int32
	XlNextMonth int32
	XlThisMonth int32
}{
	XlToday: 0,
	XlYesterday: 1,
	XlLast7Days: 2,
	XlThisWeek: 3,
	XlLastWeek: 4,
	XlLastMonth: 5,
	XlTomorrow: 6,
	XlNextWeek: 7,
	XlNextMonth: 8,
	XlThisMonth: 9,
}

// enum XlDupeUnique
var XlDupeUnique = struct {
	XlUnique int32
	XlDuplicate int32
}{
	XlUnique: 0,
	XlDuplicate: 1,
}

// enum XlTopBottom
var XlTopBottom = struct {
	XlTop10Top int32
	XlTop10Bottom int32
}{
	XlTop10Top: 1,
	XlTop10Bottom: 0,
}

// enum XlIconSet
var XlIconSet = struct {
	XlCustomSet int32
	Xl3Arrows int32
	Xl3ArrowsGray int32
	Xl3Flags int32
	Xl3TrafficLights1 int32
	Xl3TrafficLights2 int32
	Xl3Signs int32
	Xl3Symbols int32
	Xl3Symbols2 int32
	Xl4Arrows int32
	Xl4ArrowsGray int32
	Xl4RedToBlack int32
	Xl4CRV int32
	Xl4TrafficLights int32
	Xl5Arrows int32
	Xl5ArrowsGray int32
	Xl5CRV int32
	Xl5Quarters int32
	Xl3Stars int32
	Xl3Triangles int32
	Xl5Boxes int32
}{
	XlCustomSet: -1,
	Xl3Arrows: 1,
	Xl3ArrowsGray: 2,
	Xl3Flags: 3,
	Xl3TrafficLights1: 4,
	Xl3TrafficLights2: 5,
	Xl3Signs: 6,
	Xl3Symbols: 7,
	Xl3Symbols2: 8,
	Xl4Arrows: 9,
	Xl4ArrowsGray: 10,
	Xl4RedToBlack: 11,
	Xl4CRV: 12,
	Xl4TrafficLights: 13,
	Xl5Arrows: 14,
	Xl5ArrowsGray: 15,
	Xl5CRV: 16,
	Xl5Quarters: 17,
	Xl3Stars: 18,
	Xl3Triangles: 19,
	Xl5Boxes: 20,
}

// enum XlThemeFont
var XlThemeFont = struct {
	XlThemeFontNone int32
	XlThemeFontMajor int32
	XlThemeFontMinor int32
}{
	XlThemeFontNone: 0,
	XlThemeFontMajor: 1,
	XlThemeFontMinor: 2,
}

// enum XlPivotLineType
var XlPivotLineType = struct {
	XlPivotLineRegular int32
	XlPivotLineSubtotal int32
	XlPivotLineGrandTotal int32
	XlPivotLineBlank int32
}{
	XlPivotLineRegular: 0,
	XlPivotLineSubtotal: 1,
	XlPivotLineGrandTotal: 2,
	XlPivotLineBlank: 3,
}

// enum XlCheckInVersionType
var XlCheckInVersionType = struct {
	XlCheckInMinorVersion int32
	XlCheckInMajorVersion int32
	XlCheckInOverwriteVersion int32
}{
	XlCheckInMinorVersion: 0,
	XlCheckInMajorVersion: 1,
	XlCheckInOverwriteVersion: 2,
}

// enum XlPropertyDisplayedIn
var XlPropertyDisplayedIn = struct {
	XlDisplayPropertyInPivotTable int32
	XlDisplayPropertyInTooltip int32
	XlDisplayPropertyInPivotTableAndTooltip int32
}{
	XlDisplayPropertyInPivotTable: 1,
	XlDisplayPropertyInTooltip: 2,
	XlDisplayPropertyInPivotTableAndTooltip: 3,
}

// enum XlConnectionType
var XlConnectionType = struct {
	XlConnectionTypeOLEDB int32
	XlConnectionTypeODBC int32
	XlConnectionTypeXMLMAP int32
	XlConnectionTypeTEXT int32
	XlConnectionTypeWEB int32
}{
	XlConnectionTypeOLEDB: 1,
	XlConnectionTypeODBC: 2,
	XlConnectionTypeXMLMAP: 3,
	XlConnectionTypeTEXT: 4,
	XlConnectionTypeWEB: 5,
}

// enum XlActionType
var XlActionType = struct {
	XlActionTypeUrl int32
	XlActionTypeRowset int32
	XlActionTypeReport int32
	XlActionTypeDrillthrough int32
}{
	XlActionTypeUrl: 1,
	XlActionTypeRowset: 16,
	XlActionTypeReport: 128,
	XlActionTypeDrillthrough: 256,
}

// enum XlLayoutRowType
var XlLayoutRowType = struct {
	XlCompactRow int32
	XlTabularRow int32
	XlOutlineRow int32
}{
	XlCompactRow: 0,
	XlTabularRow: 1,
	XlOutlineRow: 2,
}

// enum XlMeasurementUnits
var XlMeasurementUnits = struct {
	XlInches int32
	XlCentimeters int32
	XlMillimeters int32
}{
	XlInches: 0,
	XlCentimeters: 1,
	XlMillimeters: 2,
}

// enum XlPivotFilterType
var XlPivotFilterType = struct {
	XlTopCount int32
	XlBottomCount int32
	XlTopPercent int32
	XlBottomPercent int32
	XlTopSum int32
	XlBottomSum int32
	XlValueEquals int32
	XlValueDoesNotEqual int32
	XlValueIsGreaterThan int32
	XlValueIsGreaterThanOrEqualTo int32
	XlValueIsLessThan int32
	XlValueIsLessThanOrEqualTo int32
	XlValueIsBetween int32
	XlValueIsNotBetween int32
	XlCaptionEquals int32
	XlCaptionDoesNotEqual int32
	XlCaptionBeginsWith int32
	XlCaptionDoesNotBeginWith int32
	XlCaptionEndsWith int32
	XlCaptionDoesNotEndWith int32
	XlCaptionContains int32
	XlCaptionDoesNotContain int32
	XlCaptionIsGreaterThan int32
	XlCaptionIsGreaterThanOrEqualTo int32
	XlCaptionIsLessThan int32
	XlCaptionIsLessThanOrEqualTo int32
	XlCaptionIsBetween int32
	XlCaptionIsNotBetween int32
	XlSpecificDate int32
	XlNotSpecificDate int32
	XlBefore int32
	XlBeforeOrEqualTo int32
	XlAfter int32
	XlAfterOrEqualTo int32
	XlDateBetween int32
	XlDateNotBetween int32
	XlDateTomorrow int32
	XlDateToday int32
	XlDateYesterday int32
	XlDateNextWeek int32
	XlDateThisWeek int32
	XlDateLastWeek int32
	XlDateNextMonth int32
	XlDateThisMonth int32
	XlDateLastMonth int32
	XlDateNextQuarter int32
	XlDateThisQuarter int32
	XlDateLastQuarter int32
	XlDateNextYear int32
	XlDateThisYear int32
	XlDateLastYear int32
	XlYearToDate int32
	XlAllDatesInPeriodQuarter1 int32
	XlAllDatesInPeriodQuarter2 int32
	XlAllDatesInPeriodQuarter3 int32
	XlAllDatesInPeriodQuarter4 int32
	XlAllDatesInPeriodJanuary int32
	XlAllDatesInPeriodFebruary int32
	XlAllDatesInPeriodMarch int32
	XlAllDatesInPeriodApril int32
	XlAllDatesInPeriodMay int32
	XlAllDatesInPeriodJune int32
	XlAllDatesInPeriodJuly int32
	XlAllDatesInPeriodAugust int32
	XlAllDatesInPeriodSeptember int32
	XlAllDatesInPeriodOctober int32
	XlAllDatesInPeriodNovember int32
	XlAllDatesInPeriodDecember int32
}{
	XlTopCount: 1,
	XlBottomCount: 2,
	XlTopPercent: 3,
	XlBottomPercent: 4,
	XlTopSum: 5,
	XlBottomSum: 6,
	XlValueEquals: 7,
	XlValueDoesNotEqual: 8,
	XlValueIsGreaterThan: 9,
	XlValueIsGreaterThanOrEqualTo: 10,
	XlValueIsLessThan: 11,
	XlValueIsLessThanOrEqualTo: 12,
	XlValueIsBetween: 13,
	XlValueIsNotBetween: 14,
	XlCaptionEquals: 15,
	XlCaptionDoesNotEqual: 16,
	XlCaptionBeginsWith: 17,
	XlCaptionDoesNotBeginWith: 18,
	XlCaptionEndsWith: 19,
	XlCaptionDoesNotEndWith: 20,
	XlCaptionContains: 21,
	XlCaptionDoesNotContain: 22,
	XlCaptionIsGreaterThan: 23,
	XlCaptionIsGreaterThanOrEqualTo: 24,
	XlCaptionIsLessThan: 25,
	XlCaptionIsLessThanOrEqualTo: 26,
	XlCaptionIsBetween: 27,
	XlCaptionIsNotBetween: 28,
	XlSpecificDate: 29,
	XlNotSpecificDate: 30,
	XlBefore: 31,
	XlBeforeOrEqualTo: 32,
	XlAfter: 33,
	XlAfterOrEqualTo: 34,
	XlDateBetween: 35,
	XlDateNotBetween: 36,
	XlDateTomorrow: 37,
	XlDateToday: 38,
	XlDateYesterday: 39,
	XlDateNextWeek: 40,
	XlDateThisWeek: 41,
	XlDateLastWeek: 42,
	XlDateNextMonth: 43,
	XlDateThisMonth: 44,
	XlDateLastMonth: 45,
	XlDateNextQuarter: 46,
	XlDateThisQuarter: 47,
	XlDateLastQuarter: 48,
	XlDateNextYear: 49,
	XlDateThisYear: 50,
	XlDateLastYear: 51,
	XlYearToDate: 52,
	XlAllDatesInPeriodQuarter1: 53,
	XlAllDatesInPeriodQuarter2: 54,
	XlAllDatesInPeriodQuarter3: 55,
	XlAllDatesInPeriodQuarter4: 56,
	XlAllDatesInPeriodJanuary: 57,
	XlAllDatesInPeriodFebruary: 58,
	XlAllDatesInPeriodMarch: 59,
	XlAllDatesInPeriodApril: 60,
	XlAllDatesInPeriodMay: 61,
	XlAllDatesInPeriodJune: 62,
	XlAllDatesInPeriodJuly: 63,
	XlAllDatesInPeriodAugust: 64,
	XlAllDatesInPeriodSeptember: 65,
	XlAllDatesInPeriodOctober: 66,
	XlAllDatesInPeriodNovember: 67,
	XlAllDatesInPeriodDecember: 68,
}

// enum XlCredentialsMethod
var XlCredentialsMethod = struct {
	XlCredentialsMethodIntegrated int32
	XlCredentialsMethodNone int32
	XlCredentialsMethodStored int32
}{
	XlCredentialsMethodIntegrated: 0,
	XlCredentialsMethodNone: 1,
	XlCredentialsMethodStored: 2,
}

// enum XlCubeFieldSubType
var XlCubeFieldSubType = struct {
	XlCubeHierarchy int32
	XlCubeMeasure int32
	XlCubeSet int32
	XlCubeAttribute int32
	XlCubeCalculatedMeasure int32
	XlCubeKPIValue int32
	XlCubeKPIGoal int32
	XlCubeKPIStatus int32
	XlCubeKPITrend int32
	XlCubeKPIWeight int32
}{
	XlCubeHierarchy: 1,
	XlCubeMeasure: 2,
	XlCubeSet: 3,
	XlCubeAttribute: 4,
	XlCubeCalculatedMeasure: 5,
	XlCubeKPIValue: 6,
	XlCubeKPIGoal: 7,
	XlCubeKPIStatus: 8,
	XlCubeKPITrend: 9,
	XlCubeKPIWeight: 10,
}

// enum XlSortOn
var XlSortOn = struct {
	XlSortOnValues int32
	XlSortOnCellColor int32
	XlSortOnFontColor int32
	XlSortOnIcon int32
}{
	XlSortOnValues: 0,
	XlSortOnCellColor: 1,
	XlSortOnFontColor: 2,
	XlSortOnIcon: 3,
}

// enum XlDynamicFilterCriteria
var XlDynamicFilterCriteria = struct {
	XlFilterToday int32
	XlFilterYesterday int32
	XlFilterTomorrow int32
	XlFilterThisWeek int32
	XlFilterLastWeek int32
	XlFilterNextWeek int32
	XlFilterThisMonth int32
	XlFilterLastMonth int32
	XlFilterNextMonth int32
	XlFilterThisQuarter int32
	XlFilterLastQuarter int32
	XlFilterNextQuarter int32
	XlFilterThisYear int32
	XlFilterLastYear int32
	XlFilterNextYear int32
	XlFilterYearToDate int32
	XlFilterAllDatesInPeriodQuarter1 int32
	XlFilterAllDatesInPeriodQuarter2 int32
	XlFilterAllDatesInPeriodQuarter3 int32
	XlFilterAllDatesInPeriodQuarter4 int32
	XlFilterAllDatesInPeriodJanuary int32
	XlFilterAllDatesInPeriodFebruray int32
	XlFilterAllDatesInPeriodMarch int32
	XlFilterAllDatesInPeriodApril int32
	XlFilterAllDatesInPeriodMay int32
	XlFilterAllDatesInPeriodJune int32
	XlFilterAllDatesInPeriodJuly int32
	XlFilterAllDatesInPeriodAugust int32
	XlFilterAllDatesInPeriodSeptember int32
	XlFilterAllDatesInPeriodOctober int32
	XlFilterAllDatesInPeriodNovember int32
	XlFilterAllDatesInPeriodDecember int32
	XlFilterAboveAverage int32
	XlFilterBelowAverage int32
}{
	XlFilterToday: 1,
	XlFilterYesterday: 2,
	XlFilterTomorrow: 3,
	XlFilterThisWeek: 4,
	XlFilterLastWeek: 5,
	XlFilterNextWeek: 6,
	XlFilterThisMonth: 7,
	XlFilterLastMonth: 8,
	XlFilterNextMonth: 9,
	XlFilterThisQuarter: 10,
	XlFilterLastQuarter: 11,
	XlFilterNextQuarter: 12,
	XlFilterThisYear: 13,
	XlFilterLastYear: 14,
	XlFilterNextYear: 15,
	XlFilterYearToDate: 16,
	XlFilterAllDatesInPeriodQuarter1: 17,
	XlFilterAllDatesInPeriodQuarter2: 18,
	XlFilterAllDatesInPeriodQuarter3: 19,
	XlFilterAllDatesInPeriodQuarter4: 20,
	XlFilterAllDatesInPeriodJanuary: 21,
	XlFilterAllDatesInPeriodFebruray: 22,
	XlFilterAllDatesInPeriodMarch: 23,
	XlFilterAllDatesInPeriodApril: 24,
	XlFilterAllDatesInPeriodMay: 25,
	XlFilterAllDatesInPeriodJune: 26,
	XlFilterAllDatesInPeriodJuly: 27,
	XlFilterAllDatesInPeriodAugust: 28,
	XlFilterAllDatesInPeriodSeptember: 29,
	XlFilterAllDatesInPeriodOctober: 30,
	XlFilterAllDatesInPeriodNovember: 31,
	XlFilterAllDatesInPeriodDecember: 32,
	XlFilterAboveAverage: 33,
	XlFilterBelowAverage: 34,
}

// enum XlFilterAllDatesInPeriod
var XlFilterAllDatesInPeriod = struct {
	XlFilterAllDatesInPeriodYear int32
	XlFilterAllDatesInPeriodMonth int32
	XlFilterAllDatesInPeriodDay int32
	XlFilterAllDatesInPeriodHour int32
	XlFilterAllDatesInPeriodMinute int32
	XlFilterAllDatesInPeriodSecond int32
}{
	XlFilterAllDatesInPeriodYear: 0,
	XlFilterAllDatesInPeriodMonth: 1,
	XlFilterAllDatesInPeriodDay: 2,
	XlFilterAllDatesInPeriodHour: 3,
	XlFilterAllDatesInPeriodMinute: 4,
	XlFilterAllDatesInPeriodSecond: 5,
}

// enum XlTableStyleElementType
var XlTableStyleElementType = struct {
	XlWholeTable int32
	XlHeaderRow int32
	XlTotalRow int32
	XlGrandTotalRow int32
	XlFirstColumn int32
	XlLastColumn int32
	XlGrandTotalColumn int32
	XlRowStripe1 int32
	XlRowStripe2 int32
	XlColumnStripe1 int32
	XlColumnStripe2 int32
	XlFirstHeaderCell int32
	XlLastHeaderCell int32
	XlFirstTotalCell int32
	XlLastTotalCell int32
	XlSubtotalColumn1 int32
	XlSubtotalColumn2 int32
	XlSubtotalColumn3 int32
	XlSubtotalRow1 int32
	XlSubtotalRow2 int32
	XlSubtotalRow3 int32
	XlBlankRow int32
	XlColumnSubheading1 int32
	XlColumnSubheading2 int32
	XlColumnSubheading3 int32
	XlRowSubheading1 int32
	XlRowSubheading2 int32
	XlRowSubheading3 int32
	XlPageFieldLabels int32
	XlPageFieldValues int32
	XlSlicerUnselectedItemWithData int32
	XlSlicerUnselectedItemWithNoData int32
	XlSlicerSelectedItemWithData int32
	XlSlicerSelectedItemWithNoData int32
	XlSlicerHoveredUnselectedItemWithData int32
	XlSlicerHoveredSelectedItemWithData int32
	XlSlicerHoveredUnselectedItemWithNoData int32
	XlSlicerHoveredSelectedItemWithNoData int32
}{
	XlWholeTable: 0,
	XlHeaderRow: 1,
	XlTotalRow: 2,
	XlGrandTotalRow: 2,
	XlFirstColumn: 3,
	XlLastColumn: 4,
	XlGrandTotalColumn: 4,
	XlRowStripe1: 5,
	XlRowStripe2: 6,
	XlColumnStripe1: 7,
	XlColumnStripe2: 8,
	XlFirstHeaderCell: 9,
	XlLastHeaderCell: 10,
	XlFirstTotalCell: 11,
	XlLastTotalCell: 12,
	XlSubtotalColumn1: 13,
	XlSubtotalColumn2: 14,
	XlSubtotalColumn3: 15,
	XlSubtotalRow1: 16,
	XlSubtotalRow2: 17,
	XlSubtotalRow3: 18,
	XlBlankRow: 19,
	XlColumnSubheading1: 20,
	XlColumnSubheading2: 21,
	XlColumnSubheading3: 22,
	XlRowSubheading1: 23,
	XlRowSubheading2: 24,
	XlRowSubheading3: 25,
	XlPageFieldLabels: 26,
	XlPageFieldValues: 27,
	XlSlicerUnselectedItemWithData: 28,
	XlSlicerUnselectedItemWithNoData: 29,
	XlSlicerSelectedItemWithData: 30,
	XlSlicerSelectedItemWithNoData: 31,
	XlSlicerHoveredUnselectedItemWithData: 32,
	XlSlicerHoveredSelectedItemWithData: 33,
	XlSlicerHoveredUnselectedItemWithNoData: 34,
	XlSlicerHoveredSelectedItemWithNoData: 35,
}

// enum XlPivotConditionScope
var XlPivotConditionScope = struct {
	XlSelectionScope int32
	XlFieldsScope int32
	XlDataFieldScope int32
}{
	XlSelectionScope: 0,
	XlFieldsScope: 1,
	XlDataFieldScope: 2,
}

// enum XlCalcFor
var XlCalcFor = struct {
	XlAllValues int32
	XlRowGroups int32
	XlColGroups int32
}{
	XlAllValues: 0,
	XlRowGroups: 1,
	XlColGroups: 2,
}

// enum XlThemeColor
var XlThemeColor = struct {
	XlThemeColorDark1 int32
	XlThemeColorLight1 int32
	XlThemeColorDark2 int32
	XlThemeColorLight2 int32
	XlThemeColorAccent1 int32
	XlThemeColorAccent2 int32
	XlThemeColorAccent3 int32
	XlThemeColorAccent4 int32
	XlThemeColorAccent5 int32
	XlThemeColorAccent6 int32
	XlThemeColorHyperlink int32
	XlThemeColorFollowedHyperlink int32
}{
	XlThemeColorDark1: 1,
	XlThemeColorLight1: 2,
	XlThemeColorDark2: 3,
	XlThemeColorLight2: 4,
	XlThemeColorAccent1: 5,
	XlThemeColorAccent2: 6,
	XlThemeColorAccent3: 7,
	XlThemeColorAccent4: 8,
	XlThemeColorAccent5: 9,
	XlThemeColorAccent6: 10,
	XlThemeColorHyperlink: 11,
	XlThemeColorFollowedHyperlink: 12,
}

// enum XlFixedFormatType
var XlFixedFormatType = struct {
	XlTypePDF int32
	XlTypeXPS int32
}{
	XlTypePDF: 0,
	XlTypeXPS: 1,
}

// enum XlFixedFormatQuality
var XlFixedFormatQuality = struct {
	XlQualityStandard int32
	XlQualityMinimum int32
}{
	XlQualityStandard: 0,
	XlQualityMinimum: 1,
}

// enum XlChartElementPosition
var XlChartElementPosition = struct {
	XlChartElementPositionAutomatic int32
	XlChartElementPositionCustom int32
}{
	XlChartElementPositionAutomatic: -4105,
	XlChartElementPositionCustom: -4114,
}

// enum XlGenerateTableRefs
var XlGenerateTableRefs = struct {
	XlGenerateTableRefA1 int32
	XlGenerateTableRefStruct int32
}{
	XlGenerateTableRefA1: 0,
	XlGenerateTableRefStruct: 1,
}

// enum XlGradientFillType
var XlGradientFillType = struct {
	XlGradientFillLinear int32
	XlGradientFillPath int32
}{
	XlGradientFillLinear: 0,
	XlGradientFillPath: 1,
}

// enum XlThreadMode
var XlThreadMode = struct {
	XlThreadModeAutomatic int32
	XlThreadModeManual int32
}{
	XlThreadModeAutomatic: 0,
	XlThreadModeManual: 1,
}

// enum XlOartHorizontalOverflow
var XlOartHorizontalOverflow = struct {
	XlOartHorizontalOverflowOverflow int32
	XlOartHorizontalOverflowClip int32
}{
	XlOartHorizontalOverflowOverflow: 0,
	XlOartHorizontalOverflowClip: 1,
}

// enum XlOartVerticalOverflow
var XlOartVerticalOverflow = struct {
	XlOartVerticalOverflowOverflow int32
	XlOartVerticalOverflowClip int32
	XlOartVerticalOverflowEllipsis int32
}{
	XlOartVerticalOverflowOverflow: 0,
	XlOartVerticalOverflowClip: 1,
	XlOartVerticalOverflowEllipsis: 2,
}

// enum XlSparkScale
var XlSparkScale = struct {
	XlSparkScaleGroup int32
	XlSparkScaleSingle int32
	XlSparkScaleCustom int32
}{
	XlSparkScaleGroup: 1,
	XlSparkScaleSingle: 2,
	XlSparkScaleCustom: 3,
}

// enum XlSparkType
var XlSparkType = struct {
	XlSparkLine int32
	XlSparkColumn int32
	XlSparkColumnStacked100 int32
}{
	XlSparkLine: 1,
	XlSparkColumn: 2,
	XlSparkColumnStacked100: 3,
}

// enum XlSparklineRowCol
var XlSparklineRowCol = struct {
	XlSparklineNonSquare int32
	XlSparklineRowsSquare int32
	XlSparklineColumnsSquare int32
}{
	XlSparklineNonSquare: 0,
	XlSparklineRowsSquare: 1,
	XlSparklineColumnsSquare: 2,
}

// enum XlDataBarFillType
var XlDataBarFillType = struct {
	XlDataBarFillSolid int32
	XlDataBarFillGradient int32
}{
	XlDataBarFillSolid: 0,
	XlDataBarFillGradient: 1,
}

// enum XlDataBarBorderType
var XlDataBarBorderType = struct {
	XlDataBarBorderNone int32
	XlDataBarBorderSolid int32
}{
	XlDataBarBorderNone: 0,
	XlDataBarBorderSolid: 1,
}

// enum XlDataBarAxisPosition
var XlDataBarAxisPosition = struct {
	XlDataBarAxisAutomatic int32
	XlDataBarAxisMidpoint int32
	XlDataBarAxisNone int32
}{
	XlDataBarAxisAutomatic: 0,
	XlDataBarAxisMidpoint: 1,
	XlDataBarAxisNone: 2,
}

// enum XlDataBarNegativeColorType
var XlDataBarNegativeColorType = struct {
	XlDataBarColor int32
	XlDataBarSameAsPositive int32
}{
	XlDataBarColor: 0,
	XlDataBarSameAsPositive: 1,
}

// enum XlAllocation
var XlAllocation = struct {
	XlManualAllocation int32
	XlAutomaticAllocation int32
}{
	XlManualAllocation: 1,
	XlAutomaticAllocation: 2,
}

// enum XlAllocationValue
var XlAllocationValue = struct {
	XlAllocateValue int32
	XlAllocateIncrement int32
}{
	XlAllocateValue: 1,
	XlAllocateIncrement: 2,
}

// enum XlAllocationMethod
var XlAllocationMethod = struct {
	XlEqualAllocation int32
	XlWeightedAllocation int32
}{
	XlEqualAllocation: 1,
	XlWeightedAllocation: 2,
}

// enum XlCellChangedState
var XlCellChangedState = struct {
	XlCellNotChanged int32
	XlCellChanged int32
	XlCellChangeApplied int32
}{
	XlCellNotChanged: 1,
	XlCellChanged: 2,
	XlCellChangeApplied: 3,
}

// enum XlPivotFieldRepeatLabels
var XlPivotFieldRepeatLabels = struct {
	XlDoNotRepeatLabels int32
	XlRepeatLabels int32
}{
	XlDoNotRepeatLabels: 1,
	XlRepeatLabels: 2,
}

// enum XlPieSliceIndex
var XlPieSliceIndex = struct {
	XlOuterCounterClockwisePoint int32
	XlOuterCenterPoint int32
	XlOuterClockwisePoint int32
	XlMidClockwiseRadiusPoint int32
	XlCenterPoint int32
	XlMidCounterClockwiseRadiusPoint int32
	XlInnerClockwisePoint int32
	XlInnerCenterPoint int32
	XlInnerCounterClockwisePoint int32
}{
	XlOuterCounterClockwisePoint: 1,
	XlOuterCenterPoint: 2,
	XlOuterClockwisePoint: 3,
	XlMidClockwiseRadiusPoint: 4,
	XlCenterPoint: 5,
	XlMidCounterClockwiseRadiusPoint: 6,
	XlInnerClockwisePoint: 7,
	XlInnerCenterPoint: 8,
	XlInnerCounterClockwisePoint: 9,
}

// enum XlSpanishModes
var XlSpanishModes = struct {
	XlSpanishTuteoOnly int32
	XlSpanishTuteoAndVoseo int32
	XlSpanishVoseoOnly int32
}{
	XlSpanishTuteoOnly: 0,
	XlSpanishTuteoAndVoseo: 1,
	XlSpanishVoseoOnly: 2,
}

// enum XlSlicerCrossFilterType
var XlSlicerCrossFilterType = struct {
	XlSlicerNoCrossFilter int32
	XlSlicerCrossFilterShowItemsWithDataAtTop int32
	XlSlicerCrossFilterShowItemsWithNoData int32
}{
	XlSlicerNoCrossFilter: 1,
	XlSlicerCrossFilterShowItemsWithDataAtTop: 2,
	XlSlicerCrossFilterShowItemsWithNoData: 3,
}

// enum XlSlicerSort
var XlSlicerSort = struct {
	XlSlicerSortDataSourceOrder int32
	XlSlicerSortAscending int32
	XlSlicerSortDescending int32
}{
	XlSlicerSortDataSourceOrder: 1,
	XlSlicerSortAscending: 2,
	XlSlicerSortDescending: 3,
}

// enum XlIcon
var XlIcon = struct {
	XlIconNoCellIcon int32
	XlIconGreenUpArrow int32
	XlIconYellowSideArrow int32
	XlIconRedDownArrow int32
	XlIconGrayUpArrow int32
	XlIconGraySideArrow int32
	XlIconGrayDownArrow int32
	XlIconGreenFlag int32
	XlIconYellowFlag int32
	XlIconRedFlag int32
	XlIconGreenCircle int32
	XlIconYellowCircle int32
	XlIconRedCircleWithBorder int32
	XlIconBlackCircleWithBorder int32
	XlIconGreenTrafficLight int32
	XlIconYellowTrafficLight int32
	XlIconRedTrafficLight int32
	XlIconYellowTriangle int32
	XlIconRedDiamond int32
	XlIconGreenCheckSymbol int32
	XlIconYellowExclamationSymbol int32
	XlIconRedCrossSymbol int32
	XlIconGreenCheck int32
	XlIconYellowExclamation int32
	XlIconRedCross int32
	XlIconYellowUpInclineArrow int32
	XlIconYellowDownInclineArrow int32
	XlIconGrayUpInclineArrow int32
	XlIconGrayDownInclineArrow int32
	XlIconRedCircle int32
	XlIconPinkCircle int32
	XlIconGrayCircle int32
	XlIconBlackCircle int32
	XlIconCircleWithOneWhiteQuarter int32
	XlIconCircleWithTwoWhiteQuarters int32
	XlIconCircleWithThreeWhiteQuarters int32
	XlIconWhiteCircleAllWhiteQuarters int32
	XlIcon0Bars int32
	XlIcon1Bar int32
	XlIcon2Bars int32
	XlIcon3Bars int32
	XlIcon4Bars int32
	XlIconGoldStar int32
	XlIconHalfGoldStar int32
	XlIconSilverStar int32
	XlIconGreenUpTriangle int32
	XlIconYellowDash int32
	XlIconRedDownTriangle int32
	XlIcon4FilledBoxes int32
	XlIcon3FilledBoxes int32
	XlIcon2FilledBoxes int32
	XlIcon1FilledBox int32
	XlIcon0FilledBoxes int32
}{
	XlIconNoCellIcon: -1,
	XlIconGreenUpArrow: 1,
	XlIconYellowSideArrow: 2,
	XlIconRedDownArrow: 3,
	XlIconGrayUpArrow: 4,
	XlIconGraySideArrow: 5,
	XlIconGrayDownArrow: 6,
	XlIconGreenFlag: 7,
	XlIconYellowFlag: 8,
	XlIconRedFlag: 9,
	XlIconGreenCircle: 10,
	XlIconYellowCircle: 11,
	XlIconRedCircleWithBorder: 12,
	XlIconBlackCircleWithBorder: 13,
	XlIconGreenTrafficLight: 14,
	XlIconYellowTrafficLight: 15,
	XlIconRedTrafficLight: 16,
	XlIconYellowTriangle: 17,
	XlIconRedDiamond: 18,
	XlIconGreenCheckSymbol: 19,
	XlIconYellowExclamationSymbol: 20,
	XlIconRedCrossSymbol: 21,
	XlIconGreenCheck: 22,
	XlIconYellowExclamation: 23,
	XlIconRedCross: 24,
	XlIconYellowUpInclineArrow: 25,
	XlIconYellowDownInclineArrow: 26,
	XlIconGrayUpInclineArrow: 27,
	XlIconGrayDownInclineArrow: 28,
	XlIconRedCircle: 29,
	XlIconPinkCircle: 30,
	XlIconGrayCircle: 31,
	XlIconBlackCircle: 32,
	XlIconCircleWithOneWhiteQuarter: 33,
	XlIconCircleWithTwoWhiteQuarters: 34,
	XlIconCircleWithThreeWhiteQuarters: 35,
	XlIconWhiteCircleAllWhiteQuarters: 36,
	XlIcon0Bars: 37,
	XlIcon1Bar: 38,
	XlIcon2Bars: 39,
	XlIcon3Bars: 40,
	XlIcon4Bars: 41,
	XlIconGoldStar: 42,
	XlIconHalfGoldStar: 43,
	XlIconSilverStar: 44,
	XlIconGreenUpTriangle: 45,
	XlIconYellowDash: 46,
	XlIconRedDownTriangle: 47,
	XlIcon4FilledBoxes: 48,
	XlIcon3FilledBoxes: 49,
	XlIcon2FilledBoxes: 50,
	XlIcon1FilledBox: 51,
	XlIcon0FilledBoxes: 52,
}

// enum XlProtectedViewCloseReason
var XlProtectedViewCloseReason = struct {
	XlProtectedViewCloseNormal int32
	XlProtectedViewCloseEdit int32
	XlProtectedViewCloseForced int32
}{
	XlProtectedViewCloseNormal: 0,
	XlProtectedViewCloseEdit: 1,
	XlProtectedViewCloseForced: 2,
}

// enum XlProtectedViewWindowState
var XlProtectedViewWindowState = struct {
	XlProtectedViewWindowNormal int32
	XlProtectedViewWindowMinimized int32
	XlProtectedViewWindowMaximized int32
}{
	XlProtectedViewWindowNormal: 0,
	XlProtectedViewWindowMinimized: 1,
	XlProtectedViewWindowMaximized: 2,
}

// enum XlFileValidationPivotMode
var XlFileValidationPivotMode = struct {
	XlFileValidationPivotDefault int32
	XlFileValidationPivotRun int32
	XlFileValidationPivotSkip int32
}{
	XlFileValidationPivotDefault: 0,
	XlFileValidationPivotRun: 1,
	XlFileValidationPivotSkip: 2,
}

// enum XlPieSliceLocation
var XlPieSliceLocation = struct {
	XlHorizontalCoordinate int32
	XlVerticalCoordinate int32
}{
	XlHorizontalCoordinate: 1,
	XlVerticalCoordinate: 2,
}

// enum XlPortugueseReform
var XlPortugueseReform = struct {
	XlPortuguesePreReform int32
	XlPortuguesePostReform int32
	XlPortugueseBoth int32
}{
	XlPortuguesePreReform: 1,
	XlPortuguesePostReform: 2,
	XlPortugueseBoth: 3,
}

