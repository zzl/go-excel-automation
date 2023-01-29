package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 0002086B-0000-0000-C000-000000000046
var IID_Series = syscall.GUID{0x0002086B, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Series struct {
	ole.OleClient
}

func NewSeries(pDisp *win32.IDispatch, addRef bool, scoped bool) *Series {
	if pDisp == nil {
		return nil
	}
	p := &Series{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SeriesFromVar(v ole.Variant) *Series {
	return NewSeries(v.IDispatch(), false, false)
}

func (this *Series) IID() *syscall.GUID {
	return &IID_Series
}

func (this *Series) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Series) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Series) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Series) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Series) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Series) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Series) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Series) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Series) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Series) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Series) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Series_ApplyDataLabels__OptArgs = []string{
	"Type", "LegendKey", "AutoText", "HasLeaderLines",
}

func (this *Series) ApplyDataLabels_(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Series_ApplyDataLabels__OptArgs, optArgs)
	retVal, _ := this.Call(0x00000097, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Series) AxisGroup() int32 {
	retVal, _ := this.PropGet(0x0000002f, nil)
	return retVal.LValVal()
}

func (this *Series) SetAxisGroup(rhs int32) {
	_ = this.PropPut(0x0000002f, []interface{}{rhs})
}

func (this *Series) Border() *Border {
	retVal, _ := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.IDispatch(), false, true)
}

func (this *Series) ClearFormats() ole.Variant {
	retVal, _ := this.Call(0x00000070, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Series) Copy() ole.Variant {
	retVal, _ := this.Call(0x00000227, nil)
	com.AddToScope(retVal)
	return *retVal
}

var Series_DataLabels_OptArgs = []string{
	"Index",
}

func (this *Series) DataLabels(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Series_DataLabels_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000009d, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Series) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

var Series_ErrorBar_OptArgs = []string{
	"Amount", "MinusValues",
}

func (this *Series) ErrorBar(direction int32, include int32, type_ int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Series_ErrorBar_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000098, []interface{}{direction, include, type_}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Series) ErrorBars() *ErrorBars {
	retVal, _ := this.PropGet(0x0000009f, nil)
	return NewErrorBars(retVal.IDispatch(), false, true)
}

func (this *Series) Explosion() int32 {
	retVal, _ := this.PropGet(0x000000b6, nil)
	return retVal.LValVal()
}

func (this *Series) SetExplosion(rhs int32) {
	_ = this.PropPut(0x000000b6, []interface{}{rhs})
}

func (this *Series) Formula() string {
	retVal, _ := this.PropGet(0x00000105, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Series) SetFormula(rhs string) {
	_ = this.PropPut(0x00000105, []interface{}{rhs})
}

func (this *Series) FormulaLocal() string {
	retVal, _ := this.PropGet(0x00000107, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Series) SetFormulaLocal(rhs string) {
	_ = this.PropPut(0x00000107, []interface{}{rhs})
}

func (this *Series) FormulaR1C1() string {
	retVal, _ := this.PropGet(0x00000108, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Series) SetFormulaR1C1(rhs string) {
	_ = this.PropPut(0x00000108, []interface{}{rhs})
}

func (this *Series) FormulaR1C1Local() string {
	retVal, _ := this.PropGet(0x00000109, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Series) SetFormulaR1C1Local(rhs string) {
	_ = this.PropPut(0x00000109, []interface{}{rhs})
}

func (this *Series) HasDataLabels() bool {
	retVal, _ := this.PropGet(0x0000004e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Series) SetHasDataLabels(rhs bool) {
	_ = this.PropPut(0x0000004e, []interface{}{rhs})
}

func (this *Series) HasErrorBars() bool {
	retVal, _ := this.PropGet(0x000000a0, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Series) SetHasErrorBars(rhs bool) {
	_ = this.PropPut(0x000000a0, []interface{}{rhs})
}

func (this *Series) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *Series) Fill() *ChartFillFormat {
	retVal, _ := this.PropGet(0x0000067f, nil)
	return NewChartFillFormat(retVal.IDispatch(), false, true)
}

func (this *Series) InvertIfNegative() bool {
	retVal, _ := this.PropGet(0x00000084, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Series) SetInvertIfNegative(rhs bool) {
	_ = this.PropPut(0x00000084, []interface{}{rhs})
}

func (this *Series) MarkerBackgroundColor() int32 {
	retVal, _ := this.PropGet(0x00000049, nil)
	return retVal.LValVal()
}

func (this *Series) SetMarkerBackgroundColor(rhs int32) {
	_ = this.PropPut(0x00000049, []interface{}{rhs})
}

func (this *Series) MarkerBackgroundColorIndex() int32 {
	retVal, _ := this.PropGet(0x0000004a, nil)
	return retVal.LValVal()
}

func (this *Series) SetMarkerBackgroundColorIndex(rhs int32) {
	_ = this.PropPut(0x0000004a, []interface{}{rhs})
}

func (this *Series) MarkerForegroundColor() int32 {
	retVal, _ := this.PropGet(0x0000004b, nil)
	return retVal.LValVal()
}

func (this *Series) SetMarkerForegroundColor(rhs int32) {
	_ = this.PropPut(0x0000004b, []interface{}{rhs})
}

func (this *Series) MarkerForegroundColorIndex() int32 {
	retVal, _ := this.PropGet(0x0000004c, nil)
	return retVal.LValVal()
}

func (this *Series) SetMarkerForegroundColorIndex(rhs int32) {
	_ = this.PropPut(0x0000004c, []interface{}{rhs})
}

func (this *Series) MarkerSize() int32 {
	retVal, _ := this.PropGet(0x000000e7, nil)
	return retVal.LValVal()
}

func (this *Series) SetMarkerSize(rhs int32) {
	_ = this.PropPut(0x000000e7, []interface{}{rhs})
}

func (this *Series) MarkerStyle() int32 {
	retVal, _ := this.PropGet(0x00000048, nil)
	return retVal.LValVal()
}

func (this *Series) SetMarkerStyle(rhs int32) {
	_ = this.PropPut(0x00000048, []interface{}{rhs})
}

func (this *Series) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Series) SetName(rhs string) {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *Series) Paste() ole.Variant {
	retVal, _ := this.Call(0x000000d3, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Series) PictureType() int32 {
	retVal, _ := this.PropGet(0x000000a1, nil)
	return retVal.LValVal()
}

func (this *Series) SetPictureType(rhs int32) {
	_ = this.PropPut(0x000000a1, []interface{}{rhs})
}

func (this *Series) PictureUnit() int32 {
	retVal, _ := this.PropGet(0x000000a2, nil)
	return retVal.LValVal()
}

func (this *Series) SetPictureUnit(rhs int32) {
	_ = this.PropPut(0x000000a2, []interface{}{rhs})
}

func (this *Series) PlotOrder() int32 {
	retVal, _ := this.PropGet(0x000000e4, nil)
	return retVal.LValVal()
}

func (this *Series) SetPlotOrder(rhs int32) {
	_ = this.PropPut(0x000000e4, []interface{}{rhs})
}

var Series_Points_OptArgs = []string{
	"Index",
}

func (this *Series) Points(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Series_Points_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000046, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Series) Select() ole.Variant {
	retVal, _ := this.Call(0x000000eb, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Series) Smooth() bool {
	retVal, _ := this.PropGet(0x000000a3, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Series) SetSmooth(rhs bool) {
	_ = this.PropPut(0x000000a3, []interface{}{rhs})
}

var Series_Trendlines_OptArgs = []string{
	"Index",
}

func (this *Series) Trendlines(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Series_Trendlines_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000009a, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Series) Type() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *Series) SetType(rhs int32) {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

func (this *Series) ChartType() int32 {
	retVal, _ := this.PropGet(0x00000578, nil)
	return retVal.LValVal()
}

func (this *Series) SetChartType(rhs int32) {
	_ = this.PropPut(0x00000578, []interface{}{rhs})
}

func (this *Series) ApplyCustomType(chartType int32) {
	retVal, _ := this.Call(0x00000579, []interface{}{chartType})
	_ = retVal
}

func (this *Series) Values() ole.Variant {
	retVal, _ := this.PropGet(0x000000a4, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Series) SetValues(rhs interface{}) {
	_ = this.PropPut(0x000000a4, []interface{}{rhs})
}

func (this *Series) XValues() ole.Variant {
	retVal, _ := this.PropGet(0x00000457, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Series) SetXValues(rhs interface{}) {
	_ = this.PropPut(0x00000457, []interface{}{rhs})
}

func (this *Series) BubbleSizes() ole.Variant {
	retVal, _ := this.PropGet(0x00000680, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Series) SetBubbleSizes(rhs interface{}) {
	_ = this.PropPut(0x00000680, []interface{}{rhs})
}

func (this *Series) BarShape() int32 {
	retVal, _ := this.PropGet(0x0000057b, nil)
	return retVal.LValVal()
}

func (this *Series) SetBarShape(rhs int32) {
	_ = this.PropPut(0x0000057b, []interface{}{rhs})
}

func (this *Series) ApplyPictToSides() bool {
	retVal, _ := this.PropGet(0x0000067b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Series) SetApplyPictToSides(rhs bool) {
	_ = this.PropPut(0x0000067b, []interface{}{rhs})
}

func (this *Series) ApplyPictToFront() bool {
	retVal, _ := this.PropGet(0x0000067c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Series) SetApplyPictToFront(rhs bool) {
	_ = this.PropPut(0x0000067c, []interface{}{rhs})
}

func (this *Series) ApplyPictToEnd() bool {
	retVal, _ := this.PropGet(0x0000067d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Series) SetApplyPictToEnd(rhs bool) {
	_ = this.PropPut(0x0000067d, []interface{}{rhs})
}

func (this *Series) Has3DEffect() bool {
	retVal, _ := this.PropGet(0x00000681, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Series) SetHas3DEffect(rhs bool) {
	_ = this.PropPut(0x00000681, []interface{}{rhs})
}

func (this *Series) Shadow() bool {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Series) SetShadow(rhs bool) {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *Series) HasLeaderLines() bool {
	retVal, _ := this.PropGet(0x00000572, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Series) SetHasLeaderLines(rhs bool) {
	_ = this.PropPut(0x00000572, []interface{}{rhs})
}

func (this *Series) LeaderLines() *LeaderLines {
	retVal, _ := this.PropGet(0x00000682, nil)
	return NewLeaderLines(retVal.IDispatch(), false, true)
}

var Series_ApplyDataLabels_OptArgs = []string{
	"Type", "LegendKey", "AutoText", "HasLeaderLines",
	"ShowSeriesName", "ShowCategoryName", "ShowValue", "ShowPercentage",
	"ShowBubbleSize", "Separator",
}

func (this *Series) ApplyDataLabels(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Series_ApplyDataLabels_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000782, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Series) PictureUnit2() float64 {
	retVal, _ := this.PropGet(0x00000a59, nil)
	return retVal.DblValVal()
}

func (this *Series) SetPictureUnit2(rhs float64) {
	_ = this.PropPut(0x00000a59, []interface{}{rhs})
}

func (this *Series) Format() *ChartFormat {
	retVal, _ := this.PropGet(0x00000074, nil)
	return NewChartFormat(retVal.IDispatch(), false, true)
}

func (this *Series) PlotColorIndex() int32 {
	retVal, _ := this.PropGet(0x00000b63, nil)
	return retVal.LValVal()
}

func (this *Series) InvertColor() int32 {
	retVal, _ := this.PropGet(0x00000b64, nil)
	return retVal.LValVal()
}

func (this *Series) SetInvertColor(rhs int32) {
	_ = this.PropPut(0x00000b64, []interface{}{rhs})
}

func (this *Series) InvertColorIndex() int32 {
	retVal, _ := this.PropGet(0x00000b65, nil)
	return retVal.LValVal()
}

func (this *Series) SetInvertColorIndex(rhs int32) {
	_ = this.PropPut(0x00000b65, []interface{}{rhs})
}
