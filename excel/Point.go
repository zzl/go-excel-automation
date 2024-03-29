package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 0002086A-0000-0000-C000-000000000046
var IID_Point = syscall.GUID{0x0002086A, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Point struct {
	ole.OleClient
}

func NewPoint(pDisp *win32.IDispatch, addRef bool, scoped bool) *Point {
	if pDisp == nil {
		return nil
	}
	p := &Point{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PointFromVar(v ole.Variant) *Point {
	return NewPoint(v.IDispatch(), false, false)
}

func (this *Point) IID() *syscall.GUID {
	return &IID_Point
}

func (this *Point) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Point) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Point) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Point) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Point) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Point) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Point) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Point) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Point) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Point) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Point) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Point_ApplyDataLabels__OptArgs = []string{
	"Type", "LegendKey", "AutoText", "HasLeaderLines",
}

func (this *Point) ApplyDataLabels_(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Point_ApplyDataLabels__OptArgs, optArgs)
	retVal, _ := this.Call(0x00000097, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Point) Border() *Border {
	retVal, _ := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.IDispatch(), false, true)
}

func (this *Point) ClearFormats() ole.Variant {
	retVal, _ := this.Call(0x00000070, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Point) Copy() ole.Variant {
	retVal, _ := this.Call(0x00000227, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Point) DataLabel() *DataLabel {
	retVal, _ := this.PropGet(0x0000009e, nil)
	return NewDataLabel(retVal.IDispatch(), false, true)
}

func (this *Point) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Point) Explosion() int32 {
	retVal, _ := this.PropGet(0x000000b6, nil)
	return retVal.LValVal()
}

func (this *Point) SetExplosion(rhs int32) {
	_ = this.PropPut(0x000000b6, []interface{}{rhs})
}

func (this *Point) HasDataLabel() bool {
	retVal, _ := this.PropGet(0x0000004d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Point) SetHasDataLabel(rhs bool) {
	_ = this.PropPut(0x0000004d, []interface{}{rhs})
}

func (this *Point) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *Point) InvertIfNegative() bool {
	retVal, _ := this.PropGet(0x00000084, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Point) SetInvertIfNegative(rhs bool) {
	_ = this.PropPut(0x00000084, []interface{}{rhs})
}

func (this *Point) MarkerBackgroundColor() int32 {
	retVal, _ := this.PropGet(0x00000049, nil)
	return retVal.LValVal()
}

func (this *Point) SetMarkerBackgroundColor(rhs int32) {
	_ = this.PropPut(0x00000049, []interface{}{rhs})
}

func (this *Point) MarkerBackgroundColorIndex() int32 {
	retVal, _ := this.PropGet(0x0000004a, nil)
	return retVal.LValVal()
}

func (this *Point) SetMarkerBackgroundColorIndex(rhs int32) {
	_ = this.PropPut(0x0000004a, []interface{}{rhs})
}

func (this *Point) MarkerForegroundColor() int32 {
	retVal, _ := this.PropGet(0x0000004b, nil)
	return retVal.LValVal()
}

func (this *Point) SetMarkerForegroundColor(rhs int32) {
	_ = this.PropPut(0x0000004b, []interface{}{rhs})
}

func (this *Point) MarkerForegroundColorIndex() int32 {
	retVal, _ := this.PropGet(0x0000004c, nil)
	return retVal.LValVal()
}

func (this *Point) SetMarkerForegroundColorIndex(rhs int32) {
	_ = this.PropPut(0x0000004c, []interface{}{rhs})
}

func (this *Point) MarkerSize() int32 {
	retVal, _ := this.PropGet(0x000000e7, nil)
	return retVal.LValVal()
}

func (this *Point) SetMarkerSize(rhs int32) {
	_ = this.PropPut(0x000000e7, []interface{}{rhs})
}

func (this *Point) MarkerStyle() int32 {
	retVal, _ := this.PropGet(0x00000048, nil)
	return retVal.LValVal()
}

func (this *Point) SetMarkerStyle(rhs int32) {
	_ = this.PropPut(0x00000048, []interface{}{rhs})
}

func (this *Point) Paste() ole.Variant {
	retVal, _ := this.Call(0x000000d3, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Point) PictureType() int32 {
	retVal, _ := this.PropGet(0x000000a1, nil)
	return retVal.LValVal()
}

func (this *Point) SetPictureType(rhs int32) {
	_ = this.PropPut(0x000000a1, []interface{}{rhs})
}

func (this *Point) PictureUnit() int32 {
	retVal, _ := this.PropGet(0x000000a2, nil)
	return retVal.LValVal()
}

func (this *Point) SetPictureUnit(rhs int32) {
	_ = this.PropPut(0x000000a2, []interface{}{rhs})
}

func (this *Point) Select() ole.Variant {
	retVal, _ := this.Call(0x000000eb, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Point) ApplyPictToSides() bool {
	retVal, _ := this.PropGet(0x0000067b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Point) SetApplyPictToSides(rhs bool) {
	_ = this.PropPut(0x0000067b, []interface{}{rhs})
}

func (this *Point) ApplyPictToFront() bool {
	retVal, _ := this.PropGet(0x0000067c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Point) SetApplyPictToFront(rhs bool) {
	_ = this.PropPut(0x0000067c, []interface{}{rhs})
}

func (this *Point) ApplyPictToEnd() bool {
	retVal, _ := this.PropGet(0x0000067d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Point) SetApplyPictToEnd(rhs bool) {
	_ = this.PropPut(0x0000067d, []interface{}{rhs})
}

func (this *Point) Shadow() bool {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Point) SetShadow(rhs bool) {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *Point) SecondaryPlot() bool {
	retVal, _ := this.PropGet(0x0000067e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Point) SetSecondaryPlot(rhs bool) {
	_ = this.PropPut(0x0000067e, []interface{}{rhs})
}

func (this *Point) Fill() *ChartFillFormat {
	retVal, _ := this.PropGet(0x0000067f, nil)
	return NewChartFillFormat(retVal.IDispatch(), false, true)
}

var Point_ApplyDataLabels_OptArgs = []string{
	"Type", "LegendKey", "AutoText", "HasLeaderLines",
	"ShowSeriesName", "ShowCategoryName", "ShowValue", "ShowPercentage",
	"ShowBubbleSize", "Separator",
}

func (this *Point) ApplyDataLabels(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Point_ApplyDataLabels_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000782, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Point) Has3DEffect() bool {
	retVal, _ := this.PropGet(0x00000681, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Point) SetHas3DEffect(rhs bool) {
	_ = this.PropPut(0x00000681, []interface{}{rhs})
}

func (this *Point) PictureUnit2() float64 {
	retVal, _ := this.PropGet(0x00000a59, nil)
	return retVal.DblValVal()
}

func (this *Point) SetPictureUnit2(rhs float64) {
	_ = this.PropPut(0x00000a59, []interface{}{rhs})
}

func (this *Point) Format() *ChartFormat {
	retVal, _ := this.PropGet(0x00000074, nil)
	return NewChartFormat(retVal.IDispatch(), false, true)
}

func (this *Point) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *Point) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *Point) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *Point) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *Point) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var Point_PieSliceLocation_OptArgs = []string{
	"Index",
}

func (this *Point) PieSliceLocation(loc int32, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(Point_PieSliceLocation_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000b61, []interface{}{loc}, optArgs...)
	return retVal.DblValVal()
}
