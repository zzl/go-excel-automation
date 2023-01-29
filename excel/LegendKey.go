package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 000208BC-0000-0000-C000-000000000046
var IID_LegendKey = syscall.GUID{0x000208BC, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type LegendKey struct {
	ole.OleClient
}

func NewLegendKey(pDisp *win32.IDispatch, addRef bool, scoped bool) *LegendKey {
	if pDisp == nil {
		return nil
	}
	p := &LegendKey{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func LegendKeyFromVar(v ole.Variant) *LegendKey {
	return NewLegendKey(v.IDispatch(), false, false)
}

func (this *LegendKey) IID() *syscall.GUID {
	return &IID_LegendKey
}

func (this *LegendKey) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *LegendKey) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *LegendKey) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *LegendKey) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *LegendKey) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *LegendKey) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *LegendKey) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *LegendKey) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *LegendKey) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *LegendKey) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *LegendKey) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *LegendKey) Border() *Border {
	retVal, _ := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.IDispatch(), false, true)
}

func (this *LegendKey) ClearFormats() ole.Variant {
	retVal, _ := this.Call(0x00000070, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *LegendKey) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *LegendKey) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *LegendKey) Fill() *ChartFillFormat {
	retVal, _ := this.PropGet(0x0000067f, nil)
	return NewChartFillFormat(retVal.IDispatch(), false, true)
}

func (this *LegendKey) InvertIfNegative() bool {
	retVal, _ := this.PropGet(0x00000084, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *LegendKey) SetInvertIfNegative(rhs bool) {
	_ = this.PropPut(0x00000084, []interface{}{rhs})
}

func (this *LegendKey) MarkerBackgroundColor() int32 {
	retVal, _ := this.PropGet(0x00000049, nil)
	return retVal.LValVal()
}

func (this *LegendKey) SetMarkerBackgroundColor(rhs int32) {
	_ = this.PropPut(0x00000049, []interface{}{rhs})
}

func (this *LegendKey) MarkerBackgroundColorIndex() int32 {
	retVal, _ := this.PropGet(0x0000004a, nil)
	return retVal.LValVal()
}

func (this *LegendKey) SetMarkerBackgroundColorIndex(rhs int32) {
	_ = this.PropPut(0x0000004a, []interface{}{rhs})
}

func (this *LegendKey) MarkerForegroundColor() int32 {
	retVal, _ := this.PropGet(0x0000004b, nil)
	return retVal.LValVal()
}

func (this *LegendKey) SetMarkerForegroundColor(rhs int32) {
	_ = this.PropPut(0x0000004b, []interface{}{rhs})
}

func (this *LegendKey) MarkerForegroundColorIndex() int32 {
	retVal, _ := this.PropGet(0x0000004c, nil)
	return retVal.LValVal()
}

func (this *LegendKey) SetMarkerForegroundColorIndex(rhs int32) {
	_ = this.PropPut(0x0000004c, []interface{}{rhs})
}

func (this *LegendKey) MarkerSize() int32 {
	retVal, _ := this.PropGet(0x000000e7, nil)
	return retVal.LValVal()
}

func (this *LegendKey) SetMarkerSize(rhs int32) {
	_ = this.PropPut(0x000000e7, []interface{}{rhs})
}

func (this *LegendKey) MarkerStyle() int32 {
	retVal, _ := this.PropGet(0x00000048, nil)
	return retVal.LValVal()
}

func (this *LegendKey) SetMarkerStyle(rhs int32) {
	_ = this.PropPut(0x00000048, []interface{}{rhs})
}

func (this *LegendKey) PictureType() int32 {
	retVal, _ := this.PropGet(0x000000a1, nil)
	return retVal.LValVal()
}

func (this *LegendKey) SetPictureType(rhs int32) {
	_ = this.PropPut(0x000000a1, []interface{}{rhs})
}

func (this *LegendKey) PictureUnit() int32 {
	retVal, _ := this.PropGet(0x000000a2, nil)
	return retVal.LValVal()
}

func (this *LegendKey) SetPictureUnit(rhs int32) {
	_ = this.PropPut(0x000000a2, []interface{}{rhs})
}

func (this *LegendKey) Select() ole.Variant {
	retVal, _ := this.Call(0x000000eb, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *LegendKey) Smooth() bool {
	retVal, _ := this.PropGet(0x000000a3, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *LegendKey) SetSmooth(rhs bool) {
	_ = this.PropPut(0x000000a3, []interface{}{rhs})
}

func (this *LegendKey) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *LegendKey) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *LegendKey) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *LegendKey) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *LegendKey) Shadow() bool {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *LegendKey) SetShadow(rhs bool) {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *LegendKey) PictureUnit2() float64 {
	retVal, _ := this.PropGet(0x00000a59, nil)
	return retVal.DblValVal()
}

func (this *LegendKey) SetPictureUnit2(rhs float64) {
	_ = this.PropPut(0x00000a59, []interface{}{rhs})
}

func (this *LegendKey) Format() *ChartFormat {
	retVal, _ := this.PropGet(0x00000074, nil)
	return NewChartFormat(retVal.IDispatch(), false, true)
}
