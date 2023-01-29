package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 0002084C-0000-0000-C000-000000000046
var IID_DisplayUnitLabel = syscall.GUID{0x0002084C, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type DisplayUnitLabel struct {
	ole.OleClient
}

func NewDisplayUnitLabel(pDisp *win32.IDispatch, addRef bool, scoped bool) *DisplayUnitLabel {
	if pDisp == nil {
		return nil
	}
	p := &DisplayUnitLabel{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DisplayUnitLabelFromVar(v ole.Variant) *DisplayUnitLabel {
	return NewDisplayUnitLabel(v.IDispatch(), false, false)
}

func (this *DisplayUnitLabel) IID() *syscall.GUID {
	return &IID_DisplayUnitLabel
}

func (this *DisplayUnitLabel) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DisplayUnitLabel) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *DisplayUnitLabel) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *DisplayUnitLabel) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *DisplayUnitLabel) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *DisplayUnitLabel) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *DisplayUnitLabel) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *DisplayUnitLabel) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *DisplayUnitLabel) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *DisplayUnitLabel) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *DisplayUnitLabel) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DisplayUnitLabel) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DisplayUnitLabel) Select() ole.Variant {
	retVal, _ := this.Call(0x000000eb, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayUnitLabel) Border() *Border {
	retVal, _ := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.IDispatch(), false, true)
}

func (this *DisplayUnitLabel) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayUnitLabel) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *DisplayUnitLabel) Fill() *ChartFillFormat {
	retVal, _ := this.PropGet(0x0000067f, nil)
	return NewChartFillFormat(retVal.IDispatch(), false, true)
}

func (this *DisplayUnitLabel) Caption() string {
	retVal, _ := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DisplayUnitLabel) SetCaption(rhs string) {
	_ = this.PropPut(0x0000008b, []interface{}{rhs})
}

var DisplayUnitLabel_Characters_OptArgs = []string{
	"Start", "Length",
}

func (this *DisplayUnitLabel) Characters(optArgs ...interface{}) *Characters {
	optArgs = ole.ProcessOptArgs(DisplayUnitLabel_Characters_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000025b, nil, optArgs...)
	return NewCharacters(retVal.IDispatch(), false, true)
}

func (this *DisplayUnitLabel) Font() *Font {
	retVal, _ := this.PropGet(0x00000092, nil)
	return NewFont(retVal.IDispatch(), false, true)
}

func (this *DisplayUnitLabel) HorizontalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x00000088, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayUnitLabel) SetHorizontalAlignment(rhs interface{}) {
	_ = this.PropPut(0x00000088, []interface{}{rhs})
}

func (this *DisplayUnitLabel) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *DisplayUnitLabel) SetLeft(rhs float64) {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *DisplayUnitLabel) Orientation() ole.Variant {
	retVal, _ := this.PropGet(0x00000086, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayUnitLabel) SetOrientation(rhs interface{}) {
	_ = this.PropPut(0x00000086, []interface{}{rhs})
}

func (this *DisplayUnitLabel) Shadow() bool {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DisplayUnitLabel) SetShadow(rhs bool) {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *DisplayUnitLabel) Text() string {
	retVal, _ := this.PropGet(0x0000008a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DisplayUnitLabel) SetText(rhs string) {
	_ = this.PropPut(0x0000008a, []interface{}{rhs})
}

func (this *DisplayUnitLabel) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *DisplayUnitLabel) SetTop(rhs float64) {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *DisplayUnitLabel) VerticalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x00000089, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayUnitLabel) SetVerticalAlignment(rhs interface{}) {
	_ = this.PropPut(0x00000089, []interface{}{rhs})
}

func (this *DisplayUnitLabel) ReadingOrder() int32 {
	retVal, _ := this.PropGet(0x000003cf, nil)
	return retVal.LValVal()
}

func (this *DisplayUnitLabel) SetReadingOrder(rhs int32) {
	_ = this.PropPut(0x000003cf, []interface{}{rhs})
}

func (this *DisplayUnitLabel) AutoScaleFont() ole.Variant {
	retVal, _ := this.PropGet(0x000005f5, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayUnitLabel) SetAutoScaleFont(rhs interface{}) {
	_ = this.PropPut(0x000005f5, []interface{}{rhs})
}

func (this *DisplayUnitLabel) Dummy21_() {
	retVal, _ := this.Call(0x00010015, nil)
	_ = retVal
}

func (this *DisplayUnitLabel) Position() int32 {
	retVal, _ := this.PropGet(0x00000085, nil)
	return retVal.LValVal()
}

func (this *DisplayUnitLabel) SetPosition(rhs int32) {
	_ = this.PropPut(0x00000085, []interface{}{rhs})
}

func (this *DisplayUnitLabel) Format() *ChartFormat {
	retVal, _ := this.PropGet(0x00000074, nil)
	return NewChartFormat(retVal.IDispatch(), false, true)
}

func (this *DisplayUnitLabel) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *DisplayUnitLabel) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *DisplayUnitLabel) Formula() string {
	retVal, _ := this.PropGet(0x00000105, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DisplayUnitLabel) SetFormula(rhs string) {
	_ = this.PropPut(0x00000105, []interface{}{rhs})
}

func (this *DisplayUnitLabel) FormulaR1C1() string {
	retVal, _ := this.PropGet(0x00000108, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DisplayUnitLabel) SetFormulaR1C1(rhs string) {
	_ = this.PropPut(0x00000108, []interface{}{rhs})
}

func (this *DisplayUnitLabel) FormulaLocal() string {
	retVal, _ := this.PropGet(0x00000107, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DisplayUnitLabel) SetFormulaLocal(rhs string) {
	_ = this.PropPut(0x00000107, []interface{}{rhs})
}

func (this *DisplayUnitLabel) FormulaR1C1Local() string {
	retVal, _ := this.PropGet(0x00000109, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DisplayUnitLabel) SetFormulaR1C1Local(rhs string) {
	_ = this.PropPut(0x00000109, []interface{}{rhs})
}
