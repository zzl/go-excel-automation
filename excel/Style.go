package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00020852-0000-0000-C000-000000000046
var IID_Style = syscall.GUID{0x00020852, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Style struct {
	ole.OleClient
}

func NewStyle(pDisp *win32.IDispatch, addRef bool, scoped bool) *Style {
	if pDisp == nil {
		return nil
	}
	p := &Style{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func StyleFromVar(v ole.Variant) *Style {
	return NewStyle(v.IDispatch(), false, false)
}

func (this *Style) IID() *syscall.GUID {
	return &IID_Style
}

func (this *Style) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Style) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Style) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Style) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Style) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Style) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Style) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Style) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Style) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Style) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Style) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Style) AddIndent() bool {
	retVal, _ := this.PropGet(0x00000427, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Style) SetAddIndent(rhs bool) {
	_ = this.PropPut(0x00000427, []interface{}{rhs})
}

func (this *Style) BuiltIn() bool {
	retVal, _ := this.PropGet(0x00000229, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Style) Borders() *Borders {
	retVal, _ := this.PropGet(0x000001b3, nil)
	return NewBorders(retVal.IDispatch(), false, true)
}

func (this *Style) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Style) Font() *Font {
	retVal, _ := this.PropGet(0x00000092, nil)
	return NewFont(retVal.IDispatch(), false, true)
}

func (this *Style) FormulaHidden() bool {
	retVal, _ := this.PropGet(0x00000106, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Style) SetFormulaHidden(rhs bool) {
	_ = this.PropPut(0x00000106, []interface{}{rhs})
}

func (this *Style) HorizontalAlignment() int32 {
	retVal, _ := this.PropGet(0x00000088, nil)
	return retVal.LValVal()
}

func (this *Style) SetHorizontalAlignment(rhs int32) {
	_ = this.PropPut(0x00000088, []interface{}{rhs})
}

func (this *Style) IncludeAlignment() bool {
	retVal, _ := this.PropGet(0x0000019d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Style) SetIncludeAlignment(rhs bool) {
	_ = this.PropPut(0x0000019d, []interface{}{rhs})
}

func (this *Style) IncludeBorder() bool {
	retVal, _ := this.PropGet(0x0000019e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Style) SetIncludeBorder(rhs bool) {
	_ = this.PropPut(0x0000019e, []interface{}{rhs})
}

func (this *Style) IncludeFont() bool {
	retVal, _ := this.PropGet(0x0000019f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Style) SetIncludeFont(rhs bool) {
	_ = this.PropPut(0x0000019f, []interface{}{rhs})
}

func (this *Style) IncludeNumber() bool {
	retVal, _ := this.PropGet(0x000001a0, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Style) SetIncludeNumber(rhs bool) {
	_ = this.PropPut(0x000001a0, []interface{}{rhs})
}

func (this *Style) IncludePatterns() bool {
	retVal, _ := this.PropGet(0x000001a1, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Style) SetIncludePatterns(rhs bool) {
	_ = this.PropPut(0x000001a1, []interface{}{rhs})
}

func (this *Style) IncludeProtection() bool {
	retVal, _ := this.PropGet(0x000001a2, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Style) SetIncludeProtection(rhs bool) {
	_ = this.PropPut(0x000001a2, []interface{}{rhs})
}

func (this *Style) IndentLevel() int32 {
	retVal, _ := this.PropGet(0x000000c9, nil)
	return retVal.LValVal()
}

func (this *Style) SetIndentLevel(rhs int32) {
	_ = this.PropPut(0x000000c9, []interface{}{rhs})
}

func (this *Style) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *Style) Locked() bool {
	retVal, _ := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Style) SetLocked(rhs bool) {
	_ = this.PropPut(0x0000010d, []interface{}{rhs})
}

func (this *Style) MergeCells() ole.Variant {
	retVal, _ := this.PropGet(0x000000d0, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Style) SetMergeCells(rhs interface{}) {
	_ = this.PropPut(0x000000d0, []interface{}{rhs})
}

func (this *Style) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Style) NameLocal() string {
	retVal, _ := this.PropGet(0x000003a9, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Style) NumberFormat() string {
	retVal, _ := this.PropGet(0x000000c1, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Style) SetNumberFormat(rhs string) {
	_ = this.PropPut(0x000000c1, []interface{}{rhs})
}

func (this *Style) NumberFormatLocal() string {
	retVal, _ := this.PropGet(0x00000449, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Style) SetNumberFormatLocal(rhs string) {
	_ = this.PropPut(0x00000449, []interface{}{rhs})
}

func (this *Style) Orientation() int32 {
	retVal, _ := this.PropGet(0x00000086, nil)
	return retVal.LValVal()
}

func (this *Style) SetOrientation(rhs int32) {
	_ = this.PropPut(0x00000086, []interface{}{rhs})
}

func (this *Style) ShrinkToFit() bool {
	retVal, _ := this.PropGet(0x000000d1, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Style) SetShrinkToFit(rhs bool) {
	_ = this.PropPut(0x000000d1, []interface{}{rhs})
}

func (this *Style) Value() string {
	retVal, _ := this.PropGet(0x00000006, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Style) VerticalAlignment() int32 {
	retVal, _ := this.PropGet(0x00000089, nil)
	return retVal.LValVal()
}

func (this *Style) SetVerticalAlignment(rhs int32) {
	_ = this.PropPut(0x00000089, []interface{}{rhs})
}

func (this *Style) WrapText() bool {
	retVal, _ := this.PropGet(0x00000114, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Style) SetWrapText(rhs bool) {
	_ = this.PropPut(0x00000114, []interface{}{rhs})
}

func (this *Style) Default_() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Style) ReadingOrder() int32 {
	retVal, _ := this.PropGet(0x000003cf, nil)
	return retVal.LValVal()
}

func (this *Style) SetReadingOrder(rhs int32) {
	_ = this.PropPut(0x000003cf, []interface{}{rhs})
}
