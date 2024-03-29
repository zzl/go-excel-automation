package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00020876-0000-0000-C000-000000000046
var IID_PivotItem = syscall.GUID{0x00020876, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type PivotItem struct {
	ole.OleClient
}

func NewPivotItem(pDisp *win32.IDispatch, addRef bool, scoped bool) *PivotItem {
	if pDisp == nil {
		return nil
	}
	p := &PivotItem{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PivotItemFromVar(v ole.Variant) *PivotItem {
	return NewPivotItem(v.IDispatch(), false, false)
}

func (this *PivotItem) IID() *syscall.GUID {
	return &IID_PivotItem
}

func (this *PivotItem) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *PivotItem) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *PivotItem) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *PivotItem) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *PivotItem) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *PivotItem) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *PivotItem) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *PivotItem) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *PivotItem) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *PivotItem) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *PivotItem) Parent() *PivotField {
	retVal, _ := this.PropGet(0x00000096, nil)
	return NewPivotField(retVal.IDispatch(), false, true)
}

var PivotItem_ChildItems_OptArgs = []string{
	"Index",
}

func (this *PivotItem) ChildItems(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(PivotItem_ChildItems_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x000002da, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PivotItem) DataRange() *Range {
	retVal, _ := this.PropGet(0x000002d0, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *PivotItem) Default_() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotItem) SetDefault_(rhs string) {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *PivotItem) LabelRange() *Range {
	retVal, _ := this.PropGet(0x000002cf, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *PivotItem) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotItem) SetName(rhs string) {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *PivotItem) ParentItem() *PivotItem {
	retVal, _ := this.PropGet(0x000002e5, nil)
	return NewPivotItem(retVal.IDispatch(), false, true)
}

func (this *PivotItem) ParentShowDetail() bool {
	retVal, _ := this.PropGet(0x000002e3, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotItem) Position() int32 {
	retVal, _ := this.PropGet(0x00000085, nil)
	return retVal.LValVal()
}

func (this *PivotItem) SetPosition(rhs int32) {
	_ = this.PropPut(0x00000085, []interface{}{rhs})
}

func (this *PivotItem) ShowDetail() bool {
	retVal, _ := this.PropGet(0x00000249, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotItem) SetShowDetail(rhs bool) {
	_ = this.PropPut(0x00000249, []interface{}{rhs})
}

func (this *PivotItem) SourceName() ole.Variant {
	retVal, _ := this.PropGet(0x000002d1, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PivotItem) Value() string {
	retVal, _ := this.PropGet(0x00000006, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotItem) SetValue(rhs string) {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *PivotItem) Visible() bool {
	retVal, _ := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotItem) SetVisible(rhs bool) {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

func (this *PivotItem) Delete() {
	retVal, _ := this.Call(0x00000075, nil)
	_ = retVal
}

func (this *PivotItem) IsCalculated() bool {
	retVal, _ := this.PropGet(0x000005e8, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotItem) RecordCount() int32 {
	retVal, _ := this.PropGet(0x000005c6, nil)
	return retVal.LValVal()
}

func (this *PivotItem) Formula() string {
	retVal, _ := this.PropGet(0x00000105, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotItem) SetFormula(rhs string) {
	_ = this.PropPut(0x00000105, []interface{}{rhs})
}

func (this *PivotItem) Caption() string {
	retVal, _ := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotItem) SetCaption(rhs string) {
	_ = this.PropPut(0x0000008b, []interface{}{rhs})
}

func (this *PivotItem) DrilledDown() bool {
	retVal, _ := this.PropGet(0x0000073a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotItem) SetDrilledDown(rhs bool) {
	_ = this.PropPut(0x0000073a, []interface{}{rhs})
}

func (this *PivotItem) StandardFormula() string {
	retVal, _ := this.PropGet(0x00000824, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotItem) SetStandardFormula(rhs string) {
	_ = this.PropPut(0x00000824, []interface{}{rhs})
}

func (this *PivotItem) SourceNameStandard() string {
	retVal, _ := this.PropGet(0x00000864, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotItem) DrillTo(field string) {
	retVal, _ := this.Call(0x00000a14, []interface{}{field})
	_ = retVal
}
