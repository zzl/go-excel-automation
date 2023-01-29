package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 0002445B-0000-0000-C000-000000000046
var IID_ErrorCheckingOptions = syscall.GUID{0x0002445B, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ErrorCheckingOptions struct {
	ole.OleClient
}

func NewErrorCheckingOptions(pDisp *win32.IDispatch, addRef bool, scoped bool) *ErrorCheckingOptions {
	if pDisp == nil {
		return nil
	}
	p := &ErrorCheckingOptions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ErrorCheckingOptionsFromVar(v ole.Variant) *ErrorCheckingOptions {
	return NewErrorCheckingOptions(v.IDispatch(), false, false)
}

func (this *ErrorCheckingOptions) IID() *syscall.GUID {
	return &IID_ErrorCheckingOptions
}

func (this *ErrorCheckingOptions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ErrorCheckingOptions) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *ErrorCheckingOptions) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ErrorCheckingOptions) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ErrorCheckingOptions) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *ErrorCheckingOptions) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *ErrorCheckingOptions) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *ErrorCheckingOptions) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *ErrorCheckingOptions) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ErrorCheckingOptions) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ErrorCheckingOptions) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ErrorCheckingOptions) BackgroundChecking() bool {
	retVal, _ := this.PropGet(0x00000899, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ErrorCheckingOptions) SetBackgroundChecking(rhs bool) {
	_ = this.PropPut(0x00000899, []interface{}{rhs})
}

func (this *ErrorCheckingOptions) IndicatorColorIndex() int32 {
	retVal, _ := this.PropGet(0x0000089a, nil)
	return retVal.LValVal()
}

func (this *ErrorCheckingOptions) SetIndicatorColorIndex(rhs int32) {
	_ = this.PropPut(0x0000089a, []interface{}{rhs})
}

func (this *ErrorCheckingOptions) EvaluateToError() bool {
	retVal, _ := this.PropGet(0x0000089b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ErrorCheckingOptions) SetEvaluateToError(rhs bool) {
	_ = this.PropPut(0x0000089b, []interface{}{rhs})
}

func (this *ErrorCheckingOptions) TextDate() bool {
	retVal, _ := this.PropGet(0x0000089c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ErrorCheckingOptions) SetTextDate(rhs bool) {
	_ = this.PropPut(0x0000089c, []interface{}{rhs})
}

func (this *ErrorCheckingOptions) NumberAsText() bool {
	retVal, _ := this.PropGet(0x0000089d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ErrorCheckingOptions) SetNumberAsText(rhs bool) {
	_ = this.PropPut(0x0000089d, []interface{}{rhs})
}

func (this *ErrorCheckingOptions) InconsistentFormula() bool {
	retVal, _ := this.PropGet(0x0000089e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ErrorCheckingOptions) SetInconsistentFormula(rhs bool) {
	_ = this.PropPut(0x0000089e, []interface{}{rhs})
}

func (this *ErrorCheckingOptions) OmittedCells() bool {
	retVal, _ := this.PropGet(0x0000089f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ErrorCheckingOptions) SetOmittedCells(rhs bool) {
	_ = this.PropPut(0x0000089f, []interface{}{rhs})
}

func (this *ErrorCheckingOptions) UnlockedFormulaCells() bool {
	retVal, _ := this.PropGet(0x000008a0, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ErrorCheckingOptions) SetUnlockedFormulaCells(rhs bool) {
	_ = this.PropPut(0x000008a0, []interface{}{rhs})
}

func (this *ErrorCheckingOptions) EmptyCellReferences() bool {
	retVal, _ := this.PropGet(0x000008a1, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ErrorCheckingOptions) SetEmptyCellReferences(rhs bool) {
	_ = this.PropPut(0x000008a1, []interface{}{rhs})
}

func (this *ErrorCheckingOptions) ListDataValidation() bool {
	retVal, _ := this.PropGet(0x000008f8, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ErrorCheckingOptions) SetListDataValidation(rhs bool) {
	_ = this.PropPut(0x000008f8, []interface{}{rhs})
}

func (this *ErrorCheckingOptions) InconsistentTableFormula() bool {
	retVal, _ := this.PropGet(0x00000a73, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ErrorCheckingOptions) SetInconsistentTableFormula(rhs bool) {
	_ = this.PropPut(0x00000a73, []interface{}{rhs})
}
