package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244A7-0000-0000-C000-000000000046
var IID_TableStyle = syscall.GUID{0x000244A7, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TableStyle struct {
	ole.OleClient
}

func NewTableStyle(pDisp *win32.IDispatch, addRef bool, scoped bool) *TableStyle {
	 if pDisp == nil {
		return nil;
	}
	p := &TableStyle{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TableStyleFromVar(v ole.Variant) *TableStyle {
	return NewTableStyle(v.IDispatch(), false, false)
}

func (this *TableStyle) IID() *syscall.GUID {
	return &IID_TableStyle
}

func (this *TableStyle) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TableStyle) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *TableStyle) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *TableStyle) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *TableStyle) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *TableStyle) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *TableStyle) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *TableStyle) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *TableStyle) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *TableStyle) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *TableStyle) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TableStyle) Default_() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TableStyle) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TableStyle) NameLocal() string {
	retVal, _ := this.PropGet(0x000003a9, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TableStyle) BuiltIn() bool {
	retVal, _ := this.PropGet(0x00000229, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TableStyle) TableStyleElements() *TableStyleElements {
	retVal, _ := this.PropGet(0x00000ab1, nil)
	return NewTableStyleElements(retVal.IDispatch(), false, true)
}

func (this *TableStyle) ShowAsAvailableTableStyle() bool {
	retVal, _ := this.PropGet(0x00000ab2, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TableStyle) SetShowAsAvailableTableStyle(rhs bool)  {
	_ = this.PropPut(0x00000ab2, []interface{}{rhs})
}

func (this *TableStyle) ShowAsAvailablePivotTableStyle() bool {
	retVal, _ := this.PropGet(0x00000ab3, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TableStyle) SetShowAsAvailablePivotTableStyle(rhs bool)  {
	_ = this.PropPut(0x00000ab3, []interface{}{rhs})
}

func (this *TableStyle) Delete()  {
	retVal, _ := this.Call(0x00000075, nil)
	_= retVal
}

var TableStyle_Duplicate_OptArgs= []string{
	"NewTableStyleName", 
}

func (this *TableStyle) Duplicate(optArgs ...interface{}) *TableStyle {
	optArgs = ole.ProcessOptArgs(TableStyle_Duplicate_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000040f, nil, optArgs...)
	return NewTableStyle(retVal.IDispatch(), false, true)
}

func (this *TableStyle) ShowAsAvailableSlicerStyle() bool {
	retVal, _ := this.PropGet(0x00000b82, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TableStyle) SetShowAsAvailableSlicerStyle(rhs bool)  {
	_ = this.PropPut(0x00000b82, []interface{}{rhs})
}

