package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002442F-0000-0000-C000-000000000046
var IID_Validation = syscall.GUID{0x0002442F, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Validation struct {
	ole.OleClient
}

func NewValidation(pDisp *win32.IDispatch, addRef bool, scoped bool) *Validation {
	p := &Validation{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ValidationFromVar(v ole.Variant) *Validation {
	return NewValidation(v.PdispValVal(), false, false)
}

func (this *Validation) IID() *syscall.GUID {
	return &IID_Validation
}

func (this *Validation) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Validation) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Validation) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Validation) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Validation) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Validation) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Validation) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Validation) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Validation) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Validation) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Validation) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Validation_Add_OptArgs= []string{
	"AlertStyle", "Operator", "Formula1", "Formula2", 
}

func (this *Validation) Add(type_ int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Validation_Add_OptArgs, optArgs)
	retVal := this.Call(0x000000b5, []interface{}{type_}, optArgs...)
	_= retVal
}

func (this *Validation) AlertStyle() int32 {
	retVal := this.PropGet(0x00000645, nil)
	return retVal.LValVal()
}

func (this *Validation) IgnoreBlank() bool {
	retVal := this.PropGet(0x00000646, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Validation) SetIgnoreBlank(rhs bool)  {
	retVal := this.PropPut(0x00000646, []interface{}{rhs})
	_= retVal
}

func (this *Validation) IMEMode() int32 {
	retVal := this.PropGet(0x00000647, nil)
	return retVal.LValVal()
}

func (this *Validation) SetIMEMode(rhs int32)  {
	retVal := this.PropPut(0x00000647, []interface{}{rhs})
	_= retVal
}

func (this *Validation) InCellDropdown() bool {
	retVal := this.PropGet(0x00000648, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Validation) SetInCellDropdown(rhs bool)  {
	retVal := this.PropPut(0x00000648, []interface{}{rhs})
	_= retVal
}

func (this *Validation) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

func (this *Validation) ErrorMessage() string {
	retVal := this.PropGet(0x00000649, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Validation) SetErrorMessage(rhs string)  {
	retVal := this.PropPut(0x00000649, []interface{}{rhs})
	_= retVal
}

func (this *Validation) ErrorTitle() string {
	retVal := this.PropGet(0x0000064a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Validation) SetErrorTitle(rhs string)  {
	retVal := this.PropPut(0x0000064a, []interface{}{rhs})
	_= retVal
}

func (this *Validation) InputMessage() string {
	retVal := this.PropGet(0x0000064b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Validation) SetInputMessage(rhs string)  {
	retVal := this.PropPut(0x0000064b, []interface{}{rhs})
	_= retVal
}

func (this *Validation) InputTitle() string {
	retVal := this.PropGet(0x0000064c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Validation) SetInputTitle(rhs string)  {
	retVal := this.PropPut(0x0000064c, []interface{}{rhs})
	_= retVal
}

func (this *Validation) Formula1() string {
	retVal := this.PropGet(0x0000062b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Validation) Formula2() string {
	retVal := this.PropGet(0x0000062c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var Validation_Modify_OptArgs= []string{
	"Type", "AlertStyle", "Operator", "Formula1", "Formula2", 
}

func (this *Validation) Modify(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Validation_Modify_OptArgs, optArgs)
	retVal := this.Call(0x0000062d, nil, optArgs...)
	_= retVal
}

func (this *Validation) Operator() int32 {
	retVal := this.PropGet(0x0000031d, nil)
	return retVal.LValVal()
}

func (this *Validation) ShowError() bool {
	retVal := this.PropGet(0x0000064d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Validation) SetShowError(rhs bool)  {
	retVal := this.PropPut(0x0000064d, []interface{}{rhs})
	_= retVal
}

func (this *Validation) ShowInput() bool {
	retVal := this.PropGet(0x0000064e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Validation) SetShowInput(rhs bool)  {
	retVal := this.PropPut(0x0000064e, []interface{}{rhs})
	_= retVal
}

func (this *Validation) Type() int32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *Validation) Value() bool {
	retVal := this.PropGet(0x00000006, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

