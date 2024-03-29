package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 0002442A-0000-0000-C000-000000000046
var IID_Parameter = syscall.GUID{0x0002442A, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Parameter struct {
	ole.OleClient
}

func NewParameter(pDisp *win32.IDispatch, addRef bool, scoped bool) *Parameter {
	if pDisp == nil {
		return nil
	}
	p := &Parameter{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ParameterFromVar(v ole.Variant) *Parameter {
	return NewParameter(v.IDispatch(), false, false)
}

func (this *Parameter) IID() *syscall.GUID {
	return &IID_Parameter
}

func (this *Parameter) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Parameter) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Parameter) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Parameter) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Parameter) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Parameter) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Parameter) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Parameter) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Parameter) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Parameter) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Parameter) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Parameter) DataType() int32 {
	retVal, _ := this.PropGet(0x000002d2, nil)
	return retVal.LValVal()
}

func (this *Parameter) SetDataType(rhs int32) {
	_ = this.PropPut(0x000002d2, []interface{}{rhs})
}

func (this *Parameter) Type() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *Parameter) PromptString() string {
	retVal, _ := this.PropGet(0x0000063f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Parameter) Value() ole.Variant {
	retVal, _ := this.PropGet(0x00000006, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Parameter) SourceRange() *Range {
	retVal, _ := this.PropGet(0x00000640, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Parameter) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Parameter) SetName(rhs string) {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *Parameter) SetParam(type_ int32, value interface{}) {
	retVal, _ := this.Call(0x00000641, []interface{}{type_, value})
	_ = retVal
}

func (this *Parameter) RefreshOnChange() bool {
	retVal, _ := this.PropGet(0x00000757, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Parameter) SetRefreshOnChange(rhs bool) {
	_ = this.PropPut(0x00000757, []interface{}{rhs})
}
