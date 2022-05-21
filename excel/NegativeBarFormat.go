package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244BF-0000-0000-C000-000000000046
var IID_NegativeBarFormat = syscall.GUID{0x000244BF, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type NegativeBarFormat struct {
	ole.OleClient
}

func NewNegativeBarFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *NegativeBarFormat {
	p := &NegativeBarFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func NegativeBarFormatFromVar(v ole.Variant) *NegativeBarFormat {
	return NewNegativeBarFormat(v.PdispValVal(), false, false)
}

func (this *NegativeBarFormat) IID() *syscall.GUID {
	return &IID_NegativeBarFormat
}

func (this *NegativeBarFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *NegativeBarFormat) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *NegativeBarFormat) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *NegativeBarFormat) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *NegativeBarFormat) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *NegativeBarFormat) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *NegativeBarFormat) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *NegativeBarFormat) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *NegativeBarFormat) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *NegativeBarFormat) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *NegativeBarFormat) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *NegativeBarFormat) ColorType() int32 {
	retVal := this.PropGet(0x00000893, nil)
	return retVal.LValVal()
}

func (this *NegativeBarFormat) SetColorType(rhs int32)  {
	retVal := this.PropPut(0x00000893, []interface{}{rhs})
	_= retVal
}

func (this *NegativeBarFormat) BorderColorType() int32 {
	retVal := this.PropGet(0x00000b99, nil)
	return retVal.LValVal()
}

func (this *NegativeBarFormat) SetBorderColorType(rhs int32)  {
	retVal := this.PropPut(0x00000b99, []interface{}{rhs})
	_= retVal
}

func (this *NegativeBarFormat) Color() *ole.DispatchClass {
	retVal := this.PropGet(0x00000063, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *NegativeBarFormat) BorderColor() *ole.DispatchClass {
	retVal := this.PropGet(0x00000b9a, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

