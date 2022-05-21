package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024469-0000-0000-C000-000000000046
var IID_Tab = syscall.GUID{0x00024469, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Tab struct {
	ole.OleClient
}

func NewTab(pDisp *win32.IDispatch, addRef bool, scoped bool) *Tab {
	p := &Tab{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TabFromVar(v ole.Variant) *Tab {
	return NewTab(v.PdispValVal(), false, false)
}

func (this *Tab) IID() *syscall.GUID {
	return &IID_Tab
}

func (this *Tab) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Tab) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Tab) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Tab) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Tab) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Tab) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Tab) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Tab) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Tab) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Tab) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Tab) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Tab) Color() ole.Variant {
	retVal := this.PropGet(0x00000063, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Tab) SetColor(rhs interface{})  {
	retVal := this.PropPut(0x00000063, []interface{}{rhs})
	_= retVal
}

func (this *Tab) ColorIndex() int32 {
	retVal := this.PropGet(0x00000061, nil)
	return retVal.LValVal()
}

func (this *Tab) SetColorIndex(rhs int32)  {
	retVal := this.PropPut(0x00000061, []interface{}{rhs})
	_= retVal
}

func (this *Tab) ThemeColor() int32 {
	retVal := this.PropGet(0x0000093d, nil)
	return retVal.LValVal()
}

func (this *Tab) SetThemeColor(rhs int32)  {
	retVal := this.PropPut(0x0000093d, []interface{}{rhs})
	_= retVal
}

func (this *Tab) TintAndShade() ole.Variant {
	retVal := this.PropGet(0x0000093e, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Tab) SetTintAndShade(rhs interface{})  {
	retVal := this.PropPut(0x0000093e, []interface{}{rhs})
	_= retVal
}

