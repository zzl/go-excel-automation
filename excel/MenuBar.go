package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020864-0000-0000-C000-000000000046
var IID_MenuBar = syscall.GUID{0x00020864, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type MenuBar struct {
	ole.OleClient
}

func NewMenuBar(pDisp *win32.IDispatch, addRef bool, scoped bool) *MenuBar {
	p := &MenuBar{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func MenuBarFromVar(v ole.Variant) *MenuBar {
	return NewMenuBar(v.PdispValVal(), false, false)
}

func (this *MenuBar) IID() *syscall.GUID {
	return &IID_MenuBar
}

func (this *MenuBar) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *MenuBar) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *MenuBar) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *MenuBar) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *MenuBar) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *MenuBar) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *MenuBar) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *MenuBar) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *MenuBar) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *MenuBar) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *MenuBar) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *MenuBar) Activate()  {
	retVal := this.Call(0x00000130, nil)
	_= retVal
}

func (this *MenuBar) BuiltIn() bool {
	retVal := this.PropGet(0x00000229, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *MenuBar) Caption() string {
	retVal := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *MenuBar) SetCaption(rhs string)  {
	retVal := this.PropPut(0x0000008b, []interface{}{rhs})
	_= retVal
}

func (this *MenuBar) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

func (this *MenuBar) Index() int32 {
	retVal := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *MenuBar) Menus() *Menus {
	retVal := this.PropGet(0x0000024f, nil)
	return NewMenus(retVal.PdispValVal(), false, true)
}

func (this *MenuBar) Reset()  {
	retVal := this.Call(0x0000022b, nil)
	_= retVal
}

