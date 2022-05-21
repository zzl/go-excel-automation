package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020857-0000-0000-C000-000000000046
var IID_AddIn = syscall.GUID{0x00020857, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type AddIn struct {
	ole.OleClient
}

func NewAddIn(pDisp *win32.IDispatch, addRef bool, scoped bool) *AddIn {
	p := &AddIn{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func AddInFromVar(v ole.Variant) *AddIn {
	return NewAddIn(v.PdispValVal(), false, false)
}

func (this *AddIn) IID() *syscall.GUID {
	return &IID_AddIn
}

func (this *AddIn) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *AddIn) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *AddIn) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *AddIn) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *AddIn) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *AddIn) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *AddIn) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *AddIn) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *AddIn) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *AddIn) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *AddIn) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *AddIn) Author() string {
	retVal := this.PropGet(0x0000023e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AddIn) Comments() string {
	retVal := this.PropGet(0x0000023f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AddIn) FullName() string {
	retVal := this.PropGet(0x00000121, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AddIn) Installed() bool {
	retVal := this.PropGet(0x00000226, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AddIn) SetInstalled(rhs bool)  {
	retVal := this.PropPut(0x00000226, []interface{}{rhs})
	_= retVal
}

func (this *AddIn) Keywords() string {
	retVal := this.PropGet(0x00000241, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AddIn) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AddIn) Path() string {
	retVal := this.PropGet(0x00000123, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AddIn) Subject() string {
	retVal := this.PropGet(0x000003b9, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AddIn) Title() string {
	retVal := this.PropGet(0x000000c7, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AddIn) ProgID() string {
	retVal := this.PropGet(0x000005f3, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AddIn) CLSID() string {
	retVal := this.PropGet(0x000007fb, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AddIn) IsOpen() bool {
	retVal := this.PropGet(0x00000b31, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

