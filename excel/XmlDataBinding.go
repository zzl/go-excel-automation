package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024478-0000-0000-C000-000000000046
var IID_XmlDataBinding = syscall.GUID{0x00024478, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type XmlDataBinding struct {
	ole.OleClient
}

func NewXmlDataBinding(pDisp *win32.IDispatch, addRef bool, scoped bool) *XmlDataBinding {
	p := &XmlDataBinding{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func XmlDataBindingFromVar(v ole.Variant) *XmlDataBinding {
	return NewXmlDataBinding(v.PdispValVal(), false, false)
}

func (this *XmlDataBinding) IID() *syscall.GUID {
	return &IID_XmlDataBinding
}

func (this *XmlDataBinding) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *XmlDataBinding) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *XmlDataBinding) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *XmlDataBinding) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *XmlDataBinding) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *XmlDataBinding) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *XmlDataBinding) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *XmlDataBinding) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *XmlDataBinding) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *XmlDataBinding) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *XmlDataBinding) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *XmlDataBinding) Default_() string {
	retVal := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XmlDataBinding) Refresh() int32 {
	retVal := this.Call(0x00000589, nil)
	return retVal.LValVal()
}

func (this *XmlDataBinding) LoadSettings(url string)  {
	retVal := this.Call(0x00000919, []interface{}{url})
	_= retVal
}

func (this *XmlDataBinding) ClearSettings()  {
	retVal := this.Call(0x0000091a, nil)
	_= retVal
}

func (this *XmlDataBinding) SourceUrl() string {
	retVal := this.PropGet(0x0000091b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

