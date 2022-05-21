package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002447E-0000-0000-C000-000000000046
var IID_XPath = syscall.GUID{0x0002447E, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type XPath struct {
	ole.OleClient
}

func NewXPath(pDisp *win32.IDispatch, addRef bool, scoped bool) *XPath {
	p := &XPath{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func XPathFromVar(v ole.Variant) *XPath {
	return NewXPath(v.PdispValVal(), false, false)
}

func (this *XPath) IID() *syscall.GUID {
	return &IID_XPath
}

func (this *XPath) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *XPath) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *XPath) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *XPath) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *XPath) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *XPath) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *XPath) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *XPath) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *XPath) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *XPath) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *XPath) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *XPath) Default_() string {
	retVal := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XPath) Value() string {
	retVal := this.PropGet(0x00000006, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XPath) Map() *XmlMap {
	retVal := this.PropGet(0x000008d6, nil)
	return NewXmlMap(retVal.PdispValVal(), false, true)
}

var XPath_SetValue_OptArgs= []string{
	"SelectionNamespace", "Repeating", 
}

func (this *XPath) SetValue(map_ *XmlMap, xpath string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(XPath_SetValue_OptArgs, optArgs)
	retVal := this.Call(0x00000936, []interface{}{map_, xpath}, optArgs...)
	_= retVal
}

func (this *XPath) Clear()  {
	retVal := this.Call(0x0000006f, nil)
	_= retVal
}

func (this *XPath) Repeating() bool {
	retVal := this.PropGet(0x00000938, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

