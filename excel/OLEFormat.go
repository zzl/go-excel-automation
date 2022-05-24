package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024441-0000-0000-C000-000000000046
var IID_OLEFormat = syscall.GUID{0x00024441, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type OLEFormat struct {
	ole.OleClient
}

func NewOLEFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *OLEFormat {
	 if pDisp == nil {
		return nil;
	}
	p := &OLEFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OLEFormatFromVar(v ole.Variant) *OLEFormat {
	return NewOLEFormat(v.IDispatch(), false, false)
}

func (this *OLEFormat) IID() *syscall.GUID {
	return &IID_OLEFormat
}

func (this *OLEFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OLEFormat) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *OLEFormat) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *OLEFormat) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *OLEFormat) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *OLEFormat) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *OLEFormat) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *OLEFormat) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *OLEFormat) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OLEFormat) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *OLEFormat) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OLEFormat) Activate()  {
	retVal, _ := this.Call(0x00000130, nil)
	_= retVal
}

func (this *OLEFormat) Object() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000419, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OLEFormat) ProgID() string {
	retVal, _ := this.PropGet(0x000005f3, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var OLEFormat_Verb_OptArgs= []string{
	"Verb", 
}

func (this *OLEFormat) Verb(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(OLEFormat_Verb_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000025e, nil, optArgs...)
	_= retVal
}

