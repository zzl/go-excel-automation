package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024445-0000-0000-C000-000000000046
var IID_OLEDBError = syscall.GUID{0x00024445, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type OLEDBError struct {
	ole.OleClient
}

func NewOLEDBError(pDisp *win32.IDispatch, addRef bool, scoped bool) *OLEDBError {
	 if pDisp == nil {
		return nil;
	}
	p := &OLEDBError{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OLEDBErrorFromVar(v ole.Variant) *OLEDBError {
	return NewOLEDBError(v.IDispatch(), false, false)
}

func (this *OLEDBError) IID() *syscall.GUID {
	return &IID_OLEDBError
}

func (this *OLEDBError) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OLEDBError) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *OLEDBError) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *OLEDBError) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *OLEDBError) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *OLEDBError) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *OLEDBError) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *OLEDBError) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *OLEDBError) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OLEDBError) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *OLEDBError) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OLEDBError) SqlState() string {
	retVal, _ := this.PropGet(0x00000643, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OLEDBError) ErrorString() string {
	retVal, _ := this.PropGet(0x000005d2, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OLEDBError) Native() int32 {
	retVal, _ := this.PropGet(0x00000769, nil)
	return retVal.LValVal()
}

func (this *OLEDBError) Number() int32 {
	retVal, _ := this.PropGet(0x000001c3, nil)
	return retVal.LValVal()
}

func (this *OLEDBError) Stage() int32 {
	retVal, _ := this.PropGet(0x0000076a, nil)
	return retVal.LValVal()
}

