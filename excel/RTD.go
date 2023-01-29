package excel

import (
	"github.com/zzl/go-win32api/v2/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002446E-0000-0000-C000-000000000046
var IID_RTD = syscall.GUID{0x0002446E, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type RTD struct {
	ole.OleClient
}

func NewRTD(pDisp *win32.IDispatch, addRef bool, scoped bool) *RTD {
	if pDisp == nil {
		return nil
	}
	p := &RTD{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func RTDFromVar(v ole.Variant) *RTD {
	return NewRTD(v.IDispatch(), false, false)
}

func (this *RTD) IID() *syscall.GUID {
	return &IID_RTD
}

func (this *RTD) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *RTD) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *RTD) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *RTD) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *RTD) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *RTD) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *RTD) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *RTD) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *RTD) ThrottleInterval() int32 {
	retVal, _ := this.PropGet(0x000008c0, nil)
	return retVal.LValVal()
}

func (this *RTD) SetThrottleInterval(rhs int32) {
	_ = this.PropPut(0x000008c0, []interface{}{rhs})
}

func (this *RTD) RefreshData() {
	retVal, _ := this.Call(0x000008c1, nil)
	_ = retVal
}

func (this *RTD) RestartServers() {
	retVal, _ := this.Call(0x000008c2, nil)
	_ = retVal
}

