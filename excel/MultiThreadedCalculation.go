package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244B1-0000-0000-C000-000000000046
var IID_MultiThreadedCalculation = syscall.GUID{0x000244B1, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type MultiThreadedCalculation struct {
	ole.OleClient
}

func NewMultiThreadedCalculation(pDisp *win32.IDispatch, addRef bool, scoped bool) *MultiThreadedCalculation {
	p := &MultiThreadedCalculation{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func MultiThreadedCalculationFromVar(v ole.Variant) *MultiThreadedCalculation {
	return NewMultiThreadedCalculation(v.PdispValVal(), false, false)
}

func (this *MultiThreadedCalculation) IID() *syscall.GUID {
	return &IID_MultiThreadedCalculation
}

func (this *MultiThreadedCalculation) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *MultiThreadedCalculation) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *MultiThreadedCalculation) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *MultiThreadedCalculation) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *MultiThreadedCalculation) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *MultiThreadedCalculation) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *MultiThreadedCalculation) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *MultiThreadedCalculation) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *MultiThreadedCalculation) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *MultiThreadedCalculation) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *MultiThreadedCalculation) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *MultiThreadedCalculation) Enabled() bool {
	retVal := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *MultiThreadedCalculation) SetEnabled(rhs bool)  {
	retVal := this.PropPut(0x00000258, []interface{}{rhs})
	_= retVal
}

func (this *MultiThreadedCalculation) ThreadMode() int32 {
	retVal := this.PropGet(0x00000ace, nil)
	return retVal.LValVal()
}

func (this *MultiThreadedCalculation) SetThreadMode(rhs int32)  {
	retVal := this.PropPut(0x00000ace, []interface{}{rhs})
	_= retVal
}

func (this *MultiThreadedCalculation) ThreadCount() int32 {
	retVal := this.PropGet(0x00000acf, nil)
	return retVal.LValVal()
}

func (this *MultiThreadedCalculation) SetThreadCount(rhs int32)  {
	retVal := this.PropPut(0x00000acf, []interface{}{rhs})
	_= retVal
}

