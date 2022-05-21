package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024488-0000-0000-C000-000000000046
var IID_ChartView = syscall.GUID{0x00024488, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ChartView struct {
	ole.OleClient
}

func NewChartView(pDisp *win32.IDispatch, addRef bool, scoped bool) *ChartView {
	p := &ChartView{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ChartViewFromVar(v ole.Variant) *ChartView {
	return NewChartView(v.PdispValVal(), false, false)
}

func (this *ChartView) IID() *syscall.GUID {
	return &IID_ChartView
}

func (this *ChartView) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ChartView) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *ChartView) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ChartView) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ChartView) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *ChartView) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *ChartView) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *ChartView) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *ChartView) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *ChartView) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ChartView) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ChartView) Sheet() *ole.DispatchClass {
	retVal := this.PropGet(0x000002ef, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

