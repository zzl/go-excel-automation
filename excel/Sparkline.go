package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244B9-0000-0000-C000-000000000046
var IID_Sparkline = syscall.GUID{0x000244B9, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Sparkline struct {
	ole.OleClient
}

func NewSparkline(pDisp *win32.IDispatch, addRef bool, scoped bool) *Sparkline {
	p := &Sparkline{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SparklineFromVar(v ole.Variant) *Sparkline {
	return NewSparkline(v.PdispValVal(), false, false)
}

func (this *Sparkline) IID() *syscall.GUID {
	return &IID_Sparkline
}

func (this *Sparkline) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Sparkline) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Sparkline) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Sparkline) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Sparkline) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Sparkline) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Sparkline) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Sparkline) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Sparkline) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Sparkline) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Sparkline) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Sparkline) Location() *Range {
	retVal := this.PropGet(0x00000575, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Sparkline) SetLocation(rhs *Range)  {
	retVal := this.PropPutRef(0x00000575, []interface{}{rhs})
	_= retVal
}

func (this *Sparkline) SourceData() string {
	retVal := this.PropGet(0x000002ae, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Sparkline) SetSourceData(rhs string)  {
	retVal := this.PropPut(0x000002ae, []interface{}{rhs})
	_= retVal
}

func (this *Sparkline) ModifyLocation(range_ *Range)  {
	retVal := this.Call(0x00000b85, []interface{}{range_})
	_= retVal
}

func (this *Sparkline) ModifySourceData(formula string)  {
	retVal := this.Call(0x00000b86, []interface{}{formula})
	_= retVal
}

