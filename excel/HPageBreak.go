package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024401-0000-0000-C000-000000000046
var IID_HPageBreak = syscall.GUID{0x00024401, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type HPageBreak struct {
	ole.OleClient
}

func NewHPageBreak(pDisp *win32.IDispatch, addRef bool, scoped bool) *HPageBreak {
	p := &HPageBreak{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func HPageBreakFromVar(v ole.Variant) *HPageBreak {
	return NewHPageBreak(v.PdispValVal(), false, false)
}

func (this *HPageBreak) IID() *syscall.GUID {
	return &IID_HPageBreak
}

func (this *HPageBreak) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *HPageBreak) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *HPageBreak) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *HPageBreak) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *HPageBreak) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *HPageBreak) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *HPageBreak) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *HPageBreak) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *HPageBreak) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *HPageBreak) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *HPageBreak) Parent() *Worksheet {
	retVal := this.PropGet(0x00000096, nil)
	return NewWorksheet(retVal.PdispValVal(), false, true)
}

func (this *HPageBreak) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

func (this *HPageBreak) DragOff(direction int32, regionIndex int32)  {
	retVal := this.Call(0x0000058c, []interface{}{direction, regionIndex})
	_= retVal
}

func (this *HPageBreak) Type() int32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *HPageBreak) SetType(rhs int32)  {
	retVal := this.PropPut(0x0000006c, []interface{}{rhs})
	_= retVal
}

func (this *HPageBreak) Extent() int32 {
	retVal := this.PropGet(0x0000058e, nil)
	return retVal.LValVal()
}

func (this *HPageBreak) Location() *Range {
	retVal := this.PropGet(0x00000575, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *HPageBreak) SetLocation(rhs *Range)  {
	retVal := this.PropPutRef(0x00000575, []interface{}{rhs})
	_= retVal
}

