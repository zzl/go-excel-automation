package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024402-0000-0000-C000-000000000046
var IID_VPageBreak = syscall.GUID{0x00024402, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type VPageBreak struct {
	ole.OleClient
}

func NewVPageBreak(pDisp *win32.IDispatch, addRef bool, scoped bool) *VPageBreak {
	 if pDisp == nil {
		return nil;
	}
	p := &VPageBreak{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func VPageBreakFromVar(v ole.Variant) *VPageBreak {
	return NewVPageBreak(v.IDispatch(), false, false)
}

func (this *VPageBreak) IID() *syscall.GUID {
	return &IID_VPageBreak
}

func (this *VPageBreak) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *VPageBreak) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *VPageBreak) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *VPageBreak) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *VPageBreak) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *VPageBreak) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *VPageBreak) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *VPageBreak) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *VPageBreak) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *VPageBreak) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *VPageBreak) Parent() *Worksheet {
	retVal, _ := this.PropGet(0x00000096, nil)
	return NewWorksheet(retVal.IDispatch(), false, true)
}

func (this *VPageBreak) Delete()  {
	retVal, _ := this.Call(0x00000075, nil)
	_= retVal
}

func (this *VPageBreak) DragOff(direction int32, regionIndex int32)  {
	retVal, _ := this.Call(0x0000058c, []interface{}{direction, regionIndex})
	_= retVal
}

func (this *VPageBreak) Type() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *VPageBreak) SetType(rhs int32)  {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

func (this *VPageBreak) Extent() int32 {
	retVal, _ := this.PropGet(0x0000058e, nil)
	return retVal.LValVal()
}

func (this *VPageBreak) Location() *Range {
	retVal, _ := this.PropGet(0x00000575, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *VPageBreak) SetLocation(rhs *Range)  {
	_ = this.PropPutRef(0x00000575, []interface{}{rhs})
}

