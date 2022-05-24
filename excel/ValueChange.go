package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244C0-0000-0000-C000-000000000046
var IID_ValueChange = syscall.GUID{0x000244C0, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ValueChange struct {
	ole.OleClient
}

func NewValueChange(pDisp *win32.IDispatch, addRef bool, scoped bool) *ValueChange {
	 if pDisp == nil {
		return nil;
	}
	p := &ValueChange{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ValueChangeFromVar(v ole.Variant) *ValueChange {
	return NewValueChange(v.IDispatch(), false, false)
}

func (this *ValueChange) IID() *syscall.GUID {
	return &IID_ValueChange
}

func (this *ValueChange) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ValueChange) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *ValueChange) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ValueChange) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ValueChange) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *ValueChange) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *ValueChange) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *ValueChange) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *ValueChange) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ValueChange) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ValueChange) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ValueChange) Order() int32 {
	retVal, _ := this.PropGet(0x000000c0, nil)
	return retVal.LValVal()
}

func (this *ValueChange) VisibleInPivotTable() bool {
	retVal, _ := this.PropGet(0x00000b9b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ValueChange) PivotCell() *PivotCell {
	retVal, _ := this.PropGet(0x000007dd, nil)
	return NewPivotCell(retVal.IDispatch(), false, true)
}

func (this *ValueChange) Tuple() string {
	retVal, _ := this.PropGet(0x00000b9c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ValueChange) Value() float64 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.DblValVal()
}

func (this *ValueChange) AllocationValue() int32 {
	retVal, _ := this.PropGet(0x00000b3a, nil)
	return retVal.LValVal()
}

func (this *ValueChange) AllocationMethod() int32 {
	retVal, _ := this.PropGet(0x00000b3b, nil)
	return retVal.LValVal()
}

func (this *ValueChange) AllocationWeightExpression() string {
	retVal, _ := this.PropGet(0x00000b3c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ValueChange) Delete()  {
	retVal, _ := this.Call(0x00000075, nil)
	_= retVal
}

