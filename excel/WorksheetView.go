package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024487-0000-0000-C000-000000000046
var IID_WorksheetView = syscall.GUID{0x00024487, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type WorksheetView struct {
	ole.OleClient
}

func NewWorksheetView(pDisp *win32.IDispatch, addRef bool, scoped bool) *WorksheetView {
	 if pDisp == nil {
		return nil;
	}
	p := &WorksheetView{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func WorksheetViewFromVar(v ole.Variant) *WorksheetView {
	return NewWorksheetView(v.IDispatch(), false, false)
}

func (this *WorksheetView) IID() *syscall.GUID {
	return &IID_WorksheetView
}

func (this *WorksheetView) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *WorksheetView) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *WorksheetView) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *WorksheetView) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *WorksheetView) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *WorksheetView) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *WorksheetView) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *WorksheetView) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *WorksheetView) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *WorksheetView) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *WorksheetView) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *WorksheetView) Sheet() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000002ef, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *WorksheetView) DisplayGridlines() bool {
	retVal, _ := this.PropGet(0x00000285, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WorksheetView) SetDisplayGridlines(rhs bool)  {
	_ = this.PropPut(0x00000285, []interface{}{rhs})
}

func (this *WorksheetView) DisplayFormulas() bool {
	retVal, _ := this.PropGet(0x00000284, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WorksheetView) SetDisplayFormulas(rhs bool)  {
	_ = this.PropPut(0x00000284, []interface{}{rhs})
}

func (this *WorksheetView) DisplayHeadings() bool {
	retVal, _ := this.PropGet(0x00000286, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WorksheetView) SetDisplayHeadings(rhs bool)  {
	_ = this.PropPut(0x00000286, []interface{}{rhs})
}

func (this *WorksheetView) DisplayOutline() bool {
	retVal, _ := this.PropGet(0x00000287, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WorksheetView) SetDisplayOutline(rhs bool)  {
	_ = this.PropPut(0x00000287, []interface{}{rhs})
}

func (this *WorksheetView) DisplayZeros() bool {
	retVal, _ := this.PropGet(0x00000289, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WorksheetView) SetDisplayZeros(rhs bool)  {
	_ = this.PropPut(0x00000289, []interface{}{rhs})
}

