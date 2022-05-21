package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024473-0000-0000-C000-000000000046
var IID_ListColumn = syscall.GUID{0x00024473, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ListColumn struct {
	ole.OleClient
}

func NewListColumn(pDisp *win32.IDispatch, addRef bool, scoped bool) *ListColumn {
	p := &ListColumn{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ListColumnFromVar(v ole.Variant) *ListColumn {
	return NewListColumn(v.PdispValVal(), false, false)
}

func (this *ListColumn) IID() *syscall.GUID {
	return &IID_ListColumn
}

func (this *ListColumn) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ListColumn) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *ListColumn) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ListColumn) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ListColumn) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *ListColumn) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *ListColumn) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *ListColumn) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *ListColumn) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *ListColumn) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ListColumn) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ListColumn) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

func (this *ListColumn) Default_() string {
	retVal := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ListColumn) ListDataFormat() *ListDataFormat {
	retVal := this.PropGet(0x00000911, nil)
	return NewListDataFormat(retVal.PdispValVal(), false, true)
}

func (this *ListColumn) Index() int32 {
	retVal := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *ListColumn) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ListColumn) SetName(rhs string)  {
	retVal := this.PropPut(0x0000006e, []interface{}{rhs})
	_= retVal
}

func (this *ListColumn) Range() *Range {
	retVal := this.PropGet(0x000000c5, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *ListColumn) TotalsCalculation() int32 {
	retVal := this.PropGet(0x00000912, nil)
	return retVal.LValVal()
}

func (this *ListColumn) SetTotalsCalculation(rhs int32)  {
	retVal := this.PropPut(0x00000912, []interface{}{rhs})
	_= retVal
}

func (this *ListColumn) XPath() *XPath {
	retVal := this.PropGet(0x000008d2, nil)
	return NewXPath(retVal.PdispValVal(), false, true)
}

func (this *ListColumn) SharePointFormula() string {
	retVal := this.PropGet(0x00000913, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ListColumn) DataBodyRange() *Range {
	retVal := this.PropGet(0x000002c1, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *ListColumn) Total() *Range {
	retVal := this.PropGet(0x00000a79, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

