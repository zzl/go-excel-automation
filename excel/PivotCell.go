package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024458-0000-0000-C000-000000000046
var IID_PivotCell = syscall.GUID{0x00024458, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type PivotCell struct {
	ole.OleClient
}

func NewPivotCell(pDisp *win32.IDispatch, addRef bool, scoped bool) *PivotCell {
	 if pDisp == nil {
		return nil;
	}
	p := &PivotCell{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PivotCellFromVar(v ole.Variant) *PivotCell {
	return NewPivotCell(v.IDispatch(), false, false)
}

func (this *PivotCell) IID() *syscall.GUID {
	return &IID_PivotCell
}

func (this *PivotCell) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *PivotCell) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *PivotCell) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *PivotCell) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *PivotCell) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *PivotCell) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *PivotCell) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *PivotCell) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *PivotCell) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *PivotCell) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *PivotCell) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *PivotCell) PivotCellType() int32 {
	retVal, _ := this.PropGet(0x0000088d, nil)
	return retVal.LValVal()
}

func (this *PivotCell) PivotTable() *PivotTable {
	retVal, _ := this.PropGet(0x000002cc, nil)
	return NewPivotTable(retVal.IDispatch(), false, true)
}

func (this *PivotCell) DataField() *PivotField {
	retVal, _ := this.PropGet(0x0000082b, nil)
	return NewPivotField(retVal.IDispatch(), false, true)
}

func (this *PivotCell) PivotField() *PivotField {
	retVal, _ := this.PropGet(0x000002db, nil)
	return NewPivotField(retVal.IDispatch(), false, true)
}

func (this *PivotCell) PivotItem() *PivotItem {
	retVal, _ := this.PropGet(0x000002e4, nil)
	return NewPivotItem(retVal.IDispatch(), false, true)
}

func (this *PivotCell) RowItems() *PivotItemList {
	retVal, _ := this.PropGet(0x0000088e, nil)
	return NewPivotItemList(retVal.IDispatch(), false, true)
}

func (this *PivotCell) ColumnItems() *PivotItemList {
	retVal, _ := this.PropGet(0x0000088f, nil)
	return NewPivotItemList(retVal.IDispatch(), false, true)
}

func (this *PivotCell) Range() *Range {
	retVal, _ := this.PropGet(0x000000c5, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *PivotCell) Dummy18() string {
	retVal, _ := this.PropGet(0x000008f7, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotCell) CustomSubtotalFunction() int32 {
	retVal, _ := this.PropGet(0x00000891, nil)
	return retVal.LValVal()
}

func (this *PivotCell) PivotRowLine() *PivotLine {
	retVal, _ := this.PropGet(0x00000a71, nil)
	return NewPivotLine(retVal.IDispatch(), false, true)
}

func (this *PivotCell) PivotColumnLine() *PivotLine {
	retVal, _ := this.PropGet(0x00000a72, nil)
	return NewPivotLine(retVal.IDispatch(), false, true)
}

func (this *PivotCell) AllocateChange()  {
	retVal, _ := this.Call(0x00000b70, nil)
	_= retVal
}

func (this *PivotCell) DiscardChange()  {
	retVal, _ := this.Call(0x00000b71, nil)
	_= retVal
}

func (this *PivotCell) DataSourceValue() ole.Variant {
	retVal, _ := this.PropGet(0x00000b72, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PivotCell) CellChanged() int32 {
	retVal, _ := this.PropGet(0x00000b73, nil)
	return retVal.LValVal()
}

func (this *PivotCell) MDX() string {
	retVal, _ := this.PropGet(0x0000084b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

