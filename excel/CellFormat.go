package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00024450-0000-0000-C000-000000000046
var IID_CellFormat = syscall.GUID{0x00024450, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type CellFormat struct {
	ole.OleClient
}

func NewCellFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *CellFormat {
	if pDisp == nil {
		return nil
	}
	p := &CellFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CellFormatFromVar(v ole.Variant) *CellFormat {
	return NewCellFormat(v.IDispatch(), false, false)
}

func (this *CellFormat) IID() *syscall.GUID {
	return &IID_CellFormat
}

func (this *CellFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CellFormat) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *CellFormat) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *CellFormat) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *CellFormat) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *CellFormat) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *CellFormat) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *CellFormat) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *CellFormat) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *CellFormat) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *CellFormat) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *CellFormat) Borders() *Borders {
	retVal, _ := this.PropGet(0x000001b3, nil)
	return NewBorders(retVal.IDispatch(), false, true)
}

func (this *CellFormat) SetBorders(rhs *Borders) {
	_ = this.PropPutRef(0x000001b3, []interface{}{rhs})
}

func (this *CellFormat) Font() *Font {
	retVal, _ := this.PropGet(0x00000092, nil)
	return NewFont(retVal.IDispatch(), false, true)
}

func (this *CellFormat) SetFont(rhs *Font) {
	_ = this.PropPutRef(0x00000092, []interface{}{rhs})
}

func (this *CellFormat) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *CellFormat) SetInterior(rhs *Interior) {
	_ = this.PropPutRef(0x00000081, []interface{}{rhs})
}

func (this *CellFormat) NumberFormat() ole.Variant {
	retVal, _ := this.PropGet(0x000000c1, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *CellFormat) SetNumberFormat(rhs interface{}) {
	_ = this.PropPut(0x000000c1, []interface{}{rhs})
}

func (this *CellFormat) NumberFormatLocal() ole.Variant {
	retVal, _ := this.PropGet(0x00000449, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *CellFormat) SetNumberFormatLocal(rhs interface{}) {
	_ = this.PropPut(0x00000449, []interface{}{rhs})
}

func (this *CellFormat) AddIndent() ole.Variant {
	retVal, _ := this.PropGet(0x00000427, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *CellFormat) SetAddIndent(rhs interface{}) {
	_ = this.PropPut(0x00000427, []interface{}{rhs})
}

func (this *CellFormat) IndentLevel() ole.Variant {
	retVal, _ := this.PropGet(0x000000c9, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *CellFormat) SetIndentLevel(rhs interface{}) {
	_ = this.PropPut(0x000000c9, []interface{}{rhs})
}

func (this *CellFormat) HorizontalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x00000088, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *CellFormat) SetHorizontalAlignment(rhs interface{}) {
	_ = this.PropPut(0x00000088, []interface{}{rhs})
}

func (this *CellFormat) VerticalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x00000089, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *CellFormat) SetVerticalAlignment(rhs interface{}) {
	_ = this.PropPut(0x00000089, []interface{}{rhs})
}

func (this *CellFormat) Orientation() ole.Variant {
	retVal, _ := this.PropGet(0x00000086, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *CellFormat) SetOrientation(rhs interface{}) {
	_ = this.PropPut(0x00000086, []interface{}{rhs})
}

func (this *CellFormat) ShrinkToFit() ole.Variant {
	retVal, _ := this.PropGet(0x000000d1, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *CellFormat) SetShrinkToFit(rhs interface{}) {
	_ = this.PropPut(0x000000d1, []interface{}{rhs})
}

func (this *CellFormat) WrapText() ole.Variant {
	retVal, _ := this.PropGet(0x00000114, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *CellFormat) SetWrapText(rhs interface{}) {
	_ = this.PropPut(0x00000114, []interface{}{rhs})
}

func (this *CellFormat) Locked() ole.Variant {
	retVal, _ := this.PropGet(0x0000010d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *CellFormat) SetLocked(rhs interface{}) {
	_ = this.PropPut(0x0000010d, []interface{}{rhs})
}

func (this *CellFormat) FormulaHidden() ole.Variant {
	retVal, _ := this.PropGet(0x00000106, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *CellFormat) SetFormulaHidden(rhs interface{}) {
	_ = this.PropPut(0x00000106, []interface{}{rhs})
}

func (this *CellFormat) MergeCells() ole.Variant {
	retVal, _ := this.PropGet(0x000000d0, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *CellFormat) SetMergeCells(rhs interface{}) {
	_ = this.PropPut(0x000000d0, []interface{}{rhs})
}

func (this *CellFormat) Clear() {
	retVal, _ := this.Call(0x0000006f, nil)
	_ = retVal
}
