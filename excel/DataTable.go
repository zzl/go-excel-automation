package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020843-0000-0000-C000-000000000046
var IID_DataTable = syscall.GUID{0x00020843, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type DataTable struct {
	ole.OleClient
}

func NewDataTable(pDisp *win32.IDispatch, addRef bool, scoped bool) *DataTable {
	p := &DataTable{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DataTableFromVar(v ole.Variant) *DataTable {
	return NewDataTable(v.PdispValVal(), false, false)
}

func (this *DataTable) IID() *syscall.GUID {
	return &IID_DataTable
}

func (this *DataTable) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DataTable) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *DataTable) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *DataTable) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *DataTable) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *DataTable) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *DataTable) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *DataTable) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *DataTable) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *DataTable) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *DataTable) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *DataTable) ShowLegendKey() bool {
	retVal := this.PropGet(0x000000ab, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataTable) SetShowLegendKey(rhs bool)  {
	retVal := this.PropPut(0x000000ab, []interface{}{rhs})
	_= retVal
}

func (this *DataTable) HasBorderHorizontal() bool {
	retVal := this.PropGet(0x00000687, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataTable) SetHasBorderHorizontal(rhs bool)  {
	retVal := this.PropPut(0x00000687, []interface{}{rhs})
	_= retVal
}

func (this *DataTable) HasBorderVertical() bool {
	retVal := this.PropGet(0x00000688, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataTable) SetHasBorderVertical(rhs bool)  {
	retVal := this.PropPut(0x00000688, []interface{}{rhs})
	_= retVal
}

func (this *DataTable) HasBorderOutline() bool {
	retVal := this.PropGet(0x00000689, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataTable) SetHasBorderOutline(rhs bool)  {
	retVal := this.PropPut(0x00000689, []interface{}{rhs})
	_= retVal
}

func (this *DataTable) Border() *Border {
	retVal := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.PdispValVal(), false, true)
}

func (this *DataTable) Font() *Font {
	retVal := this.PropGet(0x00000092, nil)
	return NewFont(retVal.PdispValVal(), false, true)
}

func (this *DataTable) Select()  {
	retVal := this.Call(0x000000eb, nil)
	_= retVal
}

func (this *DataTable) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

func (this *DataTable) AutoScaleFont() ole.Variant {
	retVal := this.PropGet(0x000005f5, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DataTable) SetAutoScaleFont(rhs interface{})  {
	retVal := this.PropPut(0x000005f5, []interface{}{rhs})
	_= retVal
}

func (this *DataTable) Format() *ChartFormat {
	retVal := this.PropGet(0x00000074, nil)
	return NewChartFormat(retVal.PdispValVal(), false, true)
}

