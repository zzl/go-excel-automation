package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024471-0000-0000-C000-000000000046
var IID_ListObject = syscall.GUID{0x00024471, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ListObject struct {
	ole.OleClient
}

func NewListObject(pDisp *win32.IDispatch, addRef bool, scoped bool) *ListObject {
	p := &ListObject{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ListObjectFromVar(v ole.Variant) *ListObject {
	return NewListObject(v.PdispValVal(), false, false)
}

func (this *ListObject) IID() *syscall.GUID {
	return &IID_ListObject
}

func (this *ListObject) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ListObject) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *ListObject) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ListObject) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ListObject) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *ListObject) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *ListObject) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *ListObject) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *ListObject) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *ListObject) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ListObject) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ListObject) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

func (this *ListObject) Publish(target interface{}, linkSource bool) string {
	retVal := this.Call(0x00000767, []interface{}{target, linkSource})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ListObject) Refresh()  {
	retVal := this.Call(0x00000589, nil)
	_= retVal
}

func (this *ListObject) Unlink()  {
	retVal := this.Call(0x00000904, nil)
	_= retVal
}

func (this *ListObject) Unlist()  {
	retVal := this.Call(0x00000905, nil)
	_= retVal
}

func (this *ListObject) UpdateChanges(iConflictType int32)  {
	retVal := this.Call(0x00000906, []interface{}{iConflictType})
	_= retVal
}

func (this *ListObject) Resize(range_ *Range)  {
	retVal := this.Call(0x00000100, []interface{}{range_})
	_= retVal
}

func (this *ListObject) Default_() string {
	retVal := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ListObject) Active() bool {
	retVal := this.PropGet(0x00000908, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListObject) DataBodyRange() *Range {
	retVal := this.PropGet(0x000002c1, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *ListObject) DisplayRightToLeft() bool {
	retVal := this.PropGet(0x000006ee, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListObject) HeaderRowRange() *Range {
	retVal := this.PropGet(0x00000909, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *ListObject) InsertRowRange() *Range {
	retVal := this.PropGet(0x0000090a, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *ListObject) ListColumns() *ListColumns {
	retVal := this.PropGet(0x0000090b, nil)
	return NewListColumns(retVal.PdispValVal(), false, true)
}

func (this *ListObject) ListRows() *ListRows {
	retVal := this.PropGet(0x0000090c, nil)
	return NewListRows(retVal.PdispValVal(), false, true)
}

func (this *ListObject) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ListObject) SetName(rhs string)  {
	retVal := this.PropPut(0x0000006e, []interface{}{rhs})
	_= retVal
}

func (this *ListObject) QueryTable() *QueryTable {
	retVal := this.PropGet(0x0000056a, nil)
	return NewQueryTable(retVal.PdispValVal(), false, true)
}

func (this *ListObject) Range() *Range {
	retVal := this.PropGet(0x000000c5, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *ListObject) ShowAutoFilter() bool {
	retVal := this.PropGet(0x0000090d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListObject) SetShowAutoFilter(rhs bool)  {
	retVal := this.PropPut(0x0000090d, []interface{}{rhs})
	_= retVal
}

func (this *ListObject) ShowTotals() bool {
	retVal := this.PropGet(0x0000090e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListObject) SetShowTotals(rhs bool)  {
	retVal := this.PropPut(0x0000090e, []interface{}{rhs})
	_= retVal
}

func (this *ListObject) SourceType() int32 {
	retVal := this.PropGet(0x000002ad, nil)
	return retVal.LValVal()
}

func (this *ListObject) TotalsRowRange() *Range {
	retVal := this.PropGet(0x0000090f, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *ListObject) SharePointURL() string {
	retVal := this.PropGet(0x00000910, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ListObject) XmlMap() *XmlMap {
	retVal := this.PropGet(0x000008cd, nil)
	return NewXmlMap(retVal.PdispValVal(), false, true)
}

func (this *ListObject) DisplayName() string {
	retVal := this.PropGet(0x00000a75, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ListObject) SetDisplayName(rhs string)  {
	retVal := this.PropPut(0x00000a75, []interface{}{rhs})
	_= retVal
}

func (this *ListObject) ShowHeaders() bool {
	retVal := this.PropGet(0x00000a76, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListObject) SetShowHeaders(rhs bool)  {
	retVal := this.PropPut(0x00000a76, []interface{}{rhs})
	_= retVal
}

func (this *ListObject) AutoFilter() *AutoFilter {
	retVal := this.PropGet(0x00000319, nil)
	return NewAutoFilter(retVal.PdispValVal(), false, true)
}

func (this *ListObject) TableStyle() ole.Variant {
	retVal := this.PropGet(0x000005e0, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ListObject) SetTableStyle(rhs interface{})  {
	retVal := this.PropPut(0x000005e0, []interface{}{rhs})
	_= retVal
}

func (this *ListObject) ShowTableStyleFirstColumn() bool {
	retVal := this.PropGet(0x00000a77, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListObject) SetShowTableStyleFirstColumn(rhs bool)  {
	retVal := this.PropPut(0x00000a77, []interface{}{rhs})
	_= retVal
}

func (this *ListObject) ShowTableStyleLastColumn() bool {
	retVal := this.PropGet(0x00000a03, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListObject) SetShowTableStyleLastColumn(rhs bool)  {
	retVal := this.PropPut(0x00000a03, []interface{}{rhs})
	_= retVal
}

func (this *ListObject) ShowTableStyleRowStripes() bool {
	retVal := this.PropGet(0x00000a04, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListObject) SetShowTableStyleRowStripes(rhs bool)  {
	retVal := this.PropPut(0x00000a04, []interface{}{rhs})
	_= retVal
}

func (this *ListObject) ShowTableStyleColumnStripes() bool {
	retVal := this.PropGet(0x00000a05, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListObject) SetShowTableStyleColumnStripes(rhs bool)  {
	retVal := this.PropPut(0x00000a05, []interface{}{rhs})
	_= retVal
}

func (this *ListObject) Sort() *Sort {
	retVal := this.PropGet(0x00000370, nil)
	return NewSort(retVal.PdispValVal(), false, true)
}

func (this *ListObject) Comment() string {
	retVal := this.PropGet(0x0000038e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ListObject) SetComment(rhs string)  {
	retVal := this.PropPut(0x0000038e, []interface{}{rhs})
	_= retVal
}

func (this *ListObject) ExportToVisio()  {
	retVal := this.Call(0x00000a78, nil)
	_= retVal
}

func (this *ListObject) AlternativeText() string {
	retVal := this.PropGet(0x00000763, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ListObject) SetAlternativeText(rhs string)  {
	retVal := this.PropPut(0x00000763, []interface{}{rhs})
	_= retVal
}

func (this *ListObject) Summary() string {
	retVal := this.PropGet(0x00000111, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ListObject) SetSummary(rhs string)  {
	retVal := this.PropPut(0x00000111, []interface{}{rhs})
	_= retVal
}

