package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
	"time"
)

// 00020872-0000-0000-C000-000000000046
var IID_PivotTable = syscall.GUID{0x00020872, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type PivotTable struct {
	ole.OleClient
}

func NewPivotTable(pDisp *win32.IDispatch, addRef bool, scoped bool) *PivotTable {
	p := &PivotTable{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PivotTableFromVar(v ole.Variant) *PivotTable {
	return NewPivotTable(v.PdispValVal(), false, false)
}

func (this *PivotTable) IID() *syscall.GUID {
	return &IID_PivotTable
}

func (this *PivotTable) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *PivotTable) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *PivotTable) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *PivotTable) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *PivotTable) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *PivotTable) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *PivotTable) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *PivotTable) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *PivotTable) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *PivotTable) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *PivotTable) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var PivotTable_AddFields_OptArgs= []string{
	"RowFields", "ColumnFields", "PageFields", "AddToTable", 
}

func (this *PivotTable) AddFields(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(PivotTable_AddFields_OptArgs, optArgs)
	retVal := this.Call(0x000002c4, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var PivotTable_ColumnFields_OptArgs= []string{
	"Index", 
}

func (this *PivotTable) ColumnFields(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(PivotTable_ColumnFields_OptArgs, optArgs)
	retVal := this.PropGet(0x000002c9, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *PivotTable) ColumnGrand() bool {
	retVal := this.PropGet(0x000002b6, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetColumnGrand(rhs bool)  {
	retVal := this.PropPut(0x000002b6, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) ColumnRange() *Range {
	retVal := this.PropGet(0x000002be, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

var PivotTable_ShowPages_OptArgs= []string{
	"PageField", 
}

func (this *PivotTable) ShowPages(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(PivotTable_ShowPages_OptArgs, optArgs)
	retVal := this.Call(0x000002c2, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *PivotTable) DataBodyRange() *Range {
	retVal := this.PropGet(0x000002c1, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

var PivotTable_DataFields_OptArgs= []string{
	"Index", 
}

func (this *PivotTable) DataFields(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(PivotTable_DataFields_OptArgs, optArgs)
	retVal := this.PropGet(0x000002cb, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *PivotTable) DataLabelRange() *Range {
	retVal := this.PropGet(0x000002c0, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *PivotTable) Default_() string {
	retVal := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotTable) SetDefault_(rhs string)  {
	retVal := this.PropPut(0x00000000, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) HasAutoFormat() bool {
	retVal := this.PropGet(0x000002b7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetHasAutoFormat(rhs bool)  {
	retVal := this.PropPut(0x000002b7, []interface{}{rhs})
	_= retVal
}

var PivotTable_HiddenFields_OptArgs= []string{
	"Index", 
}

func (this *PivotTable) HiddenFields(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(PivotTable_HiddenFields_OptArgs, optArgs)
	retVal := this.PropGet(0x000002c7, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *PivotTable) InnerDetail() string {
	retVal := this.PropGet(0x000002ba, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotTable) SetInnerDetail(rhs string)  {
	retVal := this.PropPut(0x000002ba, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotTable) SetName(rhs string)  {
	retVal := this.PropPut(0x0000006e, []interface{}{rhs})
	_= retVal
}

var PivotTable_PageFields_OptArgs= []string{
	"Index", 
}

func (this *PivotTable) PageFields(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(PivotTable_PageFields_OptArgs, optArgs)
	retVal := this.PropGet(0x000002ca, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *PivotTable) PageRange() *Range {
	retVal := this.PropGet(0x000002bf, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *PivotTable) PageRangeCells() *Range {
	retVal := this.PropGet(0x000005ca, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

var PivotTable_PivotFields_OptArgs= []string{
	"Index", 
}

func (this *PivotTable) PivotFields(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(PivotTable_PivotFields_OptArgs, optArgs)
	retVal := this.Call(0x000002ce, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *PivotTable) RefreshDate() time.Time {
	retVal := this.PropGet(0x000002b8, nil)
	return ole.Date(retVal.DateVal()).ToGoTime()
}

func (this *PivotTable) RefreshName() string {
	retVal := this.PropGet(0x000002b9, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotTable) RefreshTable() bool {
	retVal := this.Call(0x000002cd, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var PivotTable_RowFields_OptArgs= []string{
	"Index", 
}

func (this *PivotTable) RowFields(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(PivotTable_RowFields_OptArgs, optArgs)
	retVal := this.PropGet(0x000002c8, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *PivotTable) RowGrand() bool {
	retVal := this.PropGet(0x000002b5, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetRowGrand(rhs bool)  {
	retVal := this.PropPut(0x000002b5, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) RowRange() *Range {
	retVal := this.PropGet(0x000002bd, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *PivotTable) SaveData() bool {
	retVal := this.PropGet(0x000002b4, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetSaveData(rhs bool)  {
	retVal := this.PropPut(0x000002b4, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) SourceData() ole.Variant {
	retVal := this.PropGet(0x000002ae, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *PivotTable) SetSourceData(rhs interface{})  {
	retVal := this.PropPut(0x000002ae, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) TableRange1() *Range {
	retVal := this.PropGet(0x000002bb, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *PivotTable) TableRange2() *Range {
	retVal := this.PropGet(0x000002bc, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *PivotTable) Value() string {
	retVal := this.PropGet(0x00000006, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotTable) SetValue(rhs string)  {
	retVal := this.PropPut(0x00000006, []interface{}{rhs})
	_= retVal
}

var PivotTable_VisibleFields_OptArgs= []string{
	"Index", 
}

func (this *PivotTable) VisibleFields(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(PivotTable_VisibleFields_OptArgs, optArgs)
	retVal := this.PropGet(0x000002c6, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *PivotTable) CacheIndex() int32 {
	retVal := this.PropGet(0x000005cb, nil)
	return retVal.LValVal()
}

func (this *PivotTable) SetCacheIndex(rhs int32)  {
	retVal := this.PropPut(0x000005cb, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) CalculatedFields() *CalculatedFields {
	retVal := this.Call(0x000005cc, nil)
	return NewCalculatedFields(retVal.PdispValVal(), false, true)
}

func (this *PivotTable) DisplayErrorString() bool {
	retVal := this.PropGet(0x000005cd, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetDisplayErrorString(rhs bool)  {
	retVal := this.PropPut(0x000005cd, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) DisplayNullString() bool {
	retVal := this.PropGet(0x000005ce, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetDisplayNullString(rhs bool)  {
	retVal := this.PropPut(0x000005ce, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) EnableDrilldown() bool {
	retVal := this.PropGet(0x000005cf, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetEnableDrilldown(rhs bool)  {
	retVal := this.PropPut(0x000005cf, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) EnableFieldDialog() bool {
	retVal := this.PropGet(0x000005d0, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetEnableFieldDialog(rhs bool)  {
	retVal := this.PropPut(0x000005d0, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) EnableWizard() bool {
	retVal := this.PropGet(0x000005d1, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetEnableWizard(rhs bool)  {
	retVal := this.PropPut(0x000005d1, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) ErrorString() string {
	retVal := this.PropGet(0x000005d2, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotTable) SetErrorString(rhs string)  {
	retVal := this.PropPut(0x000005d2, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) GetData(name string) float64 {
	retVal := this.Call(0x000005d3, []interface{}{name})
	return retVal.DblValVal()
}

func (this *PivotTable) ListFormulas()  {
	retVal := this.Call(0x000005d4, nil)
	_= retVal
}

func (this *PivotTable) ManualUpdate() bool {
	retVal := this.PropGet(0x000005d5, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetManualUpdate(rhs bool)  {
	retVal := this.PropPut(0x000005d5, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) MergeLabels() bool {
	retVal := this.PropGet(0x000005d6, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetMergeLabels(rhs bool)  {
	retVal := this.PropPut(0x000005d6, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) NullString() string {
	retVal := this.PropGet(0x000005d7, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotTable) SetNullString(rhs string)  {
	retVal := this.PropPut(0x000005d7, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) PivotCache() *PivotCache {
	retVal := this.Call(0x000005d8, nil)
	return NewPivotCache(retVal.PdispValVal(), false, true)
}

func (this *PivotTable) PivotFormulas() *PivotFormulas {
	retVal := this.PropGet(0x000005d9, nil)
	return NewPivotFormulas(retVal.PdispValVal(), false, true)
}

var PivotTable_PivotTableWizard_OptArgs= []string{
	"SourceType", "SourceData", "TableDestination", "TableName", 
	"RowGrand", "ColumnGrand", "SaveData", "HasAutoFormat", 
	"AutoPage", "Reserved", "BackgroundQuery", "OptimizeCache", 
	"PageFieldOrder", "PageFieldWrapCount", "ReadData", "Connection", 
}

func (this *PivotTable) PivotTableWizard(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(PivotTable_PivotTableWizard_OptArgs, optArgs)
	retVal := this.Call(0x000002ac, nil, optArgs...)
	_= retVal
}

func (this *PivotTable) SubtotalHiddenPageItems() bool {
	retVal := this.PropGet(0x000005da, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetSubtotalHiddenPageItems(rhs bool)  {
	retVal := this.PropPut(0x000005da, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) PageFieldOrder() int32 {
	retVal := this.PropGet(0x00000595, nil)
	return retVal.LValVal()
}

func (this *PivotTable) SetPageFieldOrder(rhs int32)  {
	retVal := this.PropPut(0x00000595, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) PageFieldStyle() string {
	retVal := this.PropGet(0x000005db, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotTable) SetPageFieldStyle(rhs string)  {
	retVal := this.PropPut(0x000005db, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) PageFieldWrapCount() int32 {
	retVal := this.PropGet(0x00000596, nil)
	return retVal.LValVal()
}

func (this *PivotTable) SetPageFieldWrapCount(rhs int32)  {
	retVal := this.PropPut(0x00000596, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) PreserveFormatting() bool {
	retVal := this.PropGet(0x000005dc, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetPreserveFormatting(rhs bool)  {
	retVal := this.PropPut(0x000005dc, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) PivotSelect_(name string, mode int32)  {
	retVal := this.Call(0x00000827, []interface{}{name, mode})
	_= retVal
}

func (this *PivotTable) PivotSelection() string {
	retVal := this.PropGet(0x000005de, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotTable) SetPivotSelection(rhs string)  {
	retVal := this.PropPut(0x000005de, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) SelectionMode() int32 {
	retVal := this.PropGet(0x000005df, nil)
	return retVal.LValVal()
}

func (this *PivotTable) SetSelectionMode(rhs int32)  {
	retVal := this.PropPut(0x000005df, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) TableStyle() string {
	retVal := this.PropGet(0x000005e0, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotTable) SetTableStyle(rhs string)  {
	retVal := this.PropPut(0x000005e0, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) Tag() string {
	retVal := this.PropGet(0x000005e1, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotTable) SetTag(rhs string)  {
	retVal := this.PropPut(0x000005e1, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) Update()  {
	retVal := this.Call(0x000002a8, nil)
	_= retVal
}

func (this *PivotTable) VacatedStyle() string {
	retVal := this.PropGet(0x000005e2, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotTable) SetVacatedStyle(rhs string)  {
	retVal := this.PropPut(0x000005e2, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) Format(format int32)  {
	retVal := this.Call(0x00000074, []interface{}{format})
	_= retVal
}

func (this *PivotTable) PrintTitles() bool {
	retVal := this.PropGet(0x0000072e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetPrintTitles(rhs bool)  {
	retVal := this.PropPut(0x0000072e, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) CubeFields() *CubeFields {
	retVal := this.PropGet(0x0000072f, nil)
	return NewCubeFields(retVal.PdispValVal(), false, true)
}

func (this *PivotTable) GrandTotalName() string {
	retVal := this.PropGet(0x00000730, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotTable) SetGrandTotalName(rhs string)  {
	retVal := this.PropPut(0x00000730, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) SmallGrid() bool {
	retVal := this.PropGet(0x00000731, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetSmallGrid(rhs bool)  {
	retVal := this.PropPut(0x00000731, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) RepeatItemsOnEachPrintedPage() bool {
	retVal := this.PropGet(0x00000732, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetRepeatItemsOnEachPrintedPage(rhs bool)  {
	retVal := this.PropPut(0x00000732, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) TotalsAnnotation() bool {
	retVal := this.PropGet(0x00000733, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetTotalsAnnotation(rhs bool)  {
	retVal := this.PropPut(0x00000733, []interface{}{rhs})
	_= retVal
}

var PivotTable_PivotSelect_OptArgs= []string{
	"UseStandardName", 
}

func (this *PivotTable) PivotSelect(name string, mode int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(PivotTable_PivotSelect_OptArgs, optArgs)
	retVal := this.Call(0x000005dd, []interface{}{name, mode}, optArgs...)
	_= retVal
}

func (this *PivotTable) PivotSelectionStandard() string {
	retVal := this.PropGet(0x00000829, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotTable) SetPivotSelectionStandard(rhs string)  {
	retVal := this.PropPut(0x00000829, []interface{}{rhs})
	_= retVal
}

var PivotTable_GetPivotData_OptArgs= []string{
	"DataField", "Field1", "Item1", "Field2", 
	"Item2", "Field3", "Item3", "Field4", 
	"Item4", "Field5", "Item5", "Field6", 
	"Item6", "Field7", "Item7", "Field8", 
	"Item8", "Field9", "Item9", "Field10", 
	"Item10", "Field11", "Item11", "Field12", 
	"Item12", "Field13", "Item13", "Field14", "Item14", 
}

func (this *PivotTable) GetPivotData(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(PivotTable_GetPivotData_OptArgs, optArgs)
	retVal := this.Call(0x0000082a, nil, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *PivotTable) DataPivotField() *PivotField {
	retVal := this.PropGet(0x00000848, nil)
	return NewPivotField(retVal.PdispValVal(), false, true)
}

func (this *PivotTable) EnableDataValueEditing() bool {
	retVal := this.PropGet(0x00000849, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetEnableDataValueEditing(rhs bool)  {
	retVal := this.PropPut(0x00000849, []interface{}{rhs})
	_= retVal
}

var PivotTable_AddDataField_OptArgs= []string{
	"Caption", "Function", 
}

func (this *PivotTable) AddDataField(field *ole.DispatchClass, optArgs ...interface{}) *PivotField {
	optArgs = ole.ProcessOptArgs(PivotTable_AddDataField_OptArgs, optArgs)
	retVal := this.Call(0x0000084a, []interface{}{field}, optArgs...)
	return NewPivotField(retVal.PdispValVal(), false, true)
}

func (this *PivotTable) MDX() string {
	retVal := this.PropGet(0x0000084b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotTable) ViewCalculatedMembers() bool {
	retVal := this.PropGet(0x0000084c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetViewCalculatedMembers(rhs bool)  {
	retVal := this.PropPut(0x0000084c, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) CalculatedMembers() *CalculatedMembers {
	retVal := this.PropGet(0x0000084d, nil)
	return NewCalculatedMembers(retVal.PdispValVal(), false, true)
}

func (this *PivotTable) DisplayImmediateItems() bool {
	retVal := this.PropGet(0x0000084e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetDisplayImmediateItems(rhs bool)  {
	retVal := this.PropPut(0x0000084e, []interface{}{rhs})
	_= retVal
}

var PivotTable_Dummy15_OptArgs= []string{
	"Arg2", "Arg3", "Arg4", "Arg5", 
	"Arg6", "Arg7", "Arg8", "Arg9", 
	"Arg10", "Arg11", "Arg12", "Arg13", 
	"Arg14", "Arg15", "Arg16", "Arg17", 
	"Arg18", "Arg19", "Arg20", "Arg21", 
	"Arg22", "Arg23", "Arg24", "Arg25", 
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30", 
}

func (this *PivotTable) Dummy15(arg1 interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(PivotTable_Dummy15_OptArgs, optArgs)
	retVal := this.Call(0x0000084f, []interface{}{arg1}, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *PivotTable) EnableFieldList() bool {
	retVal := this.PropGet(0x00000850, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetEnableFieldList(rhs bool)  {
	retVal := this.PropPut(0x00000850, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) VisualTotals() bool {
	retVal := this.PropGet(0x00000851, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetVisualTotals(rhs bool)  {
	retVal := this.PropPut(0x00000851, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) ShowPageMultipleItemLabel() bool {
	retVal := this.PropGet(0x00000852, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetShowPageMultipleItemLabel(rhs bool)  {
	retVal := this.PropPut(0x00000852, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) Version() int32 {
	retVal := this.PropGet(0x00000188, nil)
	return retVal.LValVal()
}

var PivotTable_CreateCubeFile_OptArgs= []string{
	"Measures", "Levels", "Members", "Properties", 
}

func (this *PivotTable) CreateCubeFile(file string, optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(PivotTable_CreateCubeFile_OptArgs, optArgs)
	retVal := this.Call(0x00000853, []interface{}{file}, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotTable) DisplayEmptyRow() bool {
	retVal := this.PropGet(0x00000858, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetDisplayEmptyRow(rhs bool)  {
	retVal := this.PropPut(0x00000858, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) DisplayEmptyColumn() bool {
	retVal := this.PropGet(0x00000859, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetDisplayEmptyColumn(rhs bool)  {
	retVal := this.PropPut(0x00000859, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) ShowCellBackgroundFromOLAP() bool {
	retVal := this.PropGet(0x0000085a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetShowCellBackgroundFromOLAP(rhs bool)  {
	retVal := this.PropPut(0x0000085a, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) PivotColumnAxis() *PivotAxis {
	retVal := this.PropGet(0x000009f2, nil)
	return NewPivotAxis(retVal.PdispValVal(), false, true)
}

func (this *PivotTable) PivotRowAxis() *PivotAxis {
	retVal := this.PropGet(0x000009f3, nil)
	return NewPivotAxis(retVal.PdispValVal(), false, true)
}

func (this *PivotTable) ShowDrillIndicators() bool {
	retVal := this.PropGet(0x000009f4, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetShowDrillIndicators(rhs bool)  {
	retVal := this.PropPut(0x000009f4, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) PrintDrillIndicators() bool {
	retVal := this.PropGet(0x000009f5, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetPrintDrillIndicators(rhs bool)  {
	retVal := this.PropPut(0x000009f5, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) DisplayMemberPropertyTooltips() bool {
	retVal := this.PropGet(0x000009f6, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetDisplayMemberPropertyTooltips(rhs bool)  {
	retVal := this.PropPut(0x000009f6, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) DisplayContextTooltips() bool {
	retVal := this.PropGet(0x000009f7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetDisplayContextTooltips(rhs bool)  {
	retVal := this.PropPut(0x000009f7, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) ClearTable()  {
	retVal := this.Call(0x000009f8, nil)
	_= retVal
}

func (this *PivotTable) CompactRowIndent() int32 {
	retVal := this.PropGet(0x000009f9, nil)
	return retVal.LValVal()
}

func (this *PivotTable) SetCompactRowIndent(rhs int32)  {
	retVal := this.PropPut(0x000009f9, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) LayoutRowDefault() int32 {
	retVal := this.PropGet(0x000009fa, nil)
	return retVal.LValVal()
}

func (this *PivotTable) SetLayoutRowDefault(rhs int32)  {
	retVal := this.PropPut(0x000009fa, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) DisplayFieldCaptions() bool {
	retVal := this.PropGet(0x000009fb, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetDisplayFieldCaptions(rhs bool)  {
	retVal := this.PropPut(0x000009fb, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) RowAxisLayout(rowLayout int32)  {
	retVal := this.Call(0x000009fc, []interface{}{rowLayout})
	_= retVal
}

func (this *PivotTable) SubtotalLocation(location int32)  {
	retVal := this.Call(0x000009fe, []interface{}{location})
	_= retVal
}

func (this *PivotTable) ActiveFilters() *PivotFilters {
	retVal := this.PropGet(0x000009ff, nil)
	return NewPivotFilters(retVal.PdispValVal(), false, true)
}

func (this *PivotTable) InGridDropZones() bool {
	retVal := this.PropGet(0x00000a00, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetInGridDropZones(rhs bool)  {
	retVal := this.PropPut(0x00000a00, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) ClearAllFilters()  {
	retVal := this.Call(0x00000a01, nil)
	_= retVal
}

func (this *PivotTable) TableStyle2() ole.Variant {
	retVal := this.PropGet(0x00000a02, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *PivotTable) SetTableStyle2(rhs interface{})  {
	retVal := this.PropPut(0x00000a02, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) ShowTableStyleLastColumn() bool {
	retVal := this.PropGet(0x00000a03, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetShowTableStyleLastColumn(rhs bool)  {
	retVal := this.PropPut(0x00000a03, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) ShowTableStyleRowStripes() bool {
	retVal := this.PropGet(0x00000a04, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetShowTableStyleRowStripes(rhs bool)  {
	retVal := this.PropPut(0x00000a04, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) ShowTableStyleColumnStripes() bool {
	retVal := this.PropGet(0x00000a05, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetShowTableStyleColumnStripes(rhs bool)  {
	retVal := this.PropPut(0x00000a05, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) ShowTableStyleRowHeaders() bool {
	retVal := this.PropGet(0x00000a06, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetShowTableStyleRowHeaders(rhs bool)  {
	retVal := this.PropPut(0x00000a06, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) ShowTableStyleColumnHeaders() bool {
	retVal := this.PropGet(0x00000a07, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetShowTableStyleColumnHeaders(rhs bool)  {
	retVal := this.PropPut(0x00000a07, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) ConvertToFormulas(convertFilters bool)  {
	retVal := this.Call(0x00000a08, []interface{}{convertFilters})
	_= retVal
}

func (this *PivotTable) AllowMultipleFilters() bool {
	retVal := this.PropGet(0x00000a0a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetAllowMultipleFilters(rhs bool)  {
	retVal := this.PropPut(0x00000a0a, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) CompactLayoutRowHeader() string {
	retVal := this.PropGet(0x00000a0b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotTable) SetCompactLayoutRowHeader(rhs string)  {
	retVal := this.PropPut(0x00000a0b, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) CompactLayoutColumnHeader() string {
	retVal := this.PropGet(0x00000a0c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotTable) SetCompactLayoutColumnHeader(rhs string)  {
	retVal := this.PropPut(0x00000a0c, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) FieldListSortAscending() bool {
	retVal := this.PropGet(0x00000a0d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetFieldListSortAscending(rhs bool)  {
	retVal := this.PropPut(0x00000a0d, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) SortUsingCustomLists() bool {
	retVal := this.PropGet(0x00000a0e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetSortUsingCustomLists(rhs bool)  {
	retVal := this.PropPut(0x00000a0e, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) ChangeConnection(conn *WorkbookConnection)  {
	retVal := this.Call(0x00000a0f, []interface{}{conn})
	_= retVal
}

func (this *PivotTable) ChangePivotCache(pivotCache interface{})  {
	retVal := this.Call(0x00000a11, []interface{}{pivotCache})
	_= retVal
}

func (this *PivotTable) Location() string {
	retVal := this.PropGet(0x00000575, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotTable) SetLocation(rhs string)  {
	retVal := this.PropPut(0x00000575, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) EnableWriteback() bool {
	retVal := this.PropGet(0x00000b38, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetEnableWriteback(rhs bool)  {
	retVal := this.PropPut(0x00000b38, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) Allocation() int32 {
	retVal := this.PropGet(0x00000b39, nil)
	return retVal.LValVal()
}

func (this *PivotTable) SetAllocation(rhs int32)  {
	retVal := this.PropPut(0x00000b39, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) AllocationValue() int32 {
	retVal := this.PropGet(0x00000b3a, nil)
	return retVal.LValVal()
}

func (this *PivotTable) SetAllocationValue(rhs int32)  {
	retVal := this.PropPut(0x00000b3a, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) AllocationMethod() int32 {
	retVal := this.PropGet(0x00000b3b, nil)
	return retVal.LValVal()
}

func (this *PivotTable) SetAllocationMethod(rhs int32)  {
	retVal := this.PropPut(0x00000b3b, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) AllocationWeightExpression() string {
	retVal := this.PropGet(0x00000b3c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotTable) SetAllocationWeightExpression(rhs string)  {
	retVal := this.PropPut(0x00000b3c, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) AllocateChanges()  {
	retVal := this.Call(0x00000b27, nil)
	_= retVal
}

func (this *PivotTable) CommitChanges()  {
	retVal := this.Call(0x00000b3d, nil)
	_= retVal
}

func (this *PivotTable) DiscardChanges()  {
	retVal := this.Call(0x00000b28, nil)
	_= retVal
}

func (this *PivotTable) RefreshDataSourceValues()  {
	retVal := this.Call(0x00000b3e, nil)
	_= retVal
}

func (this *PivotTable) RepeatAllLabels(repeat int32)  {
	retVal := this.Call(0x00000b3f, []interface{}{repeat})
	_= retVal
}

func (this *PivotTable) ChangeList() *PivotTableChangeList {
	retVal := this.PropGet(0x00000b40, nil)
	return NewPivotTableChangeList(retVal.PdispValVal(), false, true)
}

func (this *PivotTable) Slicers() *Slicers {
	retVal := this.PropGet(0x00000b41, nil)
	return NewSlicers(retVal.PdispValVal(), false, true)
}

func (this *PivotTable) AlternativeText() string {
	retVal := this.PropGet(0x00000763, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotTable) SetAlternativeText(rhs string)  {
	retVal := this.PropPut(0x00000763, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) Summary() string {
	retVal := this.PropGet(0x00000111, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotTable) SetSummary(rhs string)  {
	retVal := this.PropPut(0x00000111, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) VisualTotalsForSets() bool {
	retVal := this.PropGet(0x00000b42, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetVisualTotalsForSets(rhs bool)  {
	retVal := this.PropPut(0x00000b42, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) ShowValuesRow() bool {
	retVal := this.PropGet(0x00000b43, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetShowValuesRow(rhs bool)  {
	retVal := this.PropPut(0x00000b43, []interface{}{rhs})
	_= retVal
}

func (this *PivotTable) CalculatedMembersInFilters() bool {
	retVal := this.PropGet(0x00000b44, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotTable) SetCalculatedMembersInFilters(rhs bool)  {
	retVal := this.PropPut(0x00000b44, []interface{}{rhs})
	_= retVal
}

