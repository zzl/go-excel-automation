package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00020874-0000-0000-C000-000000000046
var IID_PivotField = syscall.GUID{0x00020874, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type PivotField struct {
	ole.OleClient
}

func NewPivotField(pDisp *win32.IDispatch, addRef bool, scoped bool) *PivotField {
	if pDisp == nil {
		return nil
	}
	p := &PivotField{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PivotFieldFromVar(v ole.Variant) *PivotField {
	return NewPivotField(v.IDispatch(), false, false)
}

func (this *PivotField) IID() *syscall.GUID {
	return &IID_PivotField
}

func (this *PivotField) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *PivotField) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *PivotField) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *PivotField) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *PivotField) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *PivotField) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *PivotField) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *PivotField) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *PivotField) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *PivotField) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *PivotField) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *PivotField) Calculation() int32 {
	retVal, _ := this.PropGet(0x0000013c, nil)
	return retVal.LValVal()
}

func (this *PivotField) SetCalculation(rhs int32) {
	_ = this.PropPut(0x0000013c, []interface{}{rhs})
}

func (this *PivotField) ChildField() *PivotField {
	retVal, _ := this.PropGet(0x000002e0, nil)
	return NewPivotField(retVal.IDispatch(), false, true)
}

var PivotField_ChildItems_OptArgs = []string{
	"Index",
}

func (this *PivotField) ChildItems(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(PivotField_ChildItems_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x000002da, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PivotField) CurrentPage() ole.Variant {
	retVal, _ := this.PropGet(0x000002e2, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PivotField) SetCurrentPage(rhs interface{}) {
	_ = this.PropPut(0x000002e2, []interface{}{rhs})
}

func (this *PivotField) DataRange() *Range {
	retVal, _ := this.PropGet(0x000002d0, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *PivotField) DataType() int32 {
	retVal, _ := this.PropGet(0x000002d2, nil)
	return retVal.LValVal()
}

func (this *PivotField) Default_() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotField) SetDefault_(rhs string) {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *PivotField) Function() int32 {
	retVal, _ := this.PropGet(0x00000383, nil)
	return retVal.LValVal()
}

func (this *PivotField) SetFunction(rhs int32) {
	_ = this.PropPut(0x00000383, []interface{}{rhs})
}

func (this *PivotField) GroupLevel() ole.Variant {
	retVal, _ := this.PropGet(0x000002d3, nil)
	com.AddToScope(retVal)
	return *retVal
}

var PivotField_HiddenItems_OptArgs = []string{
	"Index",
}

func (this *PivotField) HiddenItems(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(PivotField_HiddenItems_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x000002d8, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PivotField) LabelRange() *Range {
	retVal, _ := this.PropGet(0x000002cf, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *PivotField) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotField) SetName(rhs string) {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *PivotField) NumberFormat() string {
	retVal, _ := this.PropGet(0x000000c1, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotField) SetNumberFormat(rhs string) {
	_ = this.PropPut(0x000000c1, []interface{}{rhs})
}

func (this *PivotField) Orientation() int32 {
	retVal, _ := this.PropGet(0x00000086, nil)
	return retVal.LValVal()
}

func (this *PivotField) SetOrientation(rhs int32) {
	_ = this.PropPut(0x00000086, []interface{}{rhs})
}

func (this *PivotField) ShowAllItems() bool {
	retVal, _ := this.PropGet(0x000001c4, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) SetShowAllItems(rhs bool) {
	_ = this.PropPut(0x000001c4, []interface{}{rhs})
}

func (this *PivotField) ParentField() *PivotField {
	retVal, _ := this.PropGet(0x000002dc, nil)
	return NewPivotField(retVal.IDispatch(), false, true)
}

var PivotField_ParentItems_OptArgs = []string{
	"Index",
}

func (this *PivotField) ParentItems(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(PivotField_ParentItems_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x000002d9, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var PivotField_PivotItems_OptArgs = []string{
	"Index",
}

func (this *PivotField) PivotItems(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(PivotField_PivotItems_OptArgs, optArgs)
	retVal, _ := this.Call(0x000002e1, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PivotField) Position() ole.Variant {
	retVal, _ := this.PropGet(0x00000085, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PivotField) SetPosition(rhs interface{}) {
	_ = this.PropPut(0x00000085, []interface{}{rhs})
}

func (this *PivotField) SourceName() string {
	retVal, _ := this.PropGet(0x000002d1, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var PivotField_Subtotals_OptArgs = []string{
	"Index",
}

func (this *PivotField) Subtotals(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(PivotField_Subtotals_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x000002dd, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var PivotField_SetSubtotals_OptArgs = []string{
	"Index",
}

func (this *PivotField) SetSubtotals(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(PivotField_SetSubtotals_OptArgs, optArgs)
	_ = this.PropPut(0x000002dd, nil, optArgs...)
}

func (this *PivotField) BaseField() ole.Variant {
	retVal, _ := this.PropGet(0x000002de, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PivotField) SetBaseField(rhs interface{}) {
	_ = this.PropPut(0x000002de, []interface{}{rhs})
}

func (this *PivotField) BaseItem() ole.Variant {
	retVal, _ := this.PropGet(0x000002df, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PivotField) SetBaseItem(rhs interface{}) {
	_ = this.PropPut(0x000002df, []interface{}{rhs})
}

func (this *PivotField) TotalLevels() ole.Variant {
	retVal, _ := this.PropGet(0x000002d4, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PivotField) Value() string {
	retVal, _ := this.PropGet(0x00000006, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotField) SetValue(rhs string) {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

var PivotField_VisibleItems_OptArgs = []string{
	"Index",
}

func (this *PivotField) VisibleItems(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(PivotField_VisibleItems_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x000002d7, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PivotField) CalculatedItems() *CalculatedItems {
	retVal, _ := this.Call(0x000005e3, nil)
	return NewCalculatedItems(retVal.IDispatch(), false, true)
}

func (this *PivotField) Delete() {
	retVal, _ := this.Call(0x00000075, nil)
	_ = retVal
}

func (this *PivotField) DragToColumn() bool {
	retVal, _ := this.PropGet(0x000005e4, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) SetDragToColumn(rhs bool) {
	_ = this.PropPut(0x000005e4, []interface{}{rhs})
}

func (this *PivotField) DragToHide() bool {
	retVal, _ := this.PropGet(0x000005e5, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) SetDragToHide(rhs bool) {
	_ = this.PropPut(0x000005e5, []interface{}{rhs})
}

func (this *PivotField) DragToPage() bool {
	retVal, _ := this.PropGet(0x000005e6, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) SetDragToPage(rhs bool) {
	_ = this.PropPut(0x000005e6, []interface{}{rhs})
}

func (this *PivotField) DragToRow() bool {
	retVal, _ := this.PropGet(0x000005e7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) SetDragToRow(rhs bool) {
	_ = this.PropPut(0x000005e7, []interface{}{rhs})
}

func (this *PivotField) DragToData() bool {
	retVal, _ := this.PropGet(0x00000734, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) SetDragToData(rhs bool) {
	_ = this.PropPut(0x00000734, []interface{}{rhs})
}

func (this *PivotField) Formula() string {
	retVal, _ := this.PropGet(0x00000105, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotField) SetFormula(rhs string) {
	_ = this.PropPut(0x00000105, []interface{}{rhs})
}

func (this *PivotField) IsCalculated() bool {
	retVal, _ := this.PropGet(0x000005e8, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) MemoryUsed() int32 {
	retVal, _ := this.PropGet(0x00000174, nil)
	return retVal.LValVal()
}

func (this *PivotField) ServerBased() bool {
	retVal, _ := this.PropGet(0x000005e9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) SetServerBased(rhs bool) {
	_ = this.PropPut(0x000005e9, []interface{}{rhs})
}

func (this *PivotField) AutoSort_(order int32, field string) {
	retVal, _ := this.Call(0x00000a13, []interface{}{order, field})
	_ = retVal
}

func (this *PivotField) AutoShow(type_ int32, range_ int32, count int32, field string) {
	retVal, _ := this.Call(0x000005eb, []interface{}{type_, range_, count, field})
	_ = retVal
}

func (this *PivotField) AutoSortOrder() int32 {
	retVal, _ := this.PropGet(0x000005ec, nil)
	return retVal.LValVal()
}

func (this *PivotField) AutoSortField() string {
	retVal, _ := this.PropGet(0x000005ed, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotField) AutoShowType() int32 {
	retVal, _ := this.PropGet(0x000005ee, nil)
	return retVal.LValVal()
}

func (this *PivotField) AutoShowRange() int32 {
	retVal, _ := this.PropGet(0x000005ef, nil)
	return retVal.LValVal()
}

func (this *PivotField) AutoShowCount() int32 {
	retVal, _ := this.PropGet(0x000005f0, nil)
	return retVal.LValVal()
}

func (this *PivotField) AutoShowField() string {
	retVal, _ := this.PropGet(0x000005f1, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotField) LayoutBlankLine() bool {
	retVal, _ := this.PropGet(0x00000735, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) SetLayoutBlankLine(rhs bool) {
	_ = this.PropPut(0x00000735, []interface{}{rhs})
}

func (this *PivotField) LayoutSubtotalLocation() int32 {
	retVal, _ := this.PropGet(0x00000736, nil)
	return retVal.LValVal()
}

func (this *PivotField) SetLayoutSubtotalLocation(rhs int32) {
	_ = this.PropPut(0x00000736, []interface{}{rhs})
}

func (this *PivotField) LayoutPageBreak() bool {
	retVal, _ := this.PropGet(0x00000737, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) SetLayoutPageBreak(rhs bool) {
	_ = this.PropPut(0x00000737, []interface{}{rhs})
}

func (this *PivotField) LayoutForm() int32 {
	retVal, _ := this.PropGet(0x00000738, nil)
	return retVal.LValVal()
}

func (this *PivotField) SetLayoutForm(rhs int32) {
	_ = this.PropPut(0x00000738, []interface{}{rhs})
}

func (this *PivotField) SubtotalName() string {
	retVal, _ := this.PropGet(0x00000739, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotField) SetSubtotalName(rhs string) {
	_ = this.PropPut(0x00000739, []interface{}{rhs})
}

func (this *PivotField) Caption() string {
	retVal, _ := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotField) SetCaption(rhs string) {
	_ = this.PropPut(0x0000008b, []interface{}{rhs})
}

func (this *PivotField) DrilledDown() bool {
	retVal, _ := this.PropGet(0x0000073a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) SetDrilledDown(rhs bool) {
	_ = this.PropPut(0x0000073a, []interface{}{rhs})
}

func (this *PivotField) CubeField() *CubeField {
	retVal, _ := this.PropGet(0x0000073b, nil)
	return NewCubeField(retVal.IDispatch(), false, true)
}

func (this *PivotField) CurrentPageName() string {
	retVal, _ := this.PropGet(0x0000073c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotField) SetCurrentPageName(rhs string) {
	_ = this.PropPut(0x0000073c, []interface{}{rhs})
}

func (this *PivotField) StandardFormula() string {
	retVal, _ := this.PropGet(0x00000824, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotField) SetStandardFormula(rhs string) {
	_ = this.PropPut(0x00000824, []interface{}{rhs})
}

func (this *PivotField) HiddenItemsList() ole.Variant {
	retVal, _ := this.PropGet(0x0000085b, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PivotField) SetHiddenItemsList(rhs interface{}) {
	_ = this.PropPut(0x0000085b, []interface{}{rhs})
}

func (this *PivotField) DatabaseSort() bool {
	retVal, _ := this.PropGet(0x0000085c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) SetDatabaseSort(rhs bool) {
	_ = this.PropPut(0x0000085c, []interface{}{rhs})
}

func (this *PivotField) IsMemberProperty() bool {
	retVal, _ := this.PropGet(0x0000085d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) PropertyParentField() *PivotField {
	retVal, _ := this.PropGet(0x0000085e, nil)
	return NewPivotField(retVal.IDispatch(), false, true)
}

func (this *PivotField) PropertyOrder() int32 {
	retVal, _ := this.PropGet(0x0000085f, nil)
	return retVal.LValVal()
}

func (this *PivotField) SetPropertyOrder(rhs int32) {
	_ = this.PropPut(0x0000085f, []interface{}{rhs})
}

func (this *PivotField) EnableItemSelection() bool {
	retVal, _ := this.PropGet(0x00000860, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) SetEnableItemSelection(rhs bool) {
	_ = this.PropPut(0x00000860, []interface{}{rhs})
}

func (this *PivotField) CurrentPageList() ole.Variant {
	retVal, _ := this.PropGet(0x00000861, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PivotField) SetCurrentPageList(rhs interface{}) {
	_ = this.PropPut(0x00000861, []interface{}{rhs})
}

var PivotField_AddPageItem_OptArgs = []string{
	"ClearList",
}

func (this *PivotField) AddPageItem(item string, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(PivotField_AddPageItem_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000862, []interface{}{item}, optArgs...)
	_ = retVal
}

func (this *PivotField) Hidden() bool {
	retVal, _ := this.PropGet(0x0000010c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) SetHidden(rhs bool) {
	_ = this.PropPut(0x0000010c, []interface{}{rhs})
}

func (this *PivotField) DrillTo(field string) {
	retVal, _ := this.Call(0x00000a14, []interface{}{field})
	_ = retVal
}

func (this *PivotField) UseMemberPropertyAsCaption() bool {
	retVal, _ := this.PropGet(0x00000a15, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) SetUseMemberPropertyAsCaption(rhs bool) {
	_ = this.PropPut(0x00000a15, []interface{}{rhs})
}

func (this *PivotField) MemberPropertyCaption() string {
	retVal, _ := this.PropGet(0x00000a16, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotField) SetMemberPropertyCaption(rhs string) {
	_ = this.PropPut(0x00000a16, []interface{}{rhs})
}

func (this *PivotField) DisplayAsTooltip() bool {
	retVal, _ := this.PropGet(0x00000a17, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) SetDisplayAsTooltip(rhs bool) {
	_ = this.PropPut(0x00000a17, []interface{}{rhs})
}

func (this *PivotField) DisplayInReport() bool {
	retVal, _ := this.PropGet(0x00000a18, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) SetDisplayInReport(rhs bool) {
	_ = this.PropPut(0x00000a18, []interface{}{rhs})
}

func (this *PivotField) DisplayAsCaption() bool {
	retVal, _ := this.PropGet(0x00000a19, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) LayoutCompactRow() bool {
	retVal, _ := this.PropGet(0x00000a1a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) SetLayoutCompactRow(rhs bool) {
	_ = this.PropPut(0x00000a1a, []interface{}{rhs})
}

func (this *PivotField) IncludeNewItemsInFilter() bool {
	retVal, _ := this.PropGet(0x00000a1b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) SetIncludeNewItemsInFilter(rhs bool) {
	_ = this.PropPut(0x00000a1b, []interface{}{rhs})
}

func (this *PivotField) VisibleItemsList() ole.Variant {
	retVal, _ := this.PropGet(0x00000a1c, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PivotField) SetVisibleItemsList(rhs interface{}) {
	_ = this.PropPut(0x00000a1c, []interface{}{rhs})
}

func (this *PivotField) PivotFilters() *PivotFilters {
	retVal, _ := this.PropGet(0x00000a1d, nil)
	return NewPivotFilters(retVal.IDispatch(), false, true)
}

func (this *PivotField) AutoSortPivotLine() *PivotLine {
	retVal, _ := this.PropGet(0x00000a1e, nil)
	return NewPivotLine(retVal.IDispatch(), false, true)
}

func (this *PivotField) AutoSortCustomSubtotal() int32 {
	retVal, _ := this.PropGet(0x00000a1f, nil)
	return retVal.LValVal()
}

func (this *PivotField) ShowingInAxis() bool {
	retVal, _ := this.PropGet(0x00000a20, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) EnableMultiplePageItems() bool {
	retVal, _ := this.PropGet(0x00000888, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) SetEnableMultiplePageItems(rhs bool) {
	_ = this.PropPut(0x00000888, []interface{}{rhs})
}

func (this *PivotField) AllItemsVisible() bool {
	retVal, _ := this.PropGet(0x00000a21, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) ClearManualFilter() {
	retVal, _ := this.Call(0x00000a22, nil)
	_ = retVal
}

func (this *PivotField) ClearAllFilters() {
	retVal, _ := this.Call(0x00000a01, nil)
	_ = retVal
}

func (this *PivotField) ClearValueFilters() {
	retVal, _ := this.Call(0x00000a23, nil)
	_ = retVal
}

func (this *PivotField) ClearLabelFilters() {
	retVal, _ := this.Call(0x00000a24, nil)
	_ = retVal
}

var PivotField_AutoSort_OptArgs = []string{
	"PivotLine", "CustomSubtotal",
}

func (this *PivotField) AutoSort(order int32, field string, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(PivotField_AutoSort_OptArgs, optArgs)
	retVal, _ := this.Call(0x000005ea, []interface{}{order, field}, optArgs...)
	_ = retVal
}

func (this *PivotField) SourceCaption() string {
	retVal, _ := this.PropGet(0x00000a27, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotField) ShowDetail() bool {
	retVal, _ := this.PropGet(0x00000249, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) SetShowDetail(rhs bool) {
	_ = this.PropPut(0x00000249, []interface{}{rhs})
}

func (this *PivotField) RepeatLabels() bool {
	retVal, _ := this.PropGet(0x00000b45, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotField) SetRepeatLabels(rhs bool) {
	_ = this.PropPut(0x00000b45, []interface{}{rhs})
}
