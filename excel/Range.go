package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00020846-0000-0000-C000-000000000046
var IID_Range = syscall.GUID{0x00020846, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Range struct {
	ole.OleClient
}

func NewRange(pDisp *win32.IDispatch, addRef bool, scoped bool) *Range {
	if pDisp == nil {
		return nil
	}
	p := &Range{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func RangeFromVar(v ole.Variant) *Range {
	return NewRange(v.IDispatch(), false, false)
}

func (this *Range) IID() *syscall.GUID {
	return &IID_Range
}

func (this *Range) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Range) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Range) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Range) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Range) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Range) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Range) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Range) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Range) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Range) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Range) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Range) Activate() ole.Variant {
	retVal, _ := this.Call(0x00000130, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) AddIndent() ole.Variant {
	retVal, _ := this.PropGet(0x00000427, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetAddIndent(rhs interface{}) {
	_ = this.PropPut(0x00000427, []interface{}{rhs})
}

var Range_Address_OptArgs = []string{
	"RowAbsolute", "ColumnAbsolute", "ReferenceStyle", "External", "RelativeTo",
}

func (this *Range) Address(optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(Range_Address_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x000000ec, nil, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var Range_AddressLocal_OptArgs = []string{
	"RowAbsolute", "ColumnAbsolute", "ReferenceStyle", "External", "RelativeTo",
}

func (this *Range) AddressLocal(optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(Range_AddressLocal_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x000001b5, nil, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var Range_AdvancedFilter_OptArgs = []string{
	"CriteriaRange", "CopyToRange", "Unique",
}

func (this *Range) AdvancedFilter(action int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_AdvancedFilter_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000036c, []interface{}{action}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Range_ApplyNames_OptArgs = []string{
	"Names", "IgnoreRelativeAbsolute", "UseRowColumnNames", "OmitColumn",
	"OmitRow", "Order", "AppendLast",
}

func (this *Range) ApplyNames(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_ApplyNames_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001b9, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) ApplyOutlineStyles() ole.Variant {
	retVal, _ := this.Call(0x000001c0, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) Areas() *Areas {
	retVal, _ := this.PropGet(0x00000238, nil)
	return NewAreas(retVal.IDispatch(), false, true)
}

func (this *Range) AutoComplete(string string) string {
	retVal, _ := this.Call(0x000004a1, []interface{}{string})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var Range_AutoFill_OptArgs = []string{
	"Type",
}

func (this *Range) AutoFill(destination *Range, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_AutoFill_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001c1, []interface{}{destination}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Range_AutoFilter_OptArgs = []string{
	"Field", "Criteria1", "Operator", "Criteria2", "VisibleDropDown",
}

func (this *Range) AutoFilter(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_AutoFilter_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000319, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) AutoFit() ole.Variant {
	retVal, _ := this.Call(0x000000ed, nil)
	com.AddToScope(retVal)
	return *retVal
}

var Range_AutoFormat_OptArgs = []string{
	"Format", "Number", "Font", "Alignment",
	"Border", "Pattern", "Width",
}

func (this *Range) AutoFormat(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_AutoFormat_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000072, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) AutoOutline() ole.Variant {
	retVal, _ := this.Call(0x0000040c, nil)
	com.AddToScope(retVal)
	return *retVal
}

var Range_BorderAround__OptArgs = []string{
	"LineStyle", "Weight", "ColorIndex", "Color",
}

func (this *Range) BorderAround_(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_BorderAround__OptArgs, optArgs)
	retVal, _ := this.Call(0x0000042b, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) Borders() *Borders {
	retVal, _ := this.PropGet(0x000001b3, nil)
	return NewBorders(retVal.IDispatch(), false, true)
}

func (this *Range) Calculate() ole.Variant {
	retVal, _ := this.Call(0x00000117, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) Cells() *Range {
	retVal, _ := this.PropGet(0x000000ee, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

var Range_Characters_OptArgs = []string{
	"Start", "Length",
}

func (this *Range) Characters(optArgs ...interface{}) *Characters {
	optArgs = ole.ProcessOptArgs(Range_Characters_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000025b, nil, optArgs...)
	return NewCharacters(retVal.IDispatch(), false, true)
}

var Range_CheckSpelling_OptArgs = []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang",
}

func (this *Range) CheckSpelling(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_CheckSpelling_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001f9, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) Clear() ole.Variant {
	retVal, _ := this.Call(0x0000006f, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) ClearContents() ole.Variant {
	retVal, _ := this.Call(0x00000071, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) ClearFormats() ole.Variant {
	retVal, _ := this.Call(0x00000070, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) ClearNotes() ole.Variant {
	retVal, _ := this.Call(0x000000ef, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) ClearOutline() ole.Variant {
	retVal, _ := this.Call(0x0000040d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) Column() int32 {
	retVal, _ := this.PropGet(0x000000f0, nil)
	return retVal.LValVal()
}

func (this *Range) ColumnDifferences(comparison interface{}) *Range {
	retVal, _ := this.Call(0x000001fe, []interface{}{comparison})
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Range) Columns() *Range {
	retVal, _ := this.PropGet(0x000000f1, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Range) ColumnWidth() ole.Variant {
	retVal, _ := this.PropGet(0x000000f2, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetColumnWidth(rhs interface{}) {
	_ = this.PropPut(0x000000f2, []interface{}{rhs})
}

var Range_Consolidate_OptArgs = []string{
	"Sources", "Function", "TopRow", "LeftColumn", "CreateLinks",
}

func (this *Range) Consolidate(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Consolidate_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001e2, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Range_Copy_OptArgs = []string{
	"Destination",
}

func (this *Range) Copy(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Copy_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000227, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Range_CopyFromRecordset_OptArgs = []string{
	"MaxRows", "MaxColumns",
}

func (this *Range) CopyFromRecordset(data *win32.IUnknown, optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Range_CopyFromRecordset_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000480, []interface{}{data}, optArgs...)
	return retVal.LValVal()
}

var Range_CopyPicture_OptArgs = []string{
	"Appearance", "Format",
}

func (this *Range) CopyPicture(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_CopyPicture_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d5, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

var Range_CreateNames_OptArgs = []string{
	"Top", "Left", "Bottom", "Right",
}

func (this *Range) CreateNames(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_CreateNames_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001c9, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Range_CreatePublisher_OptArgs = []string{
	"Edition", "Appearance", "ContainsPICT", "ContainsBIFF",
	"ContainsRTF", "ContainsVALU",
}

func (this *Range) CreatePublisher(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_CreatePublisher_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001ca, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) CurrentArray() *Range {
	retVal, _ := this.PropGet(0x000001f5, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Range) CurrentRegion() *Range {
	retVal, _ := this.PropGet(0x000000f3, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

var Range_Cut_OptArgs = []string{
	"Destination",
}

func (this *Range) Cut(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Cut_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000235, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Range_DataSeries_OptArgs = []string{
	"Rowcol", "Type", "Date", "Step",
	"Stop", "Trend",
}

func (this *Range) DataSeries(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_DataSeries_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001d0, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Range_Default__OptArgs = []string{
	"RowIndex", "ColumnIndex",
}

func (this *Range) Default_(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Default__OptArgs, optArgs)
	retVal, _ := this.PropGet(0x00000000, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Range_SetDefault__OptArgs = []string{
	"RowIndex", "ColumnIndex",
}

func (this *Range) SetDefault_(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Range_SetDefault__OptArgs, optArgs)
	_ = this.PropPut(0x00000000, nil, optArgs...)
}

var Range_Delete_OptArgs = []string{
	"Shift",
}

func (this *Range) Delete(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Delete_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000075, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) Dependents() *Range {
	retVal, _ := this.PropGet(0x0000021f, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Range) DialogBox() ole.Variant {
	retVal, _ := this.Call(0x000000f5, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) DirectDependents() *Range {
	retVal, _ := this.PropGet(0x00000221, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Range) DirectPrecedents() *Range {
	retVal, _ := this.PropGet(0x00000222, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

var Range_EditionOptions_OptArgs = []string{
	"Name", "Reference", "Appearance", "ChartSize", "Format",
}

func (this *Range) EditionOptions(type_ int32, option int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_EditionOptions_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000046b, []interface{}{type_, option}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) End(direction int32) *Range {
	retVal, _ := this.PropGet(0x000001f4, []interface{}{direction})
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Range) EntireColumn() *Range {
	retVal, _ := this.PropGet(0x000000f6, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Range) EntireRow() *Range {
	retVal, _ := this.PropGet(0x000000f7, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Range) FillDown() ole.Variant {
	retVal, _ := this.Call(0x000000f8, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) FillLeft() ole.Variant {
	retVal, _ := this.Call(0x000000f9, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) FillRight() ole.Variant {
	retVal, _ := this.Call(0x000000fa, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) FillUp() ole.Variant {
	retVal, _ := this.Call(0x000000fb, nil)
	com.AddToScope(retVal)
	return *retVal
}

var Range_Find_OptArgs = []string{
	"After", "LookIn", "LookAt", "SearchOrder",
	"SearchDirection", "MatchCase", "MatchByte", "SearchFormat",
}

func (this *Range) Find(what interface{}, optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Range_Find_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000018e, []interface{}{what}, optArgs...)
	return NewRange(retVal.IDispatch(), false, true)
}

var Range_FindNext_OptArgs = []string{
	"After",
}

func (this *Range) FindNext(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Range_FindNext_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000018f, nil, optArgs...)
	return NewRange(retVal.IDispatch(), false, true)
}

var Range_FindPrevious_OptArgs = []string{
	"After",
}

func (this *Range) FindPrevious(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Range_FindPrevious_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000190, nil, optArgs...)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Range) Font() *Font {
	retVal, _ := this.PropGet(0x00000092, nil)
	return NewFont(retVal.IDispatch(), false, true)
}

func (this *Range) Formula() ole.Variant {
	retVal, _ := this.PropGet(0x00000105, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetFormula(rhs interface{}) {
	_ = this.PropPut(0x00000105, []interface{}{rhs})
}

func (this *Range) FormulaArray() ole.Variant {
	retVal, _ := this.PropGet(0x0000024a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetFormulaArray(rhs interface{}) {
	_ = this.PropPut(0x0000024a, []interface{}{rhs})
}

func (this *Range) FormulaLabel() int32 {
	retVal, _ := this.PropGet(0x00000564, nil)
	return retVal.LValVal()
}

func (this *Range) SetFormulaLabel(rhs int32) {
	_ = this.PropPut(0x00000564, []interface{}{rhs})
}

func (this *Range) FormulaHidden() ole.Variant {
	retVal, _ := this.PropGet(0x00000106, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetFormulaHidden(rhs interface{}) {
	_ = this.PropPut(0x00000106, []interface{}{rhs})
}

func (this *Range) FormulaLocal() ole.Variant {
	retVal, _ := this.PropGet(0x00000107, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetFormulaLocal(rhs interface{}) {
	_ = this.PropPut(0x00000107, []interface{}{rhs})
}

func (this *Range) FormulaR1C1() ole.Variant {
	retVal, _ := this.PropGet(0x00000108, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetFormulaR1C1(rhs interface{}) {
	_ = this.PropPut(0x00000108, []interface{}{rhs})
}

func (this *Range) FormulaR1C1Local() ole.Variant {
	retVal, _ := this.PropGet(0x00000109, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetFormulaR1C1Local(rhs interface{}) {
	_ = this.PropPut(0x00000109, []interface{}{rhs})
}

func (this *Range) FunctionWizard() ole.Variant {
	retVal, _ := this.Call(0x0000023b, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) GoalSeek(goal interface{}, changingCell *Range) bool {
	retVal, _ := this.Call(0x000001d8, []interface{}{goal, changingCell})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var Range_Group_OptArgs = []string{
	"Start", "End", "By", "Periods",
}

func (this *Range) Group(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Group_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000002e, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) HasArray() ole.Variant {
	retVal, _ := this.PropGet(0x0000010a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) HasFormula() ole.Variant {
	retVal, _ := this.PropGet(0x0000010b, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) Height() ole.Variant {
	retVal, _ := this.PropGet(0x0000007b, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) Hidden() ole.Variant {
	retVal, _ := this.PropGet(0x0000010c, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetHidden(rhs interface{}) {
	_ = this.PropPut(0x0000010c, []interface{}{rhs})
}

func (this *Range) HorizontalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x00000088, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetHorizontalAlignment(rhs interface{}) {
	_ = this.PropPut(0x00000088, []interface{}{rhs})
}

func (this *Range) IndentLevel() ole.Variant {
	retVal, _ := this.PropGet(0x000000c9, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetIndentLevel(rhs interface{}) {
	_ = this.PropPut(0x000000c9, []interface{}{rhs})
}

func (this *Range) InsertIndent(insertAmount int32) {
	retVal, _ := this.Call(0x00000565, []interface{}{insertAmount})
	_ = retVal
}

var Range_Insert_OptArgs = []string{
	"Shift", "CopyOrigin",
}

func (this *Range) Insert(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Insert_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000fc, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

var Range_Item_OptArgs = []string{
	"ColumnIndex",
}

func (this *Range) Item(rowIndex interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Item_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x000000aa, []interface{}{rowIndex}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Range_SetItem_OptArgs = []string{
	"ColumnIndex",
}

func (this *Range) SetItem(rowIndex interface{}, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Range_SetItem_OptArgs, optArgs)
	_ = this.PropPut(0x000000aa, []interface{}{rowIndex}, optArgs...)
}

func (this *Range) Justify() ole.Variant {
	retVal, _ := this.Call(0x000001ef, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) Left() ole.Variant {
	retVal, _ := this.PropGet(0x0000007f, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) ListHeaderRows() int32 {
	retVal, _ := this.PropGet(0x000004a3, nil)
	return retVal.LValVal()
}

func (this *Range) ListNames() ole.Variant {
	retVal, _ := this.Call(0x000000fd, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) LocationInTable() int32 {
	retVal, _ := this.PropGet(0x000002b3, nil)
	return retVal.LValVal()
}

func (this *Range) Locked() ole.Variant {
	retVal, _ := this.PropGet(0x0000010d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetLocked(rhs interface{}) {
	_ = this.PropPut(0x0000010d, []interface{}{rhs})
}

var Range_Merge_OptArgs = []string{
	"Across",
}

func (this *Range) Merge(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Range_Merge_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000234, nil, optArgs...)
	_ = retVal
}

func (this *Range) UnMerge() {
	retVal, _ := this.Call(0x00000568, nil)
	_ = retVal
}

func (this *Range) MergeArea() *Range {
	retVal, _ := this.PropGet(0x00000569, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Range) MergeCells() ole.Variant {
	retVal, _ := this.PropGet(0x000000d0, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetMergeCells(rhs interface{}) {
	_ = this.PropPut(0x000000d0, []interface{}{rhs})
}

func (this *Range) Name() ole.Variant {
	retVal, _ := this.PropGet(0x0000006e, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetName(rhs interface{}) {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

var Range_NavigateArrow_OptArgs = []string{
	"TowardPrecedent", "ArrowNumber", "LinkNumber",
}

func (this *Range) NavigateArrow(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_NavigateArrow_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000408, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Range) ForEach(action func(item ole.Variant) bool) {
	pEnum := this.NewEnum_()
	var pEnumVar *win32.IEnumVARIANT
	pEnum.QueryInterface(&win32.IID_IEnumVARIANT, unsafe.Pointer(&pEnumVar))
	defer pEnumVar.Release()
	for {
		var c uint32
		var v ole.Variant
		pEnumVar.Next(1, (*win32.VARIANT)(&v), &c)
		if c == 0 {
			break
		}
		pItem := v
		ret := action(pItem)
		if !ret {
			break
		}
	}
}

func (this *Range) Next() *Range {
	retVal, _ := this.PropGet(0x000001f6, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

var Range_NoteText_OptArgs = []string{
	"Text", "Start", "Length",
}

func (this *Range) NoteText(optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(Range_NoteText_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000467, nil, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Range) NumberFormat() ole.Variant {
	retVal, _ := this.PropGet(0x000000c1, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetNumberFormat(rhs interface{}) {
	_ = this.PropPut(0x000000c1, []interface{}{rhs})
}

func (this *Range) NumberFormatLocal() ole.Variant {
	retVal, _ := this.PropGet(0x00000449, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetNumberFormatLocal(rhs interface{}) {
	_ = this.PropPut(0x00000449, []interface{}{rhs})
}

var Range_Offset_OptArgs = []string{
	"RowOffset", "ColumnOffset",
}

func (this *Range) Offset(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Range_Offset_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x000000fe, nil, optArgs...)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Range) Orientation() ole.Variant {
	retVal, _ := this.PropGet(0x00000086, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetOrientation(rhs interface{}) {
	_ = this.PropPut(0x00000086, []interface{}{rhs})
}

func (this *Range) OutlineLevel() ole.Variant {
	retVal, _ := this.PropGet(0x0000010f, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetOutlineLevel(rhs interface{}) {
	_ = this.PropPut(0x0000010f, []interface{}{rhs})
}

func (this *Range) PageBreak() int32 {
	retVal, _ := this.PropGet(0x000000ff, nil)
	return retVal.LValVal()
}

func (this *Range) SetPageBreak(rhs int32) {
	_ = this.PropPut(0x000000ff, []interface{}{rhs})
}

var Range_Parse_OptArgs = []string{
	"ParseLine", "Destination",
}

func (this *Range) Parse(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Parse_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001dd, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Range_PasteSpecial__OptArgs = []string{
	"Paste", "Operation", "SkipBlanks", "Transpose",
}

func (this *Range) PasteSpecial_(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_PasteSpecial__OptArgs, optArgs)
	retVal, _ := this.Call(0x00000403, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) PivotField() *PivotField {
	retVal, _ := this.PropGet(0x000002db, nil)
	return NewPivotField(retVal.IDispatch(), false, true)
}

func (this *Range) PivotItem() *PivotItem {
	retVal, _ := this.PropGet(0x000002e4, nil)
	return NewPivotItem(retVal.IDispatch(), false, true)
}

func (this *Range) PivotTable() *PivotTable {
	retVal, _ := this.PropGet(0x000002cc, nil)
	return NewPivotTable(retVal.IDispatch(), false, true)
}

func (this *Range) Precedents() *Range {
	retVal, _ := this.PropGet(0x00000220, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Range) PrefixCharacter() ole.Variant {
	retVal, _ := this.PropGet(0x000001f8, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) Previous() *Range {
	retVal, _ := this.PropGet(0x000001f7, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

var Range_PrintOut___OptArgs = []string{
	"From", "To", "Copies", "Preview",
	"ActivePrinter", "PrintToFile", "Collate",
}

func (this *Range) PrintOut__(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_PrintOut___OptArgs, optArgs)
	retVal, _ := this.Call(0x00000389, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Range_PrintPreview_OptArgs = []string{
	"EnableChanges",
}

func (this *Range) PrintPreview(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_PrintPreview_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000119, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) QueryTable() *QueryTable {
	retVal, _ := this.PropGet(0x0000056a, nil)
	return NewQueryTable(retVal.IDispatch(), false, true)
}

var Range_Range_OptArgs = []string{
	"Cell2",
}

func (this *Range) Range(cell1 interface{}, optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Range_Range_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x000000c5, []interface{}{cell1}, optArgs...)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Range) RemoveSubtotal() ole.Variant {
	retVal, _ := this.Call(0x00000373, nil)
	com.AddToScope(retVal)
	return *retVal
}

var Range_Replace_OptArgs = []string{
	"LookAt", "SearchOrder", "MatchCase", "MatchByte",
	"SearchFormat", "ReplaceFormat",
}

func (this *Range) Replace(what interface{}, replacement interface{}, optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(Range_Replace_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000e2, []interface{}{what, replacement}, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var Range_Resize_OptArgs = []string{
	"RowSize", "ColumnSize",
}

func (this *Range) Resize(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Range_Resize_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x00000100, nil, optArgs...)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Range) Row() int32 {
	retVal, _ := this.PropGet(0x00000101, nil)
	return retVal.LValVal()
}

func (this *Range) RowDifferences(comparison interface{}) *Range {
	retVal, _ := this.Call(0x000001ff, []interface{}{comparison})
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Range) RowHeight() ole.Variant {
	retVal, _ := this.PropGet(0x00000110, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetRowHeight(rhs interface{}) {
	_ = this.PropPut(0x00000110, []interface{}{rhs})
}

func (this *Range) Rows() *Range {
	retVal, _ := this.PropGet(0x00000102, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

var Range_Run_OptArgs = []string{
	"Arg1", "Arg2", "Arg3", "Arg4",
	"Arg5", "Arg6", "Arg7", "Arg8",
	"Arg9", "Arg10", "Arg11", "Arg12",
	"Arg13", "Arg14", "Arg15", "Arg16",
	"Arg17", "Arg18", "Arg19", "Arg20",
	"Arg21", "Arg22", "Arg23", "Arg24",
	"Arg25", "Arg26", "Arg27", "Arg28",
	"Arg29", "Arg30",
}

func (this *Range) Run(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Run_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000103, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) Select() ole.Variant {
	retVal, _ := this.Call(0x000000eb, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) Show() ole.Variant {
	retVal, _ := this.Call(0x000001f0, nil)
	com.AddToScope(retVal)
	return *retVal
}

var Range_ShowDependents_OptArgs = []string{
	"Remove",
}

func (this *Range) ShowDependents(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_ShowDependents_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000036d, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) ShowDetail() ole.Variant {
	retVal, _ := this.PropGet(0x00000249, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetShowDetail(rhs interface{}) {
	_ = this.PropPut(0x00000249, []interface{}{rhs})
}

func (this *Range) ShowErrors() ole.Variant {
	retVal, _ := this.Call(0x0000036e, nil)
	com.AddToScope(retVal)
	return *retVal
}

var Range_ShowPrecedents_OptArgs = []string{
	"Remove",
}

func (this *Range) ShowPrecedents(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_ShowPrecedents_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000036f, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) ShrinkToFit() ole.Variant {
	retVal, _ := this.PropGet(0x000000d1, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetShrinkToFit(rhs interface{}) {
	_ = this.PropPut(0x000000d1, []interface{}{rhs})
}

var Range_Sort_OptArgs = []string{
	"Key1", "Order1", "Key2", "Type",
	"Order2", "Key3", "Order3", "Header",
	"OrderCustom", "MatchCase", "Orientation", "SortMethod",
	"DataOption1", "DataOption2", "DataOption3",
}

func (this *Range) Sort(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Sort_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000370, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Range_SortSpecial_OptArgs = []string{
	"SortMethod", "Key1", "Order1", "Type",
	"Key2", "Order2", "Key3", "Order3",
	"Header", "OrderCustom", "MatchCase", "Orientation",
	"DataOption1", "DataOption2", "DataOption3",
}

func (this *Range) SortSpecial(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_SortSpecial_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000371, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SoundNote() *SoundNote {
	retVal, _ := this.PropGet(0x00000394, nil)
	return NewSoundNote(retVal.IDispatch(), false, true)
}

var Range_SpecialCells_OptArgs = []string{
	"Value",
}

func (this *Range) SpecialCells(type_ int32, optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Range_SpecialCells_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000019a, []interface{}{type_}, optArgs...)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Range) Style() ole.Variant {
	retVal, _ := this.PropGet(0x00000104, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetStyle(rhs interface{}) {
	_ = this.PropPut(0x00000104, []interface{}{rhs})
}

var Range_SubscribeTo_OptArgs = []string{
	"Format",
}

func (this *Range) SubscribeTo(edition string, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_SubscribeTo_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001e1, []interface{}{edition}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Range_Subtotal_OptArgs = []string{
	"Replace", "PageBreaks", "SummaryBelowData",
}

func (this *Range) Subtotal(groupBy int32, function int32, totalList interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Subtotal_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000372, []interface{}{groupBy, function, totalList}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) Summary() ole.Variant {
	retVal, _ := this.PropGet(0x00000111, nil)
	com.AddToScope(retVal)
	return *retVal
}

var Range_Table_OptArgs = []string{
	"RowInput", "ColumnInput",
}

func (this *Range) Table(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Table_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001f1, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) Text() ole.Variant {
	retVal, _ := this.PropGet(0x0000008a, nil)
	com.AddToScope(retVal)
	return *retVal
}

var Range_TextToColumns_OptArgs = []string{
	"Destination", "DataType", "TextQualifier", "ConsecutiveDelimiter",
	"Tab", "Semicolon", "Comma", "Space",
	"Other", "OtherChar", "FieldInfo", "DecimalSeparator",
	"ThousandsSeparator", "TrailingMinusNumbers",
}

func (this *Range) TextToColumns(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_TextToColumns_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000410, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) Top() ole.Variant {
	retVal, _ := this.PropGet(0x0000007e, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) Ungroup() ole.Variant {
	retVal, _ := this.Call(0x000000f4, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) UseStandardHeight() ole.Variant {
	retVal, _ := this.PropGet(0x00000112, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetUseStandardHeight(rhs interface{}) {
	_ = this.PropPut(0x00000112, []interface{}{rhs})
}

func (this *Range) UseStandardWidth() ole.Variant {
	retVal, _ := this.PropGet(0x00000113, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetUseStandardWidth(rhs interface{}) {
	_ = this.PropPut(0x00000113, []interface{}{rhs})
}

func (this *Range) Validation() *Validation {
	retVal, _ := this.PropGet(0x0000056b, nil)
	return NewValidation(retVal.IDispatch(), false, true)
}

var Range_Value_OptArgs = []string{
	"RangeValueDataType",
}

func (this *Range) Value(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Value_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x00000006, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Range_SetValue_OptArgs = []string{
	"RangeValueDataType",
}

func (this *Range) SetValue(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Range_SetValue_OptArgs, optArgs)
	_ = this.PropPut(0x00000006, nil, optArgs...)
}

func (this *Range) Value2() ole.Variant {
	retVal, _ := this.PropGet(0x0000056c, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetValue2(rhs interface{}) {
	_ = this.PropPut(0x0000056c, []interface{}{rhs})
}

func (this *Range) VerticalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x00000089, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetVerticalAlignment(rhs interface{}) {
	_ = this.PropPut(0x00000089, []interface{}{rhs})
}

func (this *Range) Width() ole.Variant {
	retVal, _ := this.PropGet(0x0000007a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) Worksheet() *Worksheet {
	retVal, _ := this.PropGet(0x0000015c, nil)
	return NewWorksheet(retVal.IDispatch(), false, true)
}

func (this *Range) WrapText() ole.Variant {
	retVal, _ := this.PropGet(0x00000114, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SetWrapText(rhs interface{}) {
	_ = this.PropPut(0x00000114, []interface{}{rhs})
}

var Range_AddComment_OptArgs = []string{
	"Text",
}

func (this *Range) AddComment(optArgs ...interface{}) *Comment {
	optArgs = ole.ProcessOptArgs(Range_AddComment_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000056d, nil, optArgs...)
	return NewComment(retVal.IDispatch(), false, true)
}

func (this *Range) Comment() *Comment {
	retVal, _ := this.PropGet(0x0000038e, nil)
	return NewComment(retVal.IDispatch(), false, true)
}

func (this *Range) ClearComments() {
	retVal, _ := this.Call(0x0000056e, nil)
	_ = retVal
}

func (this *Range) Phonetic() *Phonetic {
	retVal, _ := this.PropGet(0x0000056f, nil)
	return NewPhonetic(retVal.IDispatch(), false, true)
}

func (this *Range) FormatConditions() *FormatConditions {
	retVal, _ := this.PropGet(0x00000570, nil)
	return NewFormatConditions(retVal.IDispatch(), false, true)
}

func (this *Range) ReadingOrder() int32 {
	retVal, _ := this.PropGet(0x000003cf, nil)
	return retVal.LValVal()
}

func (this *Range) SetReadingOrder(rhs int32) {
	_ = this.PropPut(0x000003cf, []interface{}{rhs})
}

func (this *Range) Hyperlinks() *Hyperlinks {
	retVal, _ := this.PropGet(0x00000571, nil)
	return NewHyperlinks(retVal.IDispatch(), false, true)
}

func (this *Range) Phonetics() *Phonetics {
	retVal, _ := this.PropGet(0x00000713, nil)
	return NewPhonetics(retVal.IDispatch(), false, true)
}

func (this *Range) SetPhonetic() {
	retVal, _ := this.Call(0x00000714, nil)
	_ = retVal
}

func (this *Range) ID() string {
	retVal, _ := this.PropGet(0x00000715, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Range) SetID(rhs string) {
	_ = this.PropPut(0x00000715, []interface{}{rhs})
}

var Range_PrintOut__OptArgs = []string{
	"From", "To", "Copies", "Preview",
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName",
}

func (this *Range) PrintOut_(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_PrintOut__OptArgs, optArgs)
	retVal, _ := this.Call(0x000006ec, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) PivotCell() *PivotCell {
	retVal, _ := this.PropGet(0x000007dd, nil)
	return NewPivotCell(retVal.IDispatch(), false, true)
}

func (this *Range) Dirty() {
	retVal, _ := this.Call(0x000007de, nil)
	_ = retVal
}

func (this *Range) Errors() *Errors {
	retVal, _ := this.PropGet(0x000007df, nil)
	return NewErrors(retVal.IDispatch(), false, true)
}

func (this *Range) SmartTags() *SmartTags {
	retVal, _ := this.PropGet(0x000007e0, nil)
	return NewSmartTags(retVal.IDispatch(), false, true)
}

var Range_Speak_OptArgs = []string{
	"SpeakDirection", "SpeakFormulas",
}

func (this *Range) Speak(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Range_Speak_OptArgs, optArgs)
	retVal, _ := this.Call(0x000007e1, nil, optArgs...)
	_ = retVal
}

var Range_PasteSpecial_OptArgs = []string{
	"Paste", "Operation", "SkipBlanks", "Transpose",
}

func (this *Range) PasteSpecial(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_PasteSpecial_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000788, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) AllowEdit() bool {
	retVal, _ := this.PropGet(0x000007e4, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Range) ListObject() *ListObject {
	retVal, _ := this.PropGet(0x000008d1, nil)
	return NewListObject(retVal.IDispatch(), false, true)
}

func (this *Range) XPath() *XPath {
	retVal, _ := this.PropGet(0x000008d2, nil)
	return NewXPath(retVal.IDispatch(), false, true)
}

func (this *Range) ServerActions() *Actions {
	retVal, _ := this.PropGet(0x000009bb, nil)
	return NewActions(retVal.IDispatch(), false, true)
}

var Range_RemoveDuplicates_OptArgs = []string{
	"Columns", "Header",
}

func (this *Range) RemoveDuplicates(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Range_RemoveDuplicates_OptArgs, optArgs)
	retVal, _ := this.Call(0x000009bc, nil, optArgs...)
	_ = retVal
}

var Range_PrintOut_OptArgs = []string{
	"From", "To", "Copies", "Preview",
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName",
}

func (this *Range) PrintOut(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_PrintOut_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000939, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) MDX() string {
	retVal, _ := this.PropGet(0x0000084b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var Range_ExportAsFixedFormat_OptArgs = []string{
	"Filename", "Quality", "IncludeDocProperties", "IgnorePrintAreas",
	"From", "To", "OpenAfterPublish", "FixedFormatExtClassPtr",
}

func (this *Range) ExportAsFixedFormat(type_ int32, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Range_ExportAsFixedFormat_OptArgs, optArgs)
	retVal, _ := this.Call(0x000009bd, []interface{}{type_}, optArgs...)
	_ = retVal
}

func (this *Range) CountLarge() ole.Variant {
	retVal, _ := this.PropGet(0x000009c3, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) CalculateRowMajorOrder() ole.Variant {
	retVal, _ := this.Call(0x0000093c, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) SparklineGroups() *SparklineGroups {
	retVal, _ := this.PropGet(0x00000b25, nil)
	return NewSparklineGroups(retVal.IDispatch(), false, true)
}

func (this *Range) ClearHyperlinks() {
	retVal, _ := this.Call(0x00000b26, nil)
	_ = retVal
}

func (this *Range) DisplayFormat() *DisplayFormat {
	retVal, _ := this.PropGet(0x0000029a, nil)
	return NewDisplayFormat(retVal.IDispatch(), false, true)
}

var Range_BorderAround_OptArgs = []string{
	"LineStyle", "Weight", "ColorIndex", "Color", "ThemeColor",
}

func (this *Range) BorderAround(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_BorderAround_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000ad3, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Range) AllocateChanges() {
	retVal, _ := this.Call(0x00000b27, nil)
	_ = retVal
}

func (this *Range) DiscardChanges() {
	retVal, _ := this.Call(0x00000b28, nil)
	_ = retVal
}
