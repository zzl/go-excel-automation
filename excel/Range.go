package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
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
	return NewRange(v.PdispValVal(), false, false)
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

func (this *Range) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Range) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Range) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Range) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Range) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Range) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Range) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Range) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Range) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Range) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Range) Activate() ole.Variant {
	retVal := this.Call(0x00000130, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) AddIndent() ole.Variant {
	retVal := this.PropGet(0x00000427, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetAddIndent(rhs interface{})  {
	retVal := this.PropPut(0x00000427, []interface{}{rhs})
	_= retVal
}

var Range_Address_OptArgs= []string{
	"External", "RelativeTo", 
}

func (this *Range) Address(rowAbsolute interface{}, columnAbsolute interface{}, referenceStyle int32, optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(Range_Address_OptArgs, optArgs)
	retVal := this.PropGet(0x000000ec, []interface{}{rowAbsolute, columnAbsolute, referenceStyle}, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var Range_AddressLocal_OptArgs= []string{
	"External", "RelativeTo", 
}

func (this *Range) AddressLocal(rowAbsolute interface{}, columnAbsolute interface{}, referenceStyle int32, optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(Range_AddressLocal_OptArgs, optArgs)
	retVal := this.PropGet(0x000001b5, []interface{}{rowAbsolute, columnAbsolute, referenceStyle}, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var Range_AdvancedFilter_OptArgs= []string{
	"CriteriaRange", "CopyToRange", "Unique", 
}

func (this *Range) AdvancedFilter(action int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_AdvancedFilter_OptArgs, optArgs)
	retVal := this.Call(0x0000036c, []interface{}{action}, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Range_ApplyNames_OptArgs= []string{
	"AppendLast", 
}

func (this *Range) ApplyNames(names interface{}, ignoreRelativeAbsolute interface{}, useRowColumnNames interface{}, omitColumn interface{}, omitRow interface{}, order int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_ApplyNames_OptArgs, optArgs)
	retVal := this.Call(0x000001b9, []interface{}{names, ignoreRelativeAbsolute, useRowColumnNames, omitColumn, omitRow, order}, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) ApplyOutlineStyles() ole.Variant {
	retVal := this.Call(0x000001c0, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) Areas() *Areas {
	retVal := this.PropGet(0x00000238, nil)
	return NewAreas(retVal.PdispValVal(), false, true)
}

func (this *Range) AutoComplete(string string) string {
	retVal := this.Call(0x000004a1, []interface{}{string})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Range) AutoFill(destination *Range, type_ int32) ole.Variant {
	retVal := this.Call(0x000001c1, []interface{}{destination, type_})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Range_AutoFilter_OptArgs= []string{
	"Criteria2", "VisibleDropDown", 
}

func (this *Range) AutoFilter(field interface{}, criteria1 interface{}, operator int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_AutoFilter_OptArgs, optArgs)
	retVal := this.Call(0x00000319, []interface{}{field, criteria1, operator}, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) AutoFit() ole.Variant {
	retVal := this.Call(0x000000ed, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Range_AutoFormat_OptArgs= []string{
	"Number", "Font", "Alignment", "Border", 
	"Pattern", "Width", 
}

func (this *Range) AutoFormat(format int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_AutoFormat_OptArgs, optArgs)
	retVal := this.Call(0x00000072, []interface{}{format}, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) AutoOutline() ole.Variant {
	retVal := this.Call(0x0000040c, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Range_BorderAround__OptArgs= []string{
	"Color", 
}

func (this *Range) BorderAround_(lineStyle interface{}, weight int32, colorIndex int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_BorderAround__OptArgs, optArgs)
	retVal := this.Call(0x0000042b, []interface{}{lineStyle, weight, colorIndex}, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) Borders() *Borders {
	retVal := this.PropGet(0x000001b3, nil)
	return NewBorders(retVal.PdispValVal(), false, true)
}

func (this *Range) Calculate() ole.Variant {
	retVal := this.Call(0x00000117, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) Cells() *Range {
	retVal := this.PropGet(0x000000ee, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

var Range_Characters_OptArgs= []string{
	"Start", "Length", 
}

func (this *Range) Characters(optArgs ...interface{}) *Characters {
	optArgs = ole.ProcessOptArgs(Range_Characters_OptArgs, optArgs)
	retVal := this.PropGet(0x0000025b, nil, optArgs...)
	return NewCharacters(retVal.PdispValVal(), false, true)
}

var Range_CheckSpelling_OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang", 
}

func (this *Range) CheckSpelling(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_CheckSpelling_OptArgs, optArgs)
	retVal := this.Call(0x000001f9, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) Clear() ole.Variant {
	retVal := this.Call(0x0000006f, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) ClearContents() ole.Variant {
	retVal := this.Call(0x00000071, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) ClearFormats() ole.Variant {
	retVal := this.Call(0x00000070, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) ClearNotes() ole.Variant {
	retVal := this.Call(0x000000ef, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) ClearOutline() ole.Variant {
	retVal := this.Call(0x0000040d, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) Column() int32 {
	retVal := this.PropGet(0x000000f0, nil)
	return retVal.LValVal()
}

func (this *Range) ColumnDifferences(comparison interface{}) *Range {
	retVal := this.Call(0x000001fe, []interface{}{comparison})
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Range) Columns() *Range {
	retVal := this.PropGet(0x000000f1, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Range) ColumnWidth() ole.Variant {
	retVal := this.PropGet(0x000000f2, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetColumnWidth(rhs interface{})  {
	retVal := this.PropPut(0x000000f2, []interface{}{rhs})
	_= retVal
}

var Range_Consolidate_OptArgs= []string{
	"Sources", "Function", "TopRow", "LeftColumn", "CreateLinks", 
}

func (this *Range) Consolidate(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Consolidate_OptArgs, optArgs)
	retVal := this.Call(0x000001e2, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Range_Copy_OptArgs= []string{
	"Destination", 
}

func (this *Range) Copy(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Copy_OptArgs, optArgs)
	retVal := this.Call(0x00000227, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Range_CopyFromRecordset_OptArgs= []string{
	"MaxRows", "MaxColumns", 
}

func (this *Range) CopyFromRecordset(data *com.UnknownClass, optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Range_CopyFromRecordset_OptArgs, optArgs)
	retVal := this.Call(0x00000480, []interface{}{data}, optArgs...)
	return retVal.LValVal()
}

func (this *Range) CopyPicture(appearance int32, format int32) ole.Variant {
	retVal := this.Call(0x000000d5, []interface{}{appearance, format})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

var Range_CreateNames_OptArgs= []string{
	"Top", "Left", "Bottom", "Right", 
}

func (this *Range) CreateNames(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_CreateNames_OptArgs, optArgs)
	retVal := this.Call(0x000001c9, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Range_CreatePublisher_OptArgs= []string{
	"ContainsPICT", "ContainsBIFF", "ContainsRTF", "ContainsVALU", 
}

func (this *Range) CreatePublisher(edition interface{}, appearance int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_CreatePublisher_OptArgs, optArgs)
	retVal := this.Call(0x000001ca, []interface{}{edition, appearance}, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) CurrentArray() *Range {
	retVal := this.PropGet(0x000001f5, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Range) CurrentRegion() *Range {
	retVal := this.PropGet(0x000000f3, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

var Range_Cut_OptArgs= []string{
	"Destination", 
}

func (this *Range) Cut(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Cut_OptArgs, optArgs)
	retVal := this.Call(0x00000235, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Range_DataSeries_OptArgs= []string{
	"Step", "Stop", "Trend", 
}

func (this *Range) DataSeries(rowcol interface{}, type_ int32, date int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_DataSeries_OptArgs, optArgs)
	retVal := this.Call(0x000001d0, []interface{}{rowcol, type_, date}, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Range_Default__OptArgs= []string{
	"RowIndex", "ColumnIndex", 
}

func (this *Range) Default_(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Default__OptArgs, optArgs)
	retVal := this.PropGet(0x00000000, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Range_SetDefault__OptArgs= []string{
	"ColumnIndex", "rhs", 
}

func (this *Range) SetDefault_(rowIndex interface{}, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_SetDefault__OptArgs, optArgs)
	retVal := this.PropPut(0x00000000, []interface{}{rowIndex}, optArgs...)
	_= retVal
}

var Range_Delete_OptArgs= []string{
	"Shift", 
}

func (this *Range) Delete(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Delete_OptArgs, optArgs)
	retVal := this.Call(0x00000075, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) Dependents() *Range {
	retVal := this.PropGet(0x0000021f, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Range) DialogBox() ole.Variant {
	retVal := this.Call(0x000000f5, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) DirectDependents() *Range {
	retVal := this.PropGet(0x00000221, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Range) DirectPrecedents() *Range {
	retVal := this.PropGet(0x00000222, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

var Range_EditionOptions_OptArgs= []string{
	"Format", 
}

func (this *Range) EditionOptions(type_ int32, option int32, name interface{}, reference interface{}, appearance int32, chartSize int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_EditionOptions_OptArgs, optArgs)
	retVal := this.Call(0x0000046b, []interface{}{type_, option, name, reference, appearance, chartSize}, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) End(direction int32) *Range {
	retVal := this.PropGet(0x000001f4, []interface{}{direction})
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Range) EntireColumn() *Range {
	retVal := this.PropGet(0x000000f6, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Range) EntireRow() *Range {
	retVal := this.PropGet(0x000000f7, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Range) FillDown() ole.Variant {
	retVal := this.Call(0x000000f8, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) FillLeft() ole.Variant {
	retVal := this.Call(0x000000f9, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) FillRight() ole.Variant {
	retVal := this.Call(0x000000fa, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) FillUp() ole.Variant {
	retVal := this.Call(0x000000fb, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Range_Find_OptArgs= []string{
	"MatchCase", "MatchByte", "SearchFormat", 
}

func (this *Range) Find(what interface{}, after interface{}, lookIn interface{}, lookAt interface{}, searchOrder interface{}, searchDirection int32, optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Range_Find_OptArgs, optArgs)
	retVal := this.Call(0x0000018e, []interface{}{what, after, lookIn, lookAt, searchOrder, searchDirection}, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

var Range_FindNext_OptArgs= []string{
	"After", 
}

func (this *Range) FindNext(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Range_FindNext_OptArgs, optArgs)
	retVal := this.Call(0x0000018f, nil, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

var Range_FindPrevious_OptArgs= []string{
	"After", 
}

func (this *Range) FindPrevious(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Range_FindPrevious_OptArgs, optArgs)
	retVal := this.Call(0x00000190, nil, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Range) Font() *Font {
	retVal := this.PropGet(0x00000092, nil)
	return NewFont(retVal.PdispValVal(), false, true)
}

func (this *Range) Formula() ole.Variant {
	retVal := this.PropGet(0x00000105, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetFormula(rhs interface{})  {
	retVal := this.PropPut(0x00000105, []interface{}{rhs})
	_= retVal
}

func (this *Range) FormulaArray() ole.Variant {
	retVal := this.PropGet(0x0000024a, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetFormulaArray(rhs interface{})  {
	retVal := this.PropPut(0x0000024a, []interface{}{rhs})
	_= retVal
}

func (this *Range) FormulaLabel() int32 {
	retVal := this.PropGet(0x00000564, nil)
	return retVal.LValVal()
}

func (this *Range) SetFormulaLabel(rhs int32)  {
	retVal := this.PropPut(0x00000564, []interface{}{rhs})
	_= retVal
}

func (this *Range) FormulaHidden() ole.Variant {
	retVal := this.PropGet(0x00000106, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetFormulaHidden(rhs interface{})  {
	retVal := this.PropPut(0x00000106, []interface{}{rhs})
	_= retVal
}

func (this *Range) FormulaLocal() ole.Variant {
	retVal := this.PropGet(0x00000107, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetFormulaLocal(rhs interface{})  {
	retVal := this.PropPut(0x00000107, []interface{}{rhs})
	_= retVal
}

func (this *Range) FormulaR1C1() ole.Variant {
	retVal := this.PropGet(0x00000108, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetFormulaR1C1(rhs interface{})  {
	retVal := this.PropPut(0x00000108, []interface{}{rhs})
	_= retVal
}

func (this *Range) FormulaR1C1Local() ole.Variant {
	retVal := this.PropGet(0x00000109, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetFormulaR1C1Local(rhs interface{})  {
	retVal := this.PropPut(0x00000109, []interface{}{rhs})
	_= retVal
}

func (this *Range) FunctionWizard() ole.Variant {
	retVal := this.Call(0x0000023b, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) GoalSeek(goal interface{}, changingCell *Range) bool {
	retVal := this.Call(0x000001d8, []interface{}{goal, changingCell})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var Range_Group_OptArgs= []string{
	"Start", "End", "By", "Periods", 
}

func (this *Range) Group(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Group_OptArgs, optArgs)
	retVal := this.Call(0x0000002e, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) HasArray() ole.Variant {
	retVal := this.PropGet(0x0000010a, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) HasFormula() ole.Variant {
	retVal := this.PropGet(0x0000010b, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) Height() ole.Variant {
	retVal := this.PropGet(0x0000007b, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) Hidden() ole.Variant {
	retVal := this.PropGet(0x0000010c, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetHidden(rhs interface{})  {
	retVal := this.PropPut(0x0000010c, []interface{}{rhs})
	_= retVal
}

func (this *Range) HorizontalAlignment() ole.Variant {
	retVal := this.PropGet(0x00000088, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetHorizontalAlignment(rhs interface{})  {
	retVal := this.PropPut(0x00000088, []interface{}{rhs})
	_= retVal
}

func (this *Range) IndentLevel() ole.Variant {
	retVal := this.PropGet(0x000000c9, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetIndentLevel(rhs interface{})  {
	retVal := this.PropPut(0x000000c9, []interface{}{rhs})
	_= retVal
}

func (this *Range) InsertIndent(insertAmount int32)  {
	retVal := this.Call(0x00000565, []interface{}{insertAmount})
	_= retVal
}

var Range_Insert_OptArgs= []string{
	"Shift", "CopyOrigin", 
}

func (this *Range) Insert(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Insert_OptArgs, optArgs)
	retVal := this.Call(0x000000fc, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) Interior() *Interior {
	retVal := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.PdispValVal(), false, true)
}

var Range_Item_OptArgs= []string{
	"ColumnIndex", 
}

func (this *Range) Item(rowIndex interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Item_OptArgs, optArgs)
	retVal := this.PropGet(0x000000aa, []interface{}{rowIndex}, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Range_SetItem_OptArgs= []string{
	"rhs", 
}

func (this *Range) SetItem(rowIndex interface{}, columnIndex interface{}, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_SetItem_OptArgs, optArgs)
	retVal := this.PropPut(0x000000aa, []interface{}{rowIndex, columnIndex}, optArgs...)
	_= retVal
}

func (this *Range) Justify() ole.Variant {
	retVal := this.Call(0x000001ef, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) Left() ole.Variant {
	retVal := this.PropGet(0x0000007f, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) ListHeaderRows() int32 {
	retVal := this.PropGet(0x000004a3, nil)
	return retVal.LValVal()
}

func (this *Range) ListNames() ole.Variant {
	retVal := this.Call(0x000000fd, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) LocationInTable() int32 {
	retVal := this.PropGet(0x000002b3, nil)
	return retVal.LValVal()
}

func (this *Range) Locked() ole.Variant {
	retVal := this.PropGet(0x0000010d, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetLocked(rhs interface{})  {
	retVal := this.PropPut(0x0000010d, []interface{}{rhs})
	_= retVal
}

var Range_Merge_OptArgs= []string{
	"Across", 
}

func (this *Range) Merge(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_Merge_OptArgs, optArgs)
	retVal := this.Call(0x00000234, nil, optArgs...)
	_= retVal
}

func (this *Range) UnMerge()  {
	retVal := this.Call(0x00000568, nil)
	_= retVal
}

func (this *Range) MergeArea() *Range {
	retVal := this.PropGet(0x00000569, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Range) MergeCells() ole.Variant {
	retVal := this.PropGet(0x000000d0, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetMergeCells(rhs interface{})  {
	retVal := this.PropPut(0x000000d0, []interface{}{rhs})
	_= retVal
}

func (this *Range) Name() ole.Variant {
	retVal := this.PropGet(0x0000006e, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetName(rhs interface{})  {
	retVal := this.PropPut(0x0000006e, []interface{}{rhs})
	_= retVal
}

var Range_NavigateArrow_OptArgs= []string{
	"TowardPrecedent", "ArrowNumber", "LinkNumber", 
}

func (this *Range) NavigateArrow(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_NavigateArrow_OptArgs, optArgs)
	retVal := this.Call(0x00000408, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Range) ForEach(action func(item ole.Variant) bool) {
	pEnum := this.NewEnum_()
	var pEnumVar *win32.IEnumVARIANT
	pEnum.QueryInterface(&win32.IID_IEnumVARIANT, unsafe.Pointer(&pEnumVar))
	defer pEnumVar.Release();
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
	retVal := this.PropGet(0x000001f6, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

var Range_NoteText_OptArgs= []string{
	"Text", "Start", "Length", 
}

func (this *Range) NoteText(optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(Range_NoteText_OptArgs, optArgs)
	retVal := this.Call(0x00000467, nil, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Range) NumberFormat() ole.Variant {
	retVal := this.PropGet(0x000000c1, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetNumberFormat(rhs interface{})  {
	retVal := this.PropPut(0x000000c1, []interface{}{rhs})
	_= retVal
}

func (this *Range) NumberFormatLocal() ole.Variant {
	retVal := this.PropGet(0x00000449, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetNumberFormatLocal(rhs interface{})  {
	retVal := this.PropPut(0x00000449, []interface{}{rhs})
	_= retVal
}

var Range_Offset_OptArgs= []string{
	"RowOffset", "ColumnOffset", 
}

func (this *Range) Offset(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Range_Offset_OptArgs, optArgs)
	retVal := this.PropGet(0x000000fe, nil, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Range) Orientation() ole.Variant {
	retVal := this.PropGet(0x00000086, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetOrientation(rhs interface{})  {
	retVal := this.PropPut(0x00000086, []interface{}{rhs})
	_= retVal
}

func (this *Range) OutlineLevel() ole.Variant {
	retVal := this.PropGet(0x0000010f, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetOutlineLevel(rhs interface{})  {
	retVal := this.PropPut(0x0000010f, []interface{}{rhs})
	_= retVal
}

func (this *Range) PageBreak() int32 {
	retVal := this.PropGet(0x000000ff, nil)
	return retVal.LValVal()
}

func (this *Range) SetPageBreak(rhs int32)  {
	retVal := this.PropPut(0x000000ff, []interface{}{rhs})
	_= retVal
}

var Range_Parse_OptArgs= []string{
	"ParseLine", "Destination", 
}

func (this *Range) Parse(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Parse_OptArgs, optArgs)
	retVal := this.Call(0x000001dd, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Range_PasteSpecial__OptArgs= []string{
	"SkipBlanks", "Transpose", 
}

func (this *Range) PasteSpecial_(paste int32, operation int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_PasteSpecial__OptArgs, optArgs)
	retVal := this.Call(0x00000403, []interface{}{paste, operation}, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) PivotField() *PivotField {
	retVal := this.PropGet(0x000002db, nil)
	return NewPivotField(retVal.PdispValVal(), false, true)
}

func (this *Range) PivotItem() *PivotItem {
	retVal := this.PropGet(0x000002e4, nil)
	return NewPivotItem(retVal.PdispValVal(), false, true)
}

func (this *Range) PivotTable() *PivotTable {
	retVal := this.PropGet(0x000002cc, nil)
	return NewPivotTable(retVal.PdispValVal(), false, true)
}

func (this *Range) Precedents() *Range {
	retVal := this.PropGet(0x00000220, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Range) PrefixCharacter() ole.Variant {
	retVal := this.PropGet(0x000001f8, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) Previous() *Range {
	retVal := this.PropGet(0x000001f7, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

var Range_PrintOut___OptArgs= []string{
	"From", "To", "Copies", "Preview", 
	"ActivePrinter", "PrintToFile", "Collate", 
}

func (this *Range) PrintOut__(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_PrintOut___OptArgs, optArgs)
	retVal := this.Call(0x00000389, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Range_PrintPreview_OptArgs= []string{
	"EnableChanges", 
}

func (this *Range) PrintPreview(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_PrintPreview_OptArgs, optArgs)
	retVal := this.Call(0x00000119, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) QueryTable() *QueryTable {
	retVal := this.PropGet(0x0000056a, nil)
	return NewQueryTable(retVal.PdispValVal(), false, true)
}

var Range_Range_OptArgs= []string{
	"Cell2", 
}

func (this *Range) Range(cell1 interface{}, optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Range_Range_OptArgs, optArgs)
	retVal := this.PropGet(0x000000c5, []interface{}{cell1}, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Range) RemoveSubtotal() ole.Variant {
	retVal := this.Call(0x00000373, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Range_Replace_OptArgs= []string{
	"LookAt", "SearchOrder", "MatchCase", "MatchByte", 
	"SearchFormat", "ReplaceFormat", 
}

func (this *Range) Replace(what interface{}, replacement interface{}, optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(Range_Replace_OptArgs, optArgs)
	retVal := this.Call(0x000000e2, []interface{}{what, replacement}, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var Range_Resize_OptArgs= []string{
	"RowSize", "ColumnSize", 
}

func (this *Range) Resize(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Range_Resize_OptArgs, optArgs)
	retVal := this.PropGet(0x00000100, nil, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Range) Row() int32 {
	retVal := this.PropGet(0x00000101, nil)
	return retVal.LValVal()
}

func (this *Range) RowDifferences(comparison interface{}) *Range {
	retVal := this.Call(0x000001ff, []interface{}{comparison})
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Range) RowHeight() ole.Variant {
	retVal := this.PropGet(0x00000110, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetRowHeight(rhs interface{})  {
	retVal := this.PropPut(0x00000110, []interface{}{rhs})
	_= retVal
}

func (this *Range) Rows() *Range {
	retVal := this.PropGet(0x00000102, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

var Range_Run_OptArgs= []string{
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
	retVal := this.Call(0x00000103, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) Select() ole.Variant {
	retVal := this.Call(0x000000eb, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) Show() ole.Variant {
	retVal := this.Call(0x000001f0, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Range_ShowDependents_OptArgs= []string{
	"Remove", 
}

func (this *Range) ShowDependents(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_ShowDependents_OptArgs, optArgs)
	retVal := this.Call(0x0000036d, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) ShowDetail() ole.Variant {
	retVal := this.PropGet(0x00000249, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetShowDetail(rhs interface{})  {
	retVal := this.PropPut(0x00000249, []interface{}{rhs})
	_= retVal
}

func (this *Range) ShowErrors() ole.Variant {
	retVal := this.Call(0x0000036e, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Range_ShowPrecedents_OptArgs= []string{
	"Remove", 
}

func (this *Range) ShowPrecedents(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_ShowPrecedents_OptArgs, optArgs)
	retVal := this.Call(0x0000036f, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) ShrinkToFit() ole.Variant {
	retVal := this.PropGet(0x000000d1, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetShrinkToFit(rhs interface{})  {
	retVal := this.PropPut(0x000000d1, []interface{}{rhs})
	_= retVal
}

func (this *Range) Sort(key1 interface{}, order1 int32, key2 interface{}, type_ interface{}, order2 int32, key3 interface{}, order3 int32, header int32, orderCustom interface{}, matchCase interface{}, orientation int32, sortMethod int32, dataOption1 int32, dataOption2 int32, dataOption3 int32) ole.Variant {
	retVal := this.Call(0x00000370, []interface{}{key1, order1, key2, type_, order2, key3, order3, header, orderCustom, matchCase, orientation, sortMethod, dataOption1, dataOption2, dataOption3})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SortSpecial(sortMethod int32, key1 interface{}, order1 int32, type_ interface{}, key2 interface{}, order2 int32, key3 interface{}, order3 int32, header int32, orderCustom interface{}, matchCase interface{}, orientation int32, dataOption1 int32, dataOption2 int32, dataOption3 int32) ole.Variant {
	retVal := this.Call(0x00000371, []interface{}{sortMethod, key1, order1, type_, key2, order2, key3, order3, header, orderCustom, matchCase, orientation, dataOption1, dataOption2, dataOption3})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SoundNote() *SoundNote {
	retVal := this.PropGet(0x00000394, nil)
	return NewSoundNote(retVal.PdispValVal(), false, true)
}

var Range_SpecialCells_OptArgs= []string{
	"Value", 
}

func (this *Range) SpecialCells(type_ int32, optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Range_SpecialCells_OptArgs, optArgs)
	retVal := this.Call(0x0000019a, []interface{}{type_}, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Range) Style() ole.Variant {
	retVal := this.PropGet(0x00000104, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetStyle(rhs interface{})  {
	retVal := this.PropPut(0x00000104, []interface{}{rhs})
	_= retVal
}

func (this *Range) SubscribeTo(edition string, format int32) ole.Variant {
	retVal := this.Call(0x000001e1, []interface{}{edition, format})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) Subtotal(groupBy int32, function int32, totalList interface{}, replace interface{}, pageBreaks interface{}, summaryBelowData int32) ole.Variant {
	retVal := this.Call(0x00000372, []interface{}{groupBy, function, totalList, replace, pageBreaks, summaryBelowData})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) Summary() ole.Variant {
	retVal := this.PropGet(0x00000111, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Range_Table_OptArgs= []string{
	"RowInput", "ColumnInput", 
}

func (this *Range) Table(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Table_OptArgs, optArgs)
	retVal := this.Call(0x000001f1, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) Text() ole.Variant {
	retVal := this.PropGet(0x0000008a, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Range_TextToColumns_OptArgs= []string{
	"ConsecutiveDelimiter", "Tab", "Semicolon", "Comma", 
	"Space", "Other", "OtherChar", "FieldInfo", 
	"DecimalSeparator", "ThousandsSeparator", "TrailingMinusNumbers", 
}

func (this *Range) TextToColumns(destination interface{}, dataType int32, textQualifier int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_TextToColumns_OptArgs, optArgs)
	retVal := this.Call(0x00000410, []interface{}{destination, dataType, textQualifier}, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) Top() ole.Variant {
	retVal := this.PropGet(0x0000007e, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) Ungroup() ole.Variant {
	retVal := this.Call(0x000000f4, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) UseStandardHeight() ole.Variant {
	retVal := this.PropGet(0x00000112, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetUseStandardHeight(rhs interface{})  {
	retVal := this.PropPut(0x00000112, []interface{}{rhs})
	_= retVal
}

func (this *Range) UseStandardWidth() ole.Variant {
	retVal := this.PropGet(0x00000113, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetUseStandardWidth(rhs interface{})  {
	retVal := this.PropPut(0x00000113, []interface{}{rhs})
	_= retVal
}

func (this *Range) Validation() *Validation {
	retVal := this.PropGet(0x0000056b, nil)
	return NewValidation(retVal.PdispValVal(), false, true)
}

var Range_Value_OptArgs= []string{
	"RangeValueDataType", 
}

func (this *Range) Value(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_Value_OptArgs, optArgs)
	retVal := this.PropGet(0x00000006, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Range_SetValue_OptArgs= []string{
	"rhs", 
}

func (this *Range) SetValue(rangeValueDataType interface{}, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_SetValue_OptArgs, optArgs)
	retVal := this.PropPut(0x00000006, []interface{}{rangeValueDataType}, optArgs...)
	_= retVal
}

func (this *Range) Value2() ole.Variant {
	retVal := this.PropGet(0x0000056c, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetValue2(rhs interface{})  {
	retVal := this.PropPut(0x0000056c, []interface{}{rhs})
	_= retVal
}

func (this *Range) VerticalAlignment() ole.Variant {
	retVal := this.PropGet(0x00000089, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetVerticalAlignment(rhs interface{})  {
	retVal := this.PropPut(0x00000089, []interface{}{rhs})
	_= retVal
}

func (this *Range) Width() ole.Variant {
	retVal := this.PropGet(0x0000007a, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) Worksheet() *Worksheet {
	retVal := this.PropGet(0x0000015c, nil)
	return NewWorksheet(retVal.PdispValVal(), false, true)
}

func (this *Range) WrapText() ole.Variant {
	retVal := this.PropGet(0x00000114, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetWrapText(rhs interface{})  {
	retVal := this.PropPut(0x00000114, []interface{}{rhs})
	_= retVal
}

var Range_AddComment_OptArgs= []string{
	"Text", 
}

func (this *Range) AddComment(optArgs ...interface{}) *Comment {
	optArgs = ole.ProcessOptArgs(Range_AddComment_OptArgs, optArgs)
	retVal := this.Call(0x0000056d, nil, optArgs...)
	return NewComment(retVal.PdispValVal(), false, true)
}

func (this *Range) Comment() *Comment {
	retVal := this.PropGet(0x0000038e, nil)
	return NewComment(retVal.PdispValVal(), false, true)
}

func (this *Range) ClearComments()  {
	retVal := this.Call(0x0000056e, nil)
	_= retVal
}

func (this *Range) Phonetic() *Phonetic {
	retVal := this.PropGet(0x0000056f, nil)
	return NewPhonetic(retVal.PdispValVal(), false, true)
}

func (this *Range) FormatConditions() *FormatConditions {
	retVal := this.PropGet(0x00000570, nil)
	return NewFormatConditions(retVal.PdispValVal(), false, true)
}

func (this *Range) ReadingOrder() int32 {
	retVal := this.PropGet(0x000003cf, nil)
	return retVal.LValVal()
}

func (this *Range) SetReadingOrder(rhs int32)  {
	retVal := this.PropPut(0x000003cf, []interface{}{rhs})
	_= retVal
}

func (this *Range) Hyperlinks() *Hyperlinks {
	retVal := this.PropGet(0x00000571, nil)
	return NewHyperlinks(retVal.PdispValVal(), false, true)
}

func (this *Range) Phonetics() *Phonetics {
	retVal := this.PropGet(0x00000713, nil)
	return NewPhonetics(retVal.PdispValVal(), false, true)
}

func (this *Range) SetPhonetic()  {
	retVal := this.Call(0x00000714, nil)
	_= retVal
}

func (this *Range) ID() string {
	retVal := this.PropGet(0x00000715, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Range) SetID(rhs string)  {
	retVal := this.PropPut(0x00000715, []interface{}{rhs})
	_= retVal
}

var Range_PrintOut__OptArgs= []string{
	"From", "To", "Copies", "Preview", 
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName", 
}

func (this *Range) PrintOut_(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_PrintOut__OptArgs, optArgs)
	retVal := this.Call(0x000006ec, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) PivotCell() *PivotCell {
	retVal := this.PropGet(0x000007dd, nil)
	return NewPivotCell(retVal.PdispValVal(), false, true)
}

func (this *Range) Dirty()  {
	retVal := this.Call(0x000007de, nil)
	_= retVal
}

func (this *Range) Errors() *Errors {
	retVal := this.PropGet(0x000007df, nil)
	return NewErrors(retVal.PdispValVal(), false, true)
}

func (this *Range) SmartTags() *SmartTags {
	retVal := this.PropGet(0x000007e0, nil)
	return NewSmartTags(retVal.PdispValVal(), false, true)
}

var Range_Speak_OptArgs= []string{
	"SpeakDirection", "SpeakFormulas", 
}

func (this *Range) Speak(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_Speak_OptArgs, optArgs)
	retVal := this.Call(0x000007e1, nil, optArgs...)
	_= retVal
}

var Range_PasteSpecial_OptArgs= []string{
	"SkipBlanks", "Transpose", 
}

func (this *Range) PasteSpecial(paste int32, operation int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_PasteSpecial_OptArgs, optArgs)
	retVal := this.Call(0x00000788, []interface{}{paste, operation}, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) AllowEdit() bool {
	retVal := this.PropGet(0x000007e4, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Range) ListObject() *ListObject {
	retVal := this.PropGet(0x000008d1, nil)
	return NewListObject(retVal.PdispValVal(), false, true)
}

func (this *Range) XPath() *XPath {
	retVal := this.PropGet(0x000008d2, nil)
	return NewXPath(retVal.PdispValVal(), false, true)
}

func (this *Range) ServerActions() *Actions {
	retVal := this.PropGet(0x000009bb, nil)
	return NewActions(retVal.PdispValVal(), false, true)
}

func (this *Range) RemoveDuplicates(columns interface{}, header int32)  {
	retVal := this.Call(0x000009bc, []interface{}{columns, header})
	_= retVal
}

var Range_PrintOut_OptArgs= []string{
	"From", "To", "Copies", "Preview", 
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName", 
}

func (this *Range) PrintOut(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_PrintOut_OptArgs, optArgs)
	retVal := this.Call(0x00000939, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) MDX() string {
	retVal := this.PropGet(0x0000084b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var Range_ExportAsFixedFormat_OptArgs= []string{
	"Filename", "Quality", "IncludeDocProperties", "IgnorePrintAreas", 
	"From", "To", "OpenAfterPublish", "FixedFormatExtClassPtr", 
}

func (this *Range) ExportAsFixedFormat(type_ int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_ExportAsFixedFormat_OptArgs, optArgs)
	retVal := this.Call(0x000009bd, []interface{}{type_}, optArgs...)
	_= retVal
}

func (this *Range) CountLarge() ole.Variant {
	retVal := this.PropGet(0x000009c3, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) CalculateRowMajorOrder() ole.Variant {
	retVal := this.Call(0x0000093c, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SparklineGroups() *SparklineGroups {
	retVal := this.PropGet(0x00000b25, nil)
	return NewSparklineGroups(retVal.PdispValVal(), false, true)
}

func (this *Range) ClearHyperlinks()  {
	retVal := this.Call(0x00000b26, nil)
	_= retVal
}

func (this *Range) DisplayFormat() *DisplayFormat {
	retVal := this.PropGet(0x0000029a, nil)
	return NewDisplayFormat(retVal.PdispValVal(), false, true)
}

var Range_BorderAround_OptArgs= []string{
	"Color", "ThemeColor", 
}

func (this *Range) BorderAround(lineStyle interface{}, weight int32, colorIndex int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Range_BorderAround_OptArgs, optArgs)
	retVal := this.Call(0x00000ad3, []interface{}{lineStyle, weight, colorIndex}, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) AllocateChanges()  {
	retVal := this.Call(0x00000b27, nil)
	_= retVal
}

func (this *Range) DiscardChanges()  {
	retVal := this.Call(0x00000b28, nil)
	_= retVal
}

