package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000208D8-0000-0000-C000-000000000046
var IID_Worksheet_ = syscall.GUID{0x000208D8, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Worksheet_ struct {
	ole.OleClient
}

func NewWorksheet_(pDisp *win32.IDispatch, addRef bool, scoped bool) *Worksheet_ {
	p := &Worksheet_{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func Worksheet_FromVar(v ole.Variant) *Worksheet_ {
	return NewWorksheet_(v.PdispValVal(), false, false)
}

func (this *Worksheet_) IID() *syscall.GUID {
	return &IID_Worksheet_
}

func (this *Worksheet_) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Worksheet_) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Worksheet_) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Worksheet_) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Worksheet_) Activate()  {
	retVal := this.Call(0x00000130, nil)
	_= retVal
}

var Worksheet__Copy_OptArgs= []string{
	"Before", "After", 
}

func (this *Worksheet_) Copy(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Worksheet__Copy_OptArgs, optArgs)
	retVal := this.Call(0x00000227, nil, optArgs...)
	_= retVal
}

func (this *Worksheet_) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

func (this *Worksheet_) CodeName() string {
	retVal := this.PropGet(0x0000055d, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Worksheet_) CodeName_() string {
	retVal := this.PropGet(-2147418112, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Worksheet_) SetCodeName_(rhs string)  {
	retVal := this.PropPut(-2147418112, []interface{}{rhs})
	_= retVal
}

func (this *Worksheet_) Index() int32 {
	retVal := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

var Worksheet__Move_OptArgs= []string{
	"Before", "After", 
}

func (this *Worksheet_) Move(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Worksheet__Move_OptArgs, optArgs)
	retVal := this.Call(0x0000027d, nil, optArgs...)
	_= retVal
}

func (this *Worksheet_) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Worksheet_) SetName(rhs string)  {
	retVal := this.PropPut(0x0000006e, []interface{}{rhs})
	_= retVal
}

func (this *Worksheet_) Next() *ole.DispatchClass {
	retVal := this.PropGet(0x000001f6, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Worksheet_) OnDoubleClick() string {
	retVal := this.PropGet(0x00000274, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Worksheet_) SetOnDoubleClick(rhs string)  {
	retVal := this.PropPut(0x00000274, []interface{}{rhs})
	_= retVal
}

func (this *Worksheet_) OnSheetActivate() string {
	retVal := this.PropGet(0x00000407, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Worksheet_) SetOnSheetActivate(rhs string)  {
	retVal := this.PropPut(0x00000407, []interface{}{rhs})
	_= retVal
}

func (this *Worksheet_) OnSheetDeactivate() string {
	retVal := this.PropGet(0x00000439, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Worksheet_) SetOnSheetDeactivate(rhs string)  {
	retVal := this.PropPut(0x00000439, []interface{}{rhs})
	_= retVal
}

func (this *Worksheet_) PageSetup() *PageSetup {
	retVal := this.PropGet(0x000003e6, nil)
	return NewPageSetup(retVal.PdispValVal(), false, true)
}

func (this *Worksheet_) Previous() *ole.DispatchClass {
	retVal := this.PropGet(0x000001f7, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Worksheet__PrintOut___OptArgs= []string{
	"From", "To", "Copies", "Preview", 
	"ActivePrinter", "PrintToFile", "Collate", 
}

func (this *Worksheet_) PrintOut__(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Worksheet__PrintOut___OptArgs, optArgs)
	retVal := this.Call(0x00000389, nil, optArgs...)
	_= retVal
}

var Worksheet__PrintPreview_OptArgs= []string{
	"EnableChanges", 
}

func (this *Worksheet_) PrintPreview(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Worksheet__PrintPreview_OptArgs, optArgs)
	retVal := this.Call(0x00000119, nil, optArgs...)
	_= retVal
}

var Worksheet__Protect__OptArgs= []string{
	"Password", "DrawingObjects", "Contents", "Scenarios", "UserInterfaceOnly", 
}

func (this *Worksheet_) Protect_(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Worksheet__Protect__OptArgs, optArgs)
	retVal := this.Call(0x0000011a, nil, optArgs...)
	_= retVal
}

func (this *Worksheet_) ProtectContents() bool {
	retVal := this.PropGet(0x00000124, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Worksheet_) ProtectDrawingObjects() bool {
	retVal := this.PropGet(0x00000125, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Worksheet_) ProtectionMode() bool {
	retVal := this.PropGet(0x00000487, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Worksheet_) ProtectScenarios() bool {
	retVal := this.PropGet(0x00000126, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var Worksheet__SaveAs__OptArgs= []string{
	"FileFormat", "Password", "WriteResPassword", "ReadOnlyRecommended", 
	"CreateBackup", "AddToMru", "TextCodepage", "TextVisualLayout", 
}

func (this *Worksheet_) SaveAs_(filename string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Worksheet__SaveAs__OptArgs, optArgs)
	retVal := this.Call(0x0000011c, []interface{}{filename}, optArgs...)
	_= retVal
}

var Worksheet__Select_OptArgs= []string{
	"Replace", 
}

func (this *Worksheet_) Select(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Worksheet__Select_OptArgs, optArgs)
	retVal := this.Call(0x000000eb, nil, optArgs...)
	_= retVal
}

var Worksheet__Unprotect_OptArgs= []string{
	"Password", 
}

func (this *Worksheet_) Unprotect(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Worksheet__Unprotect_OptArgs, optArgs)
	retVal := this.Call(0x0000011d, nil, optArgs...)
	_= retVal
}

func (this *Worksheet_) Visible() int32 {
	retVal := this.PropGet(0x0000022e, nil)
	return retVal.LValVal()
}

func (this *Worksheet_) SetVisible(rhs int32)  {
	retVal := this.PropPut(0x0000022e, []interface{}{rhs})
	_= retVal
}

func (this *Worksheet_) Shapes() *Shapes {
	retVal := this.PropGet(0x00000561, nil)
	return NewShapes(retVal.PdispValVal(), false, true)
}

func (this *Worksheet_) TransitionExpEval() bool {
	retVal := this.PropGet(0x00000191, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Worksheet_) SetTransitionExpEval(rhs bool)  {
	retVal := this.PropPut(0x00000191, []interface{}{rhs})
	_= retVal
}

var Worksheet__Arcs_OptArgs= []string{
	"Index", 
}

func (this *Worksheet_) Arcs(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Worksheet__Arcs_OptArgs, optArgs)
	retVal := this.Call(0x000002f8, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Worksheet_) AutoFilterMode() bool {
	retVal := this.PropGet(0x00000318, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Worksheet_) SetAutoFilterMode(rhs bool)  {
	retVal := this.PropPut(0x00000318, []interface{}{rhs})
	_= retVal
}

func (this *Worksheet_) SetBackgroundPicture(filename string)  {
	retVal := this.Call(0x000004a4, []interface{}{filename})
	_= retVal
}

var Worksheet__Buttons_OptArgs= []string{
	"Index", 
}

func (this *Worksheet_) Buttons(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Worksheet__Buttons_OptArgs, optArgs)
	retVal := this.Call(0x0000022d, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Worksheet_) Calculate()  {
	retVal := this.Call(0x00000117, nil)
	_= retVal
}

func (this *Worksheet_) EnableCalculation() bool {
	retVal := this.PropGet(0x00000590, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Worksheet_) SetEnableCalculation(rhs bool)  {
	retVal := this.PropPut(0x00000590, []interface{}{rhs})
	_= retVal
}

func (this *Worksheet_) Cells() *Range {
	retVal := this.PropGet(0x000000ee, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

var Worksheet__ChartObjects_OptArgs= []string{
	"Index", 
}

func (this *Worksheet_) ChartObjects(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Worksheet__ChartObjects_OptArgs, optArgs)
	retVal := this.Call(0x00000424, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Worksheet__CheckBoxes_OptArgs= []string{
	"Index", 
}

func (this *Worksheet_) CheckBoxes(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Worksheet__CheckBoxes_OptArgs, optArgs)
	retVal := this.Call(0x00000338, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Worksheet__CheckSpelling_OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang", 
}

func (this *Worksheet_) CheckSpelling(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Worksheet__CheckSpelling_OptArgs, optArgs)
	retVal := this.Call(0x000001f9, nil, optArgs...)
	_= retVal
}

func (this *Worksheet_) CircularReference() *Range {
	retVal := this.PropGet(0x0000042d, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Worksheet_) ClearArrows()  {
	retVal := this.Call(0x000003ca, nil)
	_= retVal
}

func (this *Worksheet_) Columns() *Range {
	retVal := this.PropGet(0x000000f1, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Worksheet_) ConsolidationFunction() int32 {
	retVal := this.PropGet(0x00000315, nil)
	return retVal.LValVal()
}

func (this *Worksheet_) ConsolidationOptions() ole.Variant {
	retVal := this.PropGet(0x00000316, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Worksheet_) ConsolidationSources() ole.Variant {
	retVal := this.PropGet(0x00000317, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Worksheet_) DisplayAutomaticPageBreaks() bool {
	retVal := this.PropGet(0x00000283, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Worksheet_) SetDisplayAutomaticPageBreaks(rhs bool)  {
	retVal := this.PropPut(0x00000283, []interface{}{rhs})
	_= retVal
}

var Worksheet__Drawings_OptArgs= []string{
	"Index", 
}

func (this *Worksheet_) Drawings(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Worksheet__Drawings_OptArgs, optArgs)
	retVal := this.Call(0x00000304, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Worksheet__DrawingObjects_OptArgs= []string{
	"Index", 
}

func (this *Worksheet_) DrawingObjects(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Worksheet__DrawingObjects_OptArgs, optArgs)
	retVal := this.Call(0x00000058, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Worksheet__DropDowns_OptArgs= []string{
	"Index", 
}

func (this *Worksheet_) DropDowns(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Worksheet__DropDowns_OptArgs, optArgs)
	retVal := this.Call(0x00000344, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Worksheet_) EnableAutoFilter() bool {
	retVal := this.PropGet(0x00000484, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Worksheet_) SetEnableAutoFilter(rhs bool)  {
	retVal := this.PropPut(0x00000484, []interface{}{rhs})
	_= retVal
}

func (this *Worksheet_) EnableSelection() int32 {
	retVal := this.PropGet(0x00000591, nil)
	return retVal.LValVal()
}

func (this *Worksheet_) SetEnableSelection(rhs int32)  {
	retVal := this.PropPut(0x00000591, []interface{}{rhs})
	_= retVal
}

func (this *Worksheet_) EnableOutlining() bool {
	retVal := this.PropGet(0x00000485, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Worksheet_) SetEnableOutlining(rhs bool)  {
	retVal := this.PropPut(0x00000485, []interface{}{rhs})
	_= retVal
}

func (this *Worksheet_) EnablePivotTable() bool {
	retVal := this.PropGet(0x00000486, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Worksheet_) SetEnablePivotTable(rhs bool)  {
	retVal := this.PropPut(0x00000486, []interface{}{rhs})
	_= retVal
}

func (this *Worksheet_) Evaluate(name interface{}) ole.Variant {
	retVal := this.Call(0x00000001, []interface{}{name})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Worksheet_) Evaluate_(name interface{}) ole.Variant {
	retVal := this.Call(-5, []interface{}{name})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Worksheet_) FilterMode() bool {
	retVal := this.PropGet(0x00000320, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Worksheet_) ResetAllPageBreaks()  {
	retVal := this.Call(0x00000592, nil)
	_= retVal
}

var Worksheet__GroupBoxes_OptArgs= []string{
	"Index", 
}

func (this *Worksheet_) GroupBoxes(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Worksheet__GroupBoxes_OptArgs, optArgs)
	retVal := this.Call(0x00000342, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Worksheet__GroupObjects_OptArgs= []string{
	"Index", 
}

func (this *Worksheet_) GroupObjects(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Worksheet__GroupObjects_OptArgs, optArgs)
	retVal := this.Call(0x00000459, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Worksheet__Labels_OptArgs= []string{
	"Index", 
}

func (this *Worksheet_) Labels(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Worksheet__Labels_OptArgs, optArgs)
	retVal := this.Call(0x00000349, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Worksheet__Lines_OptArgs= []string{
	"Index", 
}

func (this *Worksheet_) Lines(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Worksheet__Lines_OptArgs, optArgs)
	retVal := this.Call(0x000002ff, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Worksheet__ListBoxes_OptArgs= []string{
	"Index", 
}

func (this *Worksheet_) ListBoxes(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Worksheet__ListBoxes_OptArgs, optArgs)
	retVal := this.Call(0x00000340, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Worksheet_) Names() *Names {
	retVal := this.PropGet(0x000001ba, nil)
	return NewNames(retVal.PdispValVal(), false, true)
}

var Worksheet__OLEObjects_OptArgs= []string{
	"Index", 
}

func (this *Worksheet_) OLEObjects(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Worksheet__OLEObjects_OptArgs, optArgs)
	retVal := this.Call(0x0000031f, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Worksheet_) OnCalculate() string {
	retVal := this.PropGet(0x00000271, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Worksheet_) SetOnCalculate(rhs string)  {
	retVal := this.PropPut(0x00000271, []interface{}{rhs})
	_= retVal
}

func (this *Worksheet_) OnData() string {
	retVal := this.PropGet(0x00000275, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Worksheet_) SetOnData(rhs string)  {
	retVal := this.PropPut(0x00000275, []interface{}{rhs})
	_= retVal
}

func (this *Worksheet_) OnEntry() string {
	retVal := this.PropGet(0x00000273, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Worksheet_) SetOnEntry(rhs string)  {
	retVal := this.PropPut(0x00000273, []interface{}{rhs})
	_= retVal
}

var Worksheet__OptionButtons_OptArgs= []string{
	"Index", 
}

func (this *Worksheet_) OptionButtons(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Worksheet__OptionButtons_OptArgs, optArgs)
	retVal := this.Call(0x0000033a, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Worksheet_) Outline() *Outline {
	retVal := this.PropGet(0x00000066, nil)
	return NewOutline(retVal.PdispValVal(), false, true)
}

var Worksheet__Ovals_OptArgs= []string{
	"Index", 
}

func (this *Worksheet_) Ovals(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Worksheet__Ovals_OptArgs, optArgs)
	retVal := this.Call(0x00000321, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Worksheet__Paste_OptArgs= []string{
	"Destination", "Link", 
}

func (this *Worksheet_) Paste(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Worksheet__Paste_OptArgs, optArgs)
	retVal := this.Call(0x000000d3, nil, optArgs...)
	_= retVal
}

var Worksheet__PasteSpecial__OptArgs= []string{
	"Format", "Link", "DisplayAsIcon", "IconFileName", 
	"IconIndex", "IconLabel", 
}

func (this *Worksheet_) PasteSpecial_(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Worksheet__PasteSpecial__OptArgs, optArgs)
	retVal := this.Call(0x00000403, nil, optArgs...)
	_= retVal
}

var Worksheet__Pictures_OptArgs= []string{
	"Index", 
}

func (this *Worksheet_) Pictures(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Worksheet__Pictures_OptArgs, optArgs)
	retVal := this.Call(0x00000303, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Worksheet__PivotTables_OptArgs= []string{
	"Index", 
}

func (this *Worksheet_) PivotTables(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Worksheet__PivotTables_OptArgs, optArgs)
	retVal := this.Call(0x000002b2, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Worksheet__PivotTableWizard_OptArgs= []string{
	"SourceType", "SourceData", "TableDestination", "TableName", 
	"RowGrand", "ColumnGrand", "SaveData", "HasAutoFormat", 
	"AutoPage", "Reserved", "BackgroundQuery", "OptimizeCache", 
	"PageFieldOrder", "PageFieldWrapCount", "ReadData", "Connection", 
}

func (this *Worksheet_) PivotTableWizard(optArgs ...interface{}) *PivotTable {
	optArgs = ole.ProcessOptArgs(Worksheet__PivotTableWizard_OptArgs, optArgs)
	retVal := this.Call(0x000002ac, nil, optArgs...)
	return NewPivotTable(retVal.PdispValVal(), false, true)
}

var Worksheet__Range_OptArgs= []string{
	"Cell2", 
}

func (this *Worksheet_) Range(cell1 interface{}, optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Worksheet__Range_OptArgs, optArgs)
	retVal := this.PropGet(0x000000c5, []interface{}{cell1}, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

var Worksheet__Rectangles_OptArgs= []string{
	"Index", 
}

func (this *Worksheet_) Rectangles(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Worksheet__Rectangles_OptArgs, optArgs)
	retVal := this.Call(0x00000306, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Worksheet_) Rows() *Range {
	retVal := this.PropGet(0x00000102, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

var Worksheet__Scenarios_OptArgs= []string{
	"Index", 
}

func (this *Worksheet_) Scenarios(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Worksheet__Scenarios_OptArgs, optArgs)
	retVal := this.Call(0x0000038c, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Worksheet_) ScrollArea() string {
	retVal := this.PropGet(0x00000599, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Worksheet_) SetScrollArea(rhs string)  {
	retVal := this.PropPut(0x00000599, []interface{}{rhs})
	_= retVal
}

var Worksheet__ScrollBars_OptArgs= []string{
	"Index", 
}

func (this *Worksheet_) ScrollBars(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Worksheet__ScrollBars_OptArgs, optArgs)
	retVal := this.Call(0x0000033e, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Worksheet_) ShowAllData()  {
	retVal := this.Call(0x0000031a, nil)
	_= retVal
}

func (this *Worksheet_) ShowDataForm()  {
	retVal := this.Call(0x00000199, nil)
	_= retVal
}

var Worksheet__Spinners_OptArgs= []string{
	"Index", 
}

func (this *Worksheet_) Spinners(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Worksheet__Spinners_OptArgs, optArgs)
	retVal := this.Call(0x00000346, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Worksheet_) StandardHeight() float64 {
	retVal := this.PropGet(0x00000197, nil)
	return retVal.DblValVal()
}

func (this *Worksheet_) StandardWidth() float64 {
	retVal := this.PropGet(0x00000198, nil)
	return retVal.DblValVal()
}

func (this *Worksheet_) SetStandardWidth(rhs float64)  {
	retVal := this.PropPut(0x00000198, []interface{}{rhs})
	_= retVal
}

var Worksheet__TextBoxes_OptArgs= []string{
	"Index", 
}

func (this *Worksheet_) TextBoxes(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Worksheet__TextBoxes_OptArgs, optArgs)
	retVal := this.Call(0x00000309, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Worksheet_) TransitionFormEntry() bool {
	retVal := this.PropGet(0x00000192, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Worksheet_) SetTransitionFormEntry(rhs bool)  {
	retVal := this.PropPut(0x00000192, []interface{}{rhs})
	_= retVal
}

func (this *Worksheet_) Type() int32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *Worksheet_) UsedRange() *Range {
	retVal := this.PropGet(0x0000019c, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Worksheet_) HPageBreaks() *HPageBreaks {
	retVal := this.PropGet(0x0000058a, nil)
	return NewHPageBreaks(retVal.PdispValVal(), false, true)
}

func (this *Worksheet_) VPageBreaks() *VPageBreaks {
	retVal := this.PropGet(0x0000058b, nil)
	return NewVPageBreaks(retVal.PdispValVal(), false, true)
}

func (this *Worksheet_) QueryTables() *QueryTables {
	retVal := this.PropGet(0x0000059a, nil)
	return NewQueryTables(retVal.PdispValVal(), false, true)
}

func (this *Worksheet_) DisplayPageBreaks() bool {
	retVal := this.PropGet(0x0000059b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Worksheet_) SetDisplayPageBreaks(rhs bool)  {
	retVal := this.PropPut(0x0000059b, []interface{}{rhs})
	_= retVal
}

func (this *Worksheet_) Comments() *Comments {
	retVal := this.PropGet(0x0000023f, nil)
	return NewComments(retVal.PdispValVal(), false, true)
}

func (this *Worksheet_) Hyperlinks() *Hyperlinks {
	retVal := this.PropGet(0x00000571, nil)
	return NewHyperlinks(retVal.PdispValVal(), false, true)
}

func (this *Worksheet_) ClearCircles()  {
	retVal := this.Call(0x0000059c, nil)
	_= retVal
}

func (this *Worksheet_) CircleInvalid()  {
	retVal := this.Call(0x0000059d, nil)
	_= retVal
}

func (this *Worksheet_) DisplayRightToLeft_() int32 {
	retVal := this.PropGet(0x00000288, nil)
	return retVal.LValVal()
}

func (this *Worksheet_) SetDisplayRightToLeft_(rhs int32)  {
	retVal := this.PropPut(0x00000288, []interface{}{rhs})
	_= retVal
}

func (this *Worksheet_) AutoFilter() *AutoFilter {
	retVal := this.PropGet(0x00000319, nil)
	return NewAutoFilter(retVal.PdispValVal(), false, true)
}

func (this *Worksheet_) DisplayRightToLeft() bool {
	retVal := this.PropGet(0x000006ee, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Worksheet_) SetDisplayRightToLeft(rhs bool)  {
	retVal := this.PropPut(0x000006ee, []interface{}{rhs})
	_= retVal
}

func (this *Worksheet_) Scripts() *ole.DispatchClass {
	retVal := this.PropGet(0x00000718, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Worksheet__PrintOut__OptArgs= []string{
	"From", "To", "Copies", "Preview", 
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName", 
}

func (this *Worksheet_) PrintOut_(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Worksheet__PrintOut__OptArgs, optArgs)
	retVal := this.Call(0x000006ec, nil, optArgs...)
	_= retVal
}

var Worksheet__CheckSpelling__OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang", 
	"IgnoreFinalYaa", "SpellScript", 
}

func (this *Worksheet_) CheckSpelling_(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Worksheet__CheckSpelling__OptArgs, optArgs)
	retVal := this.Call(0x00000719, nil, optArgs...)
	_= retVal
}

func (this *Worksheet_) Tab() *Tab {
	retVal := this.PropGet(0x00000411, nil)
	return NewTab(retVal.PdispValVal(), false, true)
}

func (this *Worksheet_) MailEnvelope() *ole.DispatchClass {
	retVal := this.PropGet(0x000007e5, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Worksheet__SaveAs_OptArgs= []string{
	"FileFormat", "Password", "WriteResPassword", "ReadOnlyRecommended", 
	"CreateBackup", "AddToMru", "TextCodepage", "TextVisualLayout", "Local", 
}

func (this *Worksheet_) SaveAs(filename string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Worksheet__SaveAs_OptArgs, optArgs)
	retVal := this.Call(0x00000785, []interface{}{filename}, optArgs...)
	_= retVal
}

func (this *Worksheet_) CustomProperties() *CustomProperties {
	retVal := this.PropGet(0x000007ee, nil)
	return NewCustomProperties(retVal.PdispValVal(), false, true)
}

func (this *Worksheet_) SmartTags() *SmartTags {
	retVal := this.PropGet(0x000007e0, nil)
	return NewSmartTags(retVal.PdispValVal(), false, true)
}

func (this *Worksheet_) Protection() *Protection {
	retVal := this.PropGet(0x000000b0, nil)
	return NewProtection(retVal.PdispValVal(), false, true)
}

var Worksheet__PasteSpecial_OptArgs= []string{
	"Format", "Link", "DisplayAsIcon", "IconFileName", 
	"IconIndex", "IconLabel", "NoHTMLFormatting", 
}

func (this *Worksheet_) PasteSpecial(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Worksheet__PasteSpecial_OptArgs, optArgs)
	retVal := this.Call(0x00000788, nil, optArgs...)
	_= retVal
}

var Worksheet__Protect_OptArgs= []string{
	"Password", "DrawingObjects", "Contents", "Scenarios", 
	"UserInterfaceOnly", "AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows", 
	"AllowInsertingColumns", "AllowInsertingRows", "AllowInsertingHyperlinks", "AllowDeletingColumns", 
	"AllowDeletingRows", "AllowSorting", "AllowFiltering", "AllowUsingPivotTables", 
}

func (this *Worksheet_) Protect(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Worksheet__Protect_OptArgs, optArgs)
	retVal := this.Call(0x000007ed, nil, optArgs...)
	_= retVal
}

func (this *Worksheet_) ListObjects() *ListObjects {
	retVal := this.PropGet(0x000008d3, nil)
	return NewListObjects(retVal.PdispValVal(), false, true)
}

var Worksheet__XmlDataQuery_OptArgs= []string{
	"SelectionNamespaces", "Map", 
}

func (this *Worksheet_) XmlDataQuery(xpath string, optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Worksheet__XmlDataQuery_OptArgs, optArgs)
	retVal := this.Call(0x000008d4, []interface{}{xpath}, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

var Worksheet__XmlMapQuery_OptArgs= []string{
	"SelectionNamespaces", "Map", 
}

func (this *Worksheet_) XmlMapQuery(xpath string, optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Worksheet__XmlMapQuery_OptArgs, optArgs)
	retVal := this.Call(0x000008d7, []interface{}{xpath}, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

var Worksheet__PrintOut_OptArgs= []string{
	"From", "To", "Copies", "Preview", 
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName", "IgnorePrintAreas", 
}

func (this *Worksheet_) PrintOut(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Worksheet__PrintOut_OptArgs, optArgs)
	retVal := this.Call(0x00000939, nil, optArgs...)
	_= retVal
}

func (this *Worksheet_) EnableFormatConditionsCalculation() bool {
	retVal := this.PropGet(0x000009cf, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Worksheet_) SetEnableFormatConditionsCalculation(rhs bool)  {
	retVal := this.PropPut(0x000009cf, []interface{}{rhs})
	_= retVal
}

func (this *Worksheet_) Sort() *Sort {
	retVal := this.PropGet(0x00000370, nil)
	return NewSort(retVal.PdispValVal(), false, true)
}

var Worksheet__ExportAsFixedFormat_OptArgs= []string{
	"Filename", "Quality", "IncludeDocProperties", "IgnorePrintAreas", 
	"From", "To", "OpenAfterPublish", "FixedFormatExtClassPtr", 
}

func (this *Worksheet_) ExportAsFixedFormat(type_ int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Worksheet__ExportAsFixedFormat_OptArgs, optArgs)
	retVal := this.Call(0x000009bd, []interface{}{type_}, optArgs...)
	_= retVal
}

func (this *Worksheet_) PrintedCommentPages() int32 {
	retVal := this.PropGet(0x00000b29, nil)
	return retVal.LValVal()
}

