package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000208AF-0000-0000-C000-000000000046
var IID_DialogSheet = syscall.GUID{0x000208AF, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type DialogSheet struct {
	ole.OleClient
}

func NewDialogSheet(pDisp *win32.IDispatch, addRef bool, scoped bool) *DialogSheet {
	 if pDisp == nil {
		return nil;
	}
	p := &DialogSheet{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DialogSheetFromVar(v ole.Variant) *DialogSheet {
	return NewDialogSheet(v.IDispatch(), false, false)
}

func (this *DialogSheet) IID() *syscall.GUID {
	return &IID_DialogSheet
}

func (this *DialogSheet) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DialogSheet) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *DialogSheet) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *DialogSheet) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *DialogSheet) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *DialogSheet) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *DialogSheet) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *DialogSheet) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *DialogSheet) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *DialogSheet) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *DialogSheet) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DialogSheet) Activate()  {
	retVal, _ := this.Call(0x00000130, nil)
	_= retVal
}

var DialogSheet_Copy_OptArgs= []string{
	"Before", "After", 
}

func (this *DialogSheet) Copy(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(DialogSheet_Copy_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000227, nil, optArgs...)
	_= retVal
}

func (this *DialogSheet) Delete()  {
	retVal, _ := this.Call(0x00000075, nil)
	_= retVal
}

func (this *DialogSheet) CodeName() string {
	retVal, _ := this.PropGet(0x0000055d, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DialogSheet) CodeName_() string {
	retVal, _ := this.PropGet(-2147418112, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DialogSheet) SetCodeName_(rhs string)  {
	_ = this.PropPut(-2147418112, []interface{}{rhs})
}

func (this *DialogSheet) Index() int32 {
	retVal, _ := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

var DialogSheet_Move_OptArgs= []string{
	"Before", "After", 
}

func (this *DialogSheet) Move(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(DialogSheet_Move_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000027d, nil, optArgs...)
	_= retVal
}

func (this *DialogSheet) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DialogSheet) SetName(rhs string)  {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *DialogSheet) Next() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000001f6, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DialogSheet) OnDoubleClick() string {
	retVal, _ := this.PropGet(0x00000274, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DialogSheet) SetOnDoubleClick(rhs string)  {
	_ = this.PropPut(0x00000274, []interface{}{rhs})
}

func (this *DialogSheet) OnSheetActivate() string {
	retVal, _ := this.PropGet(0x00000407, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DialogSheet) SetOnSheetActivate(rhs string)  {
	_ = this.PropPut(0x00000407, []interface{}{rhs})
}

func (this *DialogSheet) OnSheetDeactivate() string {
	retVal, _ := this.PropGet(0x00000439, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DialogSheet) SetOnSheetDeactivate(rhs string)  {
	_ = this.PropPut(0x00000439, []interface{}{rhs})
}

func (this *DialogSheet) PageSetup() *PageSetup {
	retVal, _ := this.PropGet(0x000003e6, nil)
	return NewPageSetup(retVal.IDispatch(), false, true)
}

func (this *DialogSheet) Previous() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000001f7, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var DialogSheet_PrintOut___OptArgs= []string{
	"From", "To", "Copies", "Preview", 
	"ActivePrinter", "PrintToFile", "Collate", 
}

func (this *DialogSheet) PrintOut__(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(DialogSheet_PrintOut___OptArgs, optArgs)
	retVal, _ := this.Call(0x00000389, nil, optArgs...)
	_= retVal
}

var DialogSheet_PrintPreview_OptArgs= []string{
	"EnableChanges", 
}

func (this *DialogSheet) PrintPreview(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(DialogSheet_PrintPreview_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000119, nil, optArgs...)
	_= retVal
}

var DialogSheet_Protect__OptArgs= []string{
	"Password", "DrawingObjects", "Contents", "Scenarios", "UserInterfaceOnly", 
}

func (this *DialogSheet) Protect_(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(DialogSheet_Protect__OptArgs, optArgs)
	retVal, _ := this.Call(0x0000011a, nil, optArgs...)
	_= retVal
}

func (this *DialogSheet) ProtectContents() bool {
	retVal, _ := this.PropGet(0x00000124, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DialogSheet) ProtectDrawingObjects() bool {
	retVal, _ := this.PropGet(0x00000125, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DialogSheet) ProtectionMode() bool {
	retVal, _ := this.PropGet(0x00000487, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DialogSheet) ProtectScenarios() bool {
	retVal, _ := this.PropGet(0x00000126, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var DialogSheet_SaveAs__OptArgs= []string{
	"FileFormat", "Password", "WriteResPassword", "ReadOnlyRecommended", 
	"CreateBackup", "AddToMru", "TextCodepage", "TextVisualLayout", 
}

func (this *DialogSheet) SaveAs_(filename string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(DialogSheet_SaveAs__OptArgs, optArgs)
	retVal, _ := this.Call(0x0000011c, []interface{}{filename}, optArgs...)
	_= retVal
}

var DialogSheet_Select_OptArgs= []string{
	"Replace", 
}

func (this *DialogSheet) Select(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(DialogSheet_Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000eb, nil, optArgs...)
	_= retVal
}

var DialogSheet_Unprotect_OptArgs= []string{
	"Password", 
}

func (this *DialogSheet) Unprotect(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(DialogSheet_Unprotect_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000011d, nil, optArgs...)
	_= retVal
}

func (this *DialogSheet) Visible() int32 {
	retVal, _ := this.PropGet(0x0000022e, nil)
	return retVal.LValVal()
}

func (this *DialogSheet) SetVisible(rhs int32)  {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

func (this *DialogSheet) Shapes() *Shapes {
	retVal, _ := this.PropGet(0x00000561, nil)
	return NewShapes(retVal.IDispatch(), false, true)
}

func (this *DialogSheet) Dummy29_()  {
	retVal, _ := this.Call(0x0001001d, nil)
	_= retVal
}

var DialogSheet_Arcs_OptArgs= []string{
	"Index", 
}

func (this *DialogSheet) Arcs(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(DialogSheet_Arcs_OptArgs, optArgs)
	retVal, _ := this.Call(0x000002f8, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DialogSheet) Dummy31_()  {
	retVal, _ := this.Call(0x0001001f, nil)
	_= retVal
}

func (this *DialogSheet) Dummy32_()  {
	retVal, _ := this.Call(0x00010020, nil)
	_= retVal
}

var DialogSheet_Buttons_OptArgs= []string{
	"Index", 
}

func (this *DialogSheet) Buttons(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(DialogSheet_Buttons_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000022d, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DialogSheet) Dummy34_()  {
	retVal, _ := this.Call(0x00010022, nil)
	_= retVal
}

func (this *DialogSheet) EnableCalculation() bool {
	retVal, _ := this.PropGet(0x00000590, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DialogSheet) SetEnableCalculation(rhs bool)  {
	_ = this.PropPut(0x00000590, []interface{}{rhs})
}

func (this *DialogSheet) Dummy36_()  {
	retVal, _ := this.Call(0x00010024, nil)
	_= retVal
}

var DialogSheet_ChartObjects_OptArgs= []string{
	"Index", 
}

func (this *DialogSheet) ChartObjects(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(DialogSheet_ChartObjects_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000424, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var DialogSheet_CheckBoxes_OptArgs= []string{
	"Index", 
}

func (this *DialogSheet) CheckBoxes(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(DialogSheet_CheckBoxes_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000338, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var DialogSheet_CheckSpelling_OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang", 
}

func (this *DialogSheet) CheckSpelling(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(DialogSheet_CheckSpelling_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001f9, nil, optArgs...)
	_= retVal
}

func (this *DialogSheet) Dummy40_()  {
	retVal, _ := this.Call(0x00010028, nil)
	_= retVal
}

func (this *DialogSheet) Dummy41_()  {
	retVal, _ := this.Call(0x00010029, nil)
	_= retVal
}

func (this *DialogSheet) Dummy42_()  {
	retVal, _ := this.Call(0x0001002a, nil)
	_= retVal
}

func (this *DialogSheet) Dummy43_()  {
	retVal, _ := this.Call(0x0001002b, nil)
	_= retVal
}

func (this *DialogSheet) Dummy44_()  {
	retVal, _ := this.Call(0x0001002c, nil)
	_= retVal
}

func (this *DialogSheet) Dummy45_()  {
	retVal, _ := this.Call(0x0001002d, nil)
	_= retVal
}

func (this *DialogSheet) DisplayAutomaticPageBreaks() bool {
	retVal, _ := this.PropGet(0x00000283, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DialogSheet) SetDisplayAutomaticPageBreaks(rhs bool)  {
	_ = this.PropPut(0x00000283, []interface{}{rhs})
}

var DialogSheet_Drawings_OptArgs= []string{
	"Index", 
}

func (this *DialogSheet) Drawings(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(DialogSheet_Drawings_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000304, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var DialogSheet_DrawingObjects_OptArgs= []string{
	"Index", 
}

func (this *DialogSheet) DrawingObjects(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(DialogSheet_DrawingObjects_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000058, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var DialogSheet_DropDowns_OptArgs= []string{
	"Index", 
}

func (this *DialogSheet) DropDowns(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(DialogSheet_DropDowns_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000344, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DialogSheet) EnableAutoFilter() bool {
	retVal, _ := this.PropGet(0x00000484, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DialogSheet) SetEnableAutoFilter(rhs bool)  {
	_ = this.PropPut(0x00000484, []interface{}{rhs})
}

func (this *DialogSheet) EnableSelection() int32 {
	retVal, _ := this.PropGet(0x00000591, nil)
	return retVal.LValVal()
}

func (this *DialogSheet) SetEnableSelection(rhs int32)  {
	_ = this.PropPut(0x00000591, []interface{}{rhs})
}

func (this *DialogSheet) EnableOutlining() bool {
	retVal, _ := this.PropGet(0x00000485, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DialogSheet) SetEnableOutlining(rhs bool)  {
	_ = this.PropPut(0x00000485, []interface{}{rhs})
}

func (this *DialogSheet) EnablePivotTable() bool {
	retVal, _ := this.PropGet(0x00000486, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DialogSheet) SetEnablePivotTable(rhs bool)  {
	_ = this.PropPut(0x00000486, []interface{}{rhs})
}

func (this *DialogSheet) Evaluate(name interface{}) ole.Variant {
	retVal, _ := this.Call(0x00000001, []interface{}{name})
	com.AddToScope(retVal)
	return *retVal
}

func (this *DialogSheet) Evaluate_(name interface{}) ole.Variant {
	retVal, _ := this.Call(-5, []interface{}{name})
	com.AddToScope(retVal)
	return *retVal
}

func (this *DialogSheet) Dummy56_()  {
	retVal, _ := this.Call(0x00010038, nil)
	_= retVal
}

func (this *DialogSheet) ResetAllPageBreaks()  {
	retVal, _ := this.Call(0x00000592, nil)
	_= retVal
}

var DialogSheet_GroupBoxes_OptArgs= []string{
	"Index", 
}

func (this *DialogSheet) GroupBoxes(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(DialogSheet_GroupBoxes_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000342, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var DialogSheet_GroupObjects_OptArgs= []string{
	"Index", 
}

func (this *DialogSheet) GroupObjects(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(DialogSheet_GroupObjects_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000459, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var DialogSheet_Labels_OptArgs= []string{
	"Index", 
}

func (this *DialogSheet) Labels(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(DialogSheet_Labels_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000349, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var DialogSheet_Lines_OptArgs= []string{
	"Index", 
}

func (this *DialogSheet) Lines(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(DialogSheet_Lines_OptArgs, optArgs)
	retVal, _ := this.Call(0x000002ff, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var DialogSheet_ListBoxes_OptArgs= []string{
	"Index", 
}

func (this *DialogSheet) ListBoxes(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(DialogSheet_ListBoxes_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000340, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DialogSheet) Names() *Names {
	retVal, _ := this.PropGet(0x000001ba, nil)
	return NewNames(retVal.IDispatch(), false, true)
}

var DialogSheet_OLEObjects_OptArgs= []string{
	"Index", 
}

func (this *DialogSheet) OLEObjects(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(DialogSheet_OLEObjects_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000031f, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DialogSheet) Dummy65_()  {
	retVal, _ := this.Call(0x00010041, nil)
	_= retVal
}

func (this *DialogSheet) Dummy66_()  {
	retVal, _ := this.Call(0x00010042, nil)
	_= retVal
}

func (this *DialogSheet) Dummy67_()  {
	retVal, _ := this.Call(0x00010043, nil)
	_= retVal
}

var DialogSheet_OptionButtons_OptArgs= []string{
	"Index", 
}

func (this *DialogSheet) OptionButtons(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(DialogSheet_OptionButtons_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000033a, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DialogSheet) Dummy69_()  {
	retVal, _ := this.Call(0x00010045, nil)
	_= retVal
}

var DialogSheet_Ovals_OptArgs= []string{
	"Index", 
}

func (this *DialogSheet) Ovals(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(DialogSheet_Ovals_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000321, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var DialogSheet_Paste_OptArgs= []string{
	"Destination", "Link", 
}

func (this *DialogSheet) Paste(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(DialogSheet_Paste_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d3, nil, optArgs...)
	_= retVal
}

var DialogSheet_PasteSpecial__OptArgs= []string{
	"Format", "Link", "DisplayAsIcon", "IconFileName", 
	"IconIndex", "IconLabel", 
}

func (this *DialogSheet) PasteSpecial_(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(DialogSheet_PasteSpecial__OptArgs, optArgs)
	retVal, _ := this.Call(0x00000403, nil, optArgs...)
	_= retVal
}

var DialogSheet_Pictures_OptArgs= []string{
	"Index", 
}

func (this *DialogSheet) Pictures(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(DialogSheet_Pictures_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000303, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DialogSheet) Dummy74_()  {
	retVal, _ := this.Call(0x0001004a, nil)
	_= retVal
}

func (this *DialogSheet) Dummy75_()  {
	retVal, _ := this.Call(0x0001004b, nil)
	_= retVal
}

func (this *DialogSheet) Dummy76_()  {
	retVal, _ := this.Call(0x0001004c, nil)
	_= retVal
}

var DialogSheet_Rectangles_OptArgs= []string{
	"Index", 
}

func (this *DialogSheet) Rectangles(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(DialogSheet_Rectangles_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000306, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DialogSheet) Dummy78_()  {
	retVal, _ := this.Call(0x0001004e, nil)
	_= retVal
}

func (this *DialogSheet) Dummy79_()  {
	retVal, _ := this.Call(0x0001004f, nil)
	_= retVal
}

func (this *DialogSheet) ScrollArea() string {
	retVal, _ := this.PropGet(0x00000599, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DialogSheet) SetScrollArea(rhs string)  {
	_ = this.PropPut(0x00000599, []interface{}{rhs})
}

var DialogSheet_ScrollBars_OptArgs= []string{
	"Index", 
}

func (this *DialogSheet) ScrollBars(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(DialogSheet_ScrollBars_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000033e, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DialogSheet) Dummy82_()  {
	retVal, _ := this.Call(0x00010052, nil)
	_= retVal
}

func (this *DialogSheet) Dummy83_()  {
	retVal, _ := this.Call(0x00010053, nil)
	_= retVal
}

var DialogSheet_Spinners_OptArgs= []string{
	"Index", 
}

func (this *DialogSheet) Spinners(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(DialogSheet_Spinners_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000346, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DialogSheet) Dummy85_()  {
	retVal, _ := this.Call(0x00010055, nil)
	_= retVal
}

func (this *DialogSheet) Dummy86_()  {
	retVal, _ := this.Call(0x00010056, nil)
	_= retVal
}

var DialogSheet_TextBoxes_OptArgs= []string{
	"Index", 
}

func (this *DialogSheet) TextBoxes(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(DialogSheet_TextBoxes_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000309, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DialogSheet) Dummy88_()  {
	retVal, _ := this.Call(0x00010058, nil)
	_= retVal
}

func (this *DialogSheet) Dummy89_()  {
	retVal, _ := this.Call(0x00010059, nil)
	_= retVal
}

func (this *DialogSheet) Dummy90_()  {
	retVal, _ := this.Call(0x0001005a, nil)
	_= retVal
}

func (this *DialogSheet) HPageBreaks() *HPageBreaks {
	retVal, _ := this.PropGet(0x0000058a, nil)
	return NewHPageBreaks(retVal.IDispatch(), false, true)
}

func (this *DialogSheet) VPageBreaks() *VPageBreaks {
	retVal, _ := this.PropGet(0x0000058b, nil)
	return NewVPageBreaks(retVal.IDispatch(), false, true)
}

func (this *DialogSheet) QueryTables() *QueryTables {
	retVal, _ := this.PropGet(0x0000059a, nil)
	return NewQueryTables(retVal.IDispatch(), false, true)
}

func (this *DialogSheet) DisplayPageBreaks() bool {
	retVal, _ := this.PropGet(0x0000059b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DialogSheet) SetDisplayPageBreaks(rhs bool)  {
	_ = this.PropPut(0x0000059b, []interface{}{rhs})
}

func (this *DialogSheet) Comments() *Comments {
	retVal, _ := this.PropGet(0x0000023f, nil)
	return NewComments(retVal.IDispatch(), false, true)
}

func (this *DialogSheet) Hyperlinks() *Hyperlinks {
	retVal, _ := this.PropGet(0x00000571, nil)
	return NewHyperlinks(retVal.IDispatch(), false, true)
}

func (this *DialogSheet) ClearCircles()  {
	retVal, _ := this.Call(0x0000059c, nil)
	_= retVal
}

func (this *DialogSheet) CircleInvalid()  {
	retVal, _ := this.Call(0x0000059d, nil)
	_= retVal
}

func (this *DialogSheet) DisplayRightToLeft_() int32 {
	retVal, _ := this.PropGet(0x00000288, nil)
	return retVal.LValVal()
}

func (this *DialogSheet) SetDisplayRightToLeft_(rhs int32)  {
	_ = this.PropPut(0x00000288, []interface{}{rhs})
}

func (this *DialogSheet) AutoFilter() *AutoFilter {
	retVal, _ := this.PropGet(0x00000319, nil)
	return NewAutoFilter(retVal.IDispatch(), false, true)
}

func (this *DialogSheet) DisplayRightToLeft() bool {
	retVal, _ := this.PropGet(0x000006ee, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DialogSheet) SetDisplayRightToLeft(rhs bool)  {
	_ = this.PropPut(0x000006ee, []interface{}{rhs})
}

func (this *DialogSheet) Scripts() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000718, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var DialogSheet_PrintOut__OptArgs= []string{
	"From", "To", "Copies", "Preview", 
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName", 
}

func (this *DialogSheet) PrintOut_(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(DialogSheet_PrintOut__OptArgs, optArgs)
	retVal, _ := this.Call(0x000006ec, nil, optArgs...)
	_= retVal
}

var DialogSheet_CheckSpelling__OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang", 
	"IgnoreFinalYaa", "SpellScript", 
}

func (this *DialogSheet) CheckSpelling_(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(DialogSheet_CheckSpelling__OptArgs, optArgs)
	retVal, _ := this.Call(0x00000719, nil, optArgs...)
	_= retVal
}

func (this *DialogSheet) Tab() *Tab {
	retVal, _ := this.PropGet(0x00000411, nil)
	return NewTab(retVal.IDispatch(), false, true)
}

func (this *DialogSheet) MailEnvelope() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000007e5, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var DialogSheet_SaveAs_OptArgs= []string{
	"FileFormat", "Password", "WriteResPassword", "ReadOnlyRecommended", 
	"CreateBackup", "AddToMru", "TextCodepage", "TextVisualLayout", "Local", 
}

func (this *DialogSheet) SaveAs(filename string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(DialogSheet_SaveAs_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000785, []interface{}{filename}, optArgs...)
	_= retVal
}

func (this *DialogSheet) CustomProperties() *CustomProperties {
	retVal, _ := this.PropGet(0x000007ee, nil)
	return NewCustomProperties(retVal.IDispatch(), false, true)
}

func (this *DialogSheet) SmartTags() *SmartTags {
	retVal, _ := this.PropGet(0x000007e0, nil)
	return NewSmartTags(retVal.IDispatch(), false, true)
}

func (this *DialogSheet) Protection() *Protection {
	retVal, _ := this.PropGet(0x000000b0, nil)
	return NewProtection(retVal.IDispatch(), false, true)
}

var DialogSheet_PasteSpecial_OptArgs= []string{
	"Format", "Link", "DisplayAsIcon", "IconFileName", 
	"IconIndex", "IconLabel", "NoHTMLFormatting", 
}

func (this *DialogSheet) PasteSpecial(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(DialogSheet_PasteSpecial_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000788, nil, optArgs...)
	_= retVal
}

var DialogSheet_Protect_OptArgs= []string{
	"Password", "DrawingObjects", "Contents", "Scenarios", 
	"UserInterfaceOnly", "AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows", 
	"AllowInsertingColumns", "AllowInsertingRows", "AllowInsertingHyperlinks", "AllowDeletingColumns", 
	"AllowDeletingRows", "AllowSorting", "AllowFiltering", "AllowUsingPivotTables", 
}

func (this *DialogSheet) Protect(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(DialogSheet_Protect_OptArgs, optArgs)
	retVal, _ := this.Call(0x000007ed, nil, optArgs...)
	_= retVal
}

func (this *DialogSheet) Dummy113_()  {
	retVal, _ := this.Call(0x00010071, nil)
	_= retVal
}

func (this *DialogSheet) Dummy114_()  {
	retVal, _ := this.Call(0x00010072, nil)
	_= retVal
}

func (this *DialogSheet) Dummy115_()  {
	retVal, _ := this.Call(0x00010073, nil)
	_= retVal
}

var DialogSheet_PrintOut_OptArgs= []string{
	"From", "To", "Copies", "Preview", 
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName", 
}

func (this *DialogSheet) PrintOut(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(DialogSheet_PrintOut_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000939, nil, optArgs...)
	_= retVal
}

func (this *DialogSheet) EnableFormatConditionsCalculation() bool {
	retVal, _ := this.PropGet(0x000009cf, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DialogSheet) SetEnableFormatConditionsCalculation(rhs bool)  {
	_ = this.PropPut(0x000009cf, []interface{}{rhs})
}

func (this *DialogSheet) Sort() *Sort {
	retVal, _ := this.PropGet(0x00000370, nil)
	return NewSort(retVal.IDispatch(), false, true)
}

var DialogSheet_ExportAsFixedFormat_OptArgs= []string{
	"Filename", "Quality", "IncludeDocProperties", "IgnorePrintAreas", 
	"From", "To", "OpenAfterPublish", "FixedFormatExtClassPtr", 
}

func (this *DialogSheet) ExportAsFixedFormat(type_ int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(DialogSheet_ExportAsFixedFormat_OptArgs, optArgs)
	retVal, _ := this.Call(0x000009bd, []interface{}{type_}, optArgs...)
	_= retVal
}

func (this *DialogSheet) PrintedCommentPages() int32 {
	retVal, _ := this.PropGet(0x00000b29, nil)
	return retVal.LValVal()
}

func (this *DialogSheet) DefaultButton() ole.Variant {
	retVal, _ := this.PropGet(0x00000359, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DialogSheet) SetDefaultButton(rhs interface{})  {
	_ = this.PropPut(0x00000359, []interface{}{rhs})
}

func (this *DialogSheet) DialogFrame() *DialogFrame {
	retVal, _ := this.PropGet(0x00000347, nil)
	return NewDialogFrame(retVal.IDispatch(), false, true)
}

var DialogSheet_EditBoxes_OptArgs= []string{
	"Index", 
}

func (this *DialogSheet) EditBoxes(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(DialogSheet_EditBoxes_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000033c, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DialogSheet) Focus() ole.Variant {
	retVal, _ := this.PropGet(0x0000032e, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DialogSheet) SetFocus(rhs interface{})  {
	_ = this.PropPut(0x0000032e, []interface{}{rhs})
}

var DialogSheet_Hide_OptArgs= []string{
	"Cancel", 
}

func (this *DialogSheet) Hide(optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(DialogSheet_Hide_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000032d, nil, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DialogSheet) Show() bool {
	retVal, _ := this.Call(0x000001f0, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

