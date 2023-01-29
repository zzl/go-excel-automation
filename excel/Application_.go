package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
)

// 000208D5-0000-0000-C000-000000000046
var IID_Application_ = syscall.GUID{0x000208D5, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Application_ struct {
	ole.OleClient
}

func NewApplication_(pDisp *win32.IDispatch, addRef bool, scoped bool) *Application_ {
	if pDisp == nil {
		return nil
	}
	p := &Application_{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func Application_FromVar(v ole.Variant) *Application_ {
	return NewApplication_(v.IDispatch(), false, false)
}

func (this *Application_) IID() *syscall.GUID {
	return &IID_Application_
}

func (this *Application_) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Application_) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Application_) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Application_) Parent() *Application {
	retVal, _ := this.PropGet(0x00000096, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Application_) ActiveCell() *Range {
	retVal, _ := this.PropGet(0x00000131, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Application_) ActiveChart() *Chart {
	retVal, _ := this.PropGet(0x000000b7, nil)
	return NewChart(retVal.IDispatch(), false, true)
}

func (this *Application_) ActiveDialog() *DialogSheet {
	retVal, _ := this.PropGet(0x0000032f, nil)
	return NewDialogSheet(retVal.IDispatch(), false, true)
}

func (this *Application_) ActiveMenuBar() *MenuBar {
	retVal, _ := this.PropGet(0x000002f6, nil)
	return NewMenuBar(retVal.IDispatch(), false, true)
}

func (this *Application_) ActivePrinter() string {
	retVal, _ := this.PropGet(0x00000132, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetActivePrinter(rhs string) {
	_ = this.PropPut(0x00000132, []interface{}{rhs})
}

func (this *Application_) ActiveSheet() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000133, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) ActiveWindow() *Window {
	retVal, _ := this.PropGet(0x000002f7, nil)
	return NewWindow(retVal.IDispatch(), false, true)
}

func (this *Application_) ActiveWorkbook() *Workbook {
	retVal, _ := this.PropGet(0x00000134, nil)
	return NewWorkbook(retVal.IDispatch(), false, true)
}

func (this *Application_) AddIns() *AddIns {
	retVal, _ := this.PropGet(0x00000225, nil)
	return NewAddIns(retVal.IDispatch(), false, true)
}

func (this *Application_) Assistant() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000059e, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) Calculate() {
	retVal, _ := this.Call(0x00000117, nil)
	_ = retVal
}

func (this *Application_) Cells() *Range {
	retVal, _ := this.PropGet(0x000000ee, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Application_) Charts() *Sheets {
	retVal, _ := this.PropGet(0x00000079, nil)
	return NewSheets(retVal.IDispatch(), false, true)
}

func (this *Application_) Columns() *Range {
	retVal, _ := this.PropGet(0x000000f1, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Application_) CommandBars() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000059f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) DDEAppReturnCode() int32 {
	retVal, _ := this.PropGet(0x0000014c, nil)
	return retVal.LValVal()
}

func (this *Application_) DDEExecute(channel int32, string string) {
	retVal, _ := this.Call(0x0000014d, []interface{}{channel, string})
	_ = retVal
}

func (this *Application_) DDEInitiate(app string, topic string) int32 {
	retVal, _ := this.Call(0x0000014e, []interface{}{app, topic})
	return retVal.LValVal()
}

func (this *Application_) DDEPoke(channel int32, item interface{}, data interface{}) {
	retVal, _ := this.Call(0x0000014f, []interface{}{channel, item, data})
	_ = retVal
}

func (this *Application_) DDERequest(channel int32, item string) ole.Variant {
	retVal, _ := this.Call(0x00000150, []interface{}{channel, item})
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) DDETerminate(channel int32) {
	retVal, _ := this.Call(0x00000151, []interface{}{channel})
	_ = retVal
}

func (this *Application_) DialogSheets() *Sheets {
	retVal, _ := this.PropGet(0x000002fc, nil)
	return NewSheets(retVal.IDispatch(), false, true)
}

func (this *Application_) Evaluate(name interface{}) ole.Variant {
	retVal, _ := this.Call(0x00000001, []interface{}{name})
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) Evaluate_(name interface{}) ole.Variant {
	retVal, _ := this.Call(-5, []interface{}{name})
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) ExecuteExcel4Macro(string string) ole.Variant {
	retVal, _ := this.Call(0x0000015e, []interface{}{string})
	com.AddToScope(retVal)
	return *retVal
}

var Application__Intersect_OptArgs = []string{
	"Arg3", "Arg4", "Arg5", "Arg6",
	"Arg7", "Arg8", "Arg9", "Arg10",
	"Arg11", "Arg12", "Arg13", "Arg14",
	"Arg15", "Arg16", "Arg17", "Arg18",
	"Arg19", "Arg20", "Arg21", "Arg22",
	"Arg23", "Arg24", "Arg25", "Arg26",
	"Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *Application_) Intersect(arg1 *Range, arg2 *Range, optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Application__Intersect_OptArgs, optArgs)
	retVal, _ := this.Call(0x000002fe, []interface{}{arg1, arg2}, optArgs...)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Application_) MenuBars() *MenuBars {
	retVal, _ := this.PropGet(0x0000024d, nil)
	return NewMenuBars(retVal.IDispatch(), false, true)
}

func (this *Application_) Modules() *Modules {
	retVal, _ := this.PropGet(0x00000246, nil)
	return NewModules(retVal.IDispatch(), false, true)
}

func (this *Application_) Names() *Names {
	retVal, _ := this.PropGet(0x000001ba, nil)
	return NewNames(retVal.IDispatch(), false, true)
}

var Application__Range_OptArgs = []string{
	"Cell2",
}

func (this *Application_) Range(cell1 interface{}, optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Application__Range_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x000000c5, []interface{}{cell1}, optArgs...)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Application_) Rows() *Range {
	retVal, _ := this.PropGet(0x00000102, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

var Application__Run_OptArgs = []string{
	"Macro", "Arg1", "Arg2", "Arg3",
	"Arg4", "Arg5", "Arg6", "Arg7",
	"Arg8", "Arg9", "Arg10", "Arg11",
	"Arg12", "Arg13", "Arg14", "Arg15",
	"Arg16", "Arg17", "Arg18", "Arg19",
	"Arg20", "Arg21", "Arg22", "Arg23",
	"Arg24", "Arg25", "Arg26", "Arg27",
	"Arg28", "Arg29", "Arg30",
}

func (this *Application_) Run(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Application__Run_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000103, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Application__Run2__OptArgs = []string{
	"Macro", "Arg1", "Arg2", "Arg3",
	"Arg4", "Arg5", "Arg6", "Arg7",
	"Arg8", "Arg9", "Arg10", "Arg11",
	"Arg12", "Arg13", "Arg14", "Arg15",
	"Arg16", "Arg17", "Arg18", "Arg19",
	"Arg20", "Arg21", "Arg22", "Arg23",
	"Arg24", "Arg25", "Arg26", "Arg27",
	"Arg28", "Arg29", "Arg30",
}

func (this *Application_) Run2_(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Application__Run2__OptArgs, optArgs)
	retVal, _ := this.Call(0x00000326, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) Selection() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000093, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Application__SendKeys_OptArgs = []string{
	"Wait",
}

func (this *Application_) SendKeys(keys interface{}, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Application__SendKeys_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000017f, []interface{}{keys}, optArgs...)
	_ = retVal
}

func (this *Application_) Sheets() *Sheets {
	retVal, _ := this.PropGet(0x000001e5, nil)
	return NewSheets(retVal.IDispatch(), false, true)
}

func (this *Application_) ShortcutMenus(index int32) *Menu {
	retVal, _ := this.PropGet(0x00000308, []interface{}{index})
	return NewMenu(retVal.IDispatch(), false, true)
}

func (this *Application_) ThisWorkbook() *Workbook {
	retVal, _ := this.PropGet(0x0000030a, nil)
	return NewWorkbook(retVal.IDispatch(), false, true)
}

func (this *Application_) Toolbars() *Toolbars {
	retVal, _ := this.PropGet(0x00000228, nil)
	return NewToolbars(retVal.IDispatch(), false, true)
}

var Application__Union_OptArgs = []string{
	"Arg3", "Arg4", "Arg5", "Arg6",
	"Arg7", "Arg8", "Arg9", "Arg10",
	"Arg11", "Arg12", "Arg13", "Arg14",
	"Arg15", "Arg16", "Arg17", "Arg18",
	"Arg19", "Arg20", "Arg21", "Arg22",
	"Arg23", "Arg24", "Arg25", "Arg26",
	"Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *Application_) Union(arg1 *Range, arg2 *Range, optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Application__Union_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000030b, []interface{}{arg1, arg2}, optArgs...)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Application_) Windows() *Windows {
	retVal, _ := this.PropGet(0x000001ae, nil)
	return NewWindows(retVal.IDispatch(), false, true)
}

func (this *Application_) Workbooks() *Workbooks {
	retVal, _ := this.PropGet(0x0000023c, nil)
	return NewWorkbooks(retVal.IDispatch(), false, true)
}

func (this *Application_) WorksheetFunction() *WorksheetFunction {
	retVal, _ := this.PropGet(0x000005a0, nil)
	return NewWorksheetFunction(retVal.IDispatch(), false, true)
}

func (this *Application_) Worksheets() *Sheets {
	retVal, _ := this.PropGet(0x000001ee, nil)
	return NewSheets(retVal.IDispatch(), false, true)
}

func (this *Application_) Excel4IntlMacroSheets() *Sheets {
	retVal, _ := this.PropGet(0x00000245, nil)
	return NewSheets(retVal.IDispatch(), false, true)
}

func (this *Application_) Excel4MacroSheets() *Sheets {
	retVal, _ := this.PropGet(0x00000243, nil)
	return NewSheets(retVal.IDispatch(), false, true)
}

func (this *Application_) ActivateMicrosoftApp(index int32) {
	retVal, _ := this.Call(0x00000447, []interface{}{index})
	_ = retVal
}

var Application__AddChartAutoFormat_OptArgs = []string{
	"Description",
}

func (this *Application_) AddChartAutoFormat(chart interface{}, name string, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Application__AddChartAutoFormat_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d8, []interface{}{chart, name}, optArgs...)
	_ = retVal
}

var Application__AddCustomList_OptArgs = []string{
	"ByRow",
}

func (this *Application_) AddCustomList(listArray interface{}, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Application__AddCustomList_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000030c, []interface{}{listArray}, optArgs...)
	_ = retVal
}

func (this *Application_) AlertBeforeOverwriting() bool {
	retVal, _ := this.PropGet(0x000003a2, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetAlertBeforeOverwriting(rhs bool) {
	_ = this.PropPut(0x000003a2, []interface{}{rhs})
}

func (this *Application_) AltStartupPath() string {
	retVal, _ := this.PropGet(0x00000139, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetAltStartupPath(rhs string) {
	_ = this.PropPut(0x00000139, []interface{}{rhs})
}

func (this *Application_) AskToUpdateLinks() bool {
	retVal, _ := this.PropGet(0x000003e0, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetAskToUpdateLinks(rhs bool) {
	_ = this.PropPut(0x000003e0, []interface{}{rhs})
}

func (this *Application_) EnableAnimations() bool {
	retVal, _ := this.PropGet(0x0000049c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetEnableAnimations(rhs bool) {
	_ = this.PropPut(0x0000049c, []interface{}{rhs})
}

func (this *Application_) AutoCorrect() *AutoCorrect {
	retVal, _ := this.PropGet(0x00000479, nil)
	return NewAutoCorrect(retVal.IDispatch(), false, true)
}

func (this *Application_) Build() int32 {
	retVal, _ := this.PropGet(0x0000013a, nil)
	return retVal.LValVal()
}

func (this *Application_) CalculateBeforeSave() bool {
	retVal, _ := this.PropGet(0x0000013b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetCalculateBeforeSave(rhs bool) {
	_ = this.PropPut(0x0000013b, []interface{}{rhs})
}

func (this *Application_) Calculation() int32 {
	retVal, _ := this.PropGet(0x0000013c, nil)
	return retVal.LValVal()
}

func (this *Application_) SetCalculation(rhs int32) {
	_ = this.PropPut(0x0000013c, []interface{}{rhs})
}

var Application__Caller_OptArgs = []string{
	"Index",
}

func (this *Application_) Caller(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Application__Caller_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000013d, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) CanPlaySounds() bool {
	retVal, _ := this.PropGet(0x0000013e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) CanRecordSounds() bool {
	retVal, _ := this.PropGet(0x0000013f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) Caption() string {
	retVal, _ := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetCaption(rhs string) {
	_ = this.PropPut(0x0000008b, []interface{}{rhs})
}

func (this *Application_) CellDragAndDrop() bool {
	retVal, _ := this.PropGet(0x00000140, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetCellDragAndDrop(rhs bool) {
	_ = this.PropPut(0x00000140, []interface{}{rhs})
}

func (this *Application_) CentimetersToPoints(centimeters float64) float64 {
	retVal, _ := this.Call(0x0000043e, []interface{}{centimeters})
	return retVal.DblValVal()
}

var Application__CheckSpelling_OptArgs = []string{
	"CustomDictionary", "IgnoreUppercase",
}

func (this *Application_) CheckSpelling(word string, optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(Application__CheckSpelling_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001f9, []interface{}{word}, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var Application__ClipboardFormats_OptArgs = []string{
	"Index",
}

func (this *Application_) ClipboardFormats(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Application__ClipboardFormats_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x00000141, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) DisplayClipboardWindow() bool {
	retVal, _ := this.PropGet(0x00000142, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDisplayClipboardWindow(rhs bool) {
	_ = this.PropPut(0x00000142, []interface{}{rhs})
}

func (this *Application_) ColorButtons() bool {
	retVal, _ := this.PropGet(0x0000016d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetColorButtons(rhs bool) {
	_ = this.PropPut(0x0000016d, []interface{}{rhs})
}

func (this *Application_) CommandUnderlines() int32 {
	retVal, _ := this.PropGet(0x00000143, nil)
	return retVal.LValVal()
}

func (this *Application_) SetCommandUnderlines(rhs int32) {
	_ = this.PropPut(0x00000143, []interface{}{rhs})
}

func (this *Application_) ConstrainNumeric() bool {
	retVal, _ := this.PropGet(0x00000144, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetConstrainNumeric(rhs bool) {
	_ = this.PropPut(0x00000144, []interface{}{rhs})
}

var Application__ConvertFormula_OptArgs = []string{
	"ToReferenceStyle", "ToAbsolute", "RelativeTo",
}

func (this *Application_) ConvertFormula(formula interface{}, fromReferenceStyle int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Application__ConvertFormula_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000145, []interface{}{formula, fromReferenceStyle}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) CopyObjectsWithCells() bool {
	retVal, _ := this.PropGet(0x000003df, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetCopyObjectsWithCells(rhs bool) {
	_ = this.PropPut(0x000003df, []interface{}{rhs})
}

func (this *Application_) Cursor() int32 {
	retVal, _ := this.PropGet(0x00000489, nil)
	return retVal.LValVal()
}

func (this *Application_) SetCursor(rhs int32) {
	_ = this.PropPut(0x00000489, []interface{}{rhs})
}

func (this *Application_) CustomListCount() int32 {
	retVal, _ := this.PropGet(0x00000313, nil)
	return retVal.LValVal()
}

func (this *Application_) CutCopyMode() int32 {
	retVal, _ := this.PropGet(0x0000014a, nil)
	return retVal.LValVal()
}

func (this *Application_) SetCutCopyMode(rhs int32) {
	_ = this.PropPut(0x0000014a, []interface{}{rhs})
}

func (this *Application_) DataEntryMode() int32 {
	retVal, _ := this.PropGet(0x0000014b, nil)
	return retVal.LValVal()
}

func (this *Application_) SetDataEntryMode(rhs int32) {
	_ = this.PropPut(0x0000014b, []interface{}{rhs})
}

var Application__Dummy1_OptArgs = []string{
	"Arg1", "Arg2", "Arg3", "Arg4",
}

func (this *Application_) Dummy1(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Application__Dummy1_OptArgs, optArgs)
	retVal, _ := this.Call(0x000006f6, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Application__Dummy2_OptArgs = []string{
	"Arg1", "Arg2", "Arg3", "Arg4",
	"Arg5", "Arg6", "Arg7", "Arg8",
}

func (this *Application_) Dummy2(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Application__Dummy2_OptArgs, optArgs)
	retVal, _ := this.Call(0x000006f7, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) Dummy3() ole.Variant {
	retVal, _ := this.Call(0x000006f8, nil)
	com.AddToScope(retVal)
	return *retVal
}

var Application__Dummy4_OptArgs = []string{
	"Arg1", "Arg2", "Arg3", "Arg4",
	"Arg5", "Arg6", "Arg7", "Arg8",
	"Arg9", "Arg10", "Arg11", "Arg12",
	"Arg13", "Arg14", "Arg15",
}

func (this *Application_) Dummy4(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Application__Dummy4_OptArgs, optArgs)
	retVal, _ := this.Call(0x000006f9, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Application__Dummy5_OptArgs = []string{
	"Arg1", "Arg2", "Arg3", "Arg4",
	"Arg5", "Arg6", "Arg7", "Arg8",
	"Arg9", "Arg10", "Arg11", "Arg12", "Arg13",
}

func (this *Application_) Dummy5(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Application__Dummy5_OptArgs, optArgs)
	retVal, _ := this.Call(0x000006fa, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) Dummy6() ole.Variant {
	retVal, _ := this.Call(0x000006fb, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) Dummy7() ole.Variant {
	retVal, _ := this.Call(0x000006fc, nil)
	com.AddToScope(retVal)
	return *retVal
}

var Application__Dummy8_OptArgs = []string{
	"Arg1",
}

func (this *Application_) Dummy8(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Application__Dummy8_OptArgs, optArgs)
	retVal, _ := this.Call(0x000006fd, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) Dummy9() ole.Variant {
	retVal, _ := this.Call(0x000006fe, nil)
	com.AddToScope(retVal)
	return *retVal
}

var Application__Dummy10_OptArgs = []string{
	"arg",
}

func (this *Application_) Dummy10(optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(Application__Dummy10_OptArgs, optArgs)
	retVal, _ := this.Call(0x000006ff, nil, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) Dummy11() {
	retVal, _ := this.Call(0x00000700, nil)
	_ = retVal
}

func (this *Application_) Default_() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) DefaultFilePath() string {
	retVal, _ := this.PropGet(0x0000040e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetDefaultFilePath(rhs string) {
	_ = this.PropPut(0x0000040e, []interface{}{rhs})
}

func (this *Application_) DeleteChartAutoFormat(name string) {
	retVal, _ := this.Call(0x000000d9, []interface{}{name})
	_ = retVal
}

func (this *Application_) DeleteCustomList(listNum int32) {
	retVal, _ := this.Call(0x0000030f, []interface{}{listNum})
	_ = retVal
}

func (this *Application_) Dialogs() *Dialogs {
	retVal, _ := this.PropGet(0x000002f9, nil)
	return NewDialogs(retVal.IDispatch(), false, true)
}

func (this *Application_) DisplayAlerts() bool {
	retVal, _ := this.PropGet(0x00000157, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDisplayAlerts(rhs bool) {
	_ = this.PropPut(0x00000157, []interface{}{rhs})
}

func (this *Application_) DisplayFormulaBar() bool {
	retVal, _ := this.PropGet(0x00000158, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDisplayFormulaBar(rhs bool) {
	_ = this.PropPut(0x00000158, []interface{}{rhs})
}

func (this *Application_) DisplayFullScreen() bool {
	retVal, _ := this.PropGet(0x00000425, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDisplayFullScreen(rhs bool) {
	_ = this.PropPut(0x00000425, []interface{}{rhs})
}

func (this *Application_) DisplayNoteIndicator() bool {
	retVal, _ := this.PropGet(0x00000159, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDisplayNoteIndicator(rhs bool) {
	_ = this.PropPut(0x00000159, []interface{}{rhs})
}

func (this *Application_) DisplayCommentIndicator() int32 {
	retVal, _ := this.PropGet(0x000004ac, nil)
	return retVal.LValVal()
}

func (this *Application_) SetDisplayCommentIndicator(rhs int32) {
	_ = this.PropPut(0x000004ac, []interface{}{rhs})
}

func (this *Application_) DisplayExcel4Menus() bool {
	retVal, _ := this.PropGet(0x0000039f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDisplayExcel4Menus(rhs bool) {
	_ = this.PropPut(0x0000039f, []interface{}{rhs})
}

func (this *Application_) DisplayRecentFiles() bool {
	retVal, _ := this.PropGet(0x0000039e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDisplayRecentFiles(rhs bool) {
	_ = this.PropPut(0x0000039e, []interface{}{rhs})
}

func (this *Application_) DisplayScrollBars() bool {
	retVal, _ := this.PropGet(0x0000015a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDisplayScrollBars(rhs bool) {
	_ = this.PropPut(0x0000015a, []interface{}{rhs})
}

func (this *Application_) DisplayStatusBar() bool {
	retVal, _ := this.PropGet(0x0000015b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDisplayStatusBar(rhs bool) {
	_ = this.PropPut(0x0000015b, []interface{}{rhs})
}

func (this *Application_) DoubleClick() {
	retVal, _ := this.Call(0x0000015d, nil)
	_ = retVal
}

func (this *Application_) EditDirectlyInCell() bool {
	retVal, _ := this.PropGet(0x000003a1, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetEditDirectlyInCell(rhs bool) {
	_ = this.PropPut(0x000003a1, []interface{}{rhs})
}

func (this *Application_) EnableAutoComplete() bool {
	retVal, _ := this.PropGet(0x0000049b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetEnableAutoComplete(rhs bool) {
	_ = this.PropPut(0x0000049b, []interface{}{rhs})
}

func (this *Application_) EnableCancelKey() int32 {
	retVal, _ := this.PropGet(0x00000448, nil)
	return retVal.LValVal()
}

func (this *Application_) SetEnableCancelKey(rhs int32) {
	_ = this.PropPut(0x00000448, []interface{}{rhs})
}

func (this *Application_) EnableSound() bool {
	retVal, _ := this.PropGet(0x000004ad, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetEnableSound(rhs bool) {
	_ = this.PropPut(0x000004ad, []interface{}{rhs})
}

func (this *Application_) EnableTipWizard() bool {
	retVal, _ := this.PropGet(0x00000428, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetEnableTipWizard(rhs bool) {
	_ = this.PropPut(0x00000428, []interface{}{rhs})
}

var Application__FileConverters_OptArgs = []string{
	"Index1", "Index2",
}

func (this *Application_) FileConverters(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Application__FileConverters_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x000003a3, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) FileSearch() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000004b0, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) FileFind() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000004b1, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) FindFile_() {
	retVal, _ := this.Call(0x0000042c, nil)
	_ = retVal
}

func (this *Application_) FixedDecimal() bool {
	retVal, _ := this.PropGet(0x0000015f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetFixedDecimal(rhs bool) {
	_ = this.PropPut(0x0000015f, []interface{}{rhs})
}

func (this *Application_) FixedDecimalPlaces() int32 {
	retVal, _ := this.PropGet(0x00000160, nil)
	return retVal.LValVal()
}

func (this *Application_) SetFixedDecimalPlaces(rhs int32) {
	_ = this.PropPut(0x00000160, []interface{}{rhs})
}

func (this *Application_) GetCustomListContents(listNum int32) ole.Variant {
	retVal, _ := this.Call(0x00000312, []interface{}{listNum})
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) GetCustomListNum(listArray interface{}) int32 {
	retVal, _ := this.Call(0x00000311, []interface{}{listArray})
	return retVal.LValVal()
}

var Application__GetOpenFilename_OptArgs = []string{
	"FileFilter", "FilterIndex", "Title", "ButtonText", "MultiSelect",
}

func (this *Application_) GetOpenFilename(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Application__GetOpenFilename_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000433, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Application__GetSaveAsFilename_OptArgs = []string{
	"InitialFilename", "FileFilter", "FilterIndex", "Title", "ButtonText",
}

func (this *Application_) GetSaveAsFilename(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Application__GetSaveAsFilename_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000434, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Application__Goto_OptArgs = []string{
	"Reference", "Scroll",
}

func (this *Application_) Goto(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Application__Goto_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001db, nil, optArgs...)
	_ = retVal
}

func (this *Application_) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *Application_) SetHeight(rhs float64) {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

var Application__Help_OptArgs = []string{
	"HelpFile", "HelpContextID",
}

func (this *Application_) Help(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Application__Help_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000162, nil, optArgs...)
	_ = retVal
}

func (this *Application_) IgnoreRemoteRequests() bool {
	retVal, _ := this.PropGet(0x00000164, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetIgnoreRemoteRequests(rhs bool) {
	_ = this.PropPut(0x00000164, []interface{}{rhs})
}

func (this *Application_) InchesToPoints(inches float64) float64 {
	retVal, _ := this.Call(0x0000043f, []interface{}{inches})
	return retVal.DblValVal()
}

var Application__InputBox_OptArgs = []string{
	"Title", "Default", "Left", "Top",
	"HelpFile", "HelpContextID", "Type",
}

func (this *Application_) InputBox(prompt string, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Application__InputBox_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000165, []interface{}{prompt}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) Interactive() bool {
	retVal, _ := this.PropGet(0x00000169, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetInteractive(rhs bool) {
	_ = this.PropPut(0x00000169, []interface{}{rhs})
}

var Application__International_OptArgs = []string{
	"Index",
}

func (this *Application_) International(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Application__International_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000016a, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) Iteration() bool {
	retVal, _ := this.PropGet(0x0000016b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetIteration(rhs bool) {
	_ = this.PropPut(0x0000016b, []interface{}{rhs})
}

func (this *Application_) LargeButtons() bool {
	retVal, _ := this.PropGet(0x0000016c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetLargeButtons(rhs bool) {
	_ = this.PropPut(0x0000016c, []interface{}{rhs})
}

func (this *Application_) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *Application_) SetLeft(rhs float64) {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *Application_) LibraryPath() string {
	retVal, _ := this.PropGet(0x0000016e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var Application__MacroOptions__OptArgs = []string{
	"Macro", "Description", "HasMenu", "MenuText",
	"HasShortcutKey", "ShortcutKey", "Category", "StatusBar",
	"HelpContextID", "HelpFile",
}

func (this *Application_) MacroOptions_(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Application__MacroOptions__OptArgs, optArgs)
	retVal, _ := this.Call(0x0000046f, nil, optArgs...)
	_ = retVal
}

func (this *Application_) MailLogoff() {
	retVal, _ := this.Call(0x000003b1, nil)
	_ = retVal
}

var Application__MailLogon_OptArgs = []string{
	"Name", "Password", "DownloadNewMail",
}

func (this *Application_) MailLogon(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Application__MailLogon_OptArgs, optArgs)
	retVal, _ := this.Call(0x000003af, nil, optArgs...)
	_ = retVal
}

func (this *Application_) MailSession() ole.Variant {
	retVal, _ := this.PropGet(0x000003ae, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) MailSystem() int32 {
	retVal, _ := this.PropGet(0x000003cb, nil)
	return retVal.LValVal()
}

func (this *Application_) MathCoprocessorAvailable() bool {
	retVal, _ := this.PropGet(0x0000016f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) MaxChange() float64 {
	retVal, _ := this.PropGet(0x00000170, nil)
	return retVal.DblValVal()
}

func (this *Application_) SetMaxChange(rhs float64) {
	_ = this.PropPut(0x00000170, []interface{}{rhs})
}

func (this *Application_) MaxIterations() int32 {
	retVal, _ := this.PropGet(0x00000171, nil)
	return retVal.LValVal()
}

func (this *Application_) SetMaxIterations(rhs int32) {
	_ = this.PropPut(0x00000171, []interface{}{rhs})
}

func (this *Application_) MemoryFree() int32 {
	retVal, _ := this.PropGet(0x00000172, nil)
	return retVal.LValVal()
}

func (this *Application_) MemoryTotal() int32 {
	retVal, _ := this.PropGet(0x00000173, nil)
	return retVal.LValVal()
}

func (this *Application_) MemoryUsed() int32 {
	retVal, _ := this.PropGet(0x00000174, nil)
	return retVal.LValVal()
}

func (this *Application_) MouseAvailable() bool {
	retVal, _ := this.PropGet(0x00000175, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) MoveAfterReturn() bool {
	retVal, _ := this.PropGet(0x00000176, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetMoveAfterReturn(rhs bool) {
	_ = this.PropPut(0x00000176, []interface{}{rhs})
}

func (this *Application_) MoveAfterReturnDirection() int32 {
	retVal, _ := this.PropGet(0x00000478, nil)
	return retVal.LValVal()
}

func (this *Application_) SetMoveAfterReturnDirection(rhs int32) {
	_ = this.PropPut(0x00000478, []interface{}{rhs})
}

func (this *Application_) RecentFiles() *RecentFiles {
	retVal, _ := this.PropGet(0x000004b2, nil)
	return NewRecentFiles(retVal.IDispatch(), false, true)
}

func (this *Application_) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) NextLetter() *Workbook {
	retVal, _ := this.Call(0x000003cc, nil)
	return NewWorkbook(retVal.IDispatch(), false, true)
}

func (this *Application_) NetworkTemplatesPath() string {
	retVal, _ := this.PropGet(0x00000184, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) ODBCErrors() *ODBCErrors {
	retVal, _ := this.PropGet(0x000004b3, nil)
	return NewODBCErrors(retVal.IDispatch(), false, true)
}

func (this *Application_) ODBCTimeout() int32 {
	retVal, _ := this.PropGet(0x000004b4, nil)
	return retVal.LValVal()
}

func (this *Application_) SetODBCTimeout(rhs int32) {
	_ = this.PropPut(0x000004b4, []interface{}{rhs})
}

func (this *Application_) OnCalculate() string {
	retVal, _ := this.PropGet(0x00000271, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetOnCalculate(rhs string) {
	_ = this.PropPut(0x00000271, []interface{}{rhs})
}

func (this *Application_) OnData() string {
	retVal, _ := this.PropGet(0x00000275, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetOnData(rhs string) {
	_ = this.PropPut(0x00000275, []interface{}{rhs})
}

func (this *Application_) OnDoubleClick() string {
	retVal, _ := this.PropGet(0x00000274, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetOnDoubleClick(rhs string) {
	_ = this.PropPut(0x00000274, []interface{}{rhs})
}

func (this *Application_) OnEntry() string {
	retVal, _ := this.PropGet(0x00000273, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetOnEntry(rhs string) {
	_ = this.PropPut(0x00000273, []interface{}{rhs})
}

var Application__OnKey_OptArgs = []string{
	"Procedure",
}

func (this *Application_) OnKey(key string, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Application__OnKey_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000272, []interface{}{key}, optArgs...)
	_ = retVal
}

func (this *Application_) OnRepeat(text string, procedure string) {
	retVal, _ := this.Call(0x00000301, []interface{}{text, procedure})
	_ = retVal
}

func (this *Application_) OnSheetActivate() string {
	retVal, _ := this.PropGet(0x00000407, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetOnSheetActivate(rhs string) {
	_ = this.PropPut(0x00000407, []interface{}{rhs})
}

func (this *Application_) OnSheetDeactivate() string {
	retVal, _ := this.PropGet(0x00000439, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetOnSheetDeactivate(rhs string) {
	_ = this.PropPut(0x00000439, []interface{}{rhs})
}

var Application__OnTime_OptArgs = []string{
	"LatestTime", "Schedule",
}

func (this *Application_) OnTime(earliestTime interface{}, procedure string, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Application__OnTime_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000270, []interface{}{earliestTime, procedure}, optArgs...)
	_ = retVal
}

func (this *Application_) OnUndo(text string, procedure string) {
	retVal, _ := this.Call(0x00000302, []interface{}{text, procedure})
	_ = retVal
}

func (this *Application_) OnWindow() string {
	retVal, _ := this.PropGet(0x0000026f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetOnWindow(rhs string) {
	_ = this.PropPut(0x0000026f, []interface{}{rhs})
}

func (this *Application_) OperatingSystem() string {
	retVal, _ := this.PropGet(0x00000177, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) OrganizationName() string {
	retVal, _ := this.PropGet(0x00000178, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) Path() string {
	retVal, _ := this.PropGet(0x00000123, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) PathSeparator() string {
	retVal, _ := this.PropGet(0x00000179, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var Application__PreviousSelections_OptArgs = []string{
	"Index",
}

func (this *Application_) PreviousSelections(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Application__PreviousSelections_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000017a, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) PivotTableSelection() bool {
	retVal, _ := this.PropGet(0x000004b5, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetPivotTableSelection(rhs bool) {
	_ = this.PropPut(0x000004b5, []interface{}{rhs})
}

func (this *Application_) PromptForSummaryInfo() bool {
	retVal, _ := this.PropGet(0x00000426, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetPromptForSummaryInfo(rhs bool) {
	_ = this.PropPut(0x00000426, []interface{}{rhs})
}

func (this *Application_) Quit() {
	retVal, _ := this.Call(0x0000012e, nil)
	_ = retVal
}

var Application__RecordMacro_OptArgs = []string{
	"BasicCode", "XlmCode",
}

func (this *Application_) RecordMacro(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Application__RecordMacro_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000305, nil, optArgs...)
	_ = retVal
}

func (this *Application_) RecordRelative() bool {
	retVal, _ := this.PropGet(0x0000017b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) ReferenceStyle() int32 {
	retVal, _ := this.PropGet(0x0000017c, nil)
	return retVal.LValVal()
}

func (this *Application_) SetReferenceStyle(rhs int32) {
	_ = this.PropPut(0x0000017c, []interface{}{rhs})
}

var Application__RegisteredFunctions_OptArgs = []string{
	"Index1", "Index2",
}

func (this *Application_) RegisteredFunctions(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Application__RegisteredFunctions_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x00000307, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) RegisterXLL(filename string) bool {
	retVal, _ := this.Call(0x0000001e, []interface{}{filename})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) Repeat() {
	retVal, _ := this.Call(0x0000012d, nil)
	_ = retVal
}

func (this *Application_) ResetTipWizard() {
	retVal, _ := this.Call(0x000003a0, nil)
	_ = retVal
}

func (this *Application_) RollZoom() bool {
	retVal, _ := this.PropGet(0x000004b6, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetRollZoom(rhs bool) {
	_ = this.PropPut(0x000004b6, []interface{}{rhs})
}

var Application__Save_OptArgs = []string{
	"Filename",
}

func (this *Application_) Save(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Application__Save_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000011b, nil, optArgs...)
	_ = retVal
}

var Application__SaveWorkspace_OptArgs = []string{
	"Filename",
}

func (this *Application_) SaveWorkspace(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Application__SaveWorkspace_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d4, nil, optArgs...)
	_ = retVal
}

func (this *Application_) ScreenUpdating() bool {
	retVal, _ := this.PropGet(0x0000017e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetScreenUpdating(rhs bool) {
	_ = this.PropPut(0x0000017e, []interface{}{rhs})
}

var Application__SetDefaultChart_OptArgs = []string{
	"FormatName", "Gallery",
}

func (this *Application_) SetDefaultChart(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Application__SetDefaultChart_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000db, nil, optArgs...)
	_ = retVal
}

func (this *Application_) SheetsInNewWorkbook() int32 {
	retVal, _ := this.PropGet(0x000003e1, nil)
	return retVal.LValVal()
}

func (this *Application_) SetSheetsInNewWorkbook(rhs int32) {
	_ = this.PropPut(0x000003e1, []interface{}{rhs})
}

func (this *Application_) ShowChartTipNames() bool {
	retVal, _ := this.PropGet(0x000004b7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetShowChartTipNames(rhs bool) {
	_ = this.PropPut(0x000004b7, []interface{}{rhs})
}

func (this *Application_) ShowChartTipValues() bool {
	retVal, _ := this.PropGet(0x000004b8, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetShowChartTipValues(rhs bool) {
	_ = this.PropPut(0x000004b8, []interface{}{rhs})
}

func (this *Application_) StandardFont() string {
	retVal, _ := this.PropGet(0x0000039c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetStandardFont(rhs string) {
	_ = this.PropPut(0x0000039c, []interface{}{rhs})
}

func (this *Application_) StandardFontSize() float64 {
	retVal, _ := this.PropGet(0x0000039d, nil)
	return retVal.DblValVal()
}

func (this *Application_) SetStandardFontSize(rhs float64) {
	_ = this.PropPut(0x0000039d, []interface{}{rhs})
}

func (this *Application_) StartupPath() string {
	retVal, _ := this.PropGet(0x00000181, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) StatusBar() ole.Variant {
	retVal, _ := this.PropGet(0x00000182, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) SetStatusBar(rhs interface{}) {
	_ = this.PropPut(0x00000182, []interface{}{rhs})
}

func (this *Application_) TemplatesPath() string {
	retVal, _ := this.PropGet(0x0000017d, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) ShowToolTips() bool {
	retVal, _ := this.PropGet(0x00000183, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetShowToolTips(rhs bool) {
	_ = this.PropPut(0x00000183, []interface{}{rhs})
}

func (this *Application_) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *Application_) SetTop(rhs float64) {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *Application_) DefaultSaveFormat() int32 {
	retVal, _ := this.PropGet(0x000004b9, nil)
	return retVal.LValVal()
}

func (this *Application_) SetDefaultSaveFormat(rhs int32) {
	_ = this.PropPut(0x000004b9, []interface{}{rhs})
}

func (this *Application_) TransitionMenuKey() string {
	retVal, _ := this.PropGet(0x00000136, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetTransitionMenuKey(rhs string) {
	_ = this.PropPut(0x00000136, []interface{}{rhs})
}

func (this *Application_) TransitionMenuKeyAction() int32 {
	retVal, _ := this.PropGet(0x00000137, nil)
	return retVal.LValVal()
}

func (this *Application_) SetTransitionMenuKeyAction(rhs int32) {
	_ = this.PropPut(0x00000137, []interface{}{rhs})
}

func (this *Application_) TransitionNavigKeys() bool {
	retVal, _ := this.PropGet(0x00000138, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetTransitionNavigKeys(rhs bool) {
	_ = this.PropPut(0x00000138, []interface{}{rhs})
}

func (this *Application_) Undo() {
	retVal, _ := this.Call(0x0000012f, nil)
	_ = retVal
}

func (this *Application_) UsableHeight() float64 {
	retVal, _ := this.PropGet(0x00000185, nil)
	return retVal.DblValVal()
}

func (this *Application_) UsableWidth() float64 {
	retVal, _ := this.PropGet(0x00000186, nil)
	return retVal.DblValVal()
}

func (this *Application_) UserControl() bool {
	retVal, _ := this.PropGet(0x000004ba, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetUserControl(rhs bool) {
	_ = this.PropPut(0x000004ba, []interface{}{rhs})
}

func (this *Application_) UserName() string {
	retVal, _ := this.PropGet(0x00000187, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetUserName(rhs string) {
	_ = this.PropPut(0x00000187, []interface{}{rhs})
}

func (this *Application_) Value() string {
	retVal, _ := this.PropGet(0x00000006, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) VBE() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000004bb, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) Version() string {
	retVal, _ := this.PropGet(0x00000188, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) Visible() bool {
	retVal, _ := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetVisible(rhs bool) {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

var Application__Volatile_OptArgs = []string{
	"Volatile",
}

func (this *Application_) Volatile(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Application__Volatile_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000314, nil, optArgs...)
	_ = retVal
}

func (this *Application_) Wait_(time interface{}) {
	retVal, _ := this.Call(0x00000189, []interface{}{time})
	_ = retVal
}

func (this *Application_) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *Application_) SetWidth(rhs float64) {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *Application_) WindowsForPens() bool {
	retVal, _ := this.PropGet(0x0000018b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) WindowState() int32 {
	retVal, _ := this.PropGet(0x0000018c, nil)
	return retVal.LValVal()
}

func (this *Application_) SetWindowState(rhs int32) {
	_ = this.PropPut(0x0000018c, []interface{}{rhs})
}

func (this *Application_) UILanguage() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Application_) SetUILanguage(rhs int32) {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *Application_) DefaultSheetDirection() int32 {
	retVal, _ := this.PropGet(0x000000e5, nil)
	return retVal.LValVal()
}

func (this *Application_) SetDefaultSheetDirection(rhs int32) {
	_ = this.PropPut(0x000000e5, []interface{}{rhs})
}

func (this *Application_) CursorMovement() int32 {
	retVal, _ := this.PropGet(0x000000e8, nil)
	return retVal.LValVal()
}

func (this *Application_) SetCursorMovement(rhs int32) {
	_ = this.PropPut(0x000000e8, []interface{}{rhs})
}

func (this *Application_) ControlCharacters() bool {
	retVal, _ := this.PropGet(0x000000e9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetControlCharacters(rhs bool) {
	_ = this.PropPut(0x000000e9, []interface{}{rhs})
}

var Application__WSFunction__OptArgs = []string{
	"Arg1", "Arg2", "Arg3", "Arg4",
	"Arg5", "Arg6", "Arg7", "Arg8",
	"Arg9", "Arg10", "Arg11", "Arg12",
	"Arg13", "Arg14", "Arg15", "Arg16",
	"Arg17", "Arg18", "Arg19", "Arg20",
	"Arg21", "Arg22", "Arg23", "Arg24",
	"Arg25", "Arg26", "Arg27", "Arg28",
	"Arg29", "Arg30",
}

func (this *Application_) WSFunction_(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Application__WSFunction__OptArgs, optArgs)
	retVal, _ := this.Call(0x000000a9, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) EnableEvents() bool {
	retVal, _ := this.PropGet(0x000004bc, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetEnableEvents(rhs bool) {
	_ = this.PropPut(0x000004bc, []interface{}{rhs})
}

func (this *Application_) DisplayInfoWindow() bool {
	retVal, _ := this.PropGet(0x000004bd, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDisplayInfoWindow(rhs bool) {
	_ = this.PropPut(0x000004bd, []interface{}{rhs})
}

func (this *Application_) Wait(time interface{}) bool {
	retVal, _ := this.Call(0x000006ea, []interface{}{time})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) ExtendList() bool {
	retVal, _ := this.PropGet(0x00000701, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetExtendList(rhs bool) {
	_ = this.PropPut(0x00000701, []interface{}{rhs})
}

func (this *Application_) OLEDBErrors() *OLEDBErrors {
	retVal, _ := this.PropGet(0x00000702, nil)
	return NewOLEDBErrors(retVal.IDispatch(), false, true)
}

var Application__GetPhonetic_OptArgs = []string{
	"Text",
}

func (this *Application_) GetPhonetic(optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(Application__GetPhonetic_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000703, nil, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) COMAddIns() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000704, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) DefaultWebOptions() *DefaultWebOptions {
	retVal, _ := this.PropGet(0x00000705, nil)
	return NewDefaultWebOptions(retVal.IDispatch(), false, true)
}

func (this *Application_) ProductCode() string {
	retVal, _ := this.PropGet(0x00000706, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) UserLibraryPath() string {
	retVal, _ := this.PropGet(0x00000707, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) AutoPercentEntry() bool {
	retVal, _ := this.PropGet(0x00000708, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetAutoPercentEntry(rhs bool) {
	_ = this.PropPut(0x00000708, []interface{}{rhs})
}

func (this *Application_) LanguageSettings() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000709, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) Dummy101() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000070a, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) Dummy12(p1 *PivotTable, p2 *PivotTable) {
	retVal, _ := this.Call(0x0000070b, []interface{}{p1, p2})
	_ = retVal
}

func (this *Application_) AnswerWizard() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000070c, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) CalculateFull() {
	retVal, _ := this.Call(0x0000070d, nil)
	_ = retVal
}

func (this *Application_) FindFile() bool {
	retVal, _ := this.Call(0x000006eb, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) CalculationVersion() int32 {
	retVal, _ := this.PropGet(0x0000070e, nil)
	return retVal.LValVal()
}

func (this *Application_) ShowWindowsInTaskbar() bool {
	retVal, _ := this.PropGet(0x0000070f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetShowWindowsInTaskbar(rhs bool) {
	_ = this.PropPut(0x0000070f, []interface{}{rhs})
}

func (this *Application_) FeatureInstall() int32 {
	retVal, _ := this.PropGet(0x00000710, nil)
	return retVal.LValVal()
}

func (this *Application_) SetFeatureInstall(rhs int32) {
	_ = this.PropPut(0x00000710, []interface{}{rhs})
}

func (this *Application_) Ready() bool {
	retVal, _ := this.PropGet(0x0000078c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var Application__Dummy13_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *Application_) Dummy13(arg1 interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Application__Dummy13_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000078d, []interface{}{arg1}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) FindFormat() *CellFormat {
	retVal, _ := this.PropGet(0x0000078e, nil)
	return NewCellFormat(retVal.IDispatch(), false, true)
}

func (this *Application_) SetFindFormat(rhs *CellFormat) {
	_ = this.PropPutRef(0x0000078e, []interface{}{rhs})
}

func (this *Application_) ReplaceFormat() *CellFormat {
	retVal, _ := this.PropGet(0x0000078f, nil)
	return NewCellFormat(retVal.IDispatch(), false, true)
}

func (this *Application_) SetReplaceFormat(rhs *CellFormat) {
	_ = this.PropPutRef(0x0000078f, []interface{}{rhs})
}

func (this *Application_) UsedObjects() *UsedObjects {
	retVal, _ := this.PropGet(0x00000790, nil)
	return NewUsedObjects(retVal.IDispatch(), false, true)
}

func (this *Application_) CalculationState() int32 {
	retVal, _ := this.PropGet(0x00000791, nil)
	return retVal.LValVal()
}

func (this *Application_) CalculationInterruptKey() int32 {
	retVal, _ := this.PropGet(0x00000792, nil)
	return retVal.LValVal()
}

func (this *Application_) SetCalculationInterruptKey(rhs int32) {
	_ = this.PropPut(0x00000792, []interface{}{rhs})
}

func (this *Application_) Watches() *Watches {
	retVal, _ := this.PropGet(0x00000793, nil)
	return NewWatches(retVal.IDispatch(), false, true)
}

func (this *Application_) DisplayFunctionToolTips() bool {
	retVal, _ := this.PropGet(0x00000794, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDisplayFunctionToolTips(rhs bool) {
	_ = this.PropPut(0x00000794, []interface{}{rhs})
}

func (this *Application_) AutomationSecurity() int32 {
	retVal, _ := this.PropGet(0x00000795, nil)
	return retVal.LValVal()
}

func (this *Application_) SetAutomationSecurity(rhs int32) {
	_ = this.PropPut(0x00000795, []interface{}{rhs})
}

func (this *Application_) FileDialog(fileDialogType int32) *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000796, []interface{}{fileDialogType})
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) Dummy14() {
	retVal, _ := this.Call(0x00000798, nil)
	_ = retVal
}

func (this *Application_) CalculateFullRebuild() {
	retVal, _ := this.Call(0x00000799, nil)
	_ = retVal
}

func (this *Application_) DisplayPasteOptions() bool {
	retVal, _ := this.PropGet(0x0000079a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDisplayPasteOptions(rhs bool) {
	_ = this.PropPut(0x0000079a, []interface{}{rhs})
}

func (this *Application_) DisplayInsertOptions() bool {
	retVal, _ := this.PropGet(0x0000079b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDisplayInsertOptions(rhs bool) {
	_ = this.PropPut(0x0000079b, []interface{}{rhs})
}

func (this *Application_) GenerateGetPivotData() bool {
	retVal, _ := this.PropGet(0x0000079c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetGenerateGetPivotData(rhs bool) {
	_ = this.PropPut(0x0000079c, []interface{}{rhs})
}

func (this *Application_) AutoRecover() *AutoRecover {
	retVal, _ := this.PropGet(0x0000079d, nil)
	return NewAutoRecover(retVal.IDispatch(), false, true)
}

func (this *Application_) Hwnd() int32 {
	retVal, _ := this.PropGet(0x0000079e, nil)
	return retVal.LValVal()
}

func (this *Application_) Hinstance() int32 {
	retVal, _ := this.PropGet(0x0000079f, nil)
	return retVal.LValVal()
}

var Application__CheckAbort_OptArgs = []string{
	"KeepAbort",
}

func (this *Application_) CheckAbort(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Application__CheckAbort_OptArgs, optArgs)
	retVal, _ := this.Call(0x000007a0, nil, optArgs...)
	_ = retVal
}

func (this *Application_) ErrorCheckingOptions() *ErrorCheckingOptions {
	retVal, _ := this.PropGet(0x000007a2, nil)
	return NewErrorCheckingOptions(retVal.IDispatch(), false, true)
}

func (this *Application_) AutoFormatAsYouTypeReplaceHyperlinks() bool {
	retVal, _ := this.PropGet(0x000007a3, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetAutoFormatAsYouTypeReplaceHyperlinks(rhs bool) {
	_ = this.PropPut(0x000007a3, []interface{}{rhs})
}

func (this *Application_) SmartTagRecognizers() *SmartTagRecognizers {
	retVal, _ := this.PropGet(0x000007a4, nil)
	return NewSmartTagRecognizers(retVal.IDispatch(), false, true)
}

func (this *Application_) NewWorkbook() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000061d, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) SpellingOptions() *SpellingOptions {
	retVal, _ := this.PropGet(0x000007a5, nil)
	return NewSpellingOptions(retVal.IDispatch(), false, true)
}

func (this *Application_) Speech() *Speech {
	retVal, _ := this.PropGet(0x000007a6, nil)
	return NewSpeech(retVal.IDispatch(), false, true)
}

func (this *Application_) MapPaperSize() bool {
	retVal, _ := this.PropGet(0x000007a7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetMapPaperSize(rhs bool) {
	_ = this.PropPut(0x000007a7, []interface{}{rhs})
}

func (this *Application_) ShowStartupDialog() bool {
	retVal, _ := this.PropGet(0x000007a8, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetShowStartupDialog(rhs bool) {
	_ = this.PropPut(0x000007a8, []interface{}{rhs})
}

func (this *Application_) DecimalSeparator() string {
	retVal, _ := this.PropGet(0x00000711, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetDecimalSeparator(rhs string) {
	_ = this.PropPut(0x00000711, []interface{}{rhs})
}

func (this *Application_) ThousandsSeparator() string {
	retVal, _ := this.PropGet(0x00000712, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetThousandsSeparator(rhs string) {
	_ = this.PropPut(0x00000712, []interface{}{rhs})
}

func (this *Application_) UseSystemSeparators() bool {
	retVal, _ := this.PropGet(0x000007a9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetUseSystemSeparators(rhs bool) {
	_ = this.PropPut(0x000007a9, []interface{}{rhs})
}

func (this *Application_) ThisCell() *Range {
	retVal, _ := this.PropGet(0x000007aa, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Application_) RTD() *RTD {
	retVal, _ := this.PropGet(0x000007ab, nil)
	return NewRTD(retVal.IDispatch(), false, true)
}

func (this *Application_) DisplayDocumentActionTaskPane() bool {
	retVal, _ := this.PropGet(0x000008cb, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDisplayDocumentActionTaskPane(rhs bool) {
	_ = this.PropPut(0x000008cb, []interface{}{rhs})
}

var Application__DisplayXMLSourcePane_OptArgs = []string{
	"XmlMap",
}

func (this *Application_) DisplayXMLSourcePane(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Application__DisplayXMLSourcePane_OptArgs, optArgs)
	retVal, _ := this.Call(0x000008cc, nil, optArgs...)
	_ = retVal
}

func (this *Application_) ArbitraryXMLSupportAvailable() bool {
	retVal, _ := this.PropGet(0x000008ce, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var Application__Support_OptArgs = []string{
	"arg",
}

func (this *Application_) Support(object *win32.IUnknown, id int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Application__Support_OptArgs, optArgs)
	retVal, _ := this.Call(0x000008cf, []interface{}{object, id}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) Dummy20(grfCompareFunctions int32) ole.Variant {
	retVal, _ := this.Call(0x00000945, []interface{}{grfCompareFunctions})
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) MeasurementUnit() int32 {
	retVal, _ := this.PropGet(0x00000947, nil)
	return retVal.LValVal()
}

func (this *Application_) SetMeasurementUnit(rhs int32) {
	_ = this.PropPut(0x00000947, []interface{}{rhs})
}

func (this *Application_) ShowSelectionFloaties() bool {
	retVal, _ := this.PropGet(0x00000948, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetShowSelectionFloaties(rhs bool) {
	_ = this.PropPut(0x00000948, []interface{}{rhs})
}

func (this *Application_) ShowMenuFloaties() bool {
	retVal, _ := this.PropGet(0x00000949, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetShowMenuFloaties(rhs bool) {
	_ = this.PropPut(0x00000949, []interface{}{rhs})
}

func (this *Application_) ShowDevTools() bool {
	retVal, _ := this.PropGet(0x0000094a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetShowDevTools(rhs bool) {
	_ = this.PropPut(0x0000094a, []interface{}{rhs})
}

func (this *Application_) EnableLivePreview() bool {
	retVal, _ := this.PropGet(0x0000094b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetEnableLivePreview(rhs bool) {
	_ = this.PropPut(0x0000094b, []interface{}{rhs})
}

func (this *Application_) DisplayDocumentInformationPanel() bool {
	retVal, _ := this.PropGet(0x0000094c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDisplayDocumentInformationPanel(rhs bool) {
	_ = this.PropPut(0x0000094c, []interface{}{rhs})
}

func (this *Application_) AlwaysUseClearType() bool {
	retVal, _ := this.PropGet(0x0000094d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetAlwaysUseClearType(rhs bool) {
	_ = this.PropPut(0x0000094d, []interface{}{rhs})
}

func (this *Application_) WarnOnFunctionNameConflict() bool {
	retVal, _ := this.PropGet(0x0000094e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetWarnOnFunctionNameConflict(rhs bool) {
	_ = this.PropPut(0x0000094e, []interface{}{rhs})
}

func (this *Application_) FormulaBarHeight() int32 {
	retVal, _ := this.PropGet(0x0000094f, nil)
	return retVal.LValVal()
}

func (this *Application_) SetFormulaBarHeight(rhs int32) {
	_ = this.PropPut(0x0000094f, []interface{}{rhs})
}

func (this *Application_) DisplayFormulaAutoComplete() bool {
	retVal, _ := this.PropGet(0x00000950, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDisplayFormulaAutoComplete(rhs bool) {
	_ = this.PropPut(0x00000950, []interface{}{rhs})
}

func (this *Application_) GenerateTableRefs() int32 {
	retVal, _ := this.PropGet(0x00000951, nil)
	return retVal.LValVal()
}

func (this *Application_) SetGenerateTableRefs(rhs int32) {
	_ = this.PropPut(0x00000951, []interface{}{rhs})
}

func (this *Application_) Assistance() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000952, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) CalculateUntilAsyncQueriesDone() {
	retVal, _ := this.Call(0x00000953, nil)
	_ = retVal
}

func (this *Application_) EnableLargeOperationAlert() bool {
	retVal, _ := this.PropGet(0x00000954, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetEnableLargeOperationAlert(rhs bool) {
	_ = this.PropPut(0x00000954, []interface{}{rhs})
}

func (this *Application_) LargeOperationCellThousandCount() int32 {
	retVal, _ := this.PropGet(0x00000955, nil)
	return retVal.LValVal()
}

func (this *Application_) SetLargeOperationCellThousandCount(rhs int32) {
	_ = this.PropPut(0x00000955, []interface{}{rhs})
}

func (this *Application_) DeferAsyncQueries() bool {
	retVal, _ := this.PropGet(0x00000956, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDeferAsyncQueries(rhs bool) {
	_ = this.PropPut(0x00000956, []interface{}{rhs})
}

func (this *Application_) MultiThreadedCalculation() *MultiThreadedCalculation {
	retVal, _ := this.PropGet(0x00000957, nil)
	return NewMultiThreadedCalculation(retVal.IDispatch(), false, true)
}

func (this *Application_) SharePointVersion(bstrUrl string) int32 {
	retVal, _ := this.Call(0x00000958, []interface{}{bstrUrl})
	return retVal.LValVal()
}

func (this *Application_) ActiveEncryptionSession() int32 {
	retVal, _ := this.PropGet(0x0000095a, nil)
	return retVal.LValVal()
}

func (this *Application_) HighQualityModeForGraphics() bool {
	retVal, _ := this.PropGet(0x0000095b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetHighQualityModeForGraphics(rhs bool) {
	_ = this.PropPut(0x0000095b, []interface{}{rhs})
}

func (this *Application_) FileExportConverters() *FileExportConverters {
	retVal, _ := this.PropGet(0x00000ad0, nil)
	return NewFileExportConverters(retVal.IDispatch(), false, true)
}

func (this *Application_) SmartArtLayouts() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000ad4, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) SmartArtQuickStyles() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000ad5, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) SmartArtColors() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000ad6, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) AddIns2() *AddIns2 {
	retVal, _ := this.PropGet(0x00000ad7, nil)
	return NewAddIns2(retVal.IDispatch(), false, true)
}

func (this *Application_) PrintCommunication() bool {
	retVal, _ := this.PropGet(0x00000ad8, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetPrintCommunication(rhs bool) {
	_ = this.PropPut(0x00000ad8, []interface{}{rhs})
}

var Application__MacroOptions_OptArgs = []string{
	"Macro", "Description", "HasMenu", "MenuText",
	"HasShortcutKey", "ShortcutKey", "Category", "StatusBar",
	"HelpContextID", "HelpFile", "ArgumentDescriptions",
}

func (this *Application_) MacroOptions(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Application__MacroOptions_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000ad2, nil, optArgs...)
	_ = retVal
}

func (this *Application_) UseClusterConnector() bool {
	retVal, _ := this.PropGet(0x00000ada, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetUseClusterConnector(rhs bool) {
	_ = this.PropPut(0x00000ada, []interface{}{rhs})
}

func (this *Application_) ClusterConnector() string {
	retVal, _ := this.PropGet(0x00000adb, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetClusterConnector(rhs string) {
	_ = this.PropPut(0x00000adb, []interface{}{rhs})
}

func (this *Application_) Quitting() bool {
	retVal, _ := this.PropGet(0x00000adc, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) Dummy22() bool {
	retVal, _ := this.PropGet(0x00000add, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDummy22(rhs bool) {
	_ = this.PropPut(0x00000add, []interface{}{rhs})
}

func (this *Application_) Dummy23() bool {
	retVal, _ := this.PropGet(0x00000ade, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDummy23(rhs bool) {
	_ = this.PropPut(0x00000ade, []interface{}{rhs})
}

func (this *Application_) ProtectedViewWindows() *ProtectedViewWindows {
	retVal, _ := this.PropGet(0x00000adf, nil)
	return NewProtectedViewWindows(retVal.IDispatch(), false, true)
}

func (this *Application_) ActiveProtectedViewWindow() *ProtectedViewWindow {
	retVal, _ := this.PropGet(0x00000ae0, nil)
	return NewProtectedViewWindow(retVal.IDispatch(), false, true)
}

func (this *Application_) IsSandboxed() bool {
	retVal, _ := this.PropGet(0x00000ae1, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SaveISO8601Dates() bool {
	retVal, _ := this.PropGet(0x00000ae2, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetSaveISO8601Dates(rhs bool) {
	_ = this.PropPut(0x00000ae2, []interface{}{rhs})
}

func (this *Application_) HinstancePtr() ole.Variant {
	retVal, _ := this.PropGet(0x00000ae3, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) FileValidation() int32 {
	retVal, _ := this.PropGet(0x00000ae4, nil)
	return retVal.LValVal()
}

func (this *Application_) SetFileValidation(rhs int32) {
	_ = this.PropPut(0x00000ae4, []interface{}{rhs})
}

func (this *Application_) FileValidationPivot() int32 {
	retVal, _ := this.PropGet(0x00000ae5, nil)
	return retVal.LValVal()
}

func (this *Application_) SetFileValidationPivot(rhs int32) {
	_ = this.PropPut(0x00000ae5, []interface{}{rhs})
}
