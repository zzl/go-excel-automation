package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000208D9-0000-0000-C000-000000000046
var IID_Global_ = syscall.GUID{0x000208D9, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Global_ struct {
	ole.OleClient
}

func NewGlobal_(pDisp *win32.IDispatch, addRef bool, scoped bool) *Global_ {
	p := &Global_{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func Global_FromVar(v ole.Variant) *Global_ {
	return NewGlobal_(v.PdispValVal(), false, false)
}

func (this *Global_) IID() *syscall.GUID {
	return &IID_Global_
}

func (this *Global_) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Global_) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Global_) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Global_) Parent() *Application {
	retVal := this.PropGet(0x00000096, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Global_) ActiveCell() *Range {
	retVal := this.PropGet(0x00000131, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Global_) ActiveChart() *Chart {
	retVal := this.PropGet(0x000000b7, nil)
	return NewChart(retVal.PdispValVal(), false, true)
}

func (this *Global_) ActiveDialog() *DialogSheet {
	retVal := this.PropGet(0x0000032f, nil)
	return NewDialogSheet(retVal.PdispValVal(), false, true)
}

func (this *Global_) ActiveMenuBar() *MenuBar {
	retVal := this.PropGet(0x000002f6, nil)
	return NewMenuBar(retVal.PdispValVal(), false, true)
}

func (this *Global_) ActivePrinter() string {
	retVal := this.PropGet(0x00000132, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Global_) SetActivePrinter(rhs string)  {
	retVal := this.PropPut(0x00000132, []interface{}{rhs})
	_= retVal
}

func (this *Global_) ActiveSheet() *ole.DispatchClass {
	retVal := this.PropGet(0x00000133, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Global_) ActiveWindow() *Window {
	retVal := this.PropGet(0x000002f7, nil)
	return NewWindow(retVal.PdispValVal(), false, true)
}

func (this *Global_) ActiveWorkbook() *Workbook {
	retVal := this.PropGet(0x00000134, nil)
	return NewWorkbook(retVal.PdispValVal(), false, true)
}

func (this *Global_) AddIns() *AddIns {
	retVal := this.PropGet(0x00000225, nil)
	return NewAddIns(retVal.PdispValVal(), false, true)
}

func (this *Global_) Assistant() *ole.DispatchClass {
	retVal := this.PropGet(0x0000059e, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Global_) Calculate()  {
	retVal := this.Call(0x00000117, nil)
	_= retVal
}

func (this *Global_) Cells() *Range {
	retVal := this.PropGet(0x000000ee, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Global_) Charts() *Sheets {
	retVal := this.PropGet(0x00000079, nil)
	return NewSheets(retVal.PdispValVal(), false, true)
}

func (this *Global_) Columns() *Range {
	retVal := this.PropGet(0x000000f1, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Global_) CommandBars() *ole.DispatchClass {
	retVal := this.PropGet(0x0000059f, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Global_) DDEAppReturnCode() int32 {
	retVal := this.PropGet(0x0000014c, nil)
	return retVal.LValVal()
}

func (this *Global_) DDEExecute(channel int32, string string)  {
	retVal := this.Call(0x0000014d, []interface{}{channel, string})
	_= retVal
}

func (this *Global_) DDEInitiate(app string, topic string) int32 {
	retVal := this.Call(0x0000014e, []interface{}{app, topic})
	return retVal.LValVal()
}

func (this *Global_) DDEPoke(channel int32, item interface{}, data interface{})  {
	retVal := this.Call(0x0000014f, []interface{}{channel, item, data})
	_= retVal
}

func (this *Global_) DDERequest(channel int32, item string) ole.Variant {
	retVal := this.Call(0x00000150, []interface{}{channel, item})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Global_) DDETerminate(channel int32)  {
	retVal := this.Call(0x00000151, []interface{}{channel})
	_= retVal
}

func (this *Global_) DialogSheets() *Sheets {
	retVal := this.PropGet(0x000002fc, nil)
	return NewSheets(retVal.PdispValVal(), false, true)
}

func (this *Global_) Evaluate(name interface{}) ole.Variant {
	retVal := this.Call(0x00000001, []interface{}{name})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Global_) Evaluate_(name interface{}) ole.Variant {
	retVal := this.Call(-5, []interface{}{name})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Global_) ExecuteExcel4Macro(string string) ole.Variant {
	retVal := this.Call(0x0000015e, []interface{}{string})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Global__Intersect_OptArgs= []string{
	"Arg3", "Arg4", "Arg5", "Arg6", 
	"Arg7", "Arg8", "Arg9", "Arg10", 
	"Arg11", "Arg12", "Arg13", "Arg14", 
	"Arg15", "Arg16", "Arg17", "Arg18", 
	"Arg19", "Arg20", "Arg21", "Arg22", 
	"Arg23", "Arg24", "Arg25", "Arg26", 
	"Arg27", "Arg28", "Arg29", "Arg30", 
}

func (this *Global_) Intersect(arg1 *Range, arg2 *Range, optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Global__Intersect_OptArgs, optArgs)
	retVal := this.Call(0x000002fe, []interface{}{arg1, arg2}, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Global_) MenuBars() *MenuBars {
	retVal := this.PropGet(0x0000024d, nil)
	return NewMenuBars(retVal.PdispValVal(), false, true)
}

func (this *Global_) Modules() *Modules {
	retVal := this.PropGet(0x00000246, nil)
	return NewModules(retVal.PdispValVal(), false, true)
}

func (this *Global_) Names() *Names {
	retVal := this.PropGet(0x000001ba, nil)
	return NewNames(retVal.PdispValVal(), false, true)
}

var Global__Range_OptArgs= []string{
	"Cell2", 
}

func (this *Global_) Range(cell1 interface{}, optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Global__Range_OptArgs, optArgs)
	retVal := this.PropGet(0x000000c5, []interface{}{cell1}, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Global_) Rows() *Range {
	retVal := this.PropGet(0x00000102, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

var Global__Run_OptArgs= []string{
	"Macro", "Arg1", "Arg2", "Arg3", 
	"Arg4", "Arg5", "Arg6", "Arg7", 
	"Arg8", "Arg9", "Arg10", "Arg11", 
	"Arg12", "Arg13", "Arg14", "Arg15", 
	"Arg16", "Arg17", "Arg18", "Arg19", 
	"Arg20", "Arg21", "Arg22", "Arg23", 
	"Arg24", "Arg25", "Arg26", "Arg27", 
	"Arg28", "Arg29", "Arg30", 
}

func (this *Global_) Run(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Global__Run_OptArgs, optArgs)
	retVal := this.Call(0x00000103, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Global__Run2__OptArgs= []string{
	"Macro", "Arg1", "Arg2", "Arg3", 
	"Arg4", "Arg5", "Arg6", "Arg7", 
	"Arg8", "Arg9", "Arg10", "Arg11", 
	"Arg12", "Arg13", "Arg14", "Arg15", 
	"Arg16", "Arg17", "Arg18", "Arg19", 
	"Arg20", "Arg21", "Arg22", "Arg23", 
	"Arg24", "Arg25", "Arg26", "Arg27", 
	"Arg28", "Arg29", "Arg30", 
}

func (this *Global_) Run2_(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Global__Run2__OptArgs, optArgs)
	retVal := this.Call(0x00000326, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Global_) Selection() *ole.DispatchClass {
	retVal := this.PropGet(0x00000093, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Global__SendKeys_OptArgs= []string{
	"Wait", 
}

func (this *Global_) SendKeys(keys interface{}, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Global__SendKeys_OptArgs, optArgs)
	retVal := this.Call(0x0000017f, []interface{}{keys}, optArgs...)
	_= retVal
}

func (this *Global_) Sheets() *Sheets {
	retVal := this.PropGet(0x000001e5, nil)
	return NewSheets(retVal.PdispValVal(), false, true)
}

func (this *Global_) ShortcutMenus(index int32) *Menu {
	retVal := this.PropGet(0x00000308, []interface{}{index})
	return NewMenu(retVal.PdispValVal(), false, true)
}

func (this *Global_) ThisWorkbook() *Workbook {
	retVal := this.PropGet(0x0000030a, nil)
	return NewWorkbook(retVal.PdispValVal(), false, true)
}

func (this *Global_) Toolbars() *Toolbars {
	retVal := this.PropGet(0x00000228, nil)
	return NewToolbars(retVal.PdispValVal(), false, true)
}

var Global__Union_OptArgs= []string{
	"Arg3", "Arg4", "Arg5", "Arg6", 
	"Arg7", "Arg8", "Arg9", "Arg10", 
	"Arg11", "Arg12", "Arg13", "Arg14", 
	"Arg15", "Arg16", "Arg17", "Arg18", 
	"Arg19", "Arg20", "Arg21", "Arg22", 
	"Arg23", "Arg24", "Arg25", "Arg26", 
	"Arg27", "Arg28", "Arg29", "Arg30", 
}

func (this *Global_) Union(arg1 *Range, arg2 *Range, optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Global__Union_OptArgs, optArgs)
	retVal := this.Call(0x0000030b, []interface{}{arg1, arg2}, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Global_) Windows() *Windows {
	retVal := this.PropGet(0x000001ae, nil)
	return NewWindows(retVal.PdispValVal(), false, true)
}

func (this *Global_) Workbooks() *Workbooks {
	retVal := this.PropGet(0x0000023c, nil)
	return NewWorkbooks(retVal.PdispValVal(), false, true)
}

func (this *Global_) WorksheetFunction() *WorksheetFunction {
	retVal := this.PropGet(0x000005a0, nil)
	return NewWorksheetFunction(retVal.PdispValVal(), false, true)
}

func (this *Global_) Worksheets() *Sheets {
	retVal := this.PropGet(0x000001ee, nil)
	return NewSheets(retVal.PdispValVal(), false, true)
}

func (this *Global_) Excel4IntlMacroSheets() *Sheets {
	retVal := this.PropGet(0x00000245, nil)
	return NewSheets(retVal.PdispValVal(), false, true)
}

func (this *Global_) Excel4MacroSheets() *Sheets {
	retVal := this.PropGet(0x00000243, nil)
	return NewSheets(retVal.PdispValVal(), false, true)
}

