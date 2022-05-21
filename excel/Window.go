package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020893-0000-0000-C000-000000000046
var IID_Window = syscall.GUID{0x00020893, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Window struct {
	ole.OleClient
}

func NewWindow(pDisp *win32.IDispatch, addRef bool, scoped bool) *Window {
	p := &Window{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func WindowFromVar(v ole.Variant) *Window {
	return NewWindow(v.PdispValVal(), false, false)
}

func (this *Window) IID() *syscall.GUID {
	return &IID_Window
}

func (this *Window) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Window) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Window) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Window) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Window) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Window) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Window) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Window) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Window) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Window) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Window) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Window) Activate() ole.Variant {
	retVal := this.Call(0x00000130, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Window) ActivateNext() ole.Variant {
	retVal := this.Call(0x0000045b, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Window) ActivatePrevious() ole.Variant {
	retVal := this.Call(0x0000045c, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Window) ActiveCell() *Range {
	retVal := this.PropGet(0x00000131, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Window) ActiveChart() *Chart {
	retVal := this.PropGet(0x000000b7, nil)
	return NewChart(retVal.PdispValVal(), false, true)
}

func (this *Window) ActivePane() *Pane {
	retVal := this.PropGet(0x00000282, nil)
	return NewPane(retVal.PdispValVal(), false, true)
}

func (this *Window) ActiveSheet() *ole.DispatchClass {
	retVal := this.PropGet(0x00000133, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Window) Caption() ole.Variant {
	retVal := this.PropGet(0x0000008b, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Window) SetCaption(rhs interface{})  {
	retVal := this.PropPut(0x0000008b, []interface{}{rhs})
	_= retVal
}

var Window_Close_OptArgs= []string{
	"SaveChanges", "Filename", "RouteWorkbook", 
}

func (this *Window) Close(optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(Window_Close_OptArgs, optArgs)
	retVal := this.Call(0x00000115, nil, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) DisplayFormulas() bool {
	retVal := this.PropGet(0x00000284, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetDisplayFormulas(rhs bool)  {
	retVal := this.PropPut(0x00000284, []interface{}{rhs})
	_= retVal
}

func (this *Window) DisplayGridlines() bool {
	retVal := this.PropGet(0x00000285, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetDisplayGridlines(rhs bool)  {
	retVal := this.PropPut(0x00000285, []interface{}{rhs})
	_= retVal
}

func (this *Window) DisplayHeadings() bool {
	retVal := this.PropGet(0x00000286, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetDisplayHeadings(rhs bool)  {
	retVal := this.PropPut(0x00000286, []interface{}{rhs})
	_= retVal
}

func (this *Window) DisplayHorizontalScrollBar() bool {
	retVal := this.PropGet(0x00000399, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetDisplayHorizontalScrollBar(rhs bool)  {
	retVal := this.PropPut(0x00000399, []interface{}{rhs})
	_= retVal
}

func (this *Window) DisplayOutline() bool {
	retVal := this.PropGet(0x00000287, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetDisplayOutline(rhs bool)  {
	retVal := this.PropPut(0x00000287, []interface{}{rhs})
	_= retVal
}

func (this *Window) DisplayRightToLeft_() bool {
	retVal := this.PropGet(0x00000288, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetDisplayRightToLeft_(rhs bool)  {
	retVal := this.PropPut(0x00000288, []interface{}{rhs})
	_= retVal
}

func (this *Window) DisplayVerticalScrollBar() bool {
	retVal := this.PropGet(0x0000039a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetDisplayVerticalScrollBar(rhs bool)  {
	retVal := this.PropPut(0x0000039a, []interface{}{rhs})
	_= retVal
}

func (this *Window) DisplayWorkbookTabs() bool {
	retVal := this.PropGet(0x0000039b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetDisplayWorkbookTabs(rhs bool)  {
	retVal := this.PropPut(0x0000039b, []interface{}{rhs})
	_= retVal
}

func (this *Window) DisplayZeros() bool {
	retVal := this.PropGet(0x00000289, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetDisplayZeros(rhs bool)  {
	retVal := this.PropPut(0x00000289, []interface{}{rhs})
	_= retVal
}

func (this *Window) EnableResize() bool {
	retVal := this.PropGet(0x000004a8, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetEnableResize(rhs bool)  {
	retVal := this.PropPut(0x000004a8, []interface{}{rhs})
	_= retVal
}

func (this *Window) FreezePanes() bool {
	retVal := this.PropGet(0x0000028a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetFreezePanes(rhs bool)  {
	retVal := this.PropPut(0x0000028a, []interface{}{rhs})
	_= retVal
}

func (this *Window) GridlineColor() int32 {
	retVal := this.PropGet(0x0000028b, nil)
	return retVal.LValVal()
}

func (this *Window) SetGridlineColor(rhs int32)  {
	retVal := this.PropPut(0x0000028b, []interface{}{rhs})
	_= retVal
}

func (this *Window) GridlineColorIndex() int32 {
	retVal := this.PropGet(0x0000028c, nil)
	return retVal.LValVal()
}

func (this *Window) SetGridlineColorIndex(rhs int32)  {
	retVal := this.PropPut(0x0000028c, []interface{}{rhs})
	_= retVal
}

func (this *Window) Height() float64 {
	retVal := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *Window) SetHeight(rhs float64)  {
	retVal := this.PropPut(0x0000007b, []interface{}{rhs})
	_= retVal
}

func (this *Window) Index() int32 {
	retVal := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

var Window_LargeScroll_OptArgs= []string{
	"Down", "Up", "ToRight", "ToLeft", 
}

func (this *Window) LargeScroll(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Window_LargeScroll_OptArgs, optArgs)
	retVal := this.Call(0x00000223, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Window) Left() float64 {
	retVal := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *Window) SetLeft(rhs float64)  {
	retVal := this.PropPut(0x0000007f, []interface{}{rhs})
	_= retVal
}

func (this *Window) NewWindow() *Window {
	retVal := this.Call(0x00000118, nil)
	return NewWindow(retVal.PdispValVal(), false, true)
}

func (this *Window) OnWindow() string {
	retVal := this.PropGet(0x0000026f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Window) SetOnWindow(rhs string)  {
	retVal := this.PropPut(0x0000026f, []interface{}{rhs})
	_= retVal
}

func (this *Window) Panes() *Panes {
	retVal := this.PropGet(0x0000028d, nil)
	return NewPanes(retVal.PdispValVal(), false, true)
}

var Window_PrintOut__OptArgs= []string{
	"From", "To", "Copies", "Preview", 
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName", 
}

func (this *Window) PrintOut_(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Window_PrintOut__OptArgs, optArgs)
	retVal := this.Call(0x000006ec, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Window_PrintPreview_OptArgs= []string{
	"EnableChanges", 
}

func (this *Window) PrintPreview(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Window_PrintPreview_OptArgs, optArgs)
	retVal := this.Call(0x00000119, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Window) RangeSelection() *Range {
	retVal := this.PropGet(0x000004a5, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Window) ScrollColumn() int32 {
	retVal := this.PropGet(0x0000028e, nil)
	return retVal.LValVal()
}

func (this *Window) SetScrollColumn(rhs int32)  {
	retVal := this.PropPut(0x0000028e, []interface{}{rhs})
	_= retVal
}

func (this *Window) ScrollRow() int32 {
	retVal := this.PropGet(0x0000028f, nil)
	return retVal.LValVal()
}

func (this *Window) SetScrollRow(rhs int32)  {
	retVal := this.PropPut(0x0000028f, []interface{}{rhs})
	_= retVal
}

var Window_ScrollWorkbookTabs_OptArgs= []string{
	"Sheets", "Position", 
}

func (this *Window) ScrollWorkbookTabs(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Window_ScrollWorkbookTabs_OptArgs, optArgs)
	retVal := this.Call(0x00000296, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Window) SelectedSheets() *Sheets {
	retVal := this.PropGet(0x00000290, nil)
	return NewSheets(retVal.PdispValVal(), false, true)
}

func (this *Window) Selection() *ole.DispatchClass {
	retVal := this.PropGet(0x00000093, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Window_SmallScroll_OptArgs= []string{
	"Down", "Up", "ToRight", "ToLeft", 
}

func (this *Window) SmallScroll(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Window_SmallScroll_OptArgs, optArgs)
	retVal := this.Call(0x00000224, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Window) Split() bool {
	retVal := this.PropGet(0x00000291, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetSplit(rhs bool)  {
	retVal := this.PropPut(0x00000291, []interface{}{rhs})
	_= retVal
}

func (this *Window) SplitColumn() int32 {
	retVal := this.PropGet(0x00000292, nil)
	return retVal.LValVal()
}

func (this *Window) SetSplitColumn(rhs int32)  {
	retVal := this.PropPut(0x00000292, []interface{}{rhs})
	_= retVal
}

func (this *Window) SplitHorizontal() float64 {
	retVal := this.PropGet(0x00000293, nil)
	return retVal.DblValVal()
}

func (this *Window) SetSplitHorizontal(rhs float64)  {
	retVal := this.PropPut(0x00000293, []interface{}{rhs})
	_= retVal
}

func (this *Window) SplitRow() int32 {
	retVal := this.PropGet(0x00000294, nil)
	return retVal.LValVal()
}

func (this *Window) SetSplitRow(rhs int32)  {
	retVal := this.PropPut(0x00000294, []interface{}{rhs})
	_= retVal
}

func (this *Window) SplitVertical() float64 {
	retVal := this.PropGet(0x00000295, nil)
	return retVal.DblValVal()
}

func (this *Window) SetSplitVertical(rhs float64)  {
	retVal := this.PropPut(0x00000295, []interface{}{rhs})
	_= retVal
}

func (this *Window) TabRatio() float64 {
	retVal := this.PropGet(0x000002a1, nil)
	return retVal.DblValVal()
}

func (this *Window) SetTabRatio(rhs float64)  {
	retVal := this.PropPut(0x000002a1, []interface{}{rhs})
	_= retVal
}

func (this *Window) Top() float64 {
	retVal := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *Window) SetTop(rhs float64)  {
	retVal := this.PropPut(0x0000007e, []interface{}{rhs})
	_= retVal
}

func (this *Window) Type() int32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *Window) UsableHeight() float64 {
	retVal := this.PropGet(0x00000185, nil)
	return retVal.DblValVal()
}

func (this *Window) UsableWidth() float64 {
	retVal := this.PropGet(0x00000186, nil)
	return retVal.DblValVal()
}

func (this *Window) Visible() bool {
	retVal := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetVisible(rhs bool)  {
	retVal := this.PropPut(0x0000022e, []interface{}{rhs})
	_= retVal
}

func (this *Window) VisibleRange() *Range {
	retVal := this.PropGet(0x0000045e, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Window) Width() float64 {
	retVal := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *Window) SetWidth(rhs float64)  {
	retVal := this.PropPut(0x0000007a, []interface{}{rhs})
	_= retVal
}

func (this *Window) WindowNumber() int32 {
	retVal := this.PropGet(0x0000045f, nil)
	return retVal.LValVal()
}

func (this *Window) WindowState() int32 {
	retVal := this.PropGet(0x0000018c, nil)
	return retVal.LValVal()
}

func (this *Window) SetWindowState(rhs int32)  {
	retVal := this.PropPut(0x0000018c, []interface{}{rhs})
	_= retVal
}

func (this *Window) Zoom() ole.Variant {
	retVal := this.PropGet(0x00000297, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Window) SetZoom(rhs interface{})  {
	retVal := this.PropPut(0x00000297, []interface{}{rhs})
	_= retVal
}

func (this *Window) View() int32 {
	retVal := this.PropGet(0x000004aa, nil)
	return retVal.LValVal()
}

func (this *Window) SetView(rhs int32)  {
	retVal := this.PropPut(0x000004aa, []interface{}{rhs})
	_= retVal
}

func (this *Window) DisplayRightToLeft() bool {
	retVal := this.PropGet(0x000006ee, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetDisplayRightToLeft(rhs bool)  {
	retVal := this.PropPut(0x000006ee, []interface{}{rhs})
	_= retVal
}

func (this *Window) PointsToScreenPixelsX(points int32) int32 {
	retVal := this.Call(0x000006f0, []interface{}{points})
	return retVal.LValVal()
}

func (this *Window) PointsToScreenPixelsY(points int32) int32 {
	retVal := this.Call(0x000006f1, []interface{}{points})
	return retVal.LValVal()
}

func (this *Window) RangeFromPoint(x int32, y int32) *ole.DispatchClass {
	retVal := this.Call(0x000006f2, []interface{}{x, y})
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Window_ScrollIntoView_OptArgs= []string{
	"Start", 
}

func (this *Window) ScrollIntoView(left int32, top int32, width int32, height int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Window_ScrollIntoView_OptArgs, optArgs)
	retVal := this.Call(0x000006f5, []interface{}{left, top, width, height}, optArgs...)
	_= retVal
}

func (this *Window) SheetViews() *SheetViews {
	retVal := this.PropGet(0x00000940, nil)
	return NewSheetViews(retVal.PdispValVal(), false, true)
}

func (this *Window) ActiveSheetView() *ole.DispatchClass {
	retVal := this.PropGet(0x00000941, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Window_PrintOut_OptArgs= []string{
	"From", "To", "Copies", "Preview", 
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName", 
}

func (this *Window) PrintOut(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Window_PrintOut_OptArgs, optArgs)
	retVal := this.Call(0x00000939, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Window) DisplayRuler() bool {
	retVal := this.PropGet(0x00000942, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetDisplayRuler(rhs bool)  {
	retVal := this.PropPut(0x00000942, []interface{}{rhs})
	_= retVal
}

func (this *Window) AutoFilterDateGrouping() bool {
	retVal := this.PropGet(0x00000943, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetAutoFilterDateGrouping(rhs bool)  {
	retVal := this.PropPut(0x00000943, []interface{}{rhs})
	_= retVal
}

func (this *Window) DisplayWhitespace() bool {
	retVal := this.PropGet(0x00000944, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetDisplayWhitespace(rhs bool)  {
	retVal := this.PropPut(0x00000944, []interface{}{rhs})
	_= retVal
}

