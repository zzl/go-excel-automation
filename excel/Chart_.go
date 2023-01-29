package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
)

// 000208D6-0000-0000-C000-000000000046
var IID_Chart_ = syscall.GUID{0x000208D6, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Chart_ struct {
	ole.OleClient
}

func NewChart_(pDisp *win32.IDispatch, addRef bool, scoped bool) *Chart_ {
	if pDisp == nil {
		return nil
	}
	p := &Chart_{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func Chart_FromVar(v ole.Variant) *Chart_ {
	return NewChart_(v.IDispatch(), false, false)
}

func (this *Chart_) IID() *syscall.GUID {
	return &IID_Chart_
}

func (this *Chart_) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Chart_) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Chart_) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Chart_) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart_) Activate() {
	retVal, _ := this.Call(0x00000130, nil)
	_ = retVal
}

var Chart__Copy_OptArgs = []string{
	"Before", "After",
}

func (this *Chart_) Copy(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__Copy_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000227, nil, optArgs...)
	_ = retVal
}

func (this *Chart_) Delete() {
	retVal, _ := this.Call(0x00000075, nil)
	_ = retVal
}

func (this *Chart_) CodeName() string {
	retVal, _ := this.PropGet(0x0000055d, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Chart_) CodeName_() string {
	retVal, _ := this.PropGet(-2147418112, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Chart_) SetCodeName_(rhs string) {
	_ = this.PropPut(-2147418112, []interface{}{rhs})
}

func (this *Chart_) Index() int32 {
	retVal, _ := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

var Chart__Move_OptArgs = []string{
	"Before", "After",
}

func (this *Chart_) Move(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__Move_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000027d, nil, optArgs...)
	_ = retVal
}

func (this *Chart_) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Chart_) SetName(rhs string) {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *Chart_) Next() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000001f6, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart_) OnDoubleClick() string {
	retVal, _ := this.PropGet(0x00000274, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Chart_) SetOnDoubleClick(rhs string) {
	_ = this.PropPut(0x00000274, []interface{}{rhs})
}

func (this *Chart_) OnSheetActivate() string {
	retVal, _ := this.PropGet(0x00000407, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Chart_) SetOnSheetActivate(rhs string) {
	_ = this.PropPut(0x00000407, []interface{}{rhs})
}

func (this *Chart_) OnSheetDeactivate() string {
	retVal, _ := this.PropGet(0x00000439, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Chart_) SetOnSheetDeactivate(rhs string) {
	_ = this.PropPut(0x00000439, []interface{}{rhs})
}

func (this *Chart_) PageSetup() *PageSetup {
	retVal, _ := this.PropGet(0x000003e6, nil)
	return NewPageSetup(retVal.IDispatch(), false, true)
}

func (this *Chart_) Previous() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000001f7, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart__PrintOut___OptArgs = []string{
	"From", "To", "Copies", "Preview",
	"ActivePrinter", "PrintToFile", "Collate",
}

func (this *Chart_) PrintOut__(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__PrintOut___OptArgs, optArgs)
	retVal, _ := this.Call(0x00000389, nil, optArgs...)
	_ = retVal
}

var Chart__PrintPreview_OptArgs = []string{
	"EnableChanges",
}

func (this *Chart_) PrintPreview(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__PrintPreview_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000119, nil, optArgs...)
	_ = retVal
}

var Chart__Protect__OptArgs = []string{
	"Password", "DrawingObjects", "Contents", "Scenarios", "UserInterfaceOnly",
}

func (this *Chart_) Protect_(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__Protect__OptArgs, optArgs)
	retVal, _ := this.Call(0x0000011a, nil, optArgs...)
	_ = retVal
}

func (this *Chart_) ProtectContents() bool {
	retVal, _ := this.PropGet(0x00000124, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) ProtectDrawingObjects() bool {
	retVal, _ := this.PropGet(0x00000125, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) ProtectionMode() bool {
	retVal, _ := this.PropGet(0x00000487, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) Dummy23_() {
	retVal, _ := this.Call(0x00010017, nil)
	_ = retVal
}

var Chart__SaveAs__OptArgs = []string{
	"FileFormat", "Password", "WriteResPassword", "ReadOnlyRecommended",
	"CreateBackup", "AddToMru", "TextCodepage", "TextVisualLayout",
}

func (this *Chart_) SaveAs_(filename string, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__SaveAs__OptArgs, optArgs)
	retVal, _ := this.Call(0x0000011c, []interface{}{filename}, optArgs...)
	_ = retVal
}

var Chart__Select_OptArgs = []string{
	"Replace",
}

func (this *Chart_) Select(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000eb, nil, optArgs...)
	_ = retVal
}

var Chart__Unprotect_OptArgs = []string{
	"Password",
}

func (this *Chart_) Unprotect(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__Unprotect_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000011d, nil, optArgs...)
	_ = retVal
}

func (this *Chart_) Visible() int32 {
	retVal, _ := this.PropGet(0x0000022e, nil)
	return retVal.LValVal()
}

func (this *Chart_) SetVisible(rhs int32) {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

func (this *Chart_) Shapes() *Shapes {
	retVal, _ := this.PropGet(0x00000561, nil)
	return NewShapes(retVal.IDispatch(), false, true)
}

var Chart__ApplyDataLabels__OptArgs = []string{
	"Type", "LegendKey", "AutoText", "HasLeaderLines",
}

func (this *Chart_) ApplyDataLabels_(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__ApplyDataLabels__OptArgs, optArgs)
	retVal, _ := this.Call(0x00000097, nil, optArgs...)
	_ = retVal
}

var Chart__Arcs_OptArgs = []string{
	"Index",
}

func (this *Chart_) Arcs(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__Arcs_OptArgs, optArgs)
	retVal, _ := this.Call(0x000002f8, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart_) Area3DGroup() *ChartGroup {
	retVal, _ := this.PropGet(0x00000011, nil)
	return NewChartGroup(retVal.IDispatch(), false, true)
}

var Chart__AreaGroups_OptArgs = []string{
	"Index",
}

func (this *Chart_) AreaGroups(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__AreaGroups_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000009, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart__AutoFormat_OptArgs = []string{
	"Format",
}

func (this *Chart_) AutoFormat(gallery int32, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__AutoFormat_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000072, []interface{}{gallery}, optArgs...)
	_ = retVal
}

func (this *Chart_) AutoScaling() bool {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) SetAutoScaling(rhs bool) {
	_ = this.PropPut(0x0000006b, []interface{}{rhs})
}

var Chart__Axes_OptArgs = []string{
	"Type", "AxisGroup",
}

func (this *Chart_) Axes(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__Axes_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000017, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart_) SetBackgroundPicture(filename string) {
	retVal, _ := this.Call(0x000004a4, []interface{}{filename})
	_ = retVal
}

func (this *Chart_) Bar3DGroup() *ChartGroup {
	retVal, _ := this.PropGet(0x00000012, nil)
	return NewChartGroup(retVal.IDispatch(), false, true)
}

var Chart__BarGroups_OptArgs = []string{
	"Index",
}

func (this *Chart_) BarGroups(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__BarGroups_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000a, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart__Buttons_OptArgs = []string{
	"Index",
}

func (this *Chart_) Buttons(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__Buttons_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000022d, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart_) ChartArea() *ChartArea {
	retVal, _ := this.PropGet(0x00000050, nil)
	return NewChartArea(retVal.IDispatch(), false, true)
}

var Chart__ChartGroups_OptArgs = []string{
	"Index",
}

func (this *Chart_) ChartGroups(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__ChartGroups_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000008, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart__ChartObjects_OptArgs = []string{
	"Index",
}

func (this *Chart_) ChartObjects(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__ChartObjects_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000424, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart_) ChartTitle() *ChartTitle {
	retVal, _ := this.PropGet(0x00000051, nil)
	return NewChartTitle(retVal.IDispatch(), false, true)
}

var Chart__ChartWizard_OptArgs = []string{
	"Source", "Gallery", "Format", "PlotBy",
	"CategoryLabels", "SeriesLabels", "HasLegend", "Title",
	"CategoryTitle", "ValueTitle", "ExtraTitle",
}

func (this *Chart_) ChartWizard(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__ChartWizard_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000c4, nil, optArgs...)
	_ = retVal
}

var Chart__CheckBoxes_OptArgs = []string{
	"Index",
}

func (this *Chart_) CheckBoxes(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__CheckBoxes_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000338, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart__CheckSpelling_OptArgs = []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang",
}

func (this *Chart_) CheckSpelling(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__CheckSpelling_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001f9, nil, optArgs...)
	_ = retVal
}

func (this *Chart_) Column3DGroup() *ChartGroup {
	retVal, _ := this.PropGet(0x00000013, nil)
	return NewChartGroup(retVal.IDispatch(), false, true)
}

var Chart__ColumnGroups_OptArgs = []string{
	"Index",
}

func (this *Chart_) ColumnGroups(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__ColumnGroups_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000b, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart__CopyPicture_OptArgs = []string{
	"Appearance", "Format", "Size",
}

func (this *Chart_) CopyPicture(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__CopyPicture_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d5, nil, optArgs...)
	_ = retVal
}

func (this *Chart_) Corners() *Corners {
	retVal, _ := this.PropGet(0x0000004f, nil)
	return NewCorners(retVal.IDispatch(), false, true)
}

var Chart__CreatePublisher_OptArgs = []string{
	"Edition", "Appearance", "Size", "ContainsPICT",
	"ContainsBIFF", "ContainsRTF", "ContainsVALU",
}

func (this *Chart_) CreatePublisher(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__CreatePublisher_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001ca, nil, optArgs...)
	_ = retVal
}

func (this *Chart_) DataTable() *DataTable {
	retVal, _ := this.PropGet(0x00000573, nil)
	return NewDataTable(retVal.IDispatch(), false, true)
}

func (this *Chart_) DepthPercent() int32 {
	retVal, _ := this.PropGet(0x00000030, nil)
	return retVal.LValVal()
}

func (this *Chart_) SetDepthPercent(rhs int32) {
	_ = this.PropPut(0x00000030, []interface{}{rhs})
}

func (this *Chart_) Deselect() {
	retVal, _ := this.Call(0x00000460, nil)
	_ = retVal
}

func (this *Chart_) DisplayBlanksAs() int32 {
	retVal, _ := this.PropGet(0x0000005d, nil)
	return retVal.LValVal()
}

func (this *Chart_) SetDisplayBlanksAs(rhs int32) {
	_ = this.PropPut(0x0000005d, []interface{}{rhs})
}

var Chart__DoughnutGroups_OptArgs = []string{
	"Index",
}

func (this *Chart_) DoughnutGroups(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__DoughnutGroups_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000e, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart__Drawings_OptArgs = []string{
	"Index",
}

func (this *Chart_) Drawings(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__Drawings_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000304, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart__DrawingObjects_OptArgs = []string{
	"Index",
}

func (this *Chart_) DrawingObjects(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__DrawingObjects_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000058, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart__DropDowns_OptArgs = []string{
	"Index",
}

func (this *Chart_) DropDowns(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__DropDowns_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000344, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart_) Elevation() int32 {
	retVal, _ := this.PropGet(0x00000031, nil)
	return retVal.LValVal()
}

func (this *Chart_) SetElevation(rhs int32) {
	_ = this.PropPut(0x00000031, []interface{}{rhs})
}

func (this *Chart_) Evaluate(name interface{}) ole.Variant {
	retVal, _ := this.Call(0x00000001, []interface{}{name})
	com.AddToScope(retVal)
	return *retVal
}

func (this *Chart_) Evaluate_(name interface{}) ole.Variant {
	retVal, _ := this.Call(-5, []interface{}{name})
	com.AddToScope(retVal)
	return *retVal
}

func (this *Chart_) Floor() *Floor {
	retVal, _ := this.PropGet(0x00000053, nil)
	return NewFloor(retVal.IDispatch(), false, true)
}

func (this *Chart_) GapDepth() int32 {
	retVal, _ := this.PropGet(0x00000032, nil)
	return retVal.LValVal()
}

func (this *Chart_) SetGapDepth(rhs int32) {
	_ = this.PropPut(0x00000032, []interface{}{rhs})
}

var Chart__GroupBoxes_OptArgs = []string{
	"Index",
}

func (this *Chart_) GroupBoxes(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__GroupBoxes_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000342, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart__GroupObjects_OptArgs = []string{
	"Index",
}

func (this *Chart_) GroupObjects(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__GroupObjects_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000459, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart__HasAxis_OptArgs = []string{
	"Index1", "Index2",
}

func (this *Chart_) HasAxis(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Chart__HasAxis_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x00000034, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Chart__SetHasAxis_OptArgs = []string{
	"Index1", "Index2",
}

func (this *Chart_) SetHasAxis(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__SetHasAxis_OptArgs, optArgs)
	_ = this.PropPut(0x00000034, nil, optArgs...)
}

func (this *Chart_) HasDataTable() bool {
	retVal, _ := this.PropGet(0x00000574, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) SetHasDataTable(rhs bool) {
	_ = this.PropPut(0x00000574, []interface{}{rhs})
}

func (this *Chart_) HasLegend() bool {
	retVal, _ := this.PropGet(0x00000035, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) SetHasLegend(rhs bool) {
	_ = this.PropPut(0x00000035, []interface{}{rhs})
}

func (this *Chart_) HasTitle() bool {
	retVal, _ := this.PropGet(0x00000036, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) SetHasTitle(rhs bool) {
	_ = this.PropPut(0x00000036, []interface{}{rhs})
}

func (this *Chart_) HeightPercent() int32 {
	retVal, _ := this.PropGet(0x00000037, nil)
	return retVal.LValVal()
}

func (this *Chart_) SetHeightPercent(rhs int32) {
	_ = this.PropPut(0x00000037, []interface{}{rhs})
}

func (this *Chart_) Hyperlinks() *Hyperlinks {
	retVal, _ := this.PropGet(0x00000571, nil)
	return NewHyperlinks(retVal.IDispatch(), false, true)
}

var Chart__Labels_OptArgs = []string{
	"Index",
}

func (this *Chart_) Labels(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__Labels_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000349, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart_) Legend() *Legend {
	retVal, _ := this.PropGet(0x00000054, nil)
	return NewLegend(retVal.IDispatch(), false, true)
}

func (this *Chart_) Line3DGroup() *ChartGroup {
	retVal, _ := this.PropGet(0x00000014, nil)
	return NewChartGroup(retVal.IDispatch(), false, true)
}

var Chart__LineGroups_OptArgs = []string{
	"Index",
}

func (this *Chart_) LineGroups(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__LineGroups_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000c, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart__Lines_OptArgs = []string{
	"Index",
}

func (this *Chart_) Lines(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__Lines_OptArgs, optArgs)
	retVal, _ := this.Call(0x000002ff, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart__ListBoxes_OptArgs = []string{
	"Index",
}

func (this *Chart_) ListBoxes(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__ListBoxes_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000340, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart__Location_OptArgs = []string{
	"Name",
}

func (this *Chart_) Location(where int32, optArgs ...interface{}) *Chart {
	optArgs = ole.ProcessOptArgs(Chart__Location_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000575, []interface{}{where}, optArgs...)
	return NewChart(retVal.IDispatch(), false, true)
}

var Chart__OLEObjects_OptArgs = []string{
	"Index",
}

func (this *Chart_) OLEObjects(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__OLEObjects_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000031f, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart__OptionButtons_OptArgs = []string{
	"Index",
}

func (this *Chart_) OptionButtons(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__OptionButtons_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000033a, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart__Ovals_OptArgs = []string{
	"Index",
}

func (this *Chart_) Ovals(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__Ovals_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000321, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart__Paste_OptArgs = []string{
	"Type",
}

func (this *Chart_) Paste(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__Paste_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d3, nil, optArgs...)
	_ = retVal
}

func (this *Chart_) Perspective() int32 {
	retVal, _ := this.PropGet(0x00000039, nil)
	return retVal.LValVal()
}

func (this *Chart_) SetPerspective(rhs int32) {
	_ = this.PropPut(0x00000039, []interface{}{rhs})
}

var Chart__Pictures_OptArgs = []string{
	"Index",
}

func (this *Chart_) Pictures(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__Pictures_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000303, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart_) Pie3DGroup() *ChartGroup {
	retVal, _ := this.PropGet(0x00000015, nil)
	return NewChartGroup(retVal.IDispatch(), false, true)
}

var Chart__PieGroups_OptArgs = []string{
	"Index",
}

func (this *Chart_) PieGroups(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__PieGroups_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000d, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart_) PlotArea() *PlotArea {
	retVal, _ := this.PropGet(0x00000055, nil)
	return NewPlotArea(retVal.IDispatch(), false, true)
}

func (this *Chart_) PlotVisibleOnly() bool {
	retVal, _ := this.PropGet(0x0000005c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) SetPlotVisibleOnly(rhs bool) {
	_ = this.PropPut(0x0000005c, []interface{}{rhs})
}

var Chart__RadarGroups_OptArgs = []string{
	"Index",
}

func (this *Chart_) RadarGroups(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__RadarGroups_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000f, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart__Rectangles_OptArgs = []string{
	"Index",
}

func (this *Chart_) Rectangles(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__Rectangles_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000306, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart_) RightAngleAxes() ole.Variant {
	retVal, _ := this.PropGet(0x0000003a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Chart_) SetRightAngleAxes(rhs interface{}) {
	_ = this.PropPut(0x0000003a, []interface{}{rhs})
}

func (this *Chart_) Rotation() ole.Variant {
	retVal, _ := this.PropGet(0x0000003b, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Chart_) SetRotation(rhs interface{}) {
	_ = this.PropPut(0x0000003b, []interface{}{rhs})
}

var Chart__ScrollBars_OptArgs = []string{
	"Index",
}

func (this *Chart_) ScrollBars(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__ScrollBars_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000033e, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart__SeriesCollection_OptArgs = []string{
	"Index",
}

func (this *Chart_) SeriesCollection(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__SeriesCollection_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000044, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart_) SizeWithWindow() bool {
	retVal, _ := this.PropGet(0x0000005e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) SetSizeWithWindow(rhs bool) {
	_ = this.PropPut(0x0000005e, []interface{}{rhs})
}

func (this *Chart_) ShowWindow() bool {
	retVal, _ := this.PropGet(0x00000577, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) SetShowWindow(rhs bool) {
	_ = this.PropPut(0x00000577, []interface{}{rhs})
}

var Chart__Spinners_OptArgs = []string{
	"Index",
}

func (this *Chart_) Spinners(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__Spinners_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000346, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart_) SubType() int32 {
	retVal, _ := this.PropGet(0x0000006d, nil)
	return retVal.LValVal()
}

func (this *Chart_) SetSubType(rhs int32) {
	_ = this.PropPut(0x0000006d, []interface{}{rhs})
}

func (this *Chart_) SurfaceGroup() *ChartGroup {
	retVal, _ := this.PropGet(0x00000016, nil)
	return NewChartGroup(retVal.IDispatch(), false, true)
}

var Chart__TextBoxes_OptArgs = []string{
	"Index",
}

func (this *Chart_) TextBoxes(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__TextBoxes_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000309, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart_) Type() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *Chart_) SetType(rhs int32) {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

func (this *Chart_) ChartType() int32 {
	retVal, _ := this.PropGet(0x00000578, nil)
	return retVal.LValVal()
}

func (this *Chart_) SetChartType(rhs int32) {
	_ = this.PropPut(0x00000578, []interface{}{rhs})
}

var Chart__ApplyCustomType_OptArgs = []string{
	"TypeName",
}

func (this *Chart_) ApplyCustomType(chartType int32, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__ApplyCustomType_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000579, []interface{}{chartType}, optArgs...)
	_ = retVal
}

func (this *Chart_) Walls() *Walls {
	retVal, _ := this.PropGet(0x00000056, nil)
	return NewWalls(retVal.IDispatch(), false, true)
}

func (this *Chart_) WallsAndGridlines2D() bool {
	retVal, _ := this.PropGet(0x000000d2, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) SetWallsAndGridlines2D(rhs bool) {
	_ = this.PropPut(0x000000d2, []interface{}{rhs})
}

var Chart__XYGroups_OptArgs = []string{
	"Index",
}

func (this *Chart_) XYGroups(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart__XYGroups_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000010, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart_) BarShape() int32 {
	retVal, _ := this.PropGet(0x0000057b, nil)
	return retVal.LValVal()
}

func (this *Chart_) SetBarShape(rhs int32) {
	_ = this.PropPut(0x0000057b, []interface{}{rhs})
}

func (this *Chart_) PlotBy() int32 {
	retVal, _ := this.PropGet(0x000000ca, nil)
	return retVal.LValVal()
}

func (this *Chart_) SetPlotBy(rhs int32) {
	_ = this.PropPut(0x000000ca, []interface{}{rhs})
}

func (this *Chart_) CopyChartBuild() {
	retVal, _ := this.Call(0x0000057c, nil)
	_ = retVal
}

func (this *Chart_) ProtectFormatting() bool {
	retVal, _ := this.PropGet(0x0000057d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) SetProtectFormatting(rhs bool) {
	_ = this.PropPut(0x0000057d, []interface{}{rhs})
}

func (this *Chart_) ProtectData() bool {
	retVal, _ := this.PropGet(0x0000057e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) SetProtectData(rhs bool) {
	_ = this.PropPut(0x0000057e, []interface{}{rhs})
}

func (this *Chart_) ProtectGoalSeek() bool {
	retVal, _ := this.PropGet(0x0000057f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) SetProtectGoalSeek(rhs bool) {
	_ = this.PropPut(0x0000057f, []interface{}{rhs})
}

func (this *Chart_) ProtectSelection() bool {
	retVal, _ := this.PropGet(0x00000580, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) SetProtectSelection(rhs bool) {
	_ = this.PropPut(0x00000580, []interface{}{rhs})
}

func (this *Chart_) GetChartElement(x int32, y int32, elementID *int32, arg1 *int32, arg2 *int32) {
	retVal, _ := this.Call(0x00000581, []interface{}{x, y, elementID, arg1, arg2})
	_ = retVal
}

var Chart__SetSourceData_OptArgs = []string{
	"PlotBy",
}

func (this *Chart_) SetSourceData(source *Range, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__SetSourceData_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000585, []interface{}{source}, optArgs...)
	_ = retVal
}

var Chart__Export_OptArgs = []string{
	"FilterName", "Interactive",
}

func (this *Chart_) Export(filename string, optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(Chart__Export_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000586, []interface{}{filename}, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) Refresh() {
	retVal, _ := this.Call(0x00000589, nil)
	_ = retVal
}

func (this *Chart_) PivotLayout() *PivotLayout {
	retVal, _ := this.PropGet(0x00000716, nil)
	return NewPivotLayout(retVal.IDispatch(), false, true)
}

func (this *Chart_) HasPivotFields() bool {
	retVal, _ := this.PropGet(0x00000717, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) SetHasPivotFields(rhs bool) {
	_ = this.PropPut(0x00000717, []interface{}{rhs})
}

func (this *Chart_) Scripts() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000718, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart__PrintOut__OptArgs = []string{
	"From", "To", "Copies", "Preview",
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName",
}

func (this *Chart_) PrintOut_(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__PrintOut__OptArgs, optArgs)
	retVal, _ := this.Call(0x000006ec, nil, optArgs...)
	_ = retVal
}

func (this *Chart_) Tab() *Tab {
	retVal, _ := this.PropGet(0x00000411, nil)
	return NewTab(retVal.IDispatch(), false, true)
}

func (this *Chart_) MailEnvelope() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000007e5, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart__ApplyDataLabels_OptArgs = []string{
	"Type", "LegendKey", "AutoText", "HasLeaderLines",
	"ShowSeriesName", "ShowCategoryName", "ShowValue", "ShowPercentage",
	"ShowBubbleSize", "Separator",
}

func (this *Chart_) ApplyDataLabels(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__ApplyDataLabels_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000782, nil, optArgs...)
	_ = retVal
}

var Chart__SaveAs_OptArgs = []string{
	"FileFormat", "Password", "WriteResPassword", "ReadOnlyRecommended",
	"CreateBackup", "AddToMru", "TextCodepage", "TextVisualLayout", "Local",
}

func (this *Chart_) SaveAs(filename string, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__SaveAs_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000785, []interface{}{filename}, optArgs...)
	_ = retVal
}

var Chart__Protect_OptArgs = []string{
	"Password", "DrawingObjects", "Contents", "Scenarios", "UserInterfaceOnly",
}

func (this *Chart_) Protect(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__Protect_OptArgs, optArgs)
	retVal, _ := this.Call(0x000007ed, nil, optArgs...)
	_ = retVal
}

var Chart__ApplyLayout_OptArgs = []string{
	"ChartType",
}

func (this *Chart_) ApplyLayout(layout int32, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__ApplyLayout_OptArgs, optArgs)
	retVal, _ := this.Call(0x000009c4, []interface{}{layout}, optArgs...)
	_ = retVal
}

func (this *Chart_) SetElement(element int32) {
	retVal, _ := this.Call(0x000009c6, []interface{}{element})
	_ = retVal
}

func (this *Chart_) ShowDataLabelsOverMaximum() bool {
	retVal, _ := this.PropGet(0x000009c8, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) SetShowDataLabelsOverMaximum(rhs bool) {
	_ = this.PropPut(0x000009c8, []interface{}{rhs})
}

func (this *Chart_) SideWall() *Walls {
	retVal, _ := this.PropGet(0x000009c9, nil)
	return NewWalls(retVal.IDispatch(), false, true)
}

func (this *Chart_) BackWall() *Walls {
	retVal, _ := this.PropGet(0x000009ca, nil)
	return NewWalls(retVal.IDispatch(), false, true)
}

var Chart__PrintOut_OptArgs = []string{
	"From", "To", "Copies", "Preview",
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName",
}

func (this *Chart_) PrintOut(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__PrintOut_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000939, nil, optArgs...)
	_ = retVal
}

func (this *Chart_) ApplyChartTemplate(filename string) {
	retVal, _ := this.Call(0x000009cb, []interface{}{filename})
	_ = retVal
}

func (this *Chart_) SaveChartTemplate(filename string) {
	retVal, _ := this.Call(0x000009cc, []interface{}{filename})
	_ = retVal
}

func (this *Chart_) SetDefaultChart(name interface{}) {
	retVal, _ := this.Call(0x000000db, []interface{}{name})
	_ = retVal
}

var Chart__ExportAsFixedFormat_OptArgs = []string{
	"Filename", "Quality", "IncludeDocProperties", "IgnorePrintAreas",
	"From", "To", "OpenAfterPublish", "FixedFormatExtClassPtr",
}

func (this *Chart_) ExportAsFixedFormat(type_ int32, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Chart__ExportAsFixedFormat_OptArgs, optArgs)
	retVal, _ := this.Call(0x000009bd, []interface{}{type_}, optArgs...)
	_ = retVal
}

func (this *Chart_) ChartStyle() ole.Variant {
	retVal, _ := this.PropGet(0x000009cd, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Chart_) SetChartStyle(rhs interface{}) {
	_ = this.PropPut(0x000009cd, []interface{}{rhs})
}

func (this *Chart_) ClearToMatchStyle() {
	retVal, _ := this.Call(0x000009ce, nil)
	_ = retVal
}

func (this *Chart_) PrintedCommentPages() int32 {
	retVal, _ := this.PropGet(0x00000b29, nil)
	return retVal.LValVal()
}

func (this *Chart_) Dummy24() bool {
	retVal, _ := this.PropGet(0x00000b2a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) SetDummy24(rhs bool) {
	_ = this.PropPut(0x00000b2a, []interface{}{rhs})
}

func (this *Chart_) Dummy25() bool {
	retVal, _ := this.PropGet(0x00000b2b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) SetDummy25(rhs bool) {
	_ = this.PropPut(0x00000b2b, []interface{}{rhs})
}

func (this *Chart_) ShowReportFilterFieldButtons() bool {
	retVal, _ := this.PropGet(0x00000b2c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) SetShowReportFilterFieldButtons(rhs bool) {
	_ = this.PropPut(0x00000b2c, []interface{}{rhs})
}

func (this *Chart_) ShowLegendFieldButtons() bool {
	retVal, _ := this.PropGet(0x00000b2d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) SetShowLegendFieldButtons(rhs bool) {
	_ = this.PropPut(0x00000b2d, []interface{}{rhs})
}

func (this *Chart_) ShowAxisFieldButtons() bool {
	retVal, _ := this.PropGet(0x00000b2e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) SetShowAxisFieldButtons(rhs bool) {
	_ = this.PropPut(0x00000b2e, []interface{}{rhs})
}

func (this *Chart_) ShowValueFieldButtons() bool {
	retVal, _ := this.PropGet(0x00000b2f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) SetShowValueFieldButtons(rhs bool) {
	_ = this.PropPut(0x00000b2f, []interface{}{rhs})
}

func (this *Chart_) ShowAllFieldButtons() bool {
	retVal, _ := this.PropGet(0x00000b30, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart_) SetShowAllFieldButtons(rhs bool) {
	_ = this.PropPut(0x00000b30, []interface{}{rhs})
}
