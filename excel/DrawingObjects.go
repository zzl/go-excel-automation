package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002086F-0000-0000-C000-000000000046
var IID_DrawingObjects = syscall.GUID{0x0002086F, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type DrawingObjects struct {
	ole.OleClient
}

func NewDrawingObjects(pDisp *win32.IDispatch, addRef bool, scoped bool) *DrawingObjects {
	p := &DrawingObjects{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DrawingObjectsFromVar(v ole.Variant) *DrawingObjects {
	return NewDrawingObjects(v.PdispValVal(), false, false)
}

func (this *DrawingObjects) IID() *syscall.GUID {
	return &IID_DrawingObjects
}

func (this *DrawingObjects) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DrawingObjects) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *DrawingObjects) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *DrawingObjects) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *DrawingObjects) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *DrawingObjects) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *DrawingObjects) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *DrawingObjects) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *DrawingObjects) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *DrawingObjects) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *DrawingObjects) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *DrawingObjects) Dummy3_()  {
	retVal := this.Call(0x00010003, nil)
	_= retVal
}

func (this *DrawingObjects) BringToFront() ole.Variant {
	retVal := this.Call(0x0000025a, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DrawingObjects) Copy() ole.Variant {
	retVal := this.Call(0x00000227, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DrawingObjects) CopyPicture(appearance int32, format int32) ole.Variant {
	retVal := this.Call(0x000000d5, []interface{}{appearance, format})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DrawingObjects) Cut() ole.Variant {
	retVal := this.Call(0x00000235, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DrawingObjects) Delete() ole.Variant {
	retVal := this.Call(0x00000075, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DrawingObjects) Duplicate() *ole.DispatchClass {
	retVal := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *DrawingObjects) Enabled() bool {
	retVal := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DrawingObjects) SetEnabled(rhs bool)  {
	retVal := this.PropPut(0x00000258, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) Height() float64 {
	retVal := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *DrawingObjects) SetHeight(rhs float64)  {
	retVal := this.PropPut(0x0000007b, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) Dummy12_()  {
	retVal := this.Call(0x0001000c, nil)
	_= retVal
}

func (this *DrawingObjects) Left() float64 {
	retVal := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *DrawingObjects) SetLeft(rhs float64)  {
	retVal := this.PropPut(0x0000007f, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) Locked() bool {
	retVal := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DrawingObjects) SetLocked(rhs bool)  {
	retVal := this.PropPut(0x0000010d, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) Dummy15_()  {
	retVal := this.Call(0x0001000f, nil)
	_= retVal
}

func (this *DrawingObjects) OnAction() string {
	retVal := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DrawingObjects) SetOnAction(rhs string)  {
	retVal := this.PropPut(0x00000254, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) Placement() ole.Variant {
	retVal := this.PropGet(0x00000269, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DrawingObjects) SetPlacement(rhs interface{})  {
	retVal := this.PropPut(0x00000269, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) PrintObject() bool {
	retVal := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DrawingObjects) SetPrintObject(rhs bool)  {
	retVal := this.PropPut(0x0000026a, []interface{}{rhs})
	_= retVal
}

var DrawingObjects_Select_OptArgs= []string{
	"Replace", 
}

func (this *DrawingObjects) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DrawingObjects_Select_OptArgs, optArgs)
	retVal := this.Call(0x000000eb, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DrawingObjects) SendToBack() ole.Variant {
	retVal := this.Call(0x0000025d, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DrawingObjects) Top() float64 {
	retVal := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *DrawingObjects) SetTop(rhs float64)  {
	retVal := this.PropPut(0x0000007e, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) Dummy22_()  {
	retVal := this.Call(0x00010016, nil)
	_= retVal
}

func (this *DrawingObjects) Visible() bool {
	retVal := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DrawingObjects) SetVisible(rhs bool)  {
	retVal := this.PropPut(0x0000022e, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) Width() float64 {
	retVal := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *DrawingObjects) SetWidth(rhs float64)  {
	retVal := this.PropPut(0x0000007a, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) ZOrder() int32 {
	retVal := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *DrawingObjects) ShapeRange() *ShapeRange {
	retVal := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.PdispValVal(), false, true)
}

func (this *DrawingObjects) Accelerator() ole.Variant {
	retVal := this.PropGet(0x0000034e, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DrawingObjects) SetAccelerator(rhs interface{})  {
	retVal := this.PropPut(0x0000034e, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) Dummy28_()  {
	retVal := this.Call(0x0001001c, nil)
	_= retVal
}

func (this *DrawingObjects) AddIndent() bool {
	retVal := this.PropGet(0x00000427, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DrawingObjects) SetAddIndent(rhs bool)  {
	retVal := this.PropPut(0x00000427, []interface{}{rhs})
	_= retVal
}

var DrawingObjects_AddItem_OptArgs= []string{
	"Index", 
}

func (this *DrawingObjects) AddItem(text interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DrawingObjects_AddItem_OptArgs, optArgs)
	retVal := this.Call(0x00000353, []interface{}{text}, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DrawingObjects) ArrowHeadLength() ole.Variant {
	retVal := this.PropGet(0x00000263, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DrawingObjects) SetArrowHeadLength(rhs interface{})  {
	retVal := this.PropPut(0x00000263, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) ArrowHeadStyle() ole.Variant {
	retVal := this.PropGet(0x00000264, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DrawingObjects) SetArrowHeadStyle(rhs interface{})  {
	retVal := this.PropPut(0x00000264, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) ArrowHeadWidth() ole.Variant {
	retVal := this.PropGet(0x00000265, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DrawingObjects) SetArrowHeadWidth(rhs interface{})  {
	retVal := this.PropPut(0x00000265, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) AutoSize() bool {
	retVal := this.PropGet(0x00000266, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DrawingObjects) SetAutoSize(rhs bool)  {
	retVal := this.PropPut(0x00000266, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) Border() *Border {
	retVal := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.PdispValVal(), false, true)
}

func (this *DrawingObjects) CancelButton() bool {
	retVal := this.PropGet(0x0000035a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DrawingObjects) SetCancelButton(rhs bool)  {
	retVal := this.PropPut(0x0000035a, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) Caption() string {
	retVal := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DrawingObjects) SetCaption(rhs string)  {
	retVal := this.PropPut(0x0000008b, []interface{}{rhs})
	_= retVal
}

var DrawingObjects_Characters_OptArgs= []string{
	"Start", "Length", 
}

func (this *DrawingObjects) Characters(optArgs ...interface{}) *Characters {
	optArgs = ole.ProcessOptArgs(DrawingObjects_Characters_OptArgs, optArgs)
	retVal := this.PropGet(0x0000025b, nil, optArgs...)
	return NewCharacters(retVal.PdispValVal(), false, true)
}

var DrawingObjects_CheckSpelling_OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang", 
}

func (this *DrawingObjects) CheckSpelling(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DrawingObjects_CheckSpelling_OptArgs, optArgs)
	retVal := this.Call(0x000001f9, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DrawingObjects) Default_() int32 {
	retVal := this.PropGet(0x00000000, nil)
	return retVal.LValVal()
}

func (this *DrawingObjects) SetDefault_(rhs int32)  {
	retVal := this.PropPut(0x00000000, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) DefaultButton() bool {
	retVal := this.PropGet(0x00000359, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DrawingObjects) SetDefaultButton(rhs bool)  {
	retVal := this.PropPut(0x00000359, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) DismissButton() bool {
	retVal := this.PropGet(0x0000035b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DrawingObjects) SetDismissButton(rhs bool)  {
	retVal := this.PropPut(0x0000035b, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) Display3DShading() bool {
	retVal := this.PropGet(0x00000462, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DrawingObjects) SetDisplay3DShading(rhs bool)  {
	retVal := this.PropPut(0x00000462, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) DisplayVerticalScrollBar() bool {
	retVal := this.PropGet(0x0000039a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DrawingObjects) SetDisplayVerticalScrollBar(rhs bool)  {
	retVal := this.PropPut(0x0000039a, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) DropDownLines() int32 {
	retVal := this.PropGet(0x00000350, nil)
	return retVal.LValVal()
}

func (this *DrawingObjects) SetDropDownLines(rhs int32)  {
	retVal := this.PropPut(0x00000350, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) Font() *Font {
	retVal := this.PropGet(0x00000092, nil)
	return NewFont(retVal.PdispValVal(), false, true)
}

func (this *DrawingObjects) Dummy47_()  {
	retVal := this.Call(0x0001002f, nil)
	_= retVal
}

func (this *DrawingObjects) HelpButton() bool {
	retVal := this.PropGet(0x0000035c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DrawingObjects) SetHelpButton(rhs bool)  {
	retVal := this.PropPut(0x0000035c, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) HorizontalAlignment() ole.Variant {
	retVal := this.PropGet(0x00000088, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DrawingObjects) SetHorizontalAlignment(rhs interface{})  {
	retVal := this.PropPut(0x00000088, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) InputType() int32 {
	retVal := this.PropGet(0x00000356, nil)
	return retVal.LValVal()
}

func (this *DrawingObjects) SetInputType(rhs int32)  {
	retVal := this.PropPut(0x00000356, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) Interior() *Interior {
	retVal := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.PdispValVal(), false, true)
}

func (this *DrawingObjects) LargeChange() int32 {
	retVal := this.PropGet(0x0000034d, nil)
	return retVal.LValVal()
}

func (this *DrawingObjects) SetLargeChange(rhs int32)  {
	retVal := this.PropPut(0x0000034d, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) LinkedCell() string {
	retVal := this.PropGet(0x00000422, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DrawingObjects) SetLinkedCell(rhs string)  {
	retVal := this.PropPut(0x00000422, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) Dummy54_()  {
	retVal := this.Call(0x00010036, nil)
	_= retVal
}

var DrawingObjects_List_OptArgs= []string{
	"Index", 
}

func (this *DrawingObjects) List(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DrawingObjects_List_OptArgs, optArgs)
	retVal := this.Call(0x0000035d, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DrawingObjects) Dummy56_()  {
	retVal := this.Call(0x00010038, nil)
	_= retVal
}

func (this *DrawingObjects) ListFillRange() string {
	retVal := this.PropGet(0x0000034f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DrawingObjects) SetListFillRange(rhs string)  {
	retVal := this.PropPut(0x0000034f, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) ListIndex() int32 {
	retVal := this.PropGet(0x00000352, nil)
	return retVal.LValVal()
}

func (this *DrawingObjects) SetListIndex(rhs int32)  {
	retVal := this.PropPut(0x00000352, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) LockedText() bool {
	retVal := this.PropGet(0x00000268, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DrawingObjects) SetLockedText(rhs bool)  {
	retVal := this.PropPut(0x00000268, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) Max() int32 {
	retVal := this.PropGet(0x0000034a, nil)
	return retVal.LValVal()
}

func (this *DrawingObjects) SetMax(rhs int32)  {
	retVal := this.PropPut(0x0000034a, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) Min() int32 {
	retVal := this.PropGet(0x0000034b, nil)
	return retVal.LValVal()
}

func (this *DrawingObjects) SetMin(rhs int32)  {
	retVal := this.PropPut(0x0000034b, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) MultiLine() bool {
	retVal := this.PropGet(0x00000357, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DrawingObjects) SetMultiLine(rhs bool)  {
	retVal := this.PropPut(0x00000357, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) MultiSelect() bool {
	retVal := this.PropGet(0x00000020, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DrawingObjects) SetMultiSelect(rhs bool)  {
	retVal := this.PropPut(0x00000020, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) Orientation() ole.Variant {
	retVal := this.PropGet(0x00000086, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DrawingObjects) SetOrientation(rhs interface{})  {
	retVal := this.PropPut(0x00000086, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) PhoneticAccelerator() ole.Variant {
	retVal := this.PropGet(0x00000461, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DrawingObjects) SetPhoneticAccelerator(rhs interface{})  {
	retVal := this.PropPut(0x00000461, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) RemoveAllItems() ole.Variant {
	retVal := this.Call(0x00000355, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var DrawingObjects_RemoveItem_OptArgs= []string{
	"Count", 
}

func (this *DrawingObjects) RemoveItem(index int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DrawingObjects_RemoveItem_OptArgs, optArgs)
	retVal := this.Call(0x00000354, []interface{}{index}, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var DrawingObjects_Reshape_OptArgs= []string{
	"Left", "Top", 
}

func (this *DrawingObjects) Reshape(vertex int32, insert interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DrawingObjects_Reshape_OptArgs, optArgs)
	retVal := this.Call(0x0000025c, []interface{}{vertex, insert}, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DrawingObjects) RoundedCorners() bool {
	retVal := this.PropGet(0x0000026b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DrawingObjects) SetRoundedCorners(rhs bool)  {
	retVal := this.PropPut(0x0000026b, []interface{}{rhs})
	_= retVal
}

var DrawingObjects_Selected_OptArgs= []string{
	"Index", 
}

func (this *DrawingObjects) Selected(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DrawingObjects_Selected_OptArgs, optArgs)
	retVal := this.Call(0x00000463, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DrawingObjects) Shadow() bool {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DrawingObjects) SetShadow(rhs bool)  {
	retVal := this.PropPut(0x00000067, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) SmallChange() int32 {
	retVal := this.PropGet(0x0000034c, nil)
	return retVal.LValVal()
}

func (this *DrawingObjects) SetSmallChange(rhs int32)  {
	retVal := this.PropPut(0x0000034c, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) Text() string {
	retVal := this.PropGet(0x0000008a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DrawingObjects) SetText(rhs string)  {
	retVal := this.PropPut(0x0000008a, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) Ungroup() *ole.DispatchClass {
	retVal := this.Call(0x000000f4, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *DrawingObjects) Value() int32 {
	retVal := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *DrawingObjects) SetValue(rhs int32)  {
	retVal := this.PropPut(0x00000006, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) VerticalAlignment() ole.Variant {
	retVal := this.PropGet(0x00000089, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DrawingObjects) SetVerticalAlignment(rhs interface{})  {
	retVal := this.PropPut(0x00000089, []interface{}{rhs})
	_= retVal
}

var DrawingObjects_Vertices_OptArgs= []string{
	"Index1", "Index2", 
}

func (this *DrawingObjects) Vertices(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DrawingObjects_Vertices_OptArgs, optArgs)
	retVal := this.Call(0x0000026d, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DrawingObjects) ReadingOrder() int32 {
	retVal := this.PropGet(0x000003cf, nil)
	return retVal.LValVal()
}

func (this *DrawingObjects) SetReadingOrder(rhs int32)  {
	retVal := this.PropPut(0x000003cf, []interface{}{rhs})
	_= retVal
}

func (this *DrawingObjects) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *DrawingObjects) Item(index interface{}) *ole.DispatchClass {
	retVal := this.Call(0x000000aa, []interface{}{index})
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *DrawingObjects) Group() *GroupObject {
	retVal := this.Call(0x0000002e, nil)
	return NewGroupObject(retVal.PdispValVal(), false, true)
}

var DrawingObjects_LinkCombo_OptArgs= []string{
	"Link", 
}

func (this *DrawingObjects) LinkCombo(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DrawingObjects_LinkCombo_OptArgs, optArgs)
	retVal := this.Call(0x00000358, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DrawingObjects) NewEnum_() *com.UnknownClass {
	retVal := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *DrawingObjects) ForEach(action func(item int32) bool) {
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
		pItem, _ := v.ToInt32()
		ret := action(pItem)
		if !ret {
			break
		}
	}
}

