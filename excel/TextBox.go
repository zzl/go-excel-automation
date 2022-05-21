package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000208A4-0000-0000-C000-000000000046
var IID_TextBox = syscall.GUID{0x000208A4, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TextBox struct {
	ole.OleClient
}

func NewTextBox(pDisp *win32.IDispatch, addRef bool, scoped bool) *TextBox {
	p := &TextBox{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TextBoxFromVar(v ole.Variant) *TextBox {
	return NewTextBox(v.PdispValVal(), false, false)
}

func (this *TextBox) IID() *syscall.GUID {
	return &IID_TextBox
}

func (this *TextBox) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TextBox) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *TextBox) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *TextBox) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *TextBox) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *TextBox) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *TextBox) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *TextBox) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *TextBox) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *TextBox) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *TextBox) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *TextBox) BottomRightCell() *Range {
	retVal := this.PropGet(0x00000267, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *TextBox) BringToFront() ole.Variant {
	retVal := this.Call(0x0000025a, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *TextBox) Copy() ole.Variant {
	retVal := this.Call(0x00000227, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *TextBox) CopyPicture(appearance int32, format int32) ole.Variant {
	retVal := this.Call(0x000000d5, []interface{}{appearance, format})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *TextBox) Cut() ole.Variant {
	retVal := this.Call(0x00000235, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *TextBox) Delete() ole.Variant {
	retVal := this.Call(0x00000075, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *TextBox) Duplicate() *ole.DispatchClass {
	retVal := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *TextBox) Enabled() bool {
	retVal := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TextBox) SetEnabled(rhs bool)  {
	retVal := this.PropPut(0x00000258, []interface{}{rhs})
	_= retVal
}

func (this *TextBox) Height() float64 {
	retVal := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *TextBox) SetHeight(rhs float64)  {
	retVal := this.PropPut(0x0000007b, []interface{}{rhs})
	_= retVal
}

func (this *TextBox) Index() int32 {
	retVal := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *TextBox) Left() float64 {
	retVal := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *TextBox) SetLeft(rhs float64)  {
	retVal := this.PropPut(0x0000007f, []interface{}{rhs})
	_= retVal
}

func (this *TextBox) Locked() bool {
	retVal := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TextBox) SetLocked(rhs bool)  {
	retVal := this.PropPut(0x0000010d, []interface{}{rhs})
	_= retVal
}

func (this *TextBox) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TextBox) SetName(rhs string)  {
	retVal := this.PropPut(0x0000006e, []interface{}{rhs})
	_= retVal
}

func (this *TextBox) OnAction() string {
	retVal := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TextBox) SetOnAction(rhs string)  {
	retVal := this.PropPut(0x00000254, []interface{}{rhs})
	_= retVal
}

func (this *TextBox) Placement() ole.Variant {
	retVal := this.PropGet(0x00000269, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *TextBox) SetPlacement(rhs interface{})  {
	retVal := this.PropPut(0x00000269, []interface{}{rhs})
	_= retVal
}

func (this *TextBox) PrintObject() bool {
	retVal := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TextBox) SetPrintObject(rhs bool)  {
	retVal := this.PropPut(0x0000026a, []interface{}{rhs})
	_= retVal
}

var TextBox_Select_OptArgs= []string{
	"Replace", 
}

func (this *TextBox) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(TextBox_Select_OptArgs, optArgs)
	retVal := this.Call(0x000000eb, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *TextBox) SendToBack() ole.Variant {
	retVal := this.Call(0x0000025d, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *TextBox) Top() float64 {
	retVal := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *TextBox) SetTop(rhs float64)  {
	retVal := this.PropPut(0x0000007e, []interface{}{rhs})
	_= retVal
}

func (this *TextBox) TopLeftCell() *Range {
	retVal := this.PropGet(0x0000026c, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *TextBox) Visible() bool {
	retVal := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TextBox) SetVisible(rhs bool)  {
	retVal := this.PropPut(0x0000022e, []interface{}{rhs})
	_= retVal
}

func (this *TextBox) Width() float64 {
	retVal := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *TextBox) SetWidth(rhs float64)  {
	retVal := this.PropPut(0x0000007a, []interface{}{rhs})
	_= retVal
}

func (this *TextBox) ZOrder() int32 {
	retVal := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *TextBox) ShapeRange() *ShapeRange {
	retVal := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.PdispValVal(), false, true)
}

func (this *TextBox) AddIndent() bool {
	retVal := this.PropGet(0x00000427, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TextBox) SetAddIndent(rhs bool)  {
	retVal := this.PropPut(0x00000427, []interface{}{rhs})
	_= retVal
}

func (this *TextBox) AutoScaleFont() ole.Variant {
	retVal := this.PropGet(0x000005f5, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *TextBox) SetAutoScaleFont(rhs interface{})  {
	retVal := this.PropPut(0x000005f5, []interface{}{rhs})
	_= retVal
}

func (this *TextBox) AutoSize() bool {
	retVal := this.PropGet(0x00000266, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TextBox) SetAutoSize(rhs bool)  {
	retVal := this.PropPut(0x00000266, []interface{}{rhs})
	_= retVal
}

func (this *TextBox) Caption() string {
	retVal := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TextBox) SetCaption(rhs string)  {
	retVal := this.PropPut(0x0000008b, []interface{}{rhs})
	_= retVal
}

var TextBox_Characters_OptArgs= []string{
	"Start", "Length", 
}

func (this *TextBox) Characters(optArgs ...interface{}) *Characters {
	optArgs = ole.ProcessOptArgs(TextBox_Characters_OptArgs, optArgs)
	retVal := this.PropGet(0x0000025b, nil, optArgs...)
	return NewCharacters(retVal.PdispValVal(), false, true)
}

var TextBox_CheckSpelling_OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang", 
}

func (this *TextBox) CheckSpelling(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(TextBox_CheckSpelling_OptArgs, optArgs)
	retVal := this.Call(0x000001f9, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *TextBox) Font() *Font {
	retVal := this.PropGet(0x00000092, nil)
	return NewFont(retVal.PdispValVal(), false, true)
}

func (this *TextBox) Formula() string {
	retVal := this.PropGet(0x00000105, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TextBox) SetFormula(rhs string)  {
	retVal := this.PropPut(0x00000105, []interface{}{rhs})
	_= retVal
}

func (this *TextBox) HorizontalAlignment() ole.Variant {
	retVal := this.PropGet(0x00000088, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *TextBox) SetHorizontalAlignment(rhs interface{})  {
	retVal := this.PropPut(0x00000088, []interface{}{rhs})
	_= retVal
}

func (this *TextBox) LockedText() bool {
	retVal := this.PropGet(0x00000268, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TextBox) SetLockedText(rhs bool)  {
	retVal := this.PropPut(0x00000268, []interface{}{rhs})
	_= retVal
}

func (this *TextBox) Orientation() ole.Variant {
	retVal := this.PropGet(0x00000086, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *TextBox) SetOrientation(rhs interface{})  {
	retVal := this.PropPut(0x00000086, []interface{}{rhs})
	_= retVal
}

func (this *TextBox) Text() string {
	retVal := this.PropGet(0x0000008a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TextBox) SetText(rhs string)  {
	retVal := this.PropPut(0x0000008a, []interface{}{rhs})
	_= retVal
}

func (this *TextBox) VerticalAlignment() ole.Variant {
	retVal := this.PropGet(0x00000089, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *TextBox) SetVerticalAlignment(rhs interface{})  {
	retVal := this.PropPut(0x00000089, []interface{}{rhs})
	_= retVal
}

func (this *TextBox) ReadingOrder() int32 {
	retVal := this.PropGet(0x000003cf, nil)
	return retVal.LValVal()
}

func (this *TextBox) SetReadingOrder(rhs int32)  {
	retVal := this.PropPut(0x000003cf, []interface{}{rhs})
	_= retVal
}

func (this *TextBox) Border() *Border {
	retVal := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.PdispValVal(), false, true)
}

func (this *TextBox) Interior() *Interior {
	retVal := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.PdispValVal(), false, true)
}

func (this *TextBox) RoundedCorners() bool {
	retVal := this.PropGet(0x0000026b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TextBox) SetRoundedCorners(rhs bool)  {
	retVal := this.PropPut(0x0000026b, []interface{}{rhs})
	_= retVal
}

func (this *TextBox) Shadow() bool {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TextBox) SetShadow(rhs bool)  {
	retVal := this.PropPut(0x00000067, []interface{}{rhs})
	_= retVal
}

