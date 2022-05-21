package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002089C-0000-0000-C000-000000000046
var IID_Rectangle = syscall.GUID{0x0002089C, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Rectangle struct {
	ole.OleClient
}

func NewRectangle(pDisp *win32.IDispatch, addRef bool, scoped bool) *Rectangle {
	p := &Rectangle{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func RectangleFromVar(v ole.Variant) *Rectangle {
	return NewRectangle(v.PdispValVal(), false, false)
}

func (this *Rectangle) IID() *syscall.GUID {
	return &IID_Rectangle
}

func (this *Rectangle) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Rectangle) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Rectangle) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Rectangle) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Rectangle) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Rectangle) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Rectangle) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Rectangle) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Rectangle) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Rectangle) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Rectangle) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Rectangle) BottomRightCell() *Range {
	retVal := this.PropGet(0x00000267, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Rectangle) BringToFront() ole.Variant {
	retVal := this.Call(0x0000025a, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangle) Copy() ole.Variant {
	retVal := this.Call(0x00000227, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangle) CopyPicture(appearance int32, format int32) ole.Variant {
	retVal := this.Call(0x000000d5, []interface{}{appearance, format})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangle) Cut() ole.Variant {
	retVal := this.Call(0x00000235, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangle) Delete() ole.Variant {
	retVal := this.Call(0x00000075, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangle) Duplicate() *ole.DispatchClass {
	retVal := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Rectangle) Enabled() bool {
	retVal := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Rectangle) SetEnabled(rhs bool)  {
	retVal := this.PropPut(0x00000258, []interface{}{rhs})
	_= retVal
}

func (this *Rectangle) Height() float64 {
	retVal := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *Rectangle) SetHeight(rhs float64)  {
	retVal := this.PropPut(0x0000007b, []interface{}{rhs})
	_= retVal
}

func (this *Rectangle) Index() int32 {
	retVal := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *Rectangle) Left() float64 {
	retVal := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *Rectangle) SetLeft(rhs float64)  {
	retVal := this.PropPut(0x0000007f, []interface{}{rhs})
	_= retVal
}

func (this *Rectangle) Locked() bool {
	retVal := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Rectangle) SetLocked(rhs bool)  {
	retVal := this.PropPut(0x0000010d, []interface{}{rhs})
	_= retVal
}

func (this *Rectangle) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Rectangle) SetName(rhs string)  {
	retVal := this.PropPut(0x0000006e, []interface{}{rhs})
	_= retVal
}

func (this *Rectangle) OnAction() string {
	retVal := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Rectangle) SetOnAction(rhs string)  {
	retVal := this.PropPut(0x00000254, []interface{}{rhs})
	_= retVal
}

func (this *Rectangle) Placement() ole.Variant {
	retVal := this.PropGet(0x00000269, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangle) SetPlacement(rhs interface{})  {
	retVal := this.PropPut(0x00000269, []interface{}{rhs})
	_= retVal
}

func (this *Rectangle) PrintObject() bool {
	retVal := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Rectangle) SetPrintObject(rhs bool)  {
	retVal := this.PropPut(0x0000026a, []interface{}{rhs})
	_= retVal
}

var Rectangle_Select_OptArgs= []string{
	"Replace", 
}

func (this *Rectangle) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Rectangle_Select_OptArgs, optArgs)
	retVal := this.Call(0x000000eb, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangle) SendToBack() ole.Variant {
	retVal := this.Call(0x0000025d, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangle) Top() float64 {
	retVal := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *Rectangle) SetTop(rhs float64)  {
	retVal := this.PropPut(0x0000007e, []interface{}{rhs})
	_= retVal
}

func (this *Rectangle) TopLeftCell() *Range {
	retVal := this.PropGet(0x0000026c, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Rectangle) Visible() bool {
	retVal := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Rectangle) SetVisible(rhs bool)  {
	retVal := this.PropPut(0x0000022e, []interface{}{rhs})
	_= retVal
}

func (this *Rectangle) Width() float64 {
	retVal := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *Rectangle) SetWidth(rhs float64)  {
	retVal := this.PropPut(0x0000007a, []interface{}{rhs})
	_= retVal
}

func (this *Rectangle) ZOrder() int32 {
	retVal := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *Rectangle) ShapeRange() *ShapeRange {
	retVal := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.PdispValVal(), false, true)
}

func (this *Rectangle) AddIndent() bool {
	retVal := this.PropGet(0x00000427, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Rectangle) SetAddIndent(rhs bool)  {
	retVal := this.PropPut(0x00000427, []interface{}{rhs})
	_= retVal
}

func (this *Rectangle) AutoScaleFont() ole.Variant {
	retVal := this.PropGet(0x000005f5, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangle) SetAutoScaleFont(rhs interface{})  {
	retVal := this.PropPut(0x000005f5, []interface{}{rhs})
	_= retVal
}

func (this *Rectangle) AutoSize() bool {
	retVal := this.PropGet(0x00000266, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Rectangle) SetAutoSize(rhs bool)  {
	retVal := this.PropPut(0x00000266, []interface{}{rhs})
	_= retVal
}

func (this *Rectangle) Caption() string {
	retVal := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Rectangle) SetCaption(rhs string)  {
	retVal := this.PropPut(0x0000008b, []interface{}{rhs})
	_= retVal
}

var Rectangle_Characters_OptArgs= []string{
	"Start", "Length", 
}

func (this *Rectangle) Characters(optArgs ...interface{}) *Characters {
	optArgs = ole.ProcessOptArgs(Rectangle_Characters_OptArgs, optArgs)
	retVal := this.PropGet(0x0000025b, nil, optArgs...)
	return NewCharacters(retVal.PdispValVal(), false, true)
}

var Rectangle_CheckSpelling_OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang", 
}

func (this *Rectangle) CheckSpelling(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Rectangle_CheckSpelling_OptArgs, optArgs)
	retVal := this.Call(0x000001f9, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangle) Font() *Font {
	retVal := this.PropGet(0x00000092, nil)
	return NewFont(retVal.PdispValVal(), false, true)
}

func (this *Rectangle) Formula() string {
	retVal := this.PropGet(0x00000105, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Rectangle) SetFormula(rhs string)  {
	retVal := this.PropPut(0x00000105, []interface{}{rhs})
	_= retVal
}

func (this *Rectangle) HorizontalAlignment() ole.Variant {
	retVal := this.PropGet(0x00000088, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangle) SetHorizontalAlignment(rhs interface{})  {
	retVal := this.PropPut(0x00000088, []interface{}{rhs})
	_= retVal
}

func (this *Rectangle) LockedText() bool {
	retVal := this.PropGet(0x00000268, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Rectangle) SetLockedText(rhs bool)  {
	retVal := this.PropPut(0x00000268, []interface{}{rhs})
	_= retVal
}

func (this *Rectangle) Orientation() ole.Variant {
	retVal := this.PropGet(0x00000086, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangle) SetOrientation(rhs interface{})  {
	retVal := this.PropPut(0x00000086, []interface{}{rhs})
	_= retVal
}

func (this *Rectangle) Text() string {
	retVal := this.PropGet(0x0000008a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Rectangle) SetText(rhs string)  {
	retVal := this.PropPut(0x0000008a, []interface{}{rhs})
	_= retVal
}

func (this *Rectangle) VerticalAlignment() ole.Variant {
	retVal := this.PropGet(0x00000089, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangle) SetVerticalAlignment(rhs interface{})  {
	retVal := this.PropPut(0x00000089, []interface{}{rhs})
	_= retVal
}

func (this *Rectangle) ReadingOrder() int32 {
	retVal := this.PropGet(0x000003cf, nil)
	return retVal.LValVal()
}

func (this *Rectangle) SetReadingOrder(rhs int32)  {
	retVal := this.PropPut(0x000003cf, []interface{}{rhs})
	_= retVal
}

func (this *Rectangle) Border() *Border {
	retVal := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.PdispValVal(), false, true)
}

func (this *Rectangle) Interior() *Interior {
	retVal := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.PdispValVal(), false, true)
}

func (this *Rectangle) Shadow() bool {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Rectangle) SetShadow(rhs bool)  {
	retVal := this.PropPut(0x00000067, []interface{}{rhs})
	_= retVal
}

func (this *Rectangle) RoundedCorners() bool {
	retVal := this.PropGet(0x0000026b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Rectangle) SetRoundedCorners(rhs bool)  {
	retVal := this.PropPut(0x0000026b, []interface{}{rhs})
	_= retVal
}

