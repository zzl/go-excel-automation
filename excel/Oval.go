package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002089E-0000-0000-C000-000000000046
var IID_Oval = syscall.GUID{0x0002089E, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Oval struct {
	ole.OleClient
}

func NewOval(pDisp *win32.IDispatch, addRef bool, scoped bool) *Oval {
	p := &Oval{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OvalFromVar(v ole.Variant) *Oval {
	return NewOval(v.PdispValVal(), false, false)
}

func (this *Oval) IID() *syscall.GUID {
	return &IID_Oval
}

func (this *Oval) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Oval) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Oval) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Oval) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Oval) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Oval) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Oval) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Oval) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Oval) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Oval) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Oval) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Oval) BottomRightCell() *Range {
	retVal := this.PropGet(0x00000267, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Oval) BringToFront() ole.Variant {
	retVal := this.Call(0x0000025a, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Oval) Copy() ole.Variant {
	retVal := this.Call(0x00000227, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Oval) CopyPicture(appearance int32, format int32) ole.Variant {
	retVal := this.Call(0x000000d5, []interface{}{appearance, format})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Oval) Cut() ole.Variant {
	retVal := this.Call(0x00000235, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Oval) Delete() ole.Variant {
	retVal := this.Call(0x00000075, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Oval) Duplicate() *ole.DispatchClass {
	retVal := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Oval) Enabled() bool {
	retVal := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Oval) SetEnabled(rhs bool)  {
	retVal := this.PropPut(0x00000258, []interface{}{rhs})
	_= retVal
}

func (this *Oval) Height() float64 {
	retVal := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *Oval) SetHeight(rhs float64)  {
	retVal := this.PropPut(0x0000007b, []interface{}{rhs})
	_= retVal
}

func (this *Oval) Index() int32 {
	retVal := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *Oval) Left() float64 {
	retVal := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *Oval) SetLeft(rhs float64)  {
	retVal := this.PropPut(0x0000007f, []interface{}{rhs})
	_= retVal
}

func (this *Oval) Locked() bool {
	retVal := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Oval) SetLocked(rhs bool)  {
	retVal := this.PropPut(0x0000010d, []interface{}{rhs})
	_= retVal
}

func (this *Oval) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Oval) SetName(rhs string)  {
	retVal := this.PropPut(0x0000006e, []interface{}{rhs})
	_= retVal
}

func (this *Oval) OnAction() string {
	retVal := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Oval) SetOnAction(rhs string)  {
	retVal := this.PropPut(0x00000254, []interface{}{rhs})
	_= retVal
}

func (this *Oval) Placement() ole.Variant {
	retVal := this.PropGet(0x00000269, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Oval) SetPlacement(rhs interface{})  {
	retVal := this.PropPut(0x00000269, []interface{}{rhs})
	_= retVal
}

func (this *Oval) PrintObject() bool {
	retVal := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Oval) SetPrintObject(rhs bool)  {
	retVal := this.PropPut(0x0000026a, []interface{}{rhs})
	_= retVal
}

var Oval_Select_OptArgs= []string{
	"Replace", 
}

func (this *Oval) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Oval_Select_OptArgs, optArgs)
	retVal := this.Call(0x000000eb, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Oval) SendToBack() ole.Variant {
	retVal := this.Call(0x0000025d, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Oval) Top() float64 {
	retVal := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *Oval) SetTop(rhs float64)  {
	retVal := this.PropPut(0x0000007e, []interface{}{rhs})
	_= retVal
}

func (this *Oval) TopLeftCell() *Range {
	retVal := this.PropGet(0x0000026c, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Oval) Visible() bool {
	retVal := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Oval) SetVisible(rhs bool)  {
	retVal := this.PropPut(0x0000022e, []interface{}{rhs})
	_= retVal
}

func (this *Oval) Width() float64 {
	retVal := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *Oval) SetWidth(rhs float64)  {
	retVal := this.PropPut(0x0000007a, []interface{}{rhs})
	_= retVal
}

func (this *Oval) ZOrder() int32 {
	retVal := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *Oval) ShapeRange() *ShapeRange {
	retVal := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.PdispValVal(), false, true)
}

func (this *Oval) AddIndent() bool {
	retVal := this.PropGet(0x00000427, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Oval) SetAddIndent(rhs bool)  {
	retVal := this.PropPut(0x00000427, []interface{}{rhs})
	_= retVal
}

func (this *Oval) AutoScaleFont() ole.Variant {
	retVal := this.PropGet(0x000005f5, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Oval) SetAutoScaleFont(rhs interface{})  {
	retVal := this.PropPut(0x000005f5, []interface{}{rhs})
	_= retVal
}

func (this *Oval) AutoSize() bool {
	retVal := this.PropGet(0x00000266, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Oval) SetAutoSize(rhs bool)  {
	retVal := this.PropPut(0x00000266, []interface{}{rhs})
	_= retVal
}

func (this *Oval) Caption() string {
	retVal := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Oval) SetCaption(rhs string)  {
	retVal := this.PropPut(0x0000008b, []interface{}{rhs})
	_= retVal
}

var Oval_Characters_OptArgs= []string{
	"Start", "Length", 
}

func (this *Oval) Characters(optArgs ...interface{}) *Characters {
	optArgs = ole.ProcessOptArgs(Oval_Characters_OptArgs, optArgs)
	retVal := this.PropGet(0x0000025b, nil, optArgs...)
	return NewCharacters(retVal.PdispValVal(), false, true)
}

var Oval_CheckSpelling_OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang", 
}

func (this *Oval) CheckSpelling(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Oval_CheckSpelling_OptArgs, optArgs)
	retVal := this.Call(0x000001f9, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Oval) Font() *Font {
	retVal := this.PropGet(0x00000092, nil)
	return NewFont(retVal.PdispValVal(), false, true)
}

func (this *Oval) Formula() string {
	retVal := this.PropGet(0x00000105, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Oval) SetFormula(rhs string)  {
	retVal := this.PropPut(0x00000105, []interface{}{rhs})
	_= retVal
}

func (this *Oval) HorizontalAlignment() ole.Variant {
	retVal := this.PropGet(0x00000088, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Oval) SetHorizontalAlignment(rhs interface{})  {
	retVal := this.PropPut(0x00000088, []interface{}{rhs})
	_= retVal
}

func (this *Oval) LockedText() bool {
	retVal := this.PropGet(0x00000268, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Oval) SetLockedText(rhs bool)  {
	retVal := this.PropPut(0x00000268, []interface{}{rhs})
	_= retVal
}

func (this *Oval) Orientation() ole.Variant {
	retVal := this.PropGet(0x00000086, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Oval) SetOrientation(rhs interface{})  {
	retVal := this.PropPut(0x00000086, []interface{}{rhs})
	_= retVal
}

func (this *Oval) Text() string {
	retVal := this.PropGet(0x0000008a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Oval) SetText(rhs string)  {
	retVal := this.PropPut(0x0000008a, []interface{}{rhs})
	_= retVal
}

func (this *Oval) VerticalAlignment() ole.Variant {
	retVal := this.PropGet(0x00000089, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Oval) SetVerticalAlignment(rhs interface{})  {
	retVal := this.PropPut(0x00000089, []interface{}{rhs})
	_= retVal
}

func (this *Oval) ReadingOrder() int32 {
	retVal := this.PropGet(0x000003cf, nil)
	return retVal.LValVal()
}

func (this *Oval) SetReadingOrder(rhs int32)  {
	retVal := this.PropPut(0x000003cf, []interface{}{rhs})
	_= retVal
}

func (this *Oval) Border() *Border {
	retVal := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.PdispValVal(), false, true)
}

func (this *Oval) Interior() *Interior {
	retVal := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.PdispValVal(), false, true)
}

func (this *Oval) Shadow() bool {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Oval) SetShadow(rhs bool)  {
	retVal := this.PropPut(0x00000067, []interface{}{rhs})
	_= retVal
}

