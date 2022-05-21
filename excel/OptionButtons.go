package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020882-0000-0000-C000-000000000046
var IID_OptionButtons = syscall.GUID{0x00020882, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type OptionButtons struct {
	ole.OleClient
}

func NewOptionButtons(pDisp *win32.IDispatch, addRef bool, scoped bool) *OptionButtons {
	p := &OptionButtons{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OptionButtonsFromVar(v ole.Variant) *OptionButtons {
	return NewOptionButtons(v.PdispValVal(), false, false)
}

func (this *OptionButtons) IID() *syscall.GUID {
	return &IID_OptionButtons
}

func (this *OptionButtons) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OptionButtons) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *OptionButtons) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *OptionButtons) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *OptionButtons) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *OptionButtons) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *OptionButtons) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *OptionButtons) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *OptionButtons) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OptionButtons) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *OptionButtons) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OptionButtons) Dummy3_()  {
	retVal := this.Call(0x00010003, nil)
	_= retVal
}

func (this *OptionButtons) BringToFront() ole.Variant {
	retVal := this.Call(0x0000025a, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OptionButtons) Copy() ole.Variant {
	retVal := this.Call(0x00000227, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OptionButtons) CopyPicture(appearance int32, format int32) ole.Variant {
	retVal := this.Call(0x000000d5, []interface{}{appearance, format})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OptionButtons) Cut() ole.Variant {
	retVal := this.Call(0x00000235, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OptionButtons) Delete() ole.Variant {
	retVal := this.Call(0x00000075, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OptionButtons) Duplicate() *ole.DispatchClass {
	retVal := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OptionButtons) Enabled() bool {
	retVal := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OptionButtons) SetEnabled(rhs bool)  {
	retVal := this.PropPut(0x00000258, []interface{}{rhs})
	_= retVal
}

func (this *OptionButtons) Height() float64 {
	retVal := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *OptionButtons) SetHeight(rhs float64)  {
	retVal := this.PropPut(0x0000007b, []interface{}{rhs})
	_= retVal
}

func (this *OptionButtons) Dummy12_()  {
	retVal := this.Call(0x0001000c, nil)
	_= retVal
}

func (this *OptionButtons) Left() float64 {
	retVal := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *OptionButtons) SetLeft(rhs float64)  {
	retVal := this.PropPut(0x0000007f, []interface{}{rhs})
	_= retVal
}

func (this *OptionButtons) Locked() bool {
	retVal := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OptionButtons) SetLocked(rhs bool)  {
	retVal := this.PropPut(0x0000010d, []interface{}{rhs})
	_= retVal
}

func (this *OptionButtons) Dummy15_()  {
	retVal := this.Call(0x0001000f, nil)
	_= retVal
}

func (this *OptionButtons) OnAction() string {
	retVal := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OptionButtons) SetOnAction(rhs string)  {
	retVal := this.PropPut(0x00000254, []interface{}{rhs})
	_= retVal
}

func (this *OptionButtons) Placement() ole.Variant {
	retVal := this.PropGet(0x00000269, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OptionButtons) SetPlacement(rhs interface{})  {
	retVal := this.PropPut(0x00000269, []interface{}{rhs})
	_= retVal
}

func (this *OptionButtons) PrintObject() bool {
	retVal := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OptionButtons) SetPrintObject(rhs bool)  {
	retVal := this.PropPut(0x0000026a, []interface{}{rhs})
	_= retVal
}

var OptionButtons_Select_OptArgs= []string{
	"Replace", 
}

func (this *OptionButtons) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(OptionButtons_Select_OptArgs, optArgs)
	retVal := this.Call(0x000000eb, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OptionButtons) SendToBack() ole.Variant {
	retVal := this.Call(0x0000025d, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OptionButtons) Top() float64 {
	retVal := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *OptionButtons) SetTop(rhs float64)  {
	retVal := this.PropPut(0x0000007e, []interface{}{rhs})
	_= retVal
}

func (this *OptionButtons) Dummy22_()  {
	retVal := this.Call(0x00010016, nil)
	_= retVal
}

func (this *OptionButtons) Visible() bool {
	retVal := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OptionButtons) SetVisible(rhs bool)  {
	retVal := this.PropPut(0x0000022e, []interface{}{rhs})
	_= retVal
}

func (this *OptionButtons) Width() float64 {
	retVal := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *OptionButtons) SetWidth(rhs float64)  {
	retVal := this.PropPut(0x0000007a, []interface{}{rhs})
	_= retVal
}

func (this *OptionButtons) ZOrder() int32 {
	retVal := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *OptionButtons) ShapeRange() *ShapeRange {
	retVal := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.PdispValVal(), false, true)
}

func (this *OptionButtons) Caption() string {
	retVal := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OptionButtons) SetCaption(rhs string)  {
	retVal := this.PropPut(0x0000008b, []interface{}{rhs})
	_= retVal
}

var OptionButtons_Characters_OptArgs= []string{
	"Start", "Length", 
}

func (this *OptionButtons) Characters(optArgs ...interface{}) *Characters {
	optArgs = ole.ProcessOptArgs(OptionButtons_Characters_OptArgs, optArgs)
	retVal := this.PropGet(0x0000025b, nil, optArgs...)
	return NewCharacters(retVal.PdispValVal(), false, true)
}

var OptionButtons_CheckSpelling_OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang", 
}

func (this *OptionButtons) CheckSpelling(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(OptionButtons_CheckSpelling_OptArgs, optArgs)
	retVal := this.Call(0x000001f9, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OptionButtons) LockedText() bool {
	retVal := this.PropGet(0x00000268, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OptionButtons) SetLockedText(rhs bool)  {
	retVal := this.PropPut(0x00000268, []interface{}{rhs})
	_= retVal
}

func (this *OptionButtons) Text() string {
	retVal := this.PropGet(0x0000008a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OptionButtons) SetText(rhs string)  {
	retVal := this.PropPut(0x0000008a, []interface{}{rhs})
	_= retVal
}

func (this *OptionButtons) Accelerator() ole.Variant {
	retVal := this.PropGet(0x0000034e, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OptionButtons) SetAccelerator(rhs interface{})  {
	retVal := this.PropPut(0x0000034e, []interface{}{rhs})
	_= retVal
}

func (this *OptionButtons) Border() *Border {
	retVal := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.PdispValVal(), false, true)
}

func (this *OptionButtons) Default_() int32 {
	retVal := this.PropGet(0x00000000, nil)
	return retVal.LValVal()
}

func (this *OptionButtons) SetDefault_(rhs int32)  {
	retVal := this.PropPut(0x00000000, []interface{}{rhs})
	_= retVal
}

func (this *OptionButtons) Display3DShading() bool {
	retVal := this.PropGet(0x00000462, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OptionButtons) SetDisplay3DShading(rhs bool)  {
	retVal := this.PropPut(0x00000462, []interface{}{rhs})
	_= retVal
}

func (this *OptionButtons) Interior() *Interior {
	retVal := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.PdispValVal(), false, true)
}

func (this *OptionButtons) LinkedCell() string {
	retVal := this.PropGet(0x00000422, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OptionButtons) SetLinkedCell(rhs string)  {
	retVal := this.PropPut(0x00000422, []interface{}{rhs})
	_= retVal
}

func (this *OptionButtons) PhoneticAccelerator() ole.Variant {
	retVal := this.PropGet(0x00000461, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OptionButtons) SetPhoneticAccelerator(rhs interface{})  {
	retVal := this.PropPut(0x00000461, []interface{}{rhs})
	_= retVal
}

func (this *OptionButtons) Value() ole.Variant {
	retVal := this.PropGet(0x00000006, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OptionButtons) SetValue(rhs interface{})  {
	retVal := this.PropPut(0x00000006, []interface{}{rhs})
	_= retVal
}

func (this *OptionButtons) GroupBox() *GroupBox {
	retVal := this.PropGet(0x00000341, nil)
	return NewGroupBox(retVal.PdispValVal(), false, true)
}

func (this *OptionButtons) Add(left float64, top float64, width float64, height float64) *OptionButton {
	retVal := this.Call(0x000000b5, []interface{}{left, top, width, height})
	return NewOptionButton(retVal.PdispValVal(), false, true)
}

func (this *OptionButtons) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *OptionButtons) Group() *GroupObject {
	retVal := this.Call(0x0000002e, nil)
	return NewGroupObject(retVal.PdispValVal(), false, true)
}

func (this *OptionButtons) Item(index interface{}) *ole.DispatchClass {
	retVal := this.Call(0x000000aa, []interface{}{index})
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OptionButtons) NewEnum_() *com.UnknownClass {
	retVal := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *OptionButtons) ForEach(action func(item int32) bool) {
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
