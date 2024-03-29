package excel

import (
	"github.com/zzl/go-win32api/v2/win32"
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
	if pDisp == nil {
		return nil
	}
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
	return NewOptionButtons(v.IDispatch(), false, false)
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

func (this *OptionButtons) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *OptionButtons) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *OptionButtons) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *OptionButtons) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *OptionButtons) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *OptionButtons) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *OptionButtons) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *OptionButtons) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OptionButtons) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *OptionButtons) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OptionButtons) Dummy3_() {
	retVal, _ := this.Call(0x00010003, nil)
	_ = retVal
}

func (this *OptionButtons) BringToFront() ole.Variant {
	retVal, _ := this.Call(0x0000025a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *OptionButtons) Copy() ole.Variant {
	retVal, _ := this.Call(0x00000227, nil)
	com.AddToScope(retVal)
	return *retVal
}

var OptionButtons_CopyPicture_OptArgs = []string{
	"Appearance", "Format",
}

func (this *OptionButtons) CopyPicture(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(OptionButtons_CopyPicture_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d5, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *OptionButtons) Cut() ole.Variant {
	retVal, _ := this.Call(0x00000235, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *OptionButtons) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *OptionButtons) Duplicate() *ole.DispatchClass {
	retVal, _ := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OptionButtons) Enabled() bool {
	retVal, _ := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OptionButtons) SetEnabled(rhs bool) {
	_ = this.PropPut(0x00000258, []interface{}{rhs})
}

func (this *OptionButtons) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *OptionButtons) SetHeight(rhs float64) {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *OptionButtons) Dummy12_() {
	retVal, _ := this.Call(0x0001000c, nil)
	_ = retVal
}

func (this *OptionButtons) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *OptionButtons) SetLeft(rhs float64) {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *OptionButtons) Locked() bool {
	retVal, _ := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OptionButtons) SetLocked(rhs bool) {
	_ = this.PropPut(0x0000010d, []interface{}{rhs})
}

func (this *OptionButtons) Dummy15_() {
	retVal, _ := this.Call(0x0001000f, nil)
	_ = retVal
}

func (this *OptionButtons) OnAction() string {
	retVal, _ := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OptionButtons) SetOnAction(rhs string) {
	_ = this.PropPut(0x00000254, []interface{}{rhs})
}

func (this *OptionButtons) Placement() ole.Variant {
	retVal, _ := this.PropGet(0x00000269, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *OptionButtons) SetPlacement(rhs interface{}) {
	_ = this.PropPut(0x00000269, []interface{}{rhs})
}

func (this *OptionButtons) PrintObject() bool {
	retVal, _ := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OptionButtons) SetPrintObject(rhs bool) {
	_ = this.PropPut(0x0000026a, []interface{}{rhs})
}

var OptionButtons_Select_OptArgs = []string{
	"Replace",
}

func (this *OptionButtons) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(OptionButtons_Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000eb, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *OptionButtons) SendToBack() ole.Variant {
	retVal, _ := this.Call(0x0000025d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *OptionButtons) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *OptionButtons) SetTop(rhs float64) {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *OptionButtons) Dummy22_() {
	retVal, _ := this.Call(0x00010016, nil)
	_ = retVal
}

func (this *OptionButtons) Visible() bool {
	retVal, _ := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OptionButtons) SetVisible(rhs bool) {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

func (this *OptionButtons) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *OptionButtons) SetWidth(rhs float64) {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *OptionButtons) ZOrder() int32 {
	retVal, _ := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *OptionButtons) ShapeRange() *ShapeRange {
	retVal, _ := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.IDispatch(), false, true)
}

func (this *OptionButtons) Caption() string {
	retVal, _ := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OptionButtons) SetCaption(rhs string) {
	_ = this.PropPut(0x0000008b, []interface{}{rhs})
}

var OptionButtons_Characters_OptArgs = []string{
	"Start", "Length",
}

func (this *OptionButtons) Characters(optArgs ...interface{}) *Characters {
	optArgs = ole.ProcessOptArgs(OptionButtons_Characters_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000025b, nil, optArgs...)
	return NewCharacters(retVal.IDispatch(), false, true)
}

var OptionButtons_CheckSpelling_OptArgs = []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang",
}

func (this *OptionButtons) CheckSpelling(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(OptionButtons_CheckSpelling_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001f9, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *OptionButtons) LockedText() bool {
	retVal, _ := this.PropGet(0x00000268, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OptionButtons) SetLockedText(rhs bool) {
	_ = this.PropPut(0x00000268, []interface{}{rhs})
}

func (this *OptionButtons) Text() string {
	retVal, _ := this.PropGet(0x0000008a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OptionButtons) SetText(rhs string) {
	_ = this.PropPut(0x0000008a, []interface{}{rhs})
}

func (this *OptionButtons) Accelerator() ole.Variant {
	retVal, _ := this.PropGet(0x0000034e, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *OptionButtons) SetAccelerator(rhs interface{}) {
	_ = this.PropPut(0x0000034e, []interface{}{rhs})
}

func (this *OptionButtons) Border() *Border {
	retVal, _ := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.IDispatch(), false, true)
}

func (this *OptionButtons) Default_() int32 {
	retVal, _ := this.PropGet(0x00000000, nil)
	return retVal.LValVal()
}

func (this *OptionButtons) SetDefault_(rhs int32) {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *OptionButtons) Display3DShading() bool {
	retVal, _ := this.PropGet(0x00000462, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OptionButtons) SetDisplay3DShading(rhs bool) {
	_ = this.PropPut(0x00000462, []interface{}{rhs})
}

func (this *OptionButtons) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *OptionButtons) LinkedCell() string {
	retVal, _ := this.PropGet(0x00000422, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OptionButtons) SetLinkedCell(rhs string) {
	_ = this.PropPut(0x00000422, []interface{}{rhs})
}

func (this *OptionButtons) PhoneticAccelerator() ole.Variant {
	retVal, _ := this.PropGet(0x00000461, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *OptionButtons) SetPhoneticAccelerator(rhs interface{}) {
	_ = this.PropPut(0x00000461, []interface{}{rhs})
}

func (this *OptionButtons) Value() ole.Variant {
	retVal, _ := this.PropGet(0x00000006, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *OptionButtons) SetValue(rhs interface{}) {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *OptionButtons) GroupBox() *GroupBox {
	retVal, _ := this.PropGet(0x00000341, nil)
	return NewGroupBox(retVal.IDispatch(), false, true)
}

func (this *OptionButtons) Add(left float64, top float64, width float64, height float64) *OptionButton {
	retVal, _ := this.Call(0x000000b5, []interface{}{left, top, width, height})
	return NewOptionButton(retVal.IDispatch(), false, true)
}

func (this *OptionButtons) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *OptionButtons) Group() *GroupObject {
	retVal, _ := this.Call(0x0000002e, nil)
	return NewGroupObject(retVal.IDispatch(), false, true)
}

func (this *OptionButtons) Item(index interface{}) *ole.DispatchClass {
	retVal, _ := this.Call(0x000000aa, []interface{}{index})
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OptionButtons) NewEnum_() *com.UnknownClass {
	retVal, _ := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *OptionButtons) ForEach(action func(item int32) bool) {
	pEnum := this.NewEnum_()
	var pEnumVar *win32.IEnumVARIANT
	pEnum.QueryInterface(&win32.IID_IEnumVARIANT, unsafe.Pointer(&pEnumVar))
	defer pEnumVar.Release()
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

