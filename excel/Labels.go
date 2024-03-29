package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00020891-0000-0000-C000-000000000046
var IID_Labels = syscall.GUID{0x00020891, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Labels struct {
	ole.OleClient
}

func NewLabels(pDisp *win32.IDispatch, addRef bool, scoped bool) *Labels {
	if pDisp == nil {
		return nil
	}
	p := &Labels{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func LabelsFromVar(v ole.Variant) *Labels {
	return NewLabels(v.IDispatch(), false, false)
}

func (this *Labels) IID() *syscall.GUID {
	return &IID_Labels
}

func (this *Labels) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Labels) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Labels) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Labels) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Labels) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Labels) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Labels) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Labels) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Labels) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Labels) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Labels) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Labels) Dummy3_() {
	retVal, _ := this.Call(0x00010003, nil)
	_ = retVal
}

func (this *Labels) BringToFront() ole.Variant {
	retVal, _ := this.Call(0x0000025a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Labels) Copy() ole.Variant {
	retVal, _ := this.Call(0x00000227, nil)
	com.AddToScope(retVal)
	return *retVal
}

var Labels_CopyPicture_OptArgs = []string{
	"Appearance", "Format",
}

func (this *Labels) CopyPicture(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Labels_CopyPicture_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d5, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Labels) Cut() ole.Variant {
	retVal, _ := this.Call(0x00000235, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Labels) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Labels) Duplicate() *ole.DispatchClass {
	retVal, _ := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Labels) Enabled() bool {
	retVal, _ := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Labels) SetEnabled(rhs bool) {
	_ = this.PropPut(0x00000258, []interface{}{rhs})
}

func (this *Labels) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *Labels) SetHeight(rhs float64) {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *Labels) Dummy12_() {
	retVal, _ := this.Call(0x0001000c, nil)
	_ = retVal
}

func (this *Labels) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *Labels) SetLeft(rhs float64) {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *Labels) Locked() bool {
	retVal, _ := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Labels) SetLocked(rhs bool) {
	_ = this.PropPut(0x0000010d, []interface{}{rhs})
}

func (this *Labels) Dummy15_() {
	retVal, _ := this.Call(0x0001000f, nil)
	_ = retVal
}

func (this *Labels) OnAction() string {
	retVal, _ := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Labels) SetOnAction(rhs string) {
	_ = this.PropPut(0x00000254, []interface{}{rhs})
}

func (this *Labels) Placement() ole.Variant {
	retVal, _ := this.PropGet(0x00000269, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Labels) SetPlacement(rhs interface{}) {
	_ = this.PropPut(0x00000269, []interface{}{rhs})
}

func (this *Labels) PrintObject() bool {
	retVal, _ := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Labels) SetPrintObject(rhs bool) {
	_ = this.PropPut(0x0000026a, []interface{}{rhs})
}

var Labels_Select_OptArgs = []string{
	"Replace",
}

func (this *Labels) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Labels_Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000eb, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Labels) SendToBack() ole.Variant {
	retVal, _ := this.Call(0x0000025d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Labels) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *Labels) SetTop(rhs float64) {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *Labels) Dummy22_() {
	retVal, _ := this.Call(0x00010016, nil)
	_ = retVal
}

func (this *Labels) Visible() bool {
	retVal, _ := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Labels) SetVisible(rhs bool) {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

func (this *Labels) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *Labels) SetWidth(rhs float64) {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *Labels) ZOrder() int32 {
	retVal, _ := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *Labels) ShapeRange() *ShapeRange {
	retVal, _ := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.IDispatch(), false, true)
}

func (this *Labels) Caption() string {
	retVal, _ := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Labels) SetCaption(rhs string) {
	_ = this.PropPut(0x0000008b, []interface{}{rhs})
}

var Labels_Characters_OptArgs = []string{
	"Start", "Length",
}

func (this *Labels) Characters(optArgs ...interface{}) *Characters {
	optArgs = ole.ProcessOptArgs(Labels_Characters_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000025b, nil, optArgs...)
	return NewCharacters(retVal.IDispatch(), false, true)
}

var Labels_CheckSpelling_OptArgs = []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang",
}

func (this *Labels) CheckSpelling(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Labels_CheckSpelling_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001f9, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Labels) LockedText() bool {
	retVal, _ := this.PropGet(0x00000268, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Labels) SetLockedText(rhs bool) {
	_ = this.PropPut(0x00000268, []interface{}{rhs})
}

func (this *Labels) Text() string {
	retVal, _ := this.PropGet(0x0000008a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Labels) SetText(rhs string) {
	_ = this.PropPut(0x0000008a, []interface{}{rhs})
}

func (this *Labels) Accelerator() ole.Variant {
	retVal, _ := this.PropGet(0x0000034e, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Labels) SetAccelerator(rhs interface{}) {
	_ = this.PropPut(0x0000034e, []interface{}{rhs})
}

func (this *Labels) Dummy33_() {
	retVal, _ := this.Call(0x00010021, nil)
	_ = retVal
}

func (this *Labels) PhoneticAccelerator() ole.Variant {
	retVal, _ := this.PropGet(0x00000461, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Labels) SetPhoneticAccelerator(rhs interface{}) {
	_ = this.PropPut(0x00000461, []interface{}{rhs})
}

func (this *Labels) Add(left float64, top float64, width float64, height float64) *Label {
	retVal, _ := this.Call(0x000000b5, []interface{}{left, top, width, height})
	return NewLabel(retVal.IDispatch(), false, true)
}

func (this *Labels) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *Labels) Group() *GroupObject {
	retVal, _ := this.Call(0x0000002e, nil)
	return NewGroupObject(retVal.IDispatch(), false, true)
}

func (this *Labels) Item(index interface{}) *ole.DispatchClass {
	retVal, _ := this.Call(0x000000aa, []interface{}{index})
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Labels) NewEnum_() *com.UnknownClass {
	retVal, _ := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Labels) ForEach(action func(item *ole.DispatchClass) bool) {
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
		pItem := (*ole.DispatchClass)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}
