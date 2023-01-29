package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 0002088F-0000-0000-C000-000000000046
var IID_DialogFrame = syscall.GUID{0x0002088F, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type DialogFrame struct {
	ole.OleClient
}

func NewDialogFrame(pDisp *win32.IDispatch, addRef bool, scoped bool) *DialogFrame {
	if pDisp == nil {
		return nil
	}
	p := &DialogFrame{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DialogFrameFromVar(v ole.Variant) *DialogFrame {
	return NewDialogFrame(v.IDispatch(), false, false)
}

func (this *DialogFrame) IID() *syscall.GUID {
	return &IID_DialogFrame
}

func (this *DialogFrame) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DialogFrame) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *DialogFrame) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *DialogFrame) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *DialogFrame) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *DialogFrame) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *DialogFrame) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *DialogFrame) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *DialogFrame) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *DialogFrame) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *DialogFrame) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DialogFrame) Dummy3_() {
	retVal, _ := this.Call(0x00010003, nil)
	_ = retVal
}

func (this *DialogFrame) Dummy4_() {
	retVal, _ := this.Call(0x00010004, nil)
	_ = retVal
}

func (this *DialogFrame) Dummy5_() {
	retVal, _ := this.Call(0x00010005, nil)
	_ = retVal
}

var DialogFrame_CopyPicture_OptArgs = []string{
	"Appearance", "Format",
}

func (this *DialogFrame) CopyPicture(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DialogFrame_CopyPicture_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d5, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DialogFrame) Dummy7_() {
	retVal, _ := this.Call(0x00010007, nil)
	_ = retVal
}

func (this *DialogFrame) Dummy8_() {
	retVal, _ := this.Call(0x00010008, nil)
	_ = retVal
}

func (this *DialogFrame) Dummy9_() {
	retVal, _ := this.Call(0x00010009, nil)
	_ = retVal
}

func (this *DialogFrame) Dummy10_() {
	retVal, _ := this.Call(0x0001000a, nil)
	_ = retVal
}

func (this *DialogFrame) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *DialogFrame) SetHeight(rhs float64) {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *DialogFrame) Dummy12_() {
	retVal, _ := this.Call(0x0001000c, nil)
	_ = retVal
}

func (this *DialogFrame) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *DialogFrame) SetLeft(rhs float64) {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *DialogFrame) Locked() bool {
	retVal, _ := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DialogFrame) SetLocked(rhs bool) {
	_ = this.PropPut(0x0000010d, []interface{}{rhs})
}

func (this *DialogFrame) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DialogFrame) SetName(rhs string) {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *DialogFrame) OnAction() string {
	retVal, _ := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DialogFrame) SetOnAction(rhs string) {
	_ = this.PropPut(0x00000254, []interface{}{rhs})
}

func (this *DialogFrame) Dummy17_() {
	retVal, _ := this.Call(0x00010011, nil)
	_ = retVal
}

func (this *DialogFrame) Dummy18_() {
	retVal, _ := this.Call(0x00010012, nil)
	_ = retVal
}

var DialogFrame_Select_OptArgs = []string{
	"Replace",
}

func (this *DialogFrame) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DialogFrame_Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000eb, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DialogFrame) Dummy20_() {
	retVal, _ := this.Call(0x00010014, nil)
	_ = retVal
}

func (this *DialogFrame) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *DialogFrame) SetTop(rhs float64) {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *DialogFrame) Dummy22_() {
	retVal, _ := this.Call(0x00010016, nil)
	_ = retVal
}

func (this *DialogFrame) Dummy23_() {
	retVal, _ := this.Call(0x00010017, nil)
	_ = retVal
}

func (this *DialogFrame) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *DialogFrame) SetWidth(rhs float64) {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *DialogFrame) Dummy25_() {
	retVal, _ := this.Call(0x00010019, nil)
	_ = retVal
}

func (this *DialogFrame) ShapeRange() *ShapeRange {
	retVal, _ := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.IDispatch(), false, true)
}

func (this *DialogFrame) Caption() string {
	retVal, _ := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DialogFrame) SetCaption(rhs string) {
	_ = this.PropPut(0x0000008b, []interface{}{rhs})
}

var DialogFrame_Characters_OptArgs = []string{
	"Start", "Length",
}

func (this *DialogFrame) Characters(optArgs ...interface{}) *Characters {
	optArgs = ole.ProcessOptArgs(DialogFrame_Characters_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000025b, nil, optArgs...)
	return NewCharacters(retVal.IDispatch(), false, true)
}

var DialogFrame_CheckSpelling_OptArgs = []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang",
}

func (this *DialogFrame) CheckSpelling(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DialogFrame_CheckSpelling_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001f9, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DialogFrame) LockedText() bool {
	retVal, _ := this.PropGet(0x00000268, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DialogFrame) SetLockedText(rhs bool) {
	_ = this.PropPut(0x00000268, []interface{}{rhs})
}

func (this *DialogFrame) Text() string {
	retVal, _ := this.PropGet(0x0000008a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DialogFrame) SetText(rhs string) {
	_ = this.PropPut(0x0000008a, []interface{}{rhs})
}
