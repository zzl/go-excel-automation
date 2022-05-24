package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020889-0000-0000-C000-000000000046
var IID_GroupBox = syscall.GUID{0x00020889, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type GroupBox struct {
	ole.OleClient
}

func NewGroupBox(pDisp *win32.IDispatch, addRef bool, scoped bool) *GroupBox {
	 if pDisp == nil {
		return nil;
	}
	p := &GroupBox{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func GroupBoxFromVar(v ole.Variant) *GroupBox {
	return NewGroupBox(v.IDispatch(), false, false)
}

func (this *GroupBox) IID() *syscall.GUID {
	return &IID_GroupBox
}

func (this *GroupBox) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *GroupBox) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *GroupBox) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *GroupBox) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *GroupBox) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *GroupBox) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *GroupBox) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *GroupBox) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *GroupBox) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *GroupBox) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *GroupBox) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *GroupBox) BottomRightCell() *Range {
	retVal, _ := this.PropGet(0x00000267, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *GroupBox) BringToFront() ole.Variant {
	retVal, _ := this.Call(0x0000025a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupBox) Copy() ole.Variant {
	retVal, _ := this.Call(0x00000227, nil)
	com.AddToScope(retVal)
	return *retVal
}

var GroupBox_CopyPicture_OptArgs= []string{
	"Appearance", "Format", 
}

func (this *GroupBox) CopyPicture(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(GroupBox_CopyPicture_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d5, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupBox) Cut() ole.Variant {
	retVal, _ := this.Call(0x00000235, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupBox) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupBox) Duplicate() *ole.DispatchClass {
	retVal, _ := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *GroupBox) Enabled() bool {
	retVal, _ := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupBox) SetEnabled(rhs bool)  {
	_ = this.PropPut(0x00000258, []interface{}{rhs})
}

func (this *GroupBox) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *GroupBox) SetHeight(rhs float64)  {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *GroupBox) Index() int32 {
	retVal, _ := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *GroupBox) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *GroupBox) SetLeft(rhs float64)  {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *GroupBox) Locked() bool {
	retVal, _ := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupBox) SetLocked(rhs bool)  {
	_ = this.PropPut(0x0000010d, []interface{}{rhs})
}

func (this *GroupBox) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *GroupBox) SetName(rhs string)  {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *GroupBox) OnAction() string {
	retVal, _ := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *GroupBox) SetOnAction(rhs string)  {
	_ = this.PropPut(0x00000254, []interface{}{rhs})
}

func (this *GroupBox) Placement() ole.Variant {
	retVal, _ := this.PropGet(0x00000269, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupBox) SetPlacement(rhs interface{})  {
	_ = this.PropPut(0x00000269, []interface{}{rhs})
}

func (this *GroupBox) PrintObject() bool {
	retVal, _ := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupBox) SetPrintObject(rhs bool)  {
	_ = this.PropPut(0x0000026a, []interface{}{rhs})
}

var GroupBox_Select_OptArgs= []string{
	"Replace", 
}

func (this *GroupBox) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(GroupBox_Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000eb, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupBox) SendToBack() ole.Variant {
	retVal, _ := this.Call(0x0000025d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupBox) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *GroupBox) SetTop(rhs float64)  {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *GroupBox) TopLeftCell() *Range {
	retVal, _ := this.PropGet(0x0000026c, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *GroupBox) Visible() bool {
	retVal, _ := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupBox) SetVisible(rhs bool)  {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

func (this *GroupBox) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *GroupBox) SetWidth(rhs float64)  {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *GroupBox) ZOrder() int32 {
	retVal, _ := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *GroupBox) ShapeRange() *ShapeRange {
	retVal, _ := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.IDispatch(), false, true)
}

func (this *GroupBox) Caption() string {
	retVal, _ := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *GroupBox) SetCaption(rhs string)  {
	_ = this.PropPut(0x0000008b, []interface{}{rhs})
}

var GroupBox_Characters_OptArgs= []string{
	"Start", "Length", 
}

func (this *GroupBox) Characters(optArgs ...interface{}) *Characters {
	optArgs = ole.ProcessOptArgs(GroupBox_Characters_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000025b, nil, optArgs...)
	return NewCharacters(retVal.IDispatch(), false, true)
}

var GroupBox_CheckSpelling_OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang", 
}

func (this *GroupBox) CheckSpelling(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(GroupBox_CheckSpelling_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001f9, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupBox) LockedText() bool {
	retVal, _ := this.PropGet(0x00000268, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupBox) SetLockedText(rhs bool)  {
	_ = this.PropPut(0x00000268, []interface{}{rhs})
}

func (this *GroupBox) Text() string {
	retVal, _ := this.PropGet(0x0000008a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *GroupBox) SetText(rhs string)  {
	_ = this.PropPut(0x0000008a, []interface{}{rhs})
}

func (this *GroupBox) Accelerator() ole.Variant {
	retVal, _ := this.PropGet(0x0000034e, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupBox) SetAccelerator(rhs interface{})  {
	_ = this.PropPut(0x0000034e, []interface{}{rhs})
}

func (this *GroupBox) Display3DShading() bool {
	retVal, _ := this.PropGet(0x00000462, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupBox) SetDisplay3DShading(rhs bool)  {
	_ = this.PropPut(0x00000462, []interface{}{rhs})
}

func (this *GroupBox) PhoneticAccelerator() ole.Variant {
	retVal, _ := this.PropGet(0x00000461, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupBox) SetPhoneticAccelerator(rhs interface{})  {
	_ = this.PropPut(0x00000461, []interface{}{rhs})
}

