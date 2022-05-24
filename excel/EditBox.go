package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020883-0000-0000-C000-000000000046
var IID_EditBox = syscall.GUID{0x00020883, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type EditBox struct {
	ole.OleClient
}

func NewEditBox(pDisp *win32.IDispatch, addRef bool, scoped bool) *EditBox {
	 if pDisp == nil {
		return nil;
	}
	p := &EditBox{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func EditBoxFromVar(v ole.Variant) *EditBox {
	return NewEditBox(v.IDispatch(), false, false)
}

func (this *EditBox) IID() *syscall.GUID {
	return &IID_EditBox
}

func (this *EditBox) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *EditBox) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *EditBox) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *EditBox) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *EditBox) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *EditBox) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *EditBox) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *EditBox) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *EditBox) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *EditBox) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *EditBox) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *EditBox) BottomRightCell() *Range {
	retVal, _ := this.PropGet(0x00000267, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *EditBox) BringToFront() ole.Variant {
	retVal, _ := this.Call(0x0000025a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *EditBox) Copy() ole.Variant {
	retVal, _ := this.Call(0x00000227, nil)
	com.AddToScope(retVal)
	return *retVal
}

var EditBox_CopyPicture_OptArgs= []string{
	"Appearance", "Format", 
}

func (this *EditBox) CopyPicture(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(EditBox_CopyPicture_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d5, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *EditBox) Cut() ole.Variant {
	retVal, _ := this.Call(0x00000235, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *EditBox) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *EditBox) Duplicate() *ole.DispatchClass {
	retVal, _ := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *EditBox) Enabled() bool {
	retVal, _ := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EditBox) SetEnabled(rhs bool)  {
	_ = this.PropPut(0x00000258, []interface{}{rhs})
}

func (this *EditBox) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *EditBox) SetHeight(rhs float64)  {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *EditBox) Index() int32 {
	retVal, _ := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *EditBox) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *EditBox) SetLeft(rhs float64)  {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *EditBox) Locked() bool {
	retVal, _ := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EditBox) SetLocked(rhs bool)  {
	_ = this.PropPut(0x0000010d, []interface{}{rhs})
}

func (this *EditBox) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *EditBox) SetName(rhs string)  {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *EditBox) OnAction() string {
	retVal, _ := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *EditBox) SetOnAction(rhs string)  {
	_ = this.PropPut(0x00000254, []interface{}{rhs})
}

func (this *EditBox) Placement() ole.Variant {
	retVal, _ := this.PropGet(0x00000269, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *EditBox) SetPlacement(rhs interface{})  {
	_ = this.PropPut(0x00000269, []interface{}{rhs})
}

func (this *EditBox) PrintObject() bool {
	retVal, _ := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EditBox) SetPrintObject(rhs bool)  {
	_ = this.PropPut(0x0000026a, []interface{}{rhs})
}

var EditBox_Select_OptArgs= []string{
	"Replace", 
}

func (this *EditBox) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(EditBox_Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000eb, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *EditBox) SendToBack() ole.Variant {
	retVal, _ := this.Call(0x0000025d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *EditBox) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *EditBox) SetTop(rhs float64)  {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *EditBox) TopLeftCell() *Range {
	retVal, _ := this.PropGet(0x0000026c, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *EditBox) Visible() bool {
	retVal, _ := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EditBox) SetVisible(rhs bool)  {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

func (this *EditBox) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *EditBox) SetWidth(rhs float64)  {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *EditBox) ZOrder() int32 {
	retVal, _ := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *EditBox) ShapeRange() *ShapeRange {
	retVal, _ := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.IDispatch(), false, true)
}

func (this *EditBox) Caption() string {
	retVal, _ := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *EditBox) SetCaption(rhs string)  {
	_ = this.PropPut(0x0000008b, []interface{}{rhs})
}

var EditBox_Characters_OptArgs= []string{
	"Start", "Length", 
}

func (this *EditBox) Characters(optArgs ...interface{}) *Characters {
	optArgs = ole.ProcessOptArgs(EditBox_Characters_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000025b, nil, optArgs...)
	return NewCharacters(retVal.IDispatch(), false, true)
}

var EditBox_CheckSpelling_OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang", 
}

func (this *EditBox) CheckSpelling(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(EditBox_CheckSpelling_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001f9, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *EditBox) LockedText() bool {
	retVal, _ := this.PropGet(0x00000268, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EditBox) SetLockedText(rhs bool)  {
	_ = this.PropPut(0x00000268, []interface{}{rhs})
}

func (this *EditBox) Text() string {
	retVal, _ := this.PropGet(0x0000008a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *EditBox) SetText(rhs string)  {
	_ = this.PropPut(0x0000008a, []interface{}{rhs})
}

func (this *EditBox) DisplayVerticalScrollBar() bool {
	retVal, _ := this.PropGet(0x0000039a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EditBox) SetDisplayVerticalScrollBar(rhs bool)  {
	_ = this.PropPut(0x0000039a, []interface{}{rhs})
}

func (this *EditBox) InputType() int32 {
	retVal, _ := this.PropGet(0x00000356, nil)
	return retVal.LValVal()
}

func (this *EditBox) SetInputType(rhs int32)  {
	_ = this.PropPut(0x00000356, []interface{}{rhs})
}

func (this *EditBox) LinkedObject() string {
	retVal, _ := this.PropGet(0x0000035e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *EditBox) MultiLine() bool {
	retVal, _ := this.PropGet(0x00000357, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EditBox) SetMultiLine(rhs bool)  {
	_ = this.PropPut(0x00000357, []interface{}{rhs})
}

func (this *EditBox) PasswordEdit() bool {
	retVal, _ := this.PropGet(0x0000048a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EditBox) SetPasswordEdit(rhs bool)  {
	_ = this.PropPut(0x0000048a, []interface{}{rhs})
}

