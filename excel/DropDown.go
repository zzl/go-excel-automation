package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002088B-0000-0000-C000-000000000046
var IID_DropDown = syscall.GUID{0x0002088B, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type DropDown struct {
	ole.OleClient
}

func NewDropDown(pDisp *win32.IDispatch, addRef bool, scoped bool) *DropDown {
	 if pDisp == nil {
		return nil;
	}
	p := &DropDown{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DropDownFromVar(v ole.Variant) *DropDown {
	return NewDropDown(v.IDispatch(), false, false)
}

func (this *DropDown) IID() *syscall.GUID {
	return &IID_DropDown
}

func (this *DropDown) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DropDown) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *DropDown) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *DropDown) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *DropDown) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *DropDown) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *DropDown) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *DropDown) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *DropDown) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *DropDown) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *DropDown) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DropDown) BottomRightCell() *Range {
	retVal, _ := this.PropGet(0x00000267, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *DropDown) BringToFront() ole.Variant {
	retVal, _ := this.Call(0x0000025a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DropDown) Copy() ole.Variant {
	retVal, _ := this.Call(0x00000227, nil)
	com.AddToScope(retVal)
	return *retVal
}

var DropDown_CopyPicture_OptArgs= []string{
	"Appearance", "Format", 
}

func (this *DropDown) CopyPicture(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DropDown_CopyPicture_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d5, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DropDown) Cut() ole.Variant {
	retVal, _ := this.Call(0x00000235, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DropDown) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DropDown) Duplicate() *ole.DispatchClass {
	retVal, _ := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DropDown) Enabled() bool {
	retVal, _ := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DropDown) SetEnabled(rhs bool)  {
	_ = this.PropPut(0x00000258, []interface{}{rhs})
}

func (this *DropDown) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *DropDown) SetHeight(rhs float64)  {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *DropDown) Index() int32 {
	retVal, _ := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *DropDown) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *DropDown) SetLeft(rhs float64)  {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *DropDown) Locked() bool {
	retVal, _ := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DropDown) SetLocked(rhs bool)  {
	_ = this.PropPut(0x0000010d, []interface{}{rhs})
}

func (this *DropDown) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DropDown) SetName(rhs string)  {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *DropDown) OnAction() string {
	retVal, _ := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DropDown) SetOnAction(rhs string)  {
	_ = this.PropPut(0x00000254, []interface{}{rhs})
}

func (this *DropDown) Placement() ole.Variant {
	retVal, _ := this.PropGet(0x00000269, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DropDown) SetPlacement(rhs interface{})  {
	_ = this.PropPut(0x00000269, []interface{}{rhs})
}

func (this *DropDown) PrintObject() bool {
	retVal, _ := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DropDown) SetPrintObject(rhs bool)  {
	_ = this.PropPut(0x0000026a, []interface{}{rhs})
}

var DropDown_Select_OptArgs= []string{
	"Replace", 
}

func (this *DropDown) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DropDown_Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000eb, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DropDown) SendToBack() ole.Variant {
	retVal, _ := this.Call(0x0000025d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DropDown) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *DropDown) SetTop(rhs float64)  {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *DropDown) TopLeftCell() *Range {
	retVal, _ := this.PropGet(0x0000026c, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *DropDown) Visible() bool {
	retVal, _ := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DropDown) SetVisible(rhs bool)  {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

func (this *DropDown) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *DropDown) SetWidth(rhs float64)  {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *DropDown) ZOrder() int32 {
	retVal, _ := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *DropDown) ShapeRange() *ShapeRange {
	retVal, _ := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.IDispatch(), false, true)
}

var DropDown_AddItem_OptArgs= []string{
	"Index", 
}

func (this *DropDown) AddItem(text interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DropDown_AddItem_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000353, []interface{}{text}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DropDown) Display3DShading() bool {
	retVal, _ := this.PropGet(0x00000462, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DropDown) SetDisplay3DShading(rhs bool)  {
	_ = this.PropPut(0x00000462, []interface{}{rhs})
}

func (this *DropDown) Default_() int32 {
	retVal, _ := this.PropGet(0x00000000, nil)
	return retVal.LValVal()
}

func (this *DropDown) SetDefault_(rhs int32)  {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *DropDown) LinkedCell() string {
	retVal, _ := this.PropGet(0x00000422, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DropDown) SetLinkedCell(rhs string)  {
	_ = this.PropPut(0x00000422, []interface{}{rhs})
}

func (this *DropDown) LinkedObject() ole.Variant {
	retVal, _ := this.PropGet(0x0000035e, nil)
	com.AddToScope(retVal)
	return *retVal
}

var DropDown_List_OptArgs= []string{
	"Index", 
}

func (this *DropDown) List(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DropDown_List_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000035d, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var DropDown_SetList_OptArgs= []string{
	"Index", 
}

func (this *DropDown) SetList(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(DropDown_SetList_OptArgs, optArgs)
	_ = this.PropPut(0x0000035d, nil, optArgs...)
}

func (this *DropDown) ListCount() int32 {
	retVal, _ := this.PropGet(0x00000351, nil)
	return retVal.LValVal()
}

func (this *DropDown) ListFillRange() string {
	retVal, _ := this.PropGet(0x0000034f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DropDown) SetListFillRange(rhs string)  {
	_ = this.PropPut(0x0000034f, []interface{}{rhs})
}

func (this *DropDown) ListIndex() int32 {
	retVal, _ := this.PropGet(0x00000352, nil)
	return retVal.LValVal()
}

func (this *DropDown) SetListIndex(rhs int32)  {
	_ = this.PropPut(0x00000352, []interface{}{rhs})
}

func (this *DropDown) Dummy36_()  {
	retVal, _ := this.Call(0x00010024, nil)
	_= retVal
}

func (this *DropDown) RemoveAllItems() ole.Variant {
	retVal, _ := this.Call(0x00000355, nil)
	com.AddToScope(retVal)
	return *retVal
}

var DropDown_RemoveItem_OptArgs= []string{
	"Count", 
}

func (this *DropDown) RemoveItem(index int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DropDown_RemoveItem_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000354, []interface{}{index}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var DropDown_Selected_OptArgs= []string{
	"Index", 
}

func (this *DropDown) Selected(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DropDown_Selected_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x00000463, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var DropDown_SetSelected_OptArgs= []string{
	"Index", 
}

func (this *DropDown) SetSelected(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(DropDown_SetSelected_OptArgs, optArgs)
	_ = this.PropPut(0x00000463, nil, optArgs...)
}

func (this *DropDown) Value() int32 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *DropDown) SetValue(rhs int32)  {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *DropDown) Caption() string {
	retVal, _ := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DropDown) SetCaption(rhs string)  {
	_ = this.PropPut(0x0000008b, []interface{}{rhs})
}

var DropDown_Characters_OptArgs= []string{
	"Start", "Length", 
}

func (this *DropDown) Characters(optArgs ...interface{}) *Characters {
	optArgs = ole.ProcessOptArgs(DropDown_Characters_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000025b, nil, optArgs...)
	return NewCharacters(retVal.IDispatch(), false, true)
}

func (this *DropDown) DropDownLines() int32 {
	retVal, _ := this.PropGet(0x00000350, nil)
	return retVal.LValVal()
}

func (this *DropDown) SetDropDownLines(rhs int32)  {
	_ = this.PropPut(0x00000350, []interface{}{rhs})
}

func (this *DropDown) Text() string {
	retVal, _ := this.PropGet(0x0000008a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DropDown) SetText(rhs string)  {
	_ = this.PropPut(0x0000008a, []interface{}{rhs})
}

