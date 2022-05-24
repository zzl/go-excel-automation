package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020887-0000-0000-C000-000000000046
var IID_ListBox = syscall.GUID{0x00020887, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ListBox struct {
	ole.OleClient
}

func NewListBox(pDisp *win32.IDispatch, addRef bool, scoped bool) *ListBox {
	 if pDisp == nil {
		return nil;
	}
	p := &ListBox{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ListBoxFromVar(v ole.Variant) *ListBox {
	return NewListBox(v.IDispatch(), false, false)
}

func (this *ListBox) IID() *syscall.GUID {
	return &IID_ListBox
}

func (this *ListBox) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ListBox) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *ListBox) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ListBox) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ListBox) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *ListBox) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *ListBox) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *ListBox) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *ListBox) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ListBox) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ListBox) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ListBox) BottomRightCell() *Range {
	retVal, _ := this.PropGet(0x00000267, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *ListBox) BringToFront() ole.Variant {
	retVal, _ := this.Call(0x0000025a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ListBox) Copy() ole.Variant {
	retVal, _ := this.Call(0x00000227, nil)
	com.AddToScope(retVal)
	return *retVal
}

var ListBox_CopyPicture_OptArgs= []string{
	"Appearance", "Format", 
}

func (this *ListBox) CopyPicture(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(ListBox_CopyPicture_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d5, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ListBox) Cut() ole.Variant {
	retVal, _ := this.Call(0x00000235, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ListBox) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ListBox) Duplicate() *ole.DispatchClass {
	retVal, _ := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ListBox) Enabled() bool {
	retVal, _ := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListBox) SetEnabled(rhs bool)  {
	_ = this.PropPut(0x00000258, []interface{}{rhs})
}

func (this *ListBox) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *ListBox) SetHeight(rhs float64)  {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *ListBox) Index() int32 {
	retVal, _ := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *ListBox) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *ListBox) SetLeft(rhs float64)  {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *ListBox) Locked() bool {
	retVal, _ := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListBox) SetLocked(rhs bool)  {
	_ = this.PropPut(0x0000010d, []interface{}{rhs})
}

func (this *ListBox) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ListBox) SetName(rhs string)  {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *ListBox) OnAction() string {
	retVal, _ := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ListBox) SetOnAction(rhs string)  {
	_ = this.PropPut(0x00000254, []interface{}{rhs})
}

func (this *ListBox) Placement() ole.Variant {
	retVal, _ := this.PropGet(0x00000269, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ListBox) SetPlacement(rhs interface{})  {
	_ = this.PropPut(0x00000269, []interface{}{rhs})
}

func (this *ListBox) PrintObject() bool {
	retVal, _ := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListBox) SetPrintObject(rhs bool)  {
	_ = this.PropPut(0x0000026a, []interface{}{rhs})
}

var ListBox_Select_OptArgs= []string{
	"Replace", 
}

func (this *ListBox) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(ListBox_Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000eb, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ListBox) SendToBack() ole.Variant {
	retVal, _ := this.Call(0x0000025d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ListBox) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *ListBox) SetTop(rhs float64)  {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *ListBox) TopLeftCell() *Range {
	retVal, _ := this.PropGet(0x0000026c, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *ListBox) Visible() bool {
	retVal, _ := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListBox) SetVisible(rhs bool)  {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

func (this *ListBox) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *ListBox) SetWidth(rhs float64)  {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *ListBox) ZOrder() int32 {
	retVal, _ := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *ListBox) ShapeRange() *ShapeRange {
	retVal, _ := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.IDispatch(), false, true)
}

var ListBox_AddItem_OptArgs= []string{
	"Index", 
}

func (this *ListBox) AddItem(text interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(ListBox_AddItem_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000353, []interface{}{text}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ListBox) Display3DShading() bool {
	retVal, _ := this.PropGet(0x00000462, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListBox) SetDisplay3DShading(rhs bool)  {
	_ = this.PropPut(0x00000462, []interface{}{rhs})
}

func (this *ListBox) Default_() int32 {
	retVal, _ := this.PropGet(0x00000000, nil)
	return retVal.LValVal()
}

func (this *ListBox) SetDefault_(rhs int32)  {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *ListBox) LinkedCell() string {
	retVal, _ := this.PropGet(0x00000422, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ListBox) SetLinkedCell(rhs string)  {
	_ = this.PropPut(0x00000422, []interface{}{rhs})
}

func (this *ListBox) LinkedObject() ole.Variant {
	retVal, _ := this.PropGet(0x0000035e, nil)
	com.AddToScope(retVal)
	return *retVal
}

var ListBox_List_OptArgs= []string{
	"Index", 
}

func (this *ListBox) List(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(ListBox_List_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000035d, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var ListBox_SetList_OptArgs= []string{
	"Index", 
}

func (this *ListBox) SetList(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ListBox_SetList_OptArgs, optArgs)
	_ = this.PropPut(0x0000035d, nil, optArgs...)
}

func (this *ListBox) ListCount() int32 {
	retVal, _ := this.PropGet(0x00000351, nil)
	return retVal.LValVal()
}

func (this *ListBox) ListFillRange() string {
	retVal, _ := this.PropGet(0x0000034f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ListBox) SetListFillRange(rhs string)  {
	_ = this.PropPut(0x0000034f, []interface{}{rhs})
}

func (this *ListBox) ListIndex() int32 {
	retVal, _ := this.PropGet(0x00000352, nil)
	return retVal.LValVal()
}

func (this *ListBox) SetListIndex(rhs int32)  {
	_ = this.PropPut(0x00000352, []interface{}{rhs})
}

func (this *ListBox) MultiSelect() int32 {
	retVal, _ := this.PropGet(0x00000020, nil)
	return retVal.LValVal()
}

func (this *ListBox) SetMultiSelect(rhs int32)  {
	_ = this.PropPut(0x00000020, []interface{}{rhs})
}

func (this *ListBox) RemoveAllItems() ole.Variant {
	retVal, _ := this.Call(0x00000355, nil)
	com.AddToScope(retVal)
	return *retVal
}

var ListBox_RemoveItem_OptArgs= []string{
	"Count", 
}

func (this *ListBox) RemoveItem(index int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(ListBox_RemoveItem_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000354, []interface{}{index}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var ListBox_Selected_OptArgs= []string{
	"Index", 
}

func (this *ListBox) Selected(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(ListBox_Selected_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x00000463, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var ListBox_SetSelected_OptArgs= []string{
	"Index", 
}

func (this *ListBox) SetSelected(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ListBox_SetSelected_OptArgs, optArgs)
	_ = this.PropPut(0x00000463, nil, optArgs...)
}

func (this *ListBox) Value() int32 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *ListBox) SetValue(rhs int32)  {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

