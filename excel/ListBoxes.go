package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00020888-0000-0000-C000-000000000046
var IID_ListBoxes = syscall.GUID{0x00020888, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ListBoxes struct {
	ole.OleClient
}

func NewListBoxes(pDisp *win32.IDispatch, addRef bool, scoped bool) *ListBoxes {
	if pDisp == nil {
		return nil
	}
	p := &ListBoxes{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ListBoxesFromVar(v ole.Variant) *ListBoxes {
	return NewListBoxes(v.IDispatch(), false, false)
}

func (this *ListBoxes) IID() *syscall.GUID {
	return &IID_ListBoxes
}

func (this *ListBoxes) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ListBoxes) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *ListBoxes) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ListBoxes) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ListBoxes) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *ListBoxes) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *ListBoxes) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *ListBoxes) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *ListBoxes) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ListBoxes) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ListBoxes) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ListBoxes) Dummy3_() {
	retVal, _ := this.Call(0x00010003, nil)
	_ = retVal
}

func (this *ListBoxes) BringToFront() ole.Variant {
	retVal, _ := this.Call(0x0000025a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ListBoxes) Copy() ole.Variant {
	retVal, _ := this.Call(0x00000227, nil)
	com.AddToScope(retVal)
	return *retVal
}

var ListBoxes_CopyPicture_OptArgs = []string{
	"Appearance", "Format",
}

func (this *ListBoxes) CopyPicture(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(ListBoxes_CopyPicture_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d5, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ListBoxes) Cut() ole.Variant {
	retVal, _ := this.Call(0x00000235, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ListBoxes) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ListBoxes) Duplicate() *ole.DispatchClass {
	retVal, _ := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ListBoxes) Enabled() bool {
	retVal, _ := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListBoxes) SetEnabled(rhs bool) {
	_ = this.PropPut(0x00000258, []interface{}{rhs})
}

func (this *ListBoxes) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *ListBoxes) SetHeight(rhs float64) {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *ListBoxes) Dummy12_() {
	retVal, _ := this.Call(0x0001000c, nil)
	_ = retVal
}

func (this *ListBoxes) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *ListBoxes) SetLeft(rhs float64) {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *ListBoxes) Locked() bool {
	retVal, _ := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListBoxes) SetLocked(rhs bool) {
	_ = this.PropPut(0x0000010d, []interface{}{rhs})
}

func (this *ListBoxes) Dummy15_() {
	retVal, _ := this.Call(0x0001000f, nil)
	_ = retVal
}

func (this *ListBoxes) OnAction() string {
	retVal, _ := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ListBoxes) SetOnAction(rhs string) {
	_ = this.PropPut(0x00000254, []interface{}{rhs})
}

func (this *ListBoxes) Placement() ole.Variant {
	retVal, _ := this.PropGet(0x00000269, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ListBoxes) SetPlacement(rhs interface{}) {
	_ = this.PropPut(0x00000269, []interface{}{rhs})
}

func (this *ListBoxes) PrintObject() bool {
	retVal, _ := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListBoxes) SetPrintObject(rhs bool) {
	_ = this.PropPut(0x0000026a, []interface{}{rhs})
}

var ListBoxes_Select_OptArgs = []string{
	"Replace",
}

func (this *ListBoxes) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(ListBoxes_Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000eb, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ListBoxes) SendToBack() ole.Variant {
	retVal, _ := this.Call(0x0000025d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ListBoxes) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *ListBoxes) SetTop(rhs float64) {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *ListBoxes) Dummy22_() {
	retVal, _ := this.Call(0x00010016, nil)
	_ = retVal
}

func (this *ListBoxes) Visible() bool {
	retVal, _ := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListBoxes) SetVisible(rhs bool) {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

func (this *ListBoxes) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *ListBoxes) SetWidth(rhs float64) {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *ListBoxes) ZOrder() int32 {
	retVal, _ := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *ListBoxes) ShapeRange() *ShapeRange {
	retVal, _ := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.IDispatch(), false, true)
}

var ListBoxes_AddItem_OptArgs = []string{
	"Index",
}

func (this *ListBoxes) AddItem(text interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(ListBoxes_AddItem_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000353, []interface{}{text}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ListBoxes) Display3DShading() bool {
	retVal, _ := this.PropGet(0x00000462, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListBoxes) SetDisplay3DShading(rhs bool) {
	_ = this.PropPut(0x00000462, []interface{}{rhs})
}

func (this *ListBoxes) Default_() int32 {
	retVal, _ := this.PropGet(0x00000000, nil)
	return retVal.LValVal()
}

func (this *ListBoxes) SetDefault_(rhs int32) {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *ListBoxes) LinkedCell() string {
	retVal, _ := this.PropGet(0x00000422, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ListBoxes) SetLinkedCell(rhs string) {
	_ = this.PropPut(0x00000422, []interface{}{rhs})
}

func (this *ListBoxes) Dummy31_() {
	retVal, _ := this.Call(0x0001001f, nil)
	_ = retVal
}

var ListBoxes_List_OptArgs = []string{
	"Index",
}

func (this *ListBoxes) List(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(ListBoxes_List_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000035d, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var ListBoxes_SetList_OptArgs = []string{
	"Index",
}

func (this *ListBoxes) SetList(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(ListBoxes_SetList_OptArgs, optArgs)
	_ = this.PropPut(0x0000035d, nil, optArgs...)
}

func (this *ListBoxes) Dummy33_() {
	retVal, _ := this.Call(0x00010021, nil)
	_ = retVal
}

func (this *ListBoxes) ListFillRange() string {
	retVal, _ := this.PropGet(0x0000034f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ListBoxes) SetListFillRange(rhs string) {
	_ = this.PropPut(0x0000034f, []interface{}{rhs})
}

func (this *ListBoxes) ListIndex() int32 {
	retVal, _ := this.PropGet(0x00000352, nil)
	return retVal.LValVal()
}

func (this *ListBoxes) SetListIndex(rhs int32) {
	_ = this.PropPut(0x00000352, []interface{}{rhs})
}

func (this *ListBoxes) MultiSelect() int32 {
	retVal, _ := this.PropGet(0x00000020, nil)
	return retVal.LValVal()
}

func (this *ListBoxes) SetMultiSelect(rhs int32) {
	_ = this.PropPut(0x00000020, []interface{}{rhs})
}

func (this *ListBoxes) RemoveAllItems() ole.Variant {
	retVal, _ := this.Call(0x00000355, nil)
	com.AddToScope(retVal)
	return *retVal
}

var ListBoxes_RemoveItem_OptArgs = []string{
	"Count",
}

func (this *ListBoxes) RemoveItem(index int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(ListBoxes_RemoveItem_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000354, []interface{}{index}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var ListBoxes_Selected_OptArgs = []string{
	"Index",
}

func (this *ListBoxes) Selected(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(ListBoxes_Selected_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x00000463, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var ListBoxes_SetSelected_OptArgs = []string{
	"Index",
}

func (this *ListBoxes) SetSelected(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(ListBoxes_SetSelected_OptArgs, optArgs)
	_ = this.PropPut(0x00000463, nil, optArgs...)
}

func (this *ListBoxes) Value() int32 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *ListBoxes) SetValue(rhs int32) {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *ListBoxes) Add(left float64, top float64, width float64, height float64) *ListBox {
	retVal, _ := this.Call(0x000000b5, []interface{}{left, top, width, height})
	return NewListBox(retVal.IDispatch(), false, true)
}

func (this *ListBoxes) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *ListBoxes) Group() *GroupObject {
	retVal, _ := this.Call(0x0000002e, nil)
	return NewGroupObject(retVal.IDispatch(), false, true)
}

func (this *ListBoxes) Item(index interface{}) *ole.DispatchClass {
	retVal, _ := this.Call(0x000000aa, []interface{}{index})
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ListBoxes) NewEnum_() *com.UnknownClass {
	retVal, _ := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *ListBoxes) ForEach(action func(item int32) bool) {
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
