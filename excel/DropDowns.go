package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002088C-0000-0000-C000-000000000046
var IID_DropDowns = syscall.GUID{0x0002088C, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type DropDowns struct {
	ole.OleClient
}

func NewDropDowns(pDisp *win32.IDispatch, addRef bool, scoped bool) *DropDowns {
	 if pDisp == nil {
		return nil;
	}
	p := &DropDowns{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DropDownsFromVar(v ole.Variant) *DropDowns {
	return NewDropDowns(v.IDispatch(), false, false)
}

func (this *DropDowns) IID() *syscall.GUID {
	return &IID_DropDowns
}

func (this *DropDowns) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DropDowns) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *DropDowns) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *DropDowns) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *DropDowns) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *DropDowns) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *DropDowns) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *DropDowns) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *DropDowns) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *DropDowns) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *DropDowns) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DropDowns) Dummy3_()  {
	retVal, _ := this.Call(0x00010003, nil)
	_= retVal
}

func (this *DropDowns) BringToFront() ole.Variant {
	retVal, _ := this.Call(0x0000025a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DropDowns) Copy() ole.Variant {
	retVal, _ := this.Call(0x00000227, nil)
	com.AddToScope(retVal)
	return *retVal
}

var DropDowns_CopyPicture_OptArgs= []string{
	"Appearance", "Format", 
}

func (this *DropDowns) CopyPicture(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DropDowns_CopyPicture_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d5, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DropDowns) Cut() ole.Variant {
	retVal, _ := this.Call(0x00000235, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DropDowns) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DropDowns) Duplicate() *ole.DispatchClass {
	retVal, _ := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DropDowns) Enabled() bool {
	retVal, _ := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DropDowns) SetEnabled(rhs bool)  {
	_ = this.PropPut(0x00000258, []interface{}{rhs})
}

func (this *DropDowns) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *DropDowns) SetHeight(rhs float64)  {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *DropDowns) Dummy12_()  {
	retVal, _ := this.Call(0x0001000c, nil)
	_= retVal
}

func (this *DropDowns) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *DropDowns) SetLeft(rhs float64)  {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *DropDowns) Locked() bool {
	retVal, _ := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DropDowns) SetLocked(rhs bool)  {
	_ = this.PropPut(0x0000010d, []interface{}{rhs})
}

func (this *DropDowns) Dummy15_()  {
	retVal, _ := this.Call(0x0001000f, nil)
	_= retVal
}

func (this *DropDowns) OnAction() string {
	retVal, _ := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DropDowns) SetOnAction(rhs string)  {
	_ = this.PropPut(0x00000254, []interface{}{rhs})
}

func (this *DropDowns) Placement() ole.Variant {
	retVal, _ := this.PropGet(0x00000269, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DropDowns) SetPlacement(rhs interface{})  {
	_ = this.PropPut(0x00000269, []interface{}{rhs})
}

func (this *DropDowns) PrintObject() bool {
	retVal, _ := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DropDowns) SetPrintObject(rhs bool)  {
	_ = this.PropPut(0x0000026a, []interface{}{rhs})
}

var DropDowns_Select_OptArgs= []string{
	"Replace", 
}

func (this *DropDowns) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DropDowns_Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000eb, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DropDowns) SendToBack() ole.Variant {
	retVal, _ := this.Call(0x0000025d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DropDowns) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *DropDowns) SetTop(rhs float64)  {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *DropDowns) Dummy22_()  {
	retVal, _ := this.Call(0x00010016, nil)
	_= retVal
}

func (this *DropDowns) Visible() bool {
	retVal, _ := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DropDowns) SetVisible(rhs bool)  {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

func (this *DropDowns) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *DropDowns) SetWidth(rhs float64)  {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *DropDowns) ZOrder() int32 {
	retVal, _ := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *DropDowns) ShapeRange() *ShapeRange {
	retVal, _ := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.IDispatch(), false, true)
}

var DropDowns_AddItem_OptArgs= []string{
	"Index", 
}

func (this *DropDowns) AddItem(text interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DropDowns_AddItem_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000353, []interface{}{text}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DropDowns) Display3DShading() bool {
	retVal, _ := this.PropGet(0x00000462, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DropDowns) SetDisplay3DShading(rhs bool)  {
	_ = this.PropPut(0x00000462, []interface{}{rhs})
}

func (this *DropDowns) Default_() int32 {
	retVal, _ := this.PropGet(0x00000000, nil)
	return retVal.LValVal()
}

func (this *DropDowns) SetDefault_(rhs int32)  {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *DropDowns) LinkedCell() string {
	retVal, _ := this.PropGet(0x00000422, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DropDowns) SetLinkedCell(rhs string)  {
	_ = this.PropPut(0x00000422, []interface{}{rhs})
}

func (this *DropDowns) Dummy31_()  {
	retVal, _ := this.Call(0x0001001f, nil)
	_= retVal
}

var DropDowns_List_OptArgs= []string{
	"Index", 
}

func (this *DropDowns) List(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DropDowns_List_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000035d, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var DropDowns_SetList_OptArgs= []string{
	"Index", 
}

func (this *DropDowns) SetList(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(DropDowns_SetList_OptArgs, optArgs)
	_ = this.PropPut(0x0000035d, nil, optArgs...)
}

func (this *DropDowns) Dummy33_()  {
	retVal, _ := this.Call(0x00010021, nil)
	_= retVal
}

func (this *DropDowns) ListFillRange() string {
	retVal, _ := this.PropGet(0x0000034f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DropDowns) SetListFillRange(rhs string)  {
	_ = this.PropPut(0x0000034f, []interface{}{rhs})
}

func (this *DropDowns) ListIndex() int32 {
	retVal, _ := this.PropGet(0x00000352, nil)
	return retVal.LValVal()
}

func (this *DropDowns) SetListIndex(rhs int32)  {
	_ = this.PropPut(0x00000352, []interface{}{rhs})
}

func (this *DropDowns) Dummy36_()  {
	retVal, _ := this.Call(0x00010024, nil)
	_= retVal
}

func (this *DropDowns) RemoveAllItems() ole.Variant {
	retVal, _ := this.Call(0x00000355, nil)
	com.AddToScope(retVal)
	return *retVal
}

var DropDowns_RemoveItem_OptArgs= []string{
	"Count", 
}

func (this *DropDowns) RemoveItem(index int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DropDowns_RemoveItem_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000354, []interface{}{index}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var DropDowns_Selected_OptArgs= []string{
	"Index", 
}

func (this *DropDowns) Selected(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(DropDowns_Selected_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x00000463, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var DropDowns_SetSelected_OptArgs= []string{
	"Index", 
}

func (this *DropDowns) SetSelected(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(DropDowns_SetSelected_OptArgs, optArgs)
	_ = this.PropPut(0x00000463, nil, optArgs...)
}

func (this *DropDowns) Value() int32 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *DropDowns) SetValue(rhs int32)  {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *DropDowns) Caption() string {
	retVal, _ := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DropDowns) SetCaption(rhs string)  {
	_ = this.PropPut(0x0000008b, []interface{}{rhs})
}

var DropDowns_Characters_OptArgs= []string{
	"Start", "Length", 
}

func (this *DropDowns) Characters(optArgs ...interface{}) *Characters {
	optArgs = ole.ProcessOptArgs(DropDowns_Characters_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000025b, nil, optArgs...)
	return NewCharacters(retVal.IDispatch(), false, true)
}

func (this *DropDowns) DropDownLines() int32 {
	retVal, _ := this.PropGet(0x00000350, nil)
	return retVal.LValVal()
}

func (this *DropDowns) SetDropDownLines(rhs int32)  {
	_ = this.PropPut(0x00000350, []interface{}{rhs})
}

func (this *DropDowns) Text() string {
	retVal, _ := this.PropGet(0x0000008a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DropDowns) SetText(rhs string)  {
	_ = this.PropPut(0x0000008a, []interface{}{rhs})
}

var DropDowns_Add_OptArgs= []string{
	"Editable", 
}

func (this *DropDowns) Add(left float64, top float64, width float64, height float64, optArgs ...interface{}) *DropDown {
	optArgs = ole.ProcessOptArgs(DropDowns_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b5, []interface{}{left, top, width, height}, optArgs...)
	return NewDropDown(retVal.IDispatch(), false, true)
}

func (this *DropDowns) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *DropDowns) Group() *GroupObject {
	retVal, _ := this.Call(0x0000002e, nil)
	return NewGroupObject(retVal.IDispatch(), false, true)
}

func (this *DropDowns) Item(index interface{}) *ole.DispatchClass {
	retVal, _ := this.Call(0x000000aa, []interface{}{index})
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DropDowns) NewEnum_() *com.UnknownClass {
	retVal, _ := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *DropDowns) ForEach(action func(item int32) bool) {
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

