package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020899-0000-0000-C000-000000000046
var IID_GroupObjects = syscall.GUID{0x00020899, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type GroupObjects struct {
	ole.OleClient
}

func NewGroupObjects(pDisp *win32.IDispatch, addRef bool, scoped bool) *GroupObjects {
	 if pDisp == nil {
		return nil;
	}
	p := &GroupObjects{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func GroupObjectsFromVar(v ole.Variant) *GroupObjects {
	return NewGroupObjects(v.IDispatch(), false, false)
}

func (this *GroupObjects) IID() *syscall.GUID {
	return &IID_GroupObjects
}

func (this *GroupObjects) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *GroupObjects) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *GroupObjects) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *GroupObjects) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *GroupObjects) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *GroupObjects) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *GroupObjects) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *GroupObjects) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *GroupObjects) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *GroupObjects) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *GroupObjects) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *GroupObjects) Dummy3_()  {
	retVal, _ := this.Call(0x00010003, nil)
	_= retVal
}

func (this *GroupObjects) BringToFront() ole.Variant {
	retVal, _ := this.Call(0x0000025a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObjects) Copy() ole.Variant {
	retVal, _ := this.Call(0x00000227, nil)
	com.AddToScope(retVal)
	return *retVal
}

var GroupObjects_CopyPicture_OptArgs= []string{
	"Appearance", "Format", 
}

func (this *GroupObjects) CopyPicture(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(GroupObjects_CopyPicture_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d5, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObjects) Cut() ole.Variant {
	retVal, _ := this.Call(0x00000235, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObjects) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObjects) Duplicate() *ole.DispatchClass {
	retVal, _ := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *GroupObjects) Enabled() bool {
	retVal, _ := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObjects) SetEnabled(rhs bool)  {
	_ = this.PropPut(0x00000258, []interface{}{rhs})
}

func (this *GroupObjects) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *GroupObjects) SetHeight(rhs float64)  {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *GroupObjects) Dummy12_()  {
	retVal, _ := this.Call(0x0001000c, nil)
	_= retVal
}

func (this *GroupObjects) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *GroupObjects) SetLeft(rhs float64)  {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *GroupObjects) Locked() bool {
	retVal, _ := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObjects) SetLocked(rhs bool)  {
	_ = this.PropPut(0x0000010d, []interface{}{rhs})
}

func (this *GroupObjects) Dummy15_()  {
	retVal, _ := this.Call(0x0001000f, nil)
	_= retVal
}

func (this *GroupObjects) OnAction() string {
	retVal, _ := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *GroupObjects) SetOnAction(rhs string)  {
	_ = this.PropPut(0x00000254, []interface{}{rhs})
}

func (this *GroupObjects) Placement() ole.Variant {
	retVal, _ := this.PropGet(0x00000269, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObjects) SetPlacement(rhs interface{})  {
	_ = this.PropPut(0x00000269, []interface{}{rhs})
}

func (this *GroupObjects) PrintObject() bool {
	retVal, _ := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObjects) SetPrintObject(rhs bool)  {
	_ = this.PropPut(0x0000026a, []interface{}{rhs})
}

var GroupObjects_Select_OptArgs= []string{
	"Replace", 
}

func (this *GroupObjects) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(GroupObjects_Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000eb, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObjects) SendToBack() ole.Variant {
	retVal, _ := this.Call(0x0000025d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObjects) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *GroupObjects) SetTop(rhs float64)  {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *GroupObjects) Dummy22_()  {
	retVal, _ := this.Call(0x00010016, nil)
	_= retVal
}

func (this *GroupObjects) Visible() bool {
	retVal, _ := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObjects) SetVisible(rhs bool)  {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

func (this *GroupObjects) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *GroupObjects) SetWidth(rhs float64)  {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *GroupObjects) ZOrder() int32 {
	retVal, _ := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *GroupObjects) ShapeRange() *ShapeRange {
	retVal, _ := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.IDispatch(), false, true)
}

func (this *GroupObjects) Dummy27_()  {
	retVal, _ := this.Call(0x0001001b, nil)
	_= retVal
}

func (this *GroupObjects) Dummy28_()  {
	retVal, _ := this.Call(0x0001001c, nil)
	_= retVal
}

func (this *GroupObjects) AddIndent() bool {
	retVal, _ := this.PropGet(0x00000427, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObjects) SetAddIndent(rhs bool)  {
	_ = this.PropPut(0x00000427, []interface{}{rhs})
}

func (this *GroupObjects) Dummy30_()  {
	retVal, _ := this.Call(0x0001001e, nil)
	_= retVal
}

func (this *GroupObjects) ArrowHeadLength() ole.Variant {
	retVal, _ := this.PropGet(0x00000263, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObjects) SetArrowHeadLength(rhs interface{})  {
	_ = this.PropPut(0x00000263, []interface{}{rhs})
}

func (this *GroupObjects) ArrowHeadStyle() ole.Variant {
	retVal, _ := this.PropGet(0x00000264, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObjects) SetArrowHeadStyle(rhs interface{})  {
	_ = this.PropPut(0x00000264, []interface{}{rhs})
}

func (this *GroupObjects) ArrowHeadWidth() ole.Variant {
	retVal, _ := this.PropGet(0x00000265, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObjects) SetArrowHeadWidth(rhs interface{})  {
	_ = this.PropPut(0x00000265, []interface{}{rhs})
}

func (this *GroupObjects) AutoSize() bool {
	retVal, _ := this.PropGet(0x00000266, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObjects) SetAutoSize(rhs bool)  {
	_ = this.PropPut(0x00000266, []interface{}{rhs})
}

func (this *GroupObjects) Border() *Border {
	retVal, _ := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.IDispatch(), false, true)
}

func (this *GroupObjects) Dummy36_()  {
	retVal, _ := this.Call(0x00010024, nil)
	_= retVal
}

func (this *GroupObjects) Dummy37_()  {
	retVal, _ := this.Call(0x00010025, nil)
	_= retVal
}

func (this *GroupObjects) Dummy38_()  {
	retVal, _ := this.Call(0x00010026, nil)
	_= retVal
}

var GroupObjects_CheckSpelling_OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang", 
}

func (this *GroupObjects) CheckSpelling(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(GroupObjects_CheckSpelling_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001f9, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObjects) Default_() int32 {
	retVal, _ := this.PropGet(0x00000000, nil)
	return retVal.LValVal()
}

func (this *GroupObjects) SetDefault_(rhs int32)  {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *GroupObjects) Dummy41_()  {
	retVal, _ := this.Call(0x00010029, nil)
	_= retVal
}

func (this *GroupObjects) Dummy42_()  {
	retVal, _ := this.Call(0x0001002a, nil)
	_= retVal
}

func (this *GroupObjects) Dummy43_()  {
	retVal, _ := this.Call(0x0001002b, nil)
	_= retVal
}

func (this *GroupObjects) Dummy44_()  {
	retVal, _ := this.Call(0x0001002c, nil)
	_= retVal
}

func (this *GroupObjects) Dummy45_()  {
	retVal, _ := this.Call(0x0001002d, nil)
	_= retVal
}

func (this *GroupObjects) Font() *Font {
	retVal, _ := this.PropGet(0x00000092, nil)
	return NewFont(retVal.IDispatch(), false, true)
}

func (this *GroupObjects) Dummy47_()  {
	retVal, _ := this.Call(0x0001002f, nil)
	_= retVal
}

func (this *GroupObjects) Dummy48_()  {
	retVal, _ := this.Call(0x00010030, nil)
	_= retVal
}

func (this *GroupObjects) HorizontalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x00000088, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObjects) SetHorizontalAlignment(rhs interface{})  {
	_ = this.PropPut(0x00000088, []interface{}{rhs})
}

func (this *GroupObjects) Dummy50_()  {
	retVal, _ := this.Call(0x00010032, nil)
	_= retVal
}

func (this *GroupObjects) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *GroupObjects) Dummy52_()  {
	retVal, _ := this.Call(0x00010034, nil)
	_= retVal
}

func (this *GroupObjects) Dummy53_()  {
	retVal, _ := this.Call(0x00010035, nil)
	_= retVal
}

func (this *GroupObjects) Dummy54_()  {
	retVal, _ := this.Call(0x00010036, nil)
	_= retVal
}

func (this *GroupObjects) Dummy55_()  {
	retVal, _ := this.Call(0x00010037, nil)
	_= retVal
}

func (this *GroupObjects) Dummy56_()  {
	retVal, _ := this.Call(0x00010038, nil)
	_= retVal
}

func (this *GroupObjects) Dummy57_()  {
	retVal, _ := this.Call(0x00010039, nil)
	_= retVal
}

func (this *GroupObjects) Dummy58_()  {
	retVal, _ := this.Call(0x0001003a, nil)
	_= retVal
}

func (this *GroupObjects) Dummy59_()  {
	retVal, _ := this.Call(0x0001003b, nil)
	_= retVal
}

func (this *GroupObjects) Dummy60_()  {
	retVal, _ := this.Call(0x0001003c, nil)
	_= retVal
}

func (this *GroupObjects) Dummy61_()  {
	retVal, _ := this.Call(0x0001003d, nil)
	_= retVal
}

func (this *GroupObjects) Dummy62_()  {
	retVal, _ := this.Call(0x0001003e, nil)
	_= retVal
}

func (this *GroupObjects) Dummy63_()  {
	retVal, _ := this.Call(0x0001003f, nil)
	_= retVal
}

func (this *GroupObjects) Orientation() ole.Variant {
	retVal, _ := this.PropGet(0x00000086, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObjects) SetOrientation(rhs interface{})  {
	_ = this.PropPut(0x00000086, []interface{}{rhs})
}

func (this *GroupObjects) Dummy65_()  {
	retVal, _ := this.Call(0x00010041, nil)
	_= retVal
}

func (this *GroupObjects) Dummy66_()  {
	retVal, _ := this.Call(0x00010042, nil)
	_= retVal
}

func (this *GroupObjects) Dummy67_()  {
	retVal, _ := this.Call(0x00010043, nil)
	_= retVal
}

func (this *GroupObjects) Dummy68_()  {
	retVal, _ := this.Call(0x00010044, nil)
	_= retVal
}

func (this *GroupObjects) RoundedCorners() bool {
	retVal, _ := this.PropGet(0x0000026b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObjects) SetRoundedCorners(rhs bool)  {
	_ = this.PropPut(0x0000026b, []interface{}{rhs})
}

func (this *GroupObjects) Dummy70_()  {
	retVal, _ := this.Call(0x00010046, nil)
	_= retVal
}

func (this *GroupObjects) Shadow() bool {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObjects) SetShadow(rhs bool)  {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *GroupObjects) Dummy72_()  {
	retVal, _ := this.Call(0x00010048, nil)
	_= retVal
}

func (this *GroupObjects) Dummy73_()  {
	retVal, _ := this.Call(0x00010049, nil)
	_= retVal
}

func (this *GroupObjects) Ungroup() *ole.DispatchClass {
	retVal, _ := this.Call(0x000000f4, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *GroupObjects) Dummy75_()  {
	retVal, _ := this.Call(0x0001004b, nil)
	_= retVal
}

func (this *GroupObjects) VerticalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x00000089, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObjects) SetVerticalAlignment(rhs interface{})  {
	_ = this.PropPut(0x00000089, []interface{}{rhs})
}

func (this *GroupObjects) Dummy77_()  {
	retVal, _ := this.Call(0x0001004d, nil)
	_= retVal
}

func (this *GroupObjects) ReadingOrder() int32 {
	retVal, _ := this.PropGet(0x000003cf, nil)
	return retVal.LValVal()
}

func (this *GroupObjects) SetReadingOrder(rhs int32)  {
	_ = this.PropPut(0x000003cf, []interface{}{rhs})
}

func (this *GroupObjects) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *GroupObjects) Group() *GroupObject {
	retVal, _ := this.Call(0x0000002e, nil)
	return NewGroupObject(retVal.IDispatch(), false, true)
}

func (this *GroupObjects) Item(index interface{}) *ole.DispatchClass {
	retVal, _ := this.Call(0x000000aa, []interface{}{index})
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *GroupObjects) NewEnum_() *com.UnknownClass {
	retVal, _ := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *GroupObjects) ForEach(action func(item int32) bool) {
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

