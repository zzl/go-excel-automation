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
	return NewGroupObjects(v.PdispValVal(), false, false)
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
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *GroupObjects) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *GroupObjects) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *GroupObjects) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *GroupObjects) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *GroupObjects) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *GroupObjects) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *GroupObjects) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *GroupObjects) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *GroupObjects) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *GroupObjects) Dummy3_()  {
	retVal := this.Call(0x00010003, nil)
	_= retVal
}

func (this *GroupObjects) BringToFront() ole.Variant {
	retVal := this.Call(0x0000025a, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *GroupObjects) Copy() ole.Variant {
	retVal := this.Call(0x00000227, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *GroupObjects) CopyPicture(appearance int32, format int32) ole.Variant {
	retVal := this.Call(0x000000d5, []interface{}{appearance, format})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *GroupObjects) Cut() ole.Variant {
	retVal := this.Call(0x00000235, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *GroupObjects) Delete() ole.Variant {
	retVal := this.Call(0x00000075, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *GroupObjects) Duplicate() *ole.DispatchClass {
	retVal := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *GroupObjects) Enabled() bool {
	retVal := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObjects) SetEnabled(rhs bool)  {
	retVal := this.PropPut(0x00000258, []interface{}{rhs})
	_= retVal
}

func (this *GroupObjects) Height() float64 {
	retVal := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *GroupObjects) SetHeight(rhs float64)  {
	retVal := this.PropPut(0x0000007b, []interface{}{rhs})
	_= retVal
}

func (this *GroupObjects) Dummy12_()  {
	retVal := this.Call(0x0001000c, nil)
	_= retVal
}

func (this *GroupObjects) Left() float64 {
	retVal := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *GroupObjects) SetLeft(rhs float64)  {
	retVal := this.PropPut(0x0000007f, []interface{}{rhs})
	_= retVal
}

func (this *GroupObjects) Locked() bool {
	retVal := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObjects) SetLocked(rhs bool)  {
	retVal := this.PropPut(0x0000010d, []interface{}{rhs})
	_= retVal
}

func (this *GroupObjects) Dummy15_()  {
	retVal := this.Call(0x0001000f, nil)
	_= retVal
}

func (this *GroupObjects) OnAction() string {
	retVal := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *GroupObjects) SetOnAction(rhs string)  {
	retVal := this.PropPut(0x00000254, []interface{}{rhs})
	_= retVal
}

func (this *GroupObjects) Placement() ole.Variant {
	retVal := this.PropGet(0x00000269, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *GroupObjects) SetPlacement(rhs interface{})  {
	retVal := this.PropPut(0x00000269, []interface{}{rhs})
	_= retVal
}

func (this *GroupObjects) PrintObject() bool {
	retVal := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObjects) SetPrintObject(rhs bool)  {
	retVal := this.PropPut(0x0000026a, []interface{}{rhs})
	_= retVal
}

var GroupObjects_Select_OptArgs= []string{
	"Replace", 
}

func (this *GroupObjects) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(GroupObjects_Select_OptArgs, optArgs)
	retVal := this.Call(0x000000eb, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *GroupObjects) SendToBack() ole.Variant {
	retVal := this.Call(0x0000025d, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *GroupObjects) Top() float64 {
	retVal := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *GroupObjects) SetTop(rhs float64)  {
	retVal := this.PropPut(0x0000007e, []interface{}{rhs})
	_= retVal
}

func (this *GroupObjects) Dummy22_()  {
	retVal := this.Call(0x00010016, nil)
	_= retVal
}

func (this *GroupObjects) Visible() bool {
	retVal := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObjects) SetVisible(rhs bool)  {
	retVal := this.PropPut(0x0000022e, []interface{}{rhs})
	_= retVal
}

func (this *GroupObjects) Width() float64 {
	retVal := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *GroupObjects) SetWidth(rhs float64)  {
	retVal := this.PropPut(0x0000007a, []interface{}{rhs})
	_= retVal
}

func (this *GroupObjects) ZOrder() int32 {
	retVal := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *GroupObjects) ShapeRange() *ShapeRange {
	retVal := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.PdispValVal(), false, true)
}

func (this *GroupObjects) Dummy27_()  {
	retVal := this.Call(0x0001001b, nil)
	_= retVal
}

func (this *GroupObjects) Dummy28_()  {
	retVal := this.Call(0x0001001c, nil)
	_= retVal
}

func (this *GroupObjects) AddIndent() bool {
	retVal := this.PropGet(0x00000427, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObjects) SetAddIndent(rhs bool)  {
	retVal := this.PropPut(0x00000427, []interface{}{rhs})
	_= retVal
}

func (this *GroupObjects) Dummy30_()  {
	retVal := this.Call(0x0001001e, nil)
	_= retVal
}

func (this *GroupObjects) ArrowHeadLength() ole.Variant {
	retVal := this.PropGet(0x00000263, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *GroupObjects) SetArrowHeadLength(rhs interface{})  {
	retVal := this.PropPut(0x00000263, []interface{}{rhs})
	_= retVal
}

func (this *GroupObjects) ArrowHeadStyle() ole.Variant {
	retVal := this.PropGet(0x00000264, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *GroupObjects) SetArrowHeadStyle(rhs interface{})  {
	retVal := this.PropPut(0x00000264, []interface{}{rhs})
	_= retVal
}

func (this *GroupObjects) ArrowHeadWidth() ole.Variant {
	retVal := this.PropGet(0x00000265, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *GroupObjects) SetArrowHeadWidth(rhs interface{})  {
	retVal := this.PropPut(0x00000265, []interface{}{rhs})
	_= retVal
}

func (this *GroupObjects) AutoSize() bool {
	retVal := this.PropGet(0x00000266, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObjects) SetAutoSize(rhs bool)  {
	retVal := this.PropPut(0x00000266, []interface{}{rhs})
	_= retVal
}

func (this *GroupObjects) Border() *Border {
	retVal := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.PdispValVal(), false, true)
}

func (this *GroupObjects) Dummy36_()  {
	retVal := this.Call(0x00010024, nil)
	_= retVal
}

func (this *GroupObjects) Dummy37_()  {
	retVal := this.Call(0x00010025, nil)
	_= retVal
}

func (this *GroupObjects) Dummy38_()  {
	retVal := this.Call(0x00010026, nil)
	_= retVal
}

var GroupObjects_CheckSpelling_OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang", 
}

func (this *GroupObjects) CheckSpelling(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(GroupObjects_CheckSpelling_OptArgs, optArgs)
	retVal := this.Call(0x000001f9, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *GroupObjects) Default_() int32 {
	retVal := this.PropGet(0x00000000, nil)
	return retVal.LValVal()
}

func (this *GroupObjects) SetDefault_(rhs int32)  {
	retVal := this.PropPut(0x00000000, []interface{}{rhs})
	_= retVal
}

func (this *GroupObjects) Dummy41_()  {
	retVal := this.Call(0x00010029, nil)
	_= retVal
}

func (this *GroupObjects) Dummy42_()  {
	retVal := this.Call(0x0001002a, nil)
	_= retVal
}

func (this *GroupObjects) Dummy43_()  {
	retVal := this.Call(0x0001002b, nil)
	_= retVal
}

func (this *GroupObjects) Dummy44_()  {
	retVal := this.Call(0x0001002c, nil)
	_= retVal
}

func (this *GroupObjects) Dummy45_()  {
	retVal := this.Call(0x0001002d, nil)
	_= retVal
}

func (this *GroupObjects) Font() *Font {
	retVal := this.PropGet(0x00000092, nil)
	return NewFont(retVal.PdispValVal(), false, true)
}

func (this *GroupObjects) Dummy47_()  {
	retVal := this.Call(0x0001002f, nil)
	_= retVal
}

func (this *GroupObjects) Dummy48_()  {
	retVal := this.Call(0x00010030, nil)
	_= retVal
}

func (this *GroupObjects) HorizontalAlignment() ole.Variant {
	retVal := this.PropGet(0x00000088, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *GroupObjects) SetHorizontalAlignment(rhs interface{})  {
	retVal := this.PropPut(0x00000088, []interface{}{rhs})
	_= retVal
}

func (this *GroupObjects) Dummy50_()  {
	retVal := this.Call(0x00010032, nil)
	_= retVal
}

func (this *GroupObjects) Interior() *Interior {
	retVal := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.PdispValVal(), false, true)
}

func (this *GroupObjects) Dummy52_()  {
	retVal := this.Call(0x00010034, nil)
	_= retVal
}

func (this *GroupObjects) Dummy53_()  {
	retVal := this.Call(0x00010035, nil)
	_= retVal
}

func (this *GroupObjects) Dummy54_()  {
	retVal := this.Call(0x00010036, nil)
	_= retVal
}

func (this *GroupObjects) Dummy55_()  {
	retVal := this.Call(0x00010037, nil)
	_= retVal
}

func (this *GroupObjects) Dummy56_()  {
	retVal := this.Call(0x00010038, nil)
	_= retVal
}

func (this *GroupObjects) Dummy57_()  {
	retVal := this.Call(0x00010039, nil)
	_= retVal
}

func (this *GroupObjects) Dummy58_()  {
	retVal := this.Call(0x0001003a, nil)
	_= retVal
}

func (this *GroupObjects) Dummy59_()  {
	retVal := this.Call(0x0001003b, nil)
	_= retVal
}

func (this *GroupObjects) Dummy60_()  {
	retVal := this.Call(0x0001003c, nil)
	_= retVal
}

func (this *GroupObjects) Dummy61_()  {
	retVal := this.Call(0x0001003d, nil)
	_= retVal
}

func (this *GroupObjects) Dummy62_()  {
	retVal := this.Call(0x0001003e, nil)
	_= retVal
}

func (this *GroupObjects) Dummy63_()  {
	retVal := this.Call(0x0001003f, nil)
	_= retVal
}

func (this *GroupObjects) Orientation() ole.Variant {
	retVal := this.PropGet(0x00000086, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *GroupObjects) SetOrientation(rhs interface{})  {
	retVal := this.PropPut(0x00000086, []interface{}{rhs})
	_= retVal
}

func (this *GroupObjects) Dummy65_()  {
	retVal := this.Call(0x00010041, nil)
	_= retVal
}

func (this *GroupObjects) Dummy66_()  {
	retVal := this.Call(0x00010042, nil)
	_= retVal
}

func (this *GroupObjects) Dummy67_()  {
	retVal := this.Call(0x00010043, nil)
	_= retVal
}

func (this *GroupObjects) Dummy68_()  {
	retVal := this.Call(0x00010044, nil)
	_= retVal
}

func (this *GroupObjects) RoundedCorners() bool {
	retVal := this.PropGet(0x0000026b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObjects) SetRoundedCorners(rhs bool)  {
	retVal := this.PropPut(0x0000026b, []interface{}{rhs})
	_= retVal
}

func (this *GroupObjects) Dummy70_()  {
	retVal := this.Call(0x00010046, nil)
	_= retVal
}

func (this *GroupObjects) Shadow() bool {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObjects) SetShadow(rhs bool)  {
	retVal := this.PropPut(0x00000067, []interface{}{rhs})
	_= retVal
}

func (this *GroupObjects) Dummy72_()  {
	retVal := this.Call(0x00010048, nil)
	_= retVal
}

func (this *GroupObjects) Dummy73_()  {
	retVal := this.Call(0x00010049, nil)
	_= retVal
}

func (this *GroupObjects) Ungroup() *ole.DispatchClass {
	retVal := this.Call(0x000000f4, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *GroupObjects) Dummy75_()  {
	retVal := this.Call(0x0001004b, nil)
	_= retVal
}

func (this *GroupObjects) VerticalAlignment() ole.Variant {
	retVal := this.PropGet(0x00000089, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *GroupObjects) SetVerticalAlignment(rhs interface{})  {
	retVal := this.PropPut(0x00000089, []interface{}{rhs})
	_= retVal
}

func (this *GroupObjects) Dummy77_()  {
	retVal := this.Call(0x0001004d, nil)
	_= retVal
}

func (this *GroupObjects) ReadingOrder() int32 {
	retVal := this.PropGet(0x000003cf, nil)
	return retVal.LValVal()
}

func (this *GroupObjects) SetReadingOrder(rhs int32)  {
	retVal := this.PropPut(0x000003cf, []interface{}{rhs})
	_= retVal
}

func (this *GroupObjects) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *GroupObjects) Group() *GroupObject {
	retVal := this.Call(0x0000002e, nil)
	return NewGroupObject(retVal.PdispValVal(), false, true)
}

func (this *GroupObjects) Item(index interface{}) *ole.DispatchClass {
	retVal := this.Call(0x000000aa, []interface{}{index})
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *GroupObjects) NewEnum_() *com.UnknownClass {
	retVal := this.Call(-4, nil)
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

