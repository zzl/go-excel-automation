package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002089D-0000-0000-C000-000000000046
var IID_Rectangles = syscall.GUID{0x0002089D, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Rectangles struct {
	ole.OleClient
}

func NewRectangles(pDisp *win32.IDispatch, addRef bool, scoped bool) *Rectangles {
	p := &Rectangles{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func RectanglesFromVar(v ole.Variant) *Rectangles {
	return NewRectangles(v.PdispValVal(), false, false)
}

func (this *Rectangles) IID() *syscall.GUID {
	return &IID_Rectangles
}

func (this *Rectangles) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Rectangles) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Rectangles) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Rectangles) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Rectangles) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Rectangles) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Rectangles) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Rectangles) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Rectangles) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Rectangles) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Rectangles) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Rectangles) Dummy3_()  {
	retVal := this.Call(0x00010003, nil)
	_= retVal
}

func (this *Rectangles) BringToFront() ole.Variant {
	retVal := this.Call(0x0000025a, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangles) Copy() ole.Variant {
	retVal := this.Call(0x00000227, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangles) CopyPicture(appearance int32, format int32) ole.Variant {
	retVal := this.Call(0x000000d5, []interface{}{appearance, format})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangles) Cut() ole.Variant {
	retVal := this.Call(0x00000235, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangles) Delete() ole.Variant {
	retVal := this.Call(0x00000075, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangles) Duplicate() *ole.DispatchClass {
	retVal := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Rectangles) Enabled() bool {
	retVal := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Rectangles) SetEnabled(rhs bool)  {
	retVal := this.PropPut(0x00000258, []interface{}{rhs})
	_= retVal
}

func (this *Rectangles) Height() float64 {
	retVal := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *Rectangles) SetHeight(rhs float64)  {
	retVal := this.PropPut(0x0000007b, []interface{}{rhs})
	_= retVal
}

func (this *Rectangles) Dummy12_()  {
	retVal := this.Call(0x0001000c, nil)
	_= retVal
}

func (this *Rectangles) Left() float64 {
	retVal := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *Rectangles) SetLeft(rhs float64)  {
	retVal := this.PropPut(0x0000007f, []interface{}{rhs})
	_= retVal
}

func (this *Rectangles) Locked() bool {
	retVal := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Rectangles) SetLocked(rhs bool)  {
	retVal := this.PropPut(0x0000010d, []interface{}{rhs})
	_= retVal
}

func (this *Rectangles) Dummy15_()  {
	retVal := this.Call(0x0001000f, nil)
	_= retVal
}

func (this *Rectangles) OnAction() string {
	retVal := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Rectangles) SetOnAction(rhs string)  {
	retVal := this.PropPut(0x00000254, []interface{}{rhs})
	_= retVal
}

func (this *Rectangles) Placement() ole.Variant {
	retVal := this.PropGet(0x00000269, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangles) SetPlacement(rhs interface{})  {
	retVal := this.PropPut(0x00000269, []interface{}{rhs})
	_= retVal
}

func (this *Rectangles) PrintObject() bool {
	retVal := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Rectangles) SetPrintObject(rhs bool)  {
	retVal := this.PropPut(0x0000026a, []interface{}{rhs})
	_= retVal
}

var Rectangles_Select_OptArgs= []string{
	"Replace", 
}

func (this *Rectangles) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Rectangles_Select_OptArgs, optArgs)
	retVal := this.Call(0x000000eb, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangles) SendToBack() ole.Variant {
	retVal := this.Call(0x0000025d, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangles) Top() float64 {
	retVal := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *Rectangles) SetTop(rhs float64)  {
	retVal := this.PropPut(0x0000007e, []interface{}{rhs})
	_= retVal
}

func (this *Rectangles) Dummy22_()  {
	retVal := this.Call(0x00010016, nil)
	_= retVal
}

func (this *Rectangles) Visible() bool {
	retVal := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Rectangles) SetVisible(rhs bool)  {
	retVal := this.PropPut(0x0000022e, []interface{}{rhs})
	_= retVal
}

func (this *Rectangles) Width() float64 {
	retVal := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *Rectangles) SetWidth(rhs float64)  {
	retVal := this.PropPut(0x0000007a, []interface{}{rhs})
	_= retVal
}

func (this *Rectangles) ZOrder() int32 {
	retVal := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *Rectangles) ShapeRange() *ShapeRange {
	retVal := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.PdispValVal(), false, true)
}

func (this *Rectangles) AddIndent() bool {
	retVal := this.PropGet(0x00000427, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Rectangles) SetAddIndent(rhs bool)  {
	retVal := this.PropPut(0x00000427, []interface{}{rhs})
	_= retVal
}

func (this *Rectangles) AutoScaleFont() ole.Variant {
	retVal := this.PropGet(0x000005f5, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangles) SetAutoScaleFont(rhs interface{})  {
	retVal := this.PropPut(0x000005f5, []interface{}{rhs})
	_= retVal
}

func (this *Rectangles) AutoSize() bool {
	retVal := this.PropGet(0x00000266, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Rectangles) SetAutoSize(rhs bool)  {
	retVal := this.PropPut(0x00000266, []interface{}{rhs})
	_= retVal
}

func (this *Rectangles) Caption() string {
	retVal := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Rectangles) SetCaption(rhs string)  {
	retVal := this.PropPut(0x0000008b, []interface{}{rhs})
	_= retVal
}

var Rectangles_Characters_OptArgs= []string{
	"Start", "Length", 
}

func (this *Rectangles) Characters(optArgs ...interface{}) *Characters {
	optArgs = ole.ProcessOptArgs(Rectangles_Characters_OptArgs, optArgs)
	retVal := this.PropGet(0x0000025b, nil, optArgs...)
	return NewCharacters(retVal.PdispValVal(), false, true)
}

var Rectangles_CheckSpelling_OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang", 
}

func (this *Rectangles) CheckSpelling(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Rectangles_CheckSpelling_OptArgs, optArgs)
	retVal := this.Call(0x000001f9, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangles) Font() *Font {
	retVal := this.PropGet(0x00000092, nil)
	return NewFont(retVal.PdispValVal(), false, true)
}

func (this *Rectangles) Formula() string {
	retVal := this.PropGet(0x00000105, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Rectangles) SetFormula(rhs string)  {
	retVal := this.PropPut(0x00000105, []interface{}{rhs})
	_= retVal
}

func (this *Rectangles) HorizontalAlignment() ole.Variant {
	retVal := this.PropGet(0x00000088, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangles) SetHorizontalAlignment(rhs interface{})  {
	retVal := this.PropPut(0x00000088, []interface{}{rhs})
	_= retVal
}

func (this *Rectangles) LockedText() bool {
	retVal := this.PropGet(0x00000268, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Rectangles) SetLockedText(rhs bool)  {
	retVal := this.PropPut(0x00000268, []interface{}{rhs})
	_= retVal
}

func (this *Rectangles) Orientation() ole.Variant {
	retVal := this.PropGet(0x00000086, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangles) SetOrientation(rhs interface{})  {
	retVal := this.PropPut(0x00000086, []interface{}{rhs})
	_= retVal
}

func (this *Rectangles) Text() string {
	retVal := this.PropGet(0x0000008a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Rectangles) SetText(rhs string)  {
	retVal := this.PropPut(0x0000008a, []interface{}{rhs})
	_= retVal
}

func (this *Rectangles) VerticalAlignment() ole.Variant {
	retVal := this.PropGet(0x00000089, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Rectangles) SetVerticalAlignment(rhs interface{})  {
	retVal := this.PropPut(0x00000089, []interface{}{rhs})
	_= retVal
}

func (this *Rectangles) ReadingOrder() int32 {
	retVal := this.PropGet(0x000003cf, nil)
	return retVal.LValVal()
}

func (this *Rectangles) SetReadingOrder(rhs int32)  {
	retVal := this.PropPut(0x000003cf, []interface{}{rhs})
	_= retVal
}

func (this *Rectangles) Border() *Border {
	retVal := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.PdispValVal(), false, true)
}

func (this *Rectangles) Interior() *Interior {
	retVal := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.PdispValVal(), false, true)
}

func (this *Rectangles) Shadow() bool {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Rectangles) SetShadow(rhs bool)  {
	retVal := this.PropPut(0x00000067, []interface{}{rhs})
	_= retVal
}

func (this *Rectangles) RoundedCorners() bool {
	retVal := this.PropGet(0x0000026b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Rectangles) SetRoundedCorners(rhs bool)  {
	retVal := this.PropPut(0x0000026b, []interface{}{rhs})
	_= retVal
}

func (this *Rectangles) Add(left float64, top float64, width float64, height float64) *Rectangle {
	retVal := this.Call(0x000000b5, []interface{}{left, top, width, height})
	return NewRectangle(retVal.PdispValVal(), false, true)
}

func (this *Rectangles) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *Rectangles) Group() *GroupObject {
	retVal := this.Call(0x0000002e, nil)
	return NewGroupObject(retVal.PdispValVal(), false, true)
}

func (this *Rectangles) Item(index interface{}) *ole.DispatchClass {
	retVal := this.Call(0x000000aa, []interface{}{index})
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Rectangles) NewEnum_() *com.UnknownClass {
	retVal := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Rectangles) ForEach(action func(item *ole.DispatchClass) bool) {
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
		pItem := (*ole.DispatchClass)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

