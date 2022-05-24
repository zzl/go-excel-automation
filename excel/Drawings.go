package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000208A9-0000-0000-C000-000000000046
var IID_Drawings = syscall.GUID{0x000208A9, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Drawings struct {
	ole.OleClient
}

func NewDrawings(pDisp *win32.IDispatch, addRef bool, scoped bool) *Drawings {
	 if pDisp == nil {
		return nil;
	}
	p := &Drawings{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DrawingsFromVar(v ole.Variant) *Drawings {
	return NewDrawings(v.IDispatch(), false, false)
}

func (this *Drawings) IID() *syscall.GUID {
	return &IID_Drawings
}

func (this *Drawings) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Drawings) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Drawings) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Drawings) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Drawings) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Drawings) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Drawings) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Drawings) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Drawings) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Drawings) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Drawings) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Drawings) Dummy3_()  {
	retVal, _ := this.Call(0x00010003, nil)
	_= retVal
}

func (this *Drawings) BringToFront() ole.Variant {
	retVal, _ := this.Call(0x0000025a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Drawings) Copy() ole.Variant {
	retVal, _ := this.Call(0x00000227, nil)
	com.AddToScope(retVal)
	return *retVal
}

var Drawings_CopyPicture_OptArgs= []string{
	"Appearance", "Format", 
}

func (this *Drawings) CopyPicture(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Drawings_CopyPicture_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d5, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Drawings) Cut() ole.Variant {
	retVal, _ := this.Call(0x00000235, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Drawings) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Drawings) Duplicate() *ole.DispatchClass {
	retVal, _ := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Drawings) Enabled() bool {
	retVal, _ := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Drawings) SetEnabled(rhs bool)  {
	_ = this.PropPut(0x00000258, []interface{}{rhs})
}

func (this *Drawings) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *Drawings) SetHeight(rhs float64)  {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *Drawings) Dummy12_()  {
	retVal, _ := this.Call(0x0001000c, nil)
	_= retVal
}

func (this *Drawings) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *Drawings) SetLeft(rhs float64)  {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *Drawings) Locked() bool {
	retVal, _ := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Drawings) SetLocked(rhs bool)  {
	_ = this.PropPut(0x0000010d, []interface{}{rhs})
}

func (this *Drawings) Dummy15_()  {
	retVal, _ := this.Call(0x0001000f, nil)
	_= retVal
}

func (this *Drawings) OnAction() string {
	retVal, _ := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Drawings) SetOnAction(rhs string)  {
	_ = this.PropPut(0x00000254, []interface{}{rhs})
}

func (this *Drawings) Placement() ole.Variant {
	retVal, _ := this.PropGet(0x00000269, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Drawings) SetPlacement(rhs interface{})  {
	_ = this.PropPut(0x00000269, []interface{}{rhs})
}

func (this *Drawings) PrintObject() bool {
	retVal, _ := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Drawings) SetPrintObject(rhs bool)  {
	_ = this.PropPut(0x0000026a, []interface{}{rhs})
}

var Drawings_Select_OptArgs= []string{
	"Replace", 
}

func (this *Drawings) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Drawings_Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000eb, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Drawings) SendToBack() ole.Variant {
	retVal, _ := this.Call(0x0000025d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Drawings) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *Drawings) SetTop(rhs float64)  {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *Drawings) Dummy22_()  {
	retVal, _ := this.Call(0x00010016, nil)
	_= retVal
}

func (this *Drawings) Visible() bool {
	retVal, _ := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Drawings) SetVisible(rhs bool)  {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

func (this *Drawings) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *Drawings) SetWidth(rhs float64)  {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *Drawings) ZOrder() int32 {
	retVal, _ := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *Drawings) ShapeRange() *ShapeRange {
	retVal, _ := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.IDispatch(), false, true)
}

func (this *Drawings) AddIndent() bool {
	retVal, _ := this.PropGet(0x00000427, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Drawings) SetAddIndent(rhs bool)  {
	_ = this.PropPut(0x00000427, []interface{}{rhs})
}

func (this *Drawings) AutoScaleFont() ole.Variant {
	retVal, _ := this.PropGet(0x000005f5, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Drawings) SetAutoScaleFont(rhs interface{})  {
	_ = this.PropPut(0x000005f5, []interface{}{rhs})
}

func (this *Drawings) AutoSize() bool {
	retVal, _ := this.PropGet(0x00000266, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Drawings) SetAutoSize(rhs bool)  {
	_ = this.PropPut(0x00000266, []interface{}{rhs})
}

func (this *Drawings) Caption() string {
	retVal, _ := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Drawings) SetCaption(rhs string)  {
	_ = this.PropPut(0x0000008b, []interface{}{rhs})
}

var Drawings_Characters_OptArgs= []string{
	"Start", "Length", 
}

func (this *Drawings) Characters(optArgs ...interface{}) *Characters {
	optArgs = ole.ProcessOptArgs(Drawings_Characters_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000025b, nil, optArgs...)
	return NewCharacters(retVal.IDispatch(), false, true)
}

var Drawings_CheckSpelling_OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang", 
}

func (this *Drawings) CheckSpelling(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Drawings_CheckSpelling_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001f9, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Drawings) Font() *Font {
	retVal, _ := this.PropGet(0x00000092, nil)
	return NewFont(retVal.IDispatch(), false, true)
}

func (this *Drawings) Formula() string {
	retVal, _ := this.PropGet(0x00000105, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Drawings) SetFormula(rhs string)  {
	_ = this.PropPut(0x00000105, []interface{}{rhs})
}

func (this *Drawings) HorizontalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x00000088, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Drawings) SetHorizontalAlignment(rhs interface{})  {
	_ = this.PropPut(0x00000088, []interface{}{rhs})
}

func (this *Drawings) LockedText() bool {
	retVal, _ := this.PropGet(0x00000268, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Drawings) SetLockedText(rhs bool)  {
	_ = this.PropPut(0x00000268, []interface{}{rhs})
}

func (this *Drawings) Orientation() ole.Variant {
	retVal, _ := this.PropGet(0x00000086, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Drawings) SetOrientation(rhs interface{})  {
	_ = this.PropPut(0x00000086, []interface{}{rhs})
}

func (this *Drawings) Text() string {
	retVal, _ := this.PropGet(0x0000008a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Drawings) SetText(rhs string)  {
	_ = this.PropPut(0x0000008a, []interface{}{rhs})
}

func (this *Drawings) VerticalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x00000089, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Drawings) SetVerticalAlignment(rhs interface{})  {
	_ = this.PropPut(0x00000089, []interface{}{rhs})
}

func (this *Drawings) ReadingOrder() int32 {
	retVal, _ := this.PropGet(0x000003cf, nil)
	return retVal.LValVal()
}

func (this *Drawings) SetReadingOrder(rhs int32)  {
	_ = this.PropPut(0x000003cf, []interface{}{rhs})
}

func (this *Drawings) Border() *Border {
	retVal, _ := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.IDispatch(), false, true)
}

func (this *Drawings) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *Drawings) Shadow() bool {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Drawings) SetShadow(rhs bool)  {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *Drawings) Dummy44_()  {
	retVal, _ := this.Call(0x0001002c, nil)
	_= retVal
}

var Drawings_Reshape_OptArgs= []string{
	"Left", "Top", 
}

func (this *Drawings) Reshape(vertex int32, insert bool, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Drawings_Reshape_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000025c, []interface{}{vertex, insert}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Drawings) Dummy46_()  {
	retVal, _ := this.Call(0x0001002e, nil)
	_= retVal
}

func (this *Drawings) Add(x1 float64, y1 float64, x2 float64, y2 float64, closed bool) *Drawing {
	retVal, _ := this.Call(0x000000b5, []interface{}{x1, y1, x2, y2, closed})
	return NewDrawing(retVal.IDispatch(), false, true)
}

func (this *Drawings) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *Drawings) Group() *GroupObject {
	retVal, _ := this.Call(0x0000002e, nil)
	return NewGroupObject(retVal.IDispatch(), false, true)
}

func (this *Drawings) Item(index interface{}) *ole.DispatchClass {
	retVal, _ := this.Call(0x000000aa, []interface{}{index})
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Drawings) NewEnum_() *com.UnknownClass {
	retVal, _ := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Drawings) ForEach(action func(item *ole.DispatchClass) bool) {
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

