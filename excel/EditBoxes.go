package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020884-0000-0000-C000-000000000046
var IID_EditBoxes = syscall.GUID{0x00020884, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type EditBoxes struct {
	ole.OleClient
}

func NewEditBoxes(pDisp *win32.IDispatch, addRef bool, scoped bool) *EditBoxes {
	p := &EditBoxes{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func EditBoxesFromVar(v ole.Variant) *EditBoxes {
	return NewEditBoxes(v.PdispValVal(), false, false)
}

func (this *EditBoxes) IID() *syscall.GUID {
	return &IID_EditBoxes
}

func (this *EditBoxes) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *EditBoxes) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *EditBoxes) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *EditBoxes) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *EditBoxes) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *EditBoxes) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *EditBoxes) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *EditBoxes) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *EditBoxes) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *EditBoxes) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *EditBoxes) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *EditBoxes) Dummy3_()  {
	retVal := this.Call(0x00010003, nil)
	_= retVal
}

func (this *EditBoxes) BringToFront() ole.Variant {
	retVal := this.Call(0x0000025a, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *EditBoxes) Copy() ole.Variant {
	retVal := this.Call(0x00000227, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *EditBoxes) CopyPicture(appearance int32, format int32) ole.Variant {
	retVal := this.Call(0x000000d5, []interface{}{appearance, format})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *EditBoxes) Cut() ole.Variant {
	retVal := this.Call(0x00000235, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *EditBoxes) Delete() ole.Variant {
	retVal := this.Call(0x00000075, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *EditBoxes) Duplicate() *ole.DispatchClass {
	retVal := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *EditBoxes) Enabled() bool {
	retVal := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EditBoxes) SetEnabled(rhs bool)  {
	retVal := this.PropPut(0x00000258, []interface{}{rhs})
	_= retVal
}

func (this *EditBoxes) Height() float64 {
	retVal := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *EditBoxes) SetHeight(rhs float64)  {
	retVal := this.PropPut(0x0000007b, []interface{}{rhs})
	_= retVal
}

func (this *EditBoxes) Dummy12_()  {
	retVal := this.Call(0x0001000c, nil)
	_= retVal
}

func (this *EditBoxes) Left() float64 {
	retVal := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *EditBoxes) SetLeft(rhs float64)  {
	retVal := this.PropPut(0x0000007f, []interface{}{rhs})
	_= retVal
}

func (this *EditBoxes) Locked() bool {
	retVal := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EditBoxes) SetLocked(rhs bool)  {
	retVal := this.PropPut(0x0000010d, []interface{}{rhs})
	_= retVal
}

func (this *EditBoxes) Dummy15_()  {
	retVal := this.Call(0x0001000f, nil)
	_= retVal
}

func (this *EditBoxes) OnAction() string {
	retVal := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *EditBoxes) SetOnAction(rhs string)  {
	retVal := this.PropPut(0x00000254, []interface{}{rhs})
	_= retVal
}

func (this *EditBoxes) Placement() ole.Variant {
	retVal := this.PropGet(0x00000269, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *EditBoxes) SetPlacement(rhs interface{})  {
	retVal := this.PropPut(0x00000269, []interface{}{rhs})
	_= retVal
}

func (this *EditBoxes) PrintObject() bool {
	retVal := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EditBoxes) SetPrintObject(rhs bool)  {
	retVal := this.PropPut(0x0000026a, []interface{}{rhs})
	_= retVal
}

var EditBoxes_Select_OptArgs= []string{
	"Replace", 
}

func (this *EditBoxes) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(EditBoxes_Select_OptArgs, optArgs)
	retVal := this.Call(0x000000eb, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *EditBoxes) SendToBack() ole.Variant {
	retVal := this.Call(0x0000025d, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *EditBoxes) Top() float64 {
	retVal := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *EditBoxes) SetTop(rhs float64)  {
	retVal := this.PropPut(0x0000007e, []interface{}{rhs})
	_= retVal
}

func (this *EditBoxes) Dummy22_()  {
	retVal := this.Call(0x00010016, nil)
	_= retVal
}

func (this *EditBoxes) Visible() bool {
	retVal := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EditBoxes) SetVisible(rhs bool)  {
	retVal := this.PropPut(0x0000022e, []interface{}{rhs})
	_= retVal
}

func (this *EditBoxes) Width() float64 {
	retVal := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *EditBoxes) SetWidth(rhs float64)  {
	retVal := this.PropPut(0x0000007a, []interface{}{rhs})
	_= retVal
}

func (this *EditBoxes) ZOrder() int32 {
	retVal := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *EditBoxes) ShapeRange() *ShapeRange {
	retVal := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.PdispValVal(), false, true)
}

func (this *EditBoxes) Caption() string {
	retVal := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *EditBoxes) SetCaption(rhs string)  {
	retVal := this.PropPut(0x0000008b, []interface{}{rhs})
	_= retVal
}

var EditBoxes_Characters_OptArgs= []string{
	"Start", "Length", 
}

func (this *EditBoxes) Characters(optArgs ...interface{}) *Characters {
	optArgs = ole.ProcessOptArgs(EditBoxes_Characters_OptArgs, optArgs)
	retVal := this.PropGet(0x0000025b, nil, optArgs...)
	return NewCharacters(retVal.PdispValVal(), false, true)
}

var EditBoxes_CheckSpelling_OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang", 
}

func (this *EditBoxes) CheckSpelling(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(EditBoxes_CheckSpelling_OptArgs, optArgs)
	retVal := this.Call(0x000001f9, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *EditBoxes) LockedText() bool {
	retVal := this.PropGet(0x00000268, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EditBoxes) SetLockedText(rhs bool)  {
	retVal := this.PropPut(0x00000268, []interface{}{rhs})
	_= retVal
}

func (this *EditBoxes) Text() string {
	retVal := this.PropGet(0x0000008a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *EditBoxes) SetText(rhs string)  {
	retVal := this.PropPut(0x0000008a, []interface{}{rhs})
	_= retVal
}

func (this *EditBoxes) DisplayVerticalScrollBar() bool {
	retVal := this.PropGet(0x0000039a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EditBoxes) SetDisplayVerticalScrollBar(rhs bool)  {
	retVal := this.PropPut(0x0000039a, []interface{}{rhs})
	_= retVal
}

func (this *EditBoxes) InputType() int32 {
	retVal := this.PropGet(0x00000356, nil)
	return retVal.LValVal()
}

func (this *EditBoxes) SetInputType(rhs int32)  {
	retVal := this.PropPut(0x00000356, []interface{}{rhs})
	_= retVal
}

func (this *EditBoxes) Dummy34_()  {
	retVal := this.Call(0x00010022, nil)
	_= retVal
}

func (this *EditBoxes) MultiLine() bool {
	retVal := this.PropGet(0x00000357, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EditBoxes) SetMultiLine(rhs bool)  {
	retVal := this.PropPut(0x00000357, []interface{}{rhs})
	_= retVal
}

func (this *EditBoxes) PasswordEdit() bool {
	retVal := this.PropGet(0x0000048a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EditBoxes) SetPasswordEdit(rhs bool)  {
	retVal := this.PropPut(0x0000048a, []interface{}{rhs})
	_= retVal
}

func (this *EditBoxes) Add(left float64, top float64, width float64, height float64) *EditBox {
	retVal := this.Call(0x000000b5, []interface{}{left, top, width, height})
	return NewEditBox(retVal.PdispValVal(), false, true)
}

func (this *EditBoxes) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *EditBoxes) Group() *GroupObject {
	retVal := this.Call(0x0000002e, nil)
	return NewGroupObject(retVal.PdispValVal(), false, true)
}

func (this *EditBoxes) Item(index interface{}) ole.Variant {
	retVal := this.Call(0x000000aa, []interface{}{index})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *EditBoxes) NewEnum_() *com.UnknownClass {
	retVal := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *EditBoxes) ForEach(action func(item ole.Variant) bool) {
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
		pItem := v
		ret := action(pItem)
		if !ret {
			break
		}
	}
}

