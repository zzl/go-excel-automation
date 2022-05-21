package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002088E-0000-0000-C000-000000000046
var IID_Spinners = syscall.GUID{0x0002088E, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Spinners struct {
	ole.OleClient
}

func NewSpinners(pDisp *win32.IDispatch, addRef bool, scoped bool) *Spinners {
	p := &Spinners{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SpinnersFromVar(v ole.Variant) *Spinners {
	return NewSpinners(v.PdispValVal(), false, false)
}

func (this *Spinners) IID() *syscall.GUID {
	return &IID_Spinners
}

func (this *Spinners) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Spinners) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Spinners) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Spinners) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Spinners) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Spinners) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Spinners) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Spinners) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Spinners) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Spinners) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Spinners) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Spinners) Dummy3_()  {
	retVal := this.Call(0x00010003, nil)
	_= retVal
}

func (this *Spinners) BringToFront() ole.Variant {
	retVal := this.Call(0x0000025a, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Spinners) Copy() ole.Variant {
	retVal := this.Call(0x00000227, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Spinners) CopyPicture(appearance int32, format int32) ole.Variant {
	retVal := this.Call(0x000000d5, []interface{}{appearance, format})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Spinners) Cut() ole.Variant {
	retVal := this.Call(0x00000235, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Spinners) Delete() ole.Variant {
	retVal := this.Call(0x00000075, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Spinners) Duplicate() *ole.DispatchClass {
	retVal := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Spinners) Enabled() bool {
	retVal := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Spinners) SetEnabled(rhs bool)  {
	retVal := this.PropPut(0x00000258, []interface{}{rhs})
	_= retVal
}

func (this *Spinners) Height() float64 {
	retVal := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *Spinners) SetHeight(rhs float64)  {
	retVal := this.PropPut(0x0000007b, []interface{}{rhs})
	_= retVal
}

func (this *Spinners) Dummy12_()  {
	retVal := this.Call(0x0001000c, nil)
	_= retVal
}

func (this *Spinners) Left() float64 {
	retVal := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *Spinners) SetLeft(rhs float64)  {
	retVal := this.PropPut(0x0000007f, []interface{}{rhs})
	_= retVal
}

func (this *Spinners) Locked() bool {
	retVal := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Spinners) SetLocked(rhs bool)  {
	retVal := this.PropPut(0x0000010d, []interface{}{rhs})
	_= retVal
}

func (this *Spinners) Dummy15_()  {
	retVal := this.Call(0x0001000f, nil)
	_= retVal
}

func (this *Spinners) OnAction() string {
	retVal := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Spinners) SetOnAction(rhs string)  {
	retVal := this.PropPut(0x00000254, []interface{}{rhs})
	_= retVal
}

func (this *Spinners) Placement() ole.Variant {
	retVal := this.PropGet(0x00000269, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Spinners) SetPlacement(rhs interface{})  {
	retVal := this.PropPut(0x00000269, []interface{}{rhs})
	_= retVal
}

func (this *Spinners) PrintObject() bool {
	retVal := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Spinners) SetPrintObject(rhs bool)  {
	retVal := this.PropPut(0x0000026a, []interface{}{rhs})
	_= retVal
}

var Spinners_Select_OptArgs= []string{
	"Replace", 
}

func (this *Spinners) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Spinners_Select_OptArgs, optArgs)
	retVal := this.Call(0x000000eb, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Spinners) SendToBack() ole.Variant {
	retVal := this.Call(0x0000025d, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Spinners) Top() float64 {
	retVal := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *Spinners) SetTop(rhs float64)  {
	retVal := this.PropPut(0x0000007e, []interface{}{rhs})
	_= retVal
}

func (this *Spinners) Dummy22_()  {
	retVal := this.Call(0x00010016, nil)
	_= retVal
}

func (this *Spinners) Visible() bool {
	retVal := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Spinners) SetVisible(rhs bool)  {
	retVal := this.PropPut(0x0000022e, []interface{}{rhs})
	_= retVal
}

func (this *Spinners) Width() float64 {
	retVal := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *Spinners) SetWidth(rhs float64)  {
	retVal := this.PropPut(0x0000007a, []interface{}{rhs})
	_= retVal
}

func (this *Spinners) ZOrder() int32 {
	retVal := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *Spinners) ShapeRange() *ShapeRange {
	retVal := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.PdispValVal(), false, true)
}

func (this *Spinners) Default_() int32 {
	retVal := this.PropGet(0x00000000, nil)
	return retVal.LValVal()
}

func (this *Spinners) SetDefault_(rhs int32)  {
	retVal := this.PropPut(0x00000000, []interface{}{rhs})
	_= retVal
}

func (this *Spinners) Display3DShading() bool {
	retVal := this.PropGet(0x00000462, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Spinners) SetDisplay3DShading(rhs bool)  {
	retVal := this.PropPut(0x00000462, []interface{}{rhs})
	_= retVal
}

func (this *Spinners) LinkedCell() string {
	retVal := this.PropGet(0x00000422, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Spinners) SetLinkedCell(rhs string)  {
	retVal := this.PropPut(0x00000422, []interface{}{rhs})
	_= retVal
}

func (this *Spinners) Max() int32 {
	retVal := this.PropGet(0x0000034a, nil)
	return retVal.LValVal()
}

func (this *Spinners) SetMax(rhs int32)  {
	retVal := this.PropPut(0x0000034a, []interface{}{rhs})
	_= retVal
}

func (this *Spinners) Min() int32 {
	retVal := this.PropGet(0x0000034b, nil)
	return retVal.LValVal()
}

func (this *Spinners) SetMin(rhs int32)  {
	retVal := this.PropPut(0x0000034b, []interface{}{rhs})
	_= retVal
}

func (this *Spinners) SmallChange() int32 {
	retVal := this.PropGet(0x0000034c, nil)
	return retVal.LValVal()
}

func (this *Spinners) SetSmallChange(rhs int32)  {
	retVal := this.PropPut(0x0000034c, []interface{}{rhs})
	_= retVal
}

func (this *Spinners) Value() int32 {
	retVal := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *Spinners) SetValue(rhs int32)  {
	retVal := this.PropPut(0x00000006, []interface{}{rhs})
	_= retVal
}

func (this *Spinners) Add(left float64, top float64, width float64, height float64) *Spinner {
	retVal := this.Call(0x000000b5, []interface{}{left, top, width, height})
	return NewSpinner(retVal.PdispValVal(), false, true)
}

func (this *Spinners) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *Spinners) Group() *GroupObject {
	retVal := this.Call(0x0000002e, nil)
	return NewGroupObject(retVal.PdispValVal(), false, true)
}

func (this *Spinners) Item(index interface{}) *ole.DispatchClass {
	retVal := this.Call(0x000000aa, []interface{}{index})
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Spinners) NewEnum_() *com.UnknownClass {
	retVal := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Spinners) ForEach(action func(item int32) bool) {
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

