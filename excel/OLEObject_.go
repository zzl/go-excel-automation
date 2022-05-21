package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000208A2-0000-0000-C000-000000000046
var IID_OLEObject_ = syscall.GUID{0x000208A2, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type OLEObject_ struct {
	ole.OleClient
}

func NewOLEObject_(pDisp *win32.IDispatch, addRef bool, scoped bool) *OLEObject_ {
	p := &OLEObject_{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OLEObject_FromVar(v ole.Variant) *OLEObject_ {
	return NewOLEObject_(v.PdispValVal(), false, false)
}

func (this *OLEObject_) IID() *syscall.GUID {
	return &IID_OLEObject_
}

func (this *OLEObject_) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OLEObject_) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *OLEObject_) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *OLEObject_) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *OLEObject_) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *OLEObject_) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *OLEObject_) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *OLEObject_) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *OLEObject_) Application() *Application {
	retVal := this.PropGet(-2147417964, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OLEObject_) Creator() int32 {
	retVal := this.PropGet(-2147417963, nil)
	return retVal.LValVal()
}

func (this *OLEObject_) Parent() *ole.DispatchClass {
	retVal := this.PropGet(-2147417962, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OLEObject_) BottomRightCell() *Range {
	retVal := this.PropGet(-2147417497, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *OLEObject_) BringToFront() ole.Variant {
	retVal := this.Call(-2147417510, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OLEObject_) Copy() ole.Variant {
	retVal := this.Call(-2147417561, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OLEObject_) CopyPicture(appearance int32, format int32) ole.Variant {
	retVal := this.Call(-2147417899, []interface{}{appearance, format})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OLEObject_) Cut() ole.Variant {
	retVal := this.Call(-2147417547, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OLEObject_) Delete() ole.Variant {
	retVal := this.Call(-2147417995, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OLEObject_) Duplicate() *ole.DispatchClass {
	retVal := this.Call(-2147417073, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OLEObject_) Enabled() bool {
	retVal := this.PropGet(-2147417512, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEObject_) SetEnabled(rhs bool)  {
	retVal := this.PropPut(-2147417512, []interface{}{rhs})
	_= retVal
}

func (this *OLEObject_) Height() float64 {
	retVal := this.PropGet(-2147417989, nil)
	return retVal.DblValVal()
}

func (this *OLEObject_) SetHeight(rhs float64)  {
	retVal := this.PropPut(-2147417989, []interface{}{rhs})
	_= retVal
}

func (this *OLEObject_) Index() int32 {
	retVal := this.PropGet(-2147417626, nil)
	return retVal.LValVal()
}

func (this *OLEObject_) Left() float64 {
	retVal := this.PropGet(-2147417985, nil)
	return retVal.DblValVal()
}

func (this *OLEObject_) SetLeft(rhs float64)  {
	retVal := this.PropPut(-2147417985, []interface{}{rhs})
	_= retVal
}

func (this *OLEObject_) Locked() bool {
	retVal := this.PropGet(-2147417843, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEObject_) SetLocked(rhs bool)  {
	retVal := this.PropPut(-2147417843, []interface{}{rhs})
	_= retVal
}

func (this *OLEObject_) Name() string {
	retVal := this.PropGet(-2147418002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OLEObject_) SetName(rhs string)  {
	retVal := this.PropPut(-2147418002, []interface{}{rhs})
	_= retVal
}

func (this *OLEObject_) OnAction() string {
	retVal := this.PropGet(-2147417516, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OLEObject_) SetOnAction(rhs string)  {
	retVal := this.PropPut(-2147417516, []interface{}{rhs})
	_= retVal
}

func (this *OLEObject_) Placement() ole.Variant {
	retVal := this.PropGet(-2147417495, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OLEObject_) SetPlacement(rhs interface{})  {
	retVal := this.PropPut(-2147417495, []interface{}{rhs})
	_= retVal
}

func (this *OLEObject_) PrintObject() bool {
	retVal := this.PropGet(-2147417494, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEObject_) SetPrintObject(rhs bool)  {
	retVal := this.PropPut(-2147417494, []interface{}{rhs})
	_= retVal
}

var OLEObject__Select_OptArgs= []string{
	"Replace", 
}

func (this *OLEObject_) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(OLEObject__Select_OptArgs, optArgs)
	retVal := this.Call(-2147417877, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OLEObject_) SendToBack() ole.Variant {
	retVal := this.Call(-2147417507, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OLEObject_) Top() float64 {
	retVal := this.PropGet(-2147417986, nil)
	return retVal.DblValVal()
}

func (this *OLEObject_) SetTop(rhs float64)  {
	retVal := this.PropPut(-2147417986, []interface{}{rhs})
	_= retVal
}

func (this *OLEObject_) TopLeftCell() *Range {
	retVal := this.PropGet(-2147417492, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *OLEObject_) Visible() bool {
	retVal := this.PropGet(-2147417554, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEObject_) SetVisible(rhs bool)  {
	retVal := this.PropPut(-2147417554, []interface{}{rhs})
	_= retVal
}

func (this *OLEObject_) Width() float64 {
	retVal := this.PropGet(-2147417990, nil)
	return retVal.DblValVal()
}

func (this *OLEObject_) SetWidth(rhs float64)  {
	retVal := this.PropPut(-2147417990, []interface{}{rhs})
	_= retVal
}

func (this *OLEObject_) ZOrder() int32 {
	retVal := this.PropGet(-2147417490, nil)
	return retVal.LValVal()
}

func (this *OLEObject_) ShapeRange() *ShapeRange {
	retVal := this.PropGet(-2147416584, nil)
	return NewShapeRange(retVal.PdispValVal(), false, true)
}

func (this *OLEObject_) Border() *Border {
	retVal := this.PropGet(-2147417984, nil)
	return NewBorder(retVal.PdispValVal(), false, true)
}

func (this *OLEObject_) Interior() *Interior {
	retVal := this.PropGet(-2147417983, nil)
	return NewInterior(retVal.PdispValVal(), false, true)
}

func (this *OLEObject_) Shadow() bool {
	retVal := this.PropGet(-2147418009, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEObject_) SetShadow(rhs bool)  {
	retVal := this.PropPut(-2147418009, []interface{}{rhs})
	_= retVal
}

func (this *OLEObject_) Activate() ole.Variant {
	retVal := this.Call(-2147417808, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OLEObject_) AutoLoad() bool {
	retVal := this.PropGet(-2147416926, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEObject_) SetAutoLoad(rhs bool)  {
	retVal := this.PropPut(-2147416926, []interface{}{rhs})
	_= retVal
}

func (this *OLEObject_) AutoUpdate() bool {
	retVal := this.PropGet(-2147417064, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEObject_) SetAutoUpdate(rhs bool)  {
	retVal := this.PropPut(-2147417064, []interface{}{rhs})
	_= retVal
}

func (this *OLEObject_) Object() *ole.DispatchClass {
	retVal := this.PropGet(-2147417063, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OLEObject_) OLEType() ole.Variant {
	retVal := this.PropGet(-2147417058, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OLEObject_) SourceName() string {
	retVal := this.PropGet(-2147417391, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OLEObject_) SetSourceName(rhs string)  {
	retVal := this.PropPut(-2147417391, []interface{}{rhs})
	_= retVal
}

func (this *OLEObject_) Update() ole.Variant {
	retVal := this.Call(-2147417432, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OLEObject_) Verb(verb int32) ole.Variant {
	retVal := this.Call(-2147417506, []interface{}{verb})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *OLEObject_) LinkedCell() string {
	retVal := this.PropGet(-2147417054, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OLEObject_) SetLinkedCell(rhs string)  {
	retVal := this.PropPut(-2147417054, []interface{}{rhs})
	_= retVal
}

func (this *OLEObject_) ListFillRange() string {
	retVal := this.PropGet(-2147417265, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OLEObject_) SetListFillRange(rhs string)  {
	retVal := this.PropPut(-2147417265, []interface{}{rhs})
	_= retVal
}

func (this *OLEObject_) ProgID() string {
	retVal := this.PropGet(-2147416589, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OLEObject_) AltHTML() string {
	retVal := this.PropGet(-2147416259, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OLEObject_) SetAltHTML(rhs string)  {
	retVal := this.PropPut(-2147416259, []interface{}{rhs})
	_= retVal
}

