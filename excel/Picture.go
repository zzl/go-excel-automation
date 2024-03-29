package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 000208A6-0000-0000-C000-000000000046
var IID_Picture = syscall.GUID{0x000208A6, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Picture struct {
	ole.OleClient
}

func NewPicture(pDisp *win32.IDispatch, addRef bool, scoped bool) *Picture {
	if pDisp == nil {
		return nil
	}
	p := &Picture{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PictureFromVar(v ole.Variant) *Picture {
	return NewPicture(v.IDispatch(), false, false)
}

func (this *Picture) IID() *syscall.GUID {
	return &IID_Picture
}

func (this *Picture) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Picture) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Picture) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Picture) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Picture) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Picture) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Picture) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Picture) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Picture) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Picture) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Picture) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Picture) BottomRightCell() *Range {
	retVal, _ := this.PropGet(0x00000267, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Picture) BringToFront() ole.Variant {
	retVal, _ := this.Call(0x0000025a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Picture) Copy() ole.Variant {
	retVal, _ := this.Call(0x00000227, nil)
	com.AddToScope(retVal)
	return *retVal
}

var Picture_CopyPicture_OptArgs = []string{
	"Appearance", "Format",
}

func (this *Picture) CopyPicture(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Picture_CopyPicture_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d5, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Picture) Cut() ole.Variant {
	retVal, _ := this.Call(0x00000235, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Picture) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Picture) Duplicate() *ole.DispatchClass {
	retVal, _ := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Picture) Enabled() bool {
	retVal, _ := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Picture) SetEnabled(rhs bool) {
	_ = this.PropPut(0x00000258, []interface{}{rhs})
}

func (this *Picture) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *Picture) SetHeight(rhs float64) {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *Picture) Index() int32 {
	retVal, _ := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *Picture) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *Picture) SetLeft(rhs float64) {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *Picture) Locked() bool {
	retVal, _ := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Picture) SetLocked(rhs bool) {
	_ = this.PropPut(0x0000010d, []interface{}{rhs})
}

func (this *Picture) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Picture) SetName(rhs string) {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *Picture) OnAction() string {
	retVal, _ := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Picture) SetOnAction(rhs string) {
	_ = this.PropPut(0x00000254, []interface{}{rhs})
}

func (this *Picture) Placement() ole.Variant {
	retVal, _ := this.PropGet(0x00000269, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Picture) SetPlacement(rhs interface{}) {
	_ = this.PropPut(0x00000269, []interface{}{rhs})
}

func (this *Picture) PrintObject() bool {
	retVal, _ := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Picture) SetPrintObject(rhs bool) {
	_ = this.PropPut(0x0000026a, []interface{}{rhs})
}

var Picture_Select_OptArgs = []string{
	"Replace",
}

func (this *Picture) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Picture_Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000eb, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Picture) SendToBack() ole.Variant {
	retVal, _ := this.Call(0x0000025d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Picture) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *Picture) SetTop(rhs float64) {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *Picture) TopLeftCell() *Range {
	retVal, _ := this.PropGet(0x0000026c, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Picture) Visible() bool {
	retVal, _ := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Picture) SetVisible(rhs bool) {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

func (this *Picture) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *Picture) SetWidth(rhs float64) {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *Picture) ZOrder() int32 {
	retVal, _ := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *Picture) ShapeRange() *ShapeRange {
	retVal, _ := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.IDispatch(), false, true)
}

func (this *Picture) Border() *Border {
	retVal, _ := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.IDispatch(), false, true)
}

func (this *Picture) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *Picture) Shadow() bool {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Picture) SetShadow(rhs bool) {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *Picture) Formula() string {
	retVal, _ := this.PropGet(0x00000105, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Picture) SetFormula(rhs string) {
	_ = this.PropPut(0x00000105, []interface{}{rhs})
}
