package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002089A-0000-0000-C000-000000000046
var IID_Line = syscall.GUID{0x0002089A, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Line struct {
	ole.OleClient
}

func NewLine(pDisp *win32.IDispatch, addRef bool, scoped bool) *Line {
	 if pDisp == nil {
		return nil;
	}
	p := &Line{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func LineFromVar(v ole.Variant) *Line {
	return NewLine(v.IDispatch(), false, false)
}

func (this *Line) IID() *syscall.GUID {
	return &IID_Line
}

func (this *Line) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Line) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Line) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Line) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Line) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Line) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Line) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Line) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Line) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Line) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Line) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Line) BottomRightCell() *Range {
	retVal, _ := this.PropGet(0x00000267, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Line) BringToFront() ole.Variant {
	retVal, _ := this.Call(0x0000025a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Line) Copy() ole.Variant {
	retVal, _ := this.Call(0x00000227, nil)
	com.AddToScope(retVal)
	return *retVal
}

var Line_CopyPicture_OptArgs= []string{
	"Appearance", "Format", 
}

func (this *Line) CopyPicture(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Line_CopyPicture_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d5, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Line) Cut() ole.Variant {
	retVal, _ := this.Call(0x00000235, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Line) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Line) Duplicate() *ole.DispatchClass {
	retVal, _ := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Line) Enabled() bool {
	retVal, _ := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Line) SetEnabled(rhs bool)  {
	_ = this.PropPut(0x00000258, []interface{}{rhs})
}

func (this *Line) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *Line) SetHeight(rhs float64)  {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *Line) Index() int32 {
	retVal, _ := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *Line) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *Line) SetLeft(rhs float64)  {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *Line) Locked() bool {
	retVal, _ := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Line) SetLocked(rhs bool)  {
	_ = this.PropPut(0x0000010d, []interface{}{rhs})
}

func (this *Line) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Line) SetName(rhs string)  {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *Line) OnAction() string {
	retVal, _ := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Line) SetOnAction(rhs string)  {
	_ = this.PropPut(0x00000254, []interface{}{rhs})
}

func (this *Line) Placement() ole.Variant {
	retVal, _ := this.PropGet(0x00000269, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Line) SetPlacement(rhs interface{})  {
	_ = this.PropPut(0x00000269, []interface{}{rhs})
}

func (this *Line) PrintObject() bool {
	retVal, _ := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Line) SetPrintObject(rhs bool)  {
	_ = this.PropPut(0x0000026a, []interface{}{rhs})
}

var Line_Select_OptArgs= []string{
	"Replace", 
}

func (this *Line) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Line_Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000eb, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Line) SendToBack() ole.Variant {
	retVal, _ := this.Call(0x0000025d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Line) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *Line) SetTop(rhs float64)  {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *Line) TopLeftCell() *Range {
	retVal, _ := this.PropGet(0x0000026c, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Line) Visible() bool {
	retVal, _ := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Line) SetVisible(rhs bool)  {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

func (this *Line) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *Line) SetWidth(rhs float64)  {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *Line) ZOrder() int32 {
	retVal, _ := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *Line) ShapeRange() *ShapeRange {
	retVal, _ := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.IDispatch(), false, true)
}

func (this *Line) ArrowHeadLength() ole.Variant {
	retVal, _ := this.PropGet(0x00000263, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Line) SetArrowHeadLength(rhs interface{})  {
	_ = this.PropPut(0x00000263, []interface{}{rhs})
}

func (this *Line) ArrowHeadStyle() ole.Variant {
	retVal, _ := this.PropGet(0x00000264, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Line) SetArrowHeadStyle(rhs interface{})  {
	_ = this.PropPut(0x00000264, []interface{}{rhs})
}

func (this *Line) ArrowHeadWidth() ole.Variant {
	retVal, _ := this.PropGet(0x00000265, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Line) SetArrowHeadWidth(rhs interface{})  {
	_ = this.PropPut(0x00000265, []interface{}{rhs})
}

func (this *Line) Border() *Border {
	retVal, _ := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.IDispatch(), false, true)
}

