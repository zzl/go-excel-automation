package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00020895-0000-0000-C000-000000000046
var IID_Pane = syscall.GUID{0x00020895, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Pane struct {
	ole.OleClient
}

func NewPane(pDisp *win32.IDispatch, addRef bool, scoped bool) *Pane {
	if pDisp == nil {
		return nil
	}
	p := &Pane{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PaneFromVar(v ole.Variant) *Pane {
	return NewPane(v.IDispatch(), false, false)
}

func (this *Pane) IID() *syscall.GUID {
	return &IID_Pane
}

func (this *Pane) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Pane) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Pane) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Pane) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Pane) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Pane) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Pane) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Pane) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Pane) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Pane) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Pane) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Pane) Activate() bool {
	retVal, _ := this.Call(0x00000130, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Pane) Index() int32 {
	retVal, _ := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

var Pane_LargeScroll_OptArgs = []string{
	"Down", "Up", "ToRight", "ToLeft",
}

func (this *Pane) LargeScroll(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Pane_LargeScroll_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000223, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Pane) ScrollColumn() int32 {
	retVal, _ := this.PropGet(0x0000028e, nil)
	return retVal.LValVal()
}

func (this *Pane) SetScrollColumn(rhs int32) {
	_ = this.PropPut(0x0000028e, []interface{}{rhs})
}

func (this *Pane) ScrollRow() int32 {
	retVal, _ := this.PropGet(0x0000028f, nil)
	return retVal.LValVal()
}

func (this *Pane) SetScrollRow(rhs int32) {
	_ = this.PropPut(0x0000028f, []interface{}{rhs})
}

var Pane_SmallScroll_OptArgs = []string{
	"Down", "Up", "ToRight", "ToLeft",
}

func (this *Pane) SmallScroll(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Pane_SmallScroll_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000224, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Pane) VisibleRange() *Range {
	retVal, _ := this.PropGet(0x0000045e, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

var Pane_ScrollIntoView_OptArgs = []string{
	"Start",
}

func (this *Pane) ScrollIntoView(left int32, top int32, width int32, height int32, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Pane_ScrollIntoView_OptArgs, optArgs)
	retVal, _ := this.Call(0x000006f5, []interface{}{left, top, width, height}, optArgs...)
	_ = retVal
}

func (this *Pane) PointsToScreenPixelsX(points int32) int32 {
	retVal, _ := this.Call(0x000006f0, []interface{}{points})
	return retVal.LValVal()
}

func (this *Pane) PointsToScreenPixelsY(points int32) int32 {
	retVal, _ := this.Call(0x000006f1, []interface{}{points})
	return retVal.LValVal()
}
