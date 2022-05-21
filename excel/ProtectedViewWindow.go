package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244CD-0000-0000-C000-000000000046
var IID_ProtectedViewWindow = syscall.GUID{0x000244CD, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ProtectedViewWindow struct {
	ole.OleClient
}

func NewProtectedViewWindow(pDisp *win32.IDispatch, addRef bool, scoped bool) *ProtectedViewWindow {
	p := &ProtectedViewWindow{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ProtectedViewWindowFromVar(v ole.Variant) *ProtectedViewWindow {
	return NewProtectedViewWindow(v.PdispValVal(), false, false)
}

func (this *ProtectedViewWindow) IID() *syscall.GUID {
	return &IID_ProtectedViewWindow
}

func (this *ProtectedViewWindow) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ProtectedViewWindow) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *ProtectedViewWindow) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ProtectedViewWindow) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ProtectedViewWindow) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *ProtectedViewWindow) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *ProtectedViewWindow) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *ProtectedViewWindow) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *ProtectedViewWindow) Default_() string {
	retVal := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ProtectedViewWindow) Caption() string {
	retVal := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ProtectedViewWindow) SetCaption(rhs string)  {
	retVal := this.PropPut(0x0000008b, []interface{}{rhs})
	_= retVal
}

func (this *ProtectedViewWindow) EnableResize() bool {
	retVal := this.PropGet(0x000004a8, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ProtectedViewWindow) SetEnableResize(rhs bool)  {
	retVal := this.PropPut(0x000004a8, []interface{}{rhs})
	_= retVal
}

func (this *ProtectedViewWindow) Height() float64 {
	retVal := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *ProtectedViewWindow) SetHeight(rhs float64)  {
	retVal := this.PropPut(0x0000007b, []interface{}{rhs})
	_= retVal
}

func (this *ProtectedViewWindow) Left() float64 {
	retVal := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *ProtectedViewWindow) SetLeft(rhs float64)  {
	retVal := this.PropPut(0x0000007f, []interface{}{rhs})
	_= retVal
}

func (this *ProtectedViewWindow) Top() float64 {
	retVal := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *ProtectedViewWindow) SetTop(rhs float64)  {
	retVal := this.PropPut(0x0000007e, []interface{}{rhs})
	_= retVal
}

func (this *ProtectedViewWindow) Width() float64 {
	retVal := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *ProtectedViewWindow) SetWidth(rhs float64)  {
	retVal := this.PropPut(0x0000007a, []interface{}{rhs})
	_= retVal
}

func (this *ProtectedViewWindow) Visible() bool {
	retVal := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ProtectedViewWindow) SetVisible(rhs bool)  {
	retVal := this.PropPut(0x0000022e, []interface{}{rhs})
	_= retVal
}

func (this *ProtectedViewWindow) SourceName() string {
	retVal := this.PropGet(0x000002d1, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ProtectedViewWindow) SourcePath() string {
	retVal := this.PropGet(0x00000bb1, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ProtectedViewWindow) WindowState() int32 {
	retVal := this.PropGet(0x0000018c, nil)
	return retVal.LValVal()
}

func (this *ProtectedViewWindow) SetWindowState(rhs int32)  {
	retVal := this.PropPut(0x0000018c, []interface{}{rhs})
	_= retVal
}

func (this *ProtectedViewWindow) Workbook() *Workbook {
	retVal := this.PropGet(0x000002f0, nil)
	return NewWorkbook(retVal.PdispValVal(), false, true)
}

func (this *ProtectedViewWindow) Activate()  {
	retVal := this.Call(0x00000130, nil)
	_= retVal
}

func (this *ProtectedViewWindow) Close() bool {
	retVal := this.Call(0x00000115, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var ProtectedViewWindow_Edit_OptArgs= []string{
	"WriteResPassword", "UpdateLinks", 
}

func (this *ProtectedViewWindow) Edit(optArgs ...interface{}) *Workbook {
	optArgs = ole.ProcessOptArgs(ProtectedViewWindow_Edit_OptArgs, optArgs)
	retVal := this.Call(0x00000232, nil, optArgs...)
	return NewWorkbook(retVal.PdispValVal(), false, true)
}

