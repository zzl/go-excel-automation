package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 0002085C-0000-0000-C000-000000000046
var IID_Toolbar = syscall.GUID{0x0002085C, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Toolbar struct {
	ole.OleClient
}

func NewToolbar(pDisp *win32.IDispatch, addRef bool, scoped bool) *Toolbar {
	if pDisp == nil {
		return nil
	}
	p := &Toolbar{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ToolbarFromVar(v ole.Variant) *Toolbar {
	return NewToolbar(v.IDispatch(), false, false)
}

func (this *Toolbar) IID() *syscall.GUID {
	return &IID_Toolbar
}

func (this *Toolbar) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Toolbar) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Toolbar) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Toolbar) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Toolbar) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Toolbar) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Toolbar) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Toolbar) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Toolbar) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Toolbar) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Toolbar) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Toolbar) BuiltIn() bool {
	retVal, _ := this.PropGet(0x00000229, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Toolbar) Delete() {
	retVal, _ := this.Call(0x00000075, nil)
	_ = retVal
}

func (this *Toolbar) Height() int32 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.LValVal()
}

func (this *Toolbar) SetHeight(rhs int32) {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *Toolbar) Left() int32 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.LValVal()
}

func (this *Toolbar) SetLeft(rhs int32) {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *Toolbar) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Toolbar) Position() int32 {
	retVal, _ := this.PropGet(0x00000085, nil)
	return retVal.LValVal()
}

func (this *Toolbar) SetPosition(rhs int32) {
	_ = this.PropPut(0x00000085, []interface{}{rhs})
}

func (this *Toolbar) Protection() int32 {
	retVal, _ := this.PropGet(0x000000b0, nil)
	return retVal.LValVal()
}

func (this *Toolbar) SetProtection(rhs int32) {
	_ = this.PropPut(0x000000b0, []interface{}{rhs})
}

func (this *Toolbar) Reset() {
	retVal, _ := this.Call(0x0000022b, nil)
	_ = retVal
}

func (this *Toolbar) ToolbarButtons() *ToolbarButtons {
	retVal, _ := this.PropGet(0x000003c4, nil)
	return NewToolbarButtons(retVal.IDispatch(), false, true)
}

func (this *Toolbar) Top() int32 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.LValVal()
}

func (this *Toolbar) SetTop(rhs int32) {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *Toolbar) Visible() bool {
	retVal, _ := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Toolbar) SetVisible(rhs bool) {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

func (this *Toolbar) Width() int32 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.LValVal()
}

func (this *Toolbar) SetWidth(rhs int32) {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}
