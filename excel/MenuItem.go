package excel

import (
	"github.com/zzl/go-win32api/v2/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020868-0000-0000-C000-000000000046
var IID_MenuItem = syscall.GUID{0x00020868, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type MenuItem struct {
	ole.OleClient
}

func NewMenuItem(pDisp *win32.IDispatch, addRef bool, scoped bool) *MenuItem {
	if pDisp == nil {
		return nil
	}
	p := &MenuItem{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func MenuItemFromVar(v ole.Variant) *MenuItem {
	return NewMenuItem(v.IDispatch(), false, false)
}

func (this *MenuItem) IID() *syscall.GUID {
	return &IID_MenuItem
}

func (this *MenuItem) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *MenuItem) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *MenuItem) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *MenuItem) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *MenuItem) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *MenuItem) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *MenuItem) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *MenuItem) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *MenuItem) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *MenuItem) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *MenuItem) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *MenuItem) Caption() string {
	retVal, _ := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *MenuItem) SetCaption(rhs string) {
	_ = this.PropPut(0x0000008b, []interface{}{rhs})
}

func (this *MenuItem) Checked() bool {
	retVal, _ := this.PropGet(0x00000257, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *MenuItem) SetChecked(rhs bool) {
	_ = this.PropPut(0x00000257, []interface{}{rhs})
}

func (this *MenuItem) Delete() {
	retVal, _ := this.Call(0x00000075, nil)
	_ = retVal
}

func (this *MenuItem) Enabled() bool {
	retVal, _ := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *MenuItem) SetEnabled(rhs bool) {
	_ = this.PropPut(0x00000258, []interface{}{rhs})
}

func (this *MenuItem) HelpContextID() int32 {
	retVal, _ := this.PropGet(0x00000163, nil)
	return retVal.LValVal()
}

func (this *MenuItem) SetHelpContextID(rhs int32) {
	_ = this.PropPut(0x00000163, []interface{}{rhs})
}

func (this *MenuItem) HelpFile() string {
	retVal, _ := this.PropGet(0x00000168, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *MenuItem) SetHelpFile(rhs string) {
	_ = this.PropPut(0x00000168, []interface{}{rhs})
}

func (this *MenuItem) Index() int32 {
	retVal, _ := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *MenuItem) OnAction() string {
	retVal, _ := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *MenuItem) SetOnAction(rhs string) {
	_ = this.PropPut(0x00000254, []interface{}{rhs})
}

func (this *MenuItem) StatusBar() string {
	retVal, _ := this.PropGet(0x00000182, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *MenuItem) SetStatusBar(rhs string) {
	_ = this.PropPut(0x00000182, []interface{}{rhs})
}

