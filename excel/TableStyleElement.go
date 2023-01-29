package excel

import (
	"github.com/zzl/go-win32api/v2/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244A5-0000-0000-C000-000000000046
var IID_TableStyleElement = syscall.GUID{0x000244A5, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TableStyleElement struct {
	ole.OleClient
}

func NewTableStyleElement(pDisp *win32.IDispatch, addRef bool, scoped bool) *TableStyleElement {
	if pDisp == nil {
		return nil
	}
	p := &TableStyleElement{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TableStyleElementFromVar(v ole.Variant) *TableStyleElement {
	return NewTableStyleElement(v.IDispatch(), false, false)
}

func (this *TableStyleElement) IID() *syscall.GUID {
	return &IID_TableStyleElement
}

func (this *TableStyleElement) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TableStyleElement) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *TableStyleElement) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *TableStyleElement) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *TableStyleElement) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *TableStyleElement) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *TableStyleElement) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *TableStyleElement) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *TableStyleElement) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *TableStyleElement) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *TableStyleElement) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TableStyleElement) HasFormat() bool {
	retVal, _ := this.PropGet(0x00000aaf, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TableStyleElement) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *TableStyleElement) Borders() *Borders {
	retVal, _ := this.PropGet(0x000001b3, nil)
	return NewBorders(retVal.IDispatch(), false, true)
}

func (this *TableStyleElement) Font() *Font {
	retVal, _ := this.PropGet(0x00000092, nil)
	return NewFont(retVal.IDispatch(), false, true)
}

func (this *TableStyleElement) StripeSize() int32 {
	retVal, _ := this.PropGet(0x00000ab0, nil)
	return retVal.LValVal()
}

func (this *TableStyleElement) SetStripeSize(rhs int32) {
	_ = this.PropPut(0x00000ab0, []interface{}{rhs})
}

func (this *TableStyleElement) Clear() {
	retVal, _ := this.Call(0x0000006f, nil)
	_ = retVal
}

