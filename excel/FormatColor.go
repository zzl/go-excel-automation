package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00024491-0000-0000-C000-000000000046
var IID_FormatColor = syscall.GUID{0x00024491, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type FormatColor struct {
	ole.OleClient
}

func NewFormatColor(pDisp *win32.IDispatch, addRef bool, scoped bool) *FormatColor {
	if pDisp == nil {
		return nil
	}
	p := &FormatColor{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FormatColorFromVar(v ole.Variant) *FormatColor {
	return NewFormatColor(v.IDispatch(), false, false)
}

func (this *FormatColor) IID() *syscall.GUID {
	return &IID_FormatColor
}

func (this *FormatColor) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *FormatColor) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *FormatColor) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *FormatColor) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *FormatColor) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *FormatColor) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *FormatColor) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *FormatColor) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *FormatColor) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *FormatColor) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *FormatColor) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *FormatColor) Color() ole.Variant {
	retVal, _ := this.PropGet(0x00000063, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *FormatColor) SetColor(rhs interface{}) {
	_ = this.PropPut(0x00000063, []interface{}{rhs})
}

func (this *FormatColor) ColorIndex() int32 {
	retVal, _ := this.PropGet(0x00000061, nil)
	return retVal.LValVal()
}

func (this *FormatColor) SetColorIndex(rhs int32) {
	_ = this.PropPut(0x00000061, []interface{}{rhs})
}

func (this *FormatColor) ThemeColor() ole.Variant {
	retVal, _ := this.PropGet(0x0000093d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *FormatColor) SetThemeColor(rhs interface{}) {
	_ = this.PropPut(0x0000093d, []interface{}{rhs})
}

func (this *FormatColor) TintAndShade() ole.Variant {
	retVal, _ := this.PropGet(0x0000093e, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *FormatColor) SetTintAndShade(rhs interface{}) {
	_ = this.PropPut(0x0000093e, []interface{}{rhs})
}
