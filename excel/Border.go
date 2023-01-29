package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00020854-0000-0000-C000-000000000046
var IID_Border = syscall.GUID{0x00020854, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Border struct {
	ole.OleClient
}

func NewBorder(pDisp *win32.IDispatch, addRef bool, scoped bool) *Border {
	if pDisp == nil {
		return nil
	}
	p := &Border{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func BorderFromVar(v ole.Variant) *Border {
	return NewBorder(v.IDispatch(), false, false)
}

func (this *Border) IID() *syscall.GUID {
	return &IID_Border
}

func (this *Border) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Border) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Border) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Border) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Border) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Border) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Border) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Border) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Border) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Border) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Border) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Border) Color() ole.Variant {
	retVal, _ := this.PropGet(0x00000063, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Border) SetColor(rhs interface{}) {
	_ = this.PropPut(0x00000063, []interface{}{rhs})
}

func (this *Border) ColorIndex() ole.Variant {
	retVal, _ := this.PropGet(0x00000061, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Border) SetColorIndex(rhs interface{}) {
	_ = this.PropPut(0x00000061, []interface{}{rhs})
}

func (this *Border) LineStyle() ole.Variant {
	retVal, _ := this.PropGet(0x00000077, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Border) SetLineStyle(rhs interface{}) {
	_ = this.PropPut(0x00000077, []interface{}{rhs})
}

func (this *Border) Weight() ole.Variant {
	retVal, _ := this.PropGet(0x00000078, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Border) SetWeight(rhs interface{}) {
	_ = this.PropPut(0x00000078, []interface{}{rhs})
}

func (this *Border) ThemeColor() ole.Variant {
	retVal, _ := this.PropGet(0x0000093d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Border) SetThemeColor(rhs interface{}) {
	_ = this.PropPut(0x0000093d, []interface{}{rhs})
}

func (this *Border) TintAndShade() ole.Variant {
	retVal, _ := this.PropGet(0x0000093e, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Border) SetTintAndShade(rhs interface{}) {
	_ = this.PropPut(0x0000093e, []interface{}{rhs})
}
