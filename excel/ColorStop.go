package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244AD-0000-0000-C000-000000000046
var IID_ColorStop = syscall.GUID{0x000244AD, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ColorStop struct {
	ole.OleClient
}

func NewColorStop(pDisp *win32.IDispatch, addRef bool, scoped bool) *ColorStop {
	 if pDisp == nil {
		return nil;
	}
	p := &ColorStop{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ColorStopFromVar(v ole.Variant) *ColorStop {
	return NewColorStop(v.IDispatch(), false, false)
}

func (this *ColorStop) IID() *syscall.GUID {
	return &IID_ColorStop
}

func (this *ColorStop) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ColorStop) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *ColorStop) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ColorStop) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ColorStop) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *ColorStop) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *ColorStop) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *ColorStop) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *ColorStop) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ColorStop) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ColorStop) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ColorStop) Delete()  {
	retVal, _ := this.Call(0x00000075, nil)
	_= retVal
}

func (this *ColorStop) Color() ole.Variant {
	retVal, _ := this.PropGet(0x00000063, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ColorStop) SetColor(rhs interface{})  {
	_ = this.PropPut(0x00000063, []interface{}{rhs})
}

func (this *ColorStop) ThemeColor() int32 {
	retVal, _ := this.PropGet(0x0000093d, nil)
	return retVal.LValVal()
}

func (this *ColorStop) SetThemeColor(rhs int32)  {
	_ = this.PropPut(0x0000093d, []interface{}{rhs})
}

func (this *ColorStop) TintAndShade() ole.Variant {
	retVal, _ := this.PropGet(0x0000093e, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ColorStop) SetTintAndShade(rhs interface{})  {
	_ = this.PropPut(0x0000093e, []interface{}{rhs})
}

func (this *ColorStop) Position() float64 {
	retVal, _ := this.PropGet(0x00000085, nil)
	return retVal.DblValVal()
}

func (this *ColorStop) SetPosition(rhs float64)  {
	_ = this.PropPut(0x00000085, []interface{}{rhs})
}

