package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020870-0000-0000-C000-000000000046
var IID_Interior = syscall.GUID{0x00020870, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Interior struct {
	ole.OleClient
}

func NewInterior(pDisp *win32.IDispatch, addRef bool, scoped bool) *Interior {
	p := &Interior{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func InteriorFromVar(v ole.Variant) *Interior {
	return NewInterior(v.PdispValVal(), false, false)
}

func (this *Interior) IID() *syscall.GUID {
	return &IID_Interior
}

func (this *Interior) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Interior) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Interior) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Interior) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Interior) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Interior) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Interior) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Interior) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Interior) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Interior) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Interior) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Interior) Color() ole.Variant {
	retVal := this.PropGet(0x00000063, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Interior) SetColor(rhs interface{})  {
	retVal := this.PropPut(0x00000063, []interface{}{rhs})
	_= retVal
}

func (this *Interior) ColorIndex() ole.Variant {
	retVal := this.PropGet(0x00000061, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Interior) SetColorIndex(rhs interface{})  {
	retVal := this.PropPut(0x00000061, []interface{}{rhs})
	_= retVal
}

func (this *Interior) InvertIfNegative() ole.Variant {
	retVal := this.PropGet(0x00000084, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Interior) SetInvertIfNegative(rhs interface{})  {
	retVal := this.PropPut(0x00000084, []interface{}{rhs})
	_= retVal
}

func (this *Interior) Pattern() ole.Variant {
	retVal := this.PropGet(0x0000005f, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Interior) SetPattern(rhs interface{})  {
	retVal := this.PropPut(0x0000005f, []interface{}{rhs})
	_= retVal
}

func (this *Interior) PatternColor() ole.Variant {
	retVal := this.PropGet(0x00000064, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Interior) SetPatternColor(rhs interface{})  {
	retVal := this.PropPut(0x00000064, []interface{}{rhs})
	_= retVal
}

func (this *Interior) PatternColorIndex() ole.Variant {
	retVal := this.PropGet(0x00000062, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Interior) SetPatternColorIndex(rhs interface{})  {
	retVal := this.PropPut(0x00000062, []interface{}{rhs})
	_= retVal
}

func (this *Interior) ThemeColor() ole.Variant {
	retVal := this.PropGet(0x0000093d, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Interior) SetThemeColor(rhs interface{})  {
	retVal := this.PropPut(0x0000093d, []interface{}{rhs})
	_= retVal
}

func (this *Interior) TintAndShade() ole.Variant {
	retVal := this.PropGet(0x0000093e, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Interior) SetTintAndShade(rhs interface{})  {
	retVal := this.PropPut(0x0000093e, []interface{}{rhs})
	_= retVal
}

func (this *Interior) PatternThemeColor() ole.Variant {
	retVal := this.PropGet(0x00000a53, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Interior) SetPatternThemeColor(rhs interface{})  {
	retVal := this.PropPut(0x00000a53, []interface{}{rhs})
	_= retVal
}

func (this *Interior) PatternTintAndShade() ole.Variant {
	retVal := this.PropGet(0x00000a54, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Interior) SetPatternTintAndShade(rhs interface{})  {
	retVal := this.PropPut(0x00000a54, []interface{}{rhs})
	_= retVal
}

func (this *Interior) Gradient() *ole.DispatchClass {
	retVal := this.PropGet(0x00000a55, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

