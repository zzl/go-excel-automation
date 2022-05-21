package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024493-0000-0000-C000-000000000046
var IID_ColorScale = syscall.GUID{0x00024493, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ColorScale struct {
	ole.OleClient
}

func NewColorScale(pDisp *win32.IDispatch, addRef bool, scoped bool) *ColorScale {
	p := &ColorScale{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ColorScaleFromVar(v ole.Variant) *ColorScale {
	return NewColorScale(v.PdispValVal(), false, false)
}

func (this *ColorScale) IID() *syscall.GUID {
	return &IID_ColorScale
}

func (this *ColorScale) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ColorScale) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *ColorScale) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ColorScale) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ColorScale) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *ColorScale) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *ColorScale) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *ColorScale) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *ColorScale) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *ColorScale) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ColorScale) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ColorScale) Priority() int32 {
	retVal := this.PropGet(0x000003d9, nil)
	return retVal.LValVal()
}

func (this *ColorScale) SetPriority(rhs int32)  {
	retVal := this.PropPut(0x000003d9, []interface{}{rhs})
	_= retVal
}

func (this *ColorScale) StopIfTrue() bool {
	retVal := this.PropGet(0x00000a41, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ColorScale) AppliesTo() *Range {
	retVal := this.PropGet(0x00000a42, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *ColorScale) Formula() string {
	retVal := this.PropGet(0x00000105, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ColorScale) SetFormula(rhs string)  {
	retVal := this.PropPut(0x00000105, []interface{}{rhs})
	_= retVal
}

func (this *ColorScale) Type() int32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *ColorScale) SetFirstPriority()  {
	retVal := this.Call(0x00000a45, nil)
	_= retVal
}

func (this *ColorScale) SetLastPriority()  {
	retVal := this.Call(0x00000a46, nil)
	_= retVal
}

func (this *ColorScale) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

func (this *ColorScale) ModifyAppliesToRange(range_ *Range)  {
	retVal := this.Call(0x00000a43, []interface{}{range_})
	_= retVal
}

func (this *ColorScale) PTCondition() bool {
	retVal := this.PropGet(0x00000a47, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ColorScale) ScopeType() int32 {
	retVal := this.PropGet(0x00000a37, nil)
	return retVal.LValVal()
}

func (this *ColorScale) SetScopeType(rhs int32)  {
	retVal := this.PropPut(0x00000a37, []interface{}{rhs})
	_= retVal
}

func (this *ColorScale) ColorScaleCriteria() *ColorScaleCriteria {
	retVal := this.PropGet(0x00000a9c, nil)
	return NewColorScaleCriteria(retVal.PdispValVal(), false, true)
}

