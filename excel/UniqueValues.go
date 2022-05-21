package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002449F-0000-0000-C000-000000000046
var IID_UniqueValues = syscall.GUID{0x0002449F, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type UniqueValues struct {
	ole.OleClient
}

func NewUniqueValues(pDisp *win32.IDispatch, addRef bool, scoped bool) *UniqueValues {
	p := &UniqueValues{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func UniqueValuesFromVar(v ole.Variant) *UniqueValues {
	return NewUniqueValues(v.PdispValVal(), false, false)
}

func (this *UniqueValues) IID() *syscall.GUID {
	return &IID_UniqueValues
}

func (this *UniqueValues) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *UniqueValues) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *UniqueValues) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *UniqueValues) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *UniqueValues) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *UniqueValues) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *UniqueValues) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *UniqueValues) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *UniqueValues) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *UniqueValues) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *UniqueValues) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *UniqueValues) Priority() int32 {
	retVal := this.PropGet(0x000003d9, nil)
	return retVal.LValVal()
}

func (this *UniqueValues) SetPriority(rhs int32)  {
	retVal := this.PropPut(0x000003d9, []interface{}{rhs})
	_= retVal
}

func (this *UniqueValues) StopIfTrue() bool {
	retVal := this.PropGet(0x00000a41, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *UniqueValues) SetStopIfTrue(rhs bool)  {
	retVal := this.PropPut(0x00000a41, []interface{}{rhs})
	_= retVal
}

func (this *UniqueValues) AppliesTo() *Range {
	retVal := this.PropGet(0x00000a42, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *UniqueValues) DupeUnique() int32 {
	retVal := this.PropGet(0x00000aad, nil)
	return retVal.LValVal()
}

func (this *UniqueValues) SetDupeUnique(rhs int32)  {
	retVal := this.PropPut(0x00000aad, []interface{}{rhs})
	_= retVal
}

func (this *UniqueValues) Interior() *Interior {
	retVal := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.PdispValVal(), false, true)
}

func (this *UniqueValues) Borders() *Borders {
	retVal := this.PropGet(0x000001b3, nil)
	return NewBorders(retVal.PdispValVal(), false, true)
}

func (this *UniqueValues) Font() *Font {
	retVal := this.PropGet(0x00000092, nil)
	return NewFont(retVal.PdispValVal(), false, true)
}

func (this *UniqueValues) Type() int32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *UniqueValues) NumberFormat() ole.Variant {
	retVal := this.PropGet(0x000000c1, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *UniqueValues) SetNumberFormat(rhs interface{})  {
	retVal := this.PropPut(0x000000c1, []interface{}{rhs})
	_= retVal
}

func (this *UniqueValues) SetFirstPriority()  {
	retVal := this.Call(0x00000a45, nil)
	_= retVal
}

func (this *UniqueValues) SetLastPriority()  {
	retVal := this.Call(0x00000a46, nil)
	_= retVal
}

func (this *UniqueValues) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

func (this *UniqueValues) ModifyAppliesToRange(range_ *Range)  {
	retVal := this.Call(0x00000a43, []interface{}{range_})
	_= retVal
}

func (this *UniqueValues) PTCondition() bool {
	retVal := this.PropGet(0x00000a47, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *UniqueValues) ScopeType() int32 {
	retVal := this.PropGet(0x00000a37, nil)
	return retVal.LValVal()
}

func (this *UniqueValues) SetScopeType(rhs int32)  {
	retVal := this.PropPut(0x00000a37, []interface{}{rhs})
	_= retVal
}

