package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024497-0000-0000-C000-000000000046
var IID_IconSetCondition = syscall.GUID{0x00024497, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IconSetCondition struct {
	ole.OleClient
}

func NewIconSetCondition(pDisp *win32.IDispatch, addRef bool, scoped bool) *IconSetCondition {
	p := &IconSetCondition{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func IconSetConditionFromVar(v ole.Variant) *IconSetCondition {
	return NewIconSetCondition(v.PdispValVal(), false, false)
}

func (this *IconSetCondition) IID() *syscall.GUID {
	return &IID_IconSetCondition
}

func (this *IconSetCondition) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *IconSetCondition) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *IconSetCondition) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *IconSetCondition) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *IconSetCondition) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *IconSetCondition) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *IconSetCondition) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *IconSetCondition) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *IconSetCondition) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *IconSetCondition) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *IconSetCondition) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *IconSetCondition) Priority() int32 {
	retVal := this.PropGet(0x000003d9, nil)
	return retVal.LValVal()
}

func (this *IconSetCondition) SetPriority(rhs int32)  {
	retVal := this.PropPut(0x000003d9, []interface{}{rhs})
	_= retVal
}

func (this *IconSetCondition) StopIfTrue() bool {
	retVal := this.PropGet(0x00000a41, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *IconSetCondition) AppliesTo() *Range {
	retVal := this.PropGet(0x00000a42, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *IconSetCondition) Type() int32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *IconSetCondition) ModifyAppliesToRange(range_ *Range)  {
	retVal := this.Call(0x00000a43, []interface{}{range_})
	_= retVal
}

func (this *IconSetCondition) PTCondition() bool {
	retVal := this.PropGet(0x00000a47, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *IconSetCondition) ScopeType() int32 {
	retVal := this.PropGet(0x00000a37, nil)
	return retVal.LValVal()
}

func (this *IconSetCondition) SetScopeType(rhs int32)  {
	retVal := this.PropPut(0x00000a37, []interface{}{rhs})
	_= retVal
}

func (this *IconSetCondition) SetFirstPriority()  {
	retVal := this.Call(0x00000a45, nil)
	_= retVal
}

func (this *IconSetCondition) SetLastPriority()  {
	retVal := this.Call(0x00000a46, nil)
	_= retVal
}

func (this *IconSetCondition) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

func (this *IconSetCondition) ReverseOrder() bool {
	retVal := this.PropGet(0x00000aa3, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *IconSetCondition) SetReverseOrder(rhs bool)  {
	retVal := this.PropPut(0x00000aa3, []interface{}{rhs})
	_= retVal
}

func (this *IconSetCondition) PercentileValues() bool {
	retVal := this.PropGet(0x00000aa4, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *IconSetCondition) SetPercentileValues(rhs bool)  {
	retVal := this.PropPut(0x00000aa4, []interface{}{rhs})
	_= retVal
}

func (this *IconSetCondition) ShowIconOnly() bool {
	retVal := this.PropGet(0x00000aa5, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *IconSetCondition) SetShowIconOnly(rhs bool)  {
	retVal := this.PropPut(0x00000aa5, []interface{}{rhs})
	_= retVal
}

func (this *IconSetCondition) Formula() string {
	retVal := this.PropGet(0x00000105, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *IconSetCondition) SetFormula(rhs string)  {
	retVal := this.PropPut(0x00000105, []interface{}{rhs})
	_= retVal
}

func (this *IconSetCondition) IconSet() ole.Variant {
	retVal := this.PropGet(0x00000aa6, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *IconSetCondition) SetIconSet(rhs interface{})  {
	retVal := this.PropPut(0x00000aa6, []interface{}{rhs})
	_= retVal
}

func (this *IconSetCondition) IconCriteria() *IconCriteria {
	retVal := this.PropGet(0x00000aa7, nil)
	return NewIconCriteria(retVal.PdispValVal(), false, true)
}
