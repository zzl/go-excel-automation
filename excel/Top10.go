package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002449D-0000-0000-C000-000000000046
var IID_Top10 = syscall.GUID{0x0002449D, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Top10 struct {
	ole.OleClient
}

func NewTop10(pDisp *win32.IDispatch, addRef bool, scoped bool) *Top10 {
	p := &Top10{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func Top10FromVar(v ole.Variant) *Top10 {
	return NewTop10(v.PdispValVal(), false, false)
}

func (this *Top10) IID() *syscall.GUID {
	return &IID_Top10
}

func (this *Top10) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Top10) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Top10) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Top10) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Top10) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Top10) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Top10) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Top10) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Top10) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Top10) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Top10) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Top10) Priority() int32 {
	retVal := this.PropGet(0x000003d9, nil)
	return retVal.LValVal()
}

func (this *Top10) SetPriority(rhs int32)  {
	retVal := this.PropPut(0x000003d9, []interface{}{rhs})
	_= retVal
}

func (this *Top10) StopIfTrue() bool {
	retVal := this.PropGet(0x00000a41, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Top10) SetStopIfTrue(rhs bool)  {
	retVal := this.PropPut(0x00000a41, []interface{}{rhs})
	_= retVal
}

func (this *Top10) AppliesTo() *Range {
	retVal := this.PropGet(0x00000a42, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Top10) TopBottom() int32 {
	retVal := this.PropGet(0x00000aa8, nil)
	return retVal.LValVal()
}

func (this *Top10) SetTopBottom(rhs int32)  {
	retVal := this.PropPut(0x00000aa8, []interface{}{rhs})
	_= retVal
}

func (this *Top10) Rank() int32 {
	retVal := this.PropGet(0x0000050a, nil)
	return retVal.LValVal()
}

func (this *Top10) SetRank(rhs int32)  {
	retVal := this.PropPut(0x0000050a, []interface{}{rhs})
	_= retVal
}

func (this *Top10) Percent() bool {
	retVal := this.PropGet(0x00000aa9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Top10) SetPercent(rhs bool)  {
	retVal := this.PropPut(0x00000aa9, []interface{}{rhs})
	_= retVal
}

func (this *Top10) Interior() *Interior {
	retVal := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.PdispValVal(), false, true)
}

func (this *Top10) Borders() *Borders {
	retVal := this.PropGet(0x000001b3, nil)
	return NewBorders(retVal.PdispValVal(), false, true)
}

func (this *Top10) Font() *Font {
	retVal := this.PropGet(0x00000092, nil)
	return NewFont(retVal.PdispValVal(), false, true)
}

func (this *Top10) Type() int32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *Top10) NumberFormat() ole.Variant {
	retVal := this.PropGet(0x000000c1, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Top10) SetNumberFormat(rhs interface{})  {
	retVal := this.PropPut(0x000000c1, []interface{}{rhs})
	_= retVal
}

func (this *Top10) SetFirstPriority()  {
	retVal := this.Call(0x00000a45, nil)
	_= retVal
}

func (this *Top10) SetLastPriority()  {
	retVal := this.Call(0x00000a46, nil)
	_= retVal
}

func (this *Top10) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

func (this *Top10) ModifyAppliesToRange(range_ *Range)  {
	retVal := this.Call(0x00000a43, []interface{}{range_})
	_= retVal
}

func (this *Top10) PTCondition() bool {
	retVal := this.PropGet(0x00000a47, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Top10) ScopeType() int32 {
	retVal := this.PropGet(0x00000a37, nil)
	return retVal.LValVal()
}

func (this *Top10) SetScopeType(rhs int32)  {
	retVal := this.PropPut(0x00000a37, []interface{}{rhs})
	_= retVal
}

func (this *Top10) CalcFor() int32 {
	retVal := this.PropGet(0x00000aaa, nil)
	return retVal.LValVal()
}

func (this *Top10) SetCalcFor(rhs int32)  {
	retVal := this.PropPut(0x00000aaa, []interface{}{rhs})
	_= retVal
}

