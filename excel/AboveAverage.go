package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 0002449E-0000-0000-C000-000000000046
var IID_AboveAverage = syscall.GUID{0x0002449E, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type AboveAverage struct {
	ole.OleClient
}

func NewAboveAverage(pDisp *win32.IDispatch, addRef bool, scoped bool) *AboveAverage {
	if pDisp == nil {
		return nil
	}
	p := &AboveAverage{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func AboveAverageFromVar(v ole.Variant) *AboveAverage {
	return NewAboveAverage(v.IDispatch(), false, false)
}

func (this *AboveAverage) IID() *syscall.GUID {
	return &IID_AboveAverage
}

func (this *AboveAverage) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *AboveAverage) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *AboveAverage) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *AboveAverage) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *AboveAverage) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *AboveAverage) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *AboveAverage) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *AboveAverage) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *AboveAverage) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *AboveAverage) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *AboveAverage) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *AboveAverage) Priority() int32 {
	retVal, _ := this.PropGet(0x000003d9, nil)
	return retVal.LValVal()
}

func (this *AboveAverage) SetPriority(rhs int32) {
	_ = this.PropPut(0x000003d9, []interface{}{rhs})
}

func (this *AboveAverage) StopIfTrue() bool {
	retVal, _ := this.PropGet(0x00000a41, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AboveAverage) SetStopIfTrue(rhs bool) {
	_ = this.PropPut(0x00000a41, []interface{}{rhs})
}

func (this *AboveAverage) AppliesTo() *Range {
	retVal, _ := this.PropGet(0x00000a42, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *AboveAverage) AboveBelow() int32 {
	retVal, _ := this.PropGet(0x00000aab, nil)
	return retVal.LValVal()
}

func (this *AboveAverage) SetAboveBelow(rhs int32) {
	_ = this.PropPut(0x00000aab, []interface{}{rhs})
}

func (this *AboveAverage) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *AboveAverage) Borders() *Borders {
	retVal, _ := this.PropGet(0x000001b3, nil)
	return NewBorders(retVal.IDispatch(), false, true)
}

func (this *AboveAverage) Font() *Font {
	retVal, _ := this.PropGet(0x00000092, nil)
	return NewFont(retVal.IDispatch(), false, true)
}

func (this *AboveAverage) Type() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *AboveAverage) NumberFormat() ole.Variant {
	retVal, _ := this.PropGet(0x000000c1, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *AboveAverage) SetNumberFormat(rhs interface{}) {
	_ = this.PropPut(0x000000c1, []interface{}{rhs})
}

func (this *AboveAverage) SetFirstPriority() {
	retVal, _ := this.Call(0x00000a45, nil)
	_ = retVal
}

func (this *AboveAverage) SetLastPriority() {
	retVal, _ := this.Call(0x00000a46, nil)
	_ = retVal
}

func (this *AboveAverage) Delete() {
	retVal, _ := this.Call(0x00000075, nil)
	_ = retVal
}

func (this *AboveAverage) ModifyAppliesToRange(range_ *Range) {
	retVal, _ := this.Call(0x00000a43, []interface{}{range_})
	_ = retVal
}

func (this *AboveAverage) PTCondition() bool {
	retVal, _ := this.PropGet(0x00000a47, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AboveAverage) ScopeType() int32 {
	retVal, _ := this.PropGet(0x00000a37, nil)
	return retVal.LValVal()
}

func (this *AboveAverage) SetScopeType(rhs int32) {
	_ = this.PropPut(0x00000a37, []interface{}{rhs})
}

func (this *AboveAverage) CalcFor() int32 {
	retVal, _ := this.PropGet(0x00000aaa, nil)
	return retVal.LValVal()
}

func (this *AboveAverage) SetCalcFor(rhs int32) {
	_ = this.PropPut(0x00000aaa, []interface{}{rhs})
}

func (this *AboveAverage) NumStdDev() int32 {
	retVal, _ := this.PropGet(0x00000aac, nil)
	return retVal.LValVal()
}

func (this *AboveAverage) SetNumStdDev(rhs int32) {
	_ = this.PropPut(0x00000aac, []interface{}{rhs})
}
