package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00024496-0000-0000-C000-000000000046
var IID_Databar = syscall.GUID{0x00024496, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Databar struct {
	ole.OleClient
}

func NewDatabar(pDisp *win32.IDispatch, addRef bool, scoped bool) *Databar {
	if pDisp == nil {
		return nil
	}
	p := &Databar{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DatabarFromVar(v ole.Variant) *Databar {
	return NewDatabar(v.IDispatch(), false, false)
}

func (this *Databar) IID() *syscall.GUID {
	return &IID_Databar
}

func (this *Databar) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Databar) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Databar) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Databar) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Databar) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Databar) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Databar) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Databar) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Databar) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Databar) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Databar) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Databar) Priority() int32 {
	retVal, _ := this.PropGet(0x000003d9, nil)
	return retVal.LValVal()
}

func (this *Databar) SetPriority(rhs int32) {
	_ = this.PropPut(0x000003d9, []interface{}{rhs})
}

func (this *Databar) StopIfTrue() bool {
	retVal, _ := this.PropGet(0x00000a41, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Databar) AppliesTo() *Range {
	retVal, _ := this.PropGet(0x00000a42, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Databar) MinPoint() *ConditionValue {
	retVal, _ := this.PropGet(0x00000a9e, nil)
	return NewConditionValue(retVal.IDispatch(), false, true)
}

func (this *Databar) MaxPoint() *ConditionValue {
	retVal, _ := this.PropGet(0x00000a9f, nil)
	return NewConditionValue(retVal.IDispatch(), false, true)
}

func (this *Databar) PercentMin() int32 {
	retVal, _ := this.PropGet(0x00000aa0, nil)
	return retVal.LValVal()
}

func (this *Databar) SetPercentMin(rhs int32) {
	_ = this.PropPut(0x00000aa0, []interface{}{rhs})
}

func (this *Databar) PercentMax() int32 {
	retVal, _ := this.PropGet(0x00000aa1, nil)
	return retVal.LValVal()
}

func (this *Databar) SetPercentMax(rhs int32) {
	_ = this.PropPut(0x00000aa1, []interface{}{rhs})
}

func (this *Databar) BarColor() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000aa2, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Databar) ShowValue() bool {
	retVal, _ := this.PropGet(0x000007e8, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Databar) SetShowValue(rhs bool) {
	_ = this.PropPut(0x000007e8, []interface{}{rhs})
}

func (this *Databar) Formula() string {
	retVal, _ := this.PropGet(0x00000105, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Databar) SetFormula(rhs string) {
	_ = this.PropPut(0x00000105, []interface{}{rhs})
}

func (this *Databar) Type() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *Databar) SetFirstPriority() {
	retVal, _ := this.Call(0x00000a45, nil)
	_ = retVal
}

func (this *Databar) SetLastPriority() {
	retVal, _ := this.Call(0x00000a46, nil)
	_ = retVal
}

func (this *Databar) Delete() {
	retVal, _ := this.Call(0x00000075, nil)
	_ = retVal
}

func (this *Databar) ModifyAppliesToRange(range_ *Range) {
	retVal, _ := this.Call(0x00000a43, []interface{}{range_})
	_ = retVal
}

func (this *Databar) PTCondition() bool {
	retVal, _ := this.PropGet(0x00000a47, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Databar) ScopeType() int32 {
	retVal, _ := this.PropGet(0x00000a37, nil)
	return retVal.LValVal()
}

func (this *Databar) SetScopeType(rhs int32) {
	_ = this.PropPut(0x00000a37, []interface{}{rhs})
}

func (this *Databar) Direction() int32 {
	retVal, _ := this.PropGet(0x000000a8, nil)
	return retVal.LValVal()
}

func (this *Databar) SetDirection(rhs int32) {
	_ = this.PropPut(0x000000a8, []interface{}{rhs})
}

func (this *Databar) BarFillType() int32 {
	retVal, _ := this.PropGet(0x00000b7d, nil)
	return retVal.LValVal()
}

func (this *Databar) SetBarFillType(rhs int32) {
	_ = this.PropPut(0x00000b7d, []interface{}{rhs})
}

func (this *Databar) AxisPosition() int32 {
	retVal, _ := this.PropGet(0x00000b7e, nil)
	return retVal.LValVal()
}

func (this *Databar) SetAxisPosition(rhs int32) {
	_ = this.PropPut(0x00000b7e, []interface{}{rhs})
}

func (this *Databar) AxisColor() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000b7f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Databar) BarBorder() *DataBarBorder {
	retVal, _ := this.PropGet(0x00000b80, nil)
	return NewDataBarBorder(retVal.IDispatch(), false, true)
}

func (this *Databar) NegativeBarFormat() *NegativeBarFormat {
	retVal, _ := this.PropGet(0x00000b81, nil)
	return NewNegativeBarFormat(retVal.IDispatch(), false, true)
}
