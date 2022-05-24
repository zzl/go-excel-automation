package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024425-0000-0000-C000-000000000046
var IID_FormatCondition = syscall.GUID{0x00024425, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type FormatCondition struct {
	ole.OleClient
}

func NewFormatCondition(pDisp *win32.IDispatch, addRef bool, scoped bool) *FormatCondition {
	 if pDisp == nil {
		return nil;
	}
	p := &FormatCondition{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FormatConditionFromVar(v ole.Variant) *FormatCondition {
	return NewFormatCondition(v.IDispatch(), false, false)
}

func (this *FormatCondition) IID() *syscall.GUID {
	return &IID_FormatCondition
}

func (this *FormatCondition) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *FormatCondition) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *FormatCondition) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *FormatCondition) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *FormatCondition) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *FormatCondition) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *FormatCondition) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *FormatCondition) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *FormatCondition) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *FormatCondition) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *FormatCondition) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var FormatCondition_Modify__OptArgs= []string{
	"Operator", "Formula1", "Formula2", 
}

func (this *FormatCondition) Modify_(type_ int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(FormatCondition_Modify__OptArgs, optArgs)
	retVal, _ := this.Call(0x00000a3f, []interface{}{type_}, optArgs...)
	_= retVal
}

func (this *FormatCondition) Type() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *FormatCondition) Operator() int32 {
	retVal, _ := this.PropGet(0x0000031d, nil)
	return retVal.LValVal()
}

func (this *FormatCondition) Formula1() string {
	retVal, _ := this.PropGet(0x0000062b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *FormatCondition) Formula2() string {
	retVal, _ := this.PropGet(0x0000062c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *FormatCondition) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *FormatCondition) Borders() *Borders {
	retVal, _ := this.PropGet(0x000001b3, nil)
	return NewBorders(retVal.IDispatch(), false, true)
}

func (this *FormatCondition) Font() *Font {
	retVal, _ := this.PropGet(0x00000092, nil)
	return NewFont(retVal.IDispatch(), false, true)
}

func (this *FormatCondition) Delete()  {
	retVal, _ := this.Call(0x00000075, nil)
	_= retVal
}

var FormatCondition_Modify_OptArgs= []string{
	"Operator", "Formula1", "Formula2", "String", "Operator2", 
}

func (this *FormatCondition) Modify(type_ int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(FormatCondition_Modify_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000062d, []interface{}{type_}, optArgs...)
	_= retVal
}

func (this *FormatCondition) Text() string {
	retVal, _ := this.PropGet(0x0000008a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *FormatCondition) SetText(rhs string)  {
	_ = this.PropPut(0x0000008a, []interface{}{rhs})
}

func (this *FormatCondition) TextOperator() int32 {
	retVal, _ := this.PropGet(0x00000a35, nil)
	return retVal.LValVal()
}

func (this *FormatCondition) SetTextOperator(rhs int32)  {
	_ = this.PropPut(0x00000a35, []interface{}{rhs})
}

func (this *FormatCondition) DateOperator() int32 {
	retVal, _ := this.PropGet(0x00000a36, nil)
	return retVal.LValVal()
}

func (this *FormatCondition) SetDateOperator(rhs int32)  {
	_ = this.PropPut(0x00000a36, []interface{}{rhs})
}

func (this *FormatCondition) NumberFormat() ole.Variant {
	retVal, _ := this.PropGet(0x000000c1, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *FormatCondition) SetNumberFormat(rhs interface{})  {
	_ = this.PropPut(0x000000c1, []interface{}{rhs})
}

func (this *FormatCondition) Priority() int32 {
	retVal, _ := this.PropGet(0x000003d9, nil)
	return retVal.LValVal()
}

func (this *FormatCondition) SetPriority(rhs int32)  {
	_ = this.PropPut(0x000003d9, []interface{}{rhs})
}

func (this *FormatCondition) StopIfTrue() bool {
	retVal, _ := this.PropGet(0x00000a41, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *FormatCondition) SetStopIfTrue(rhs bool)  {
	_ = this.PropPut(0x00000a41, []interface{}{rhs})
}

func (this *FormatCondition) AppliesTo() *Range {
	retVal, _ := this.PropGet(0x00000a42, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *FormatCondition) ModifyAppliesToRange(range_ *Range)  {
	retVal, _ := this.Call(0x00000a43, []interface{}{range_})
	_= retVal
}

func (this *FormatCondition) SetFirstPriority()  {
	retVal, _ := this.Call(0x00000a45, nil)
	_= retVal
}

func (this *FormatCondition) SetLastPriority()  {
	retVal, _ := this.Call(0x00000a46, nil)
	_= retVal
}

func (this *FormatCondition) PTCondition() bool {
	retVal, _ := this.PropGet(0x00000a47, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *FormatCondition) ScopeType() int32 {
	retVal, _ := this.PropGet(0x00000a37, nil)
	return retVal.LValVal()
}

func (this *FormatCondition) SetScopeType(rhs int32)  {
	_ = this.PropPut(0x00000a37, []interface{}{rhs})
}

