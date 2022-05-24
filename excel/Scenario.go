package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020897-0000-0000-C000-000000000046
var IID_Scenario = syscall.GUID{0x00020897, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Scenario struct {
	ole.OleClient
}

func NewScenario(pDisp *win32.IDispatch, addRef bool, scoped bool) *Scenario {
	 if pDisp == nil {
		return nil;
	}
	p := &Scenario{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ScenarioFromVar(v ole.Variant) *Scenario {
	return NewScenario(v.IDispatch(), false, false)
}

func (this *Scenario) IID() *syscall.GUID {
	return &IID_Scenario
}

func (this *Scenario) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Scenario) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Scenario) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Scenario) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Scenario) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Scenario) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Scenario) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Scenario) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Scenario) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Scenario) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Scenario) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Scenario_ChangeScenario_OptArgs= []string{
	"Values", 
}

func (this *Scenario) ChangeScenario(changingCells interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Scenario_ChangeScenario_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000390, []interface{}{changingCells}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Scenario) ChangingCells() *Range {
	retVal, _ := this.PropGet(0x0000038f, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Scenario) Comment() string {
	retVal, _ := this.PropGet(0x0000038e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Scenario) SetComment(rhs string)  {
	_ = this.PropPut(0x0000038e, []interface{}{rhs})
}

func (this *Scenario) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Scenario) Hidden() bool {
	retVal, _ := this.PropGet(0x0000010c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Scenario) SetHidden(rhs bool)  {
	_ = this.PropPut(0x0000010c, []interface{}{rhs})
}

func (this *Scenario) Index() int32 {
	retVal, _ := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *Scenario) Locked() bool {
	retVal, _ := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Scenario) SetLocked(rhs bool)  {
	_ = this.PropPut(0x0000010d, []interface{}{rhs})
}

func (this *Scenario) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Scenario) SetName(rhs string)  {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *Scenario) Show() ole.Variant {
	retVal, _ := this.Call(0x000001f0, nil)
	com.AddToScope(retVal)
	return *retVal
}

var Scenario_Values_OptArgs= []string{
	"Index", 
}

func (this *Scenario) Values(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Scenario_Values_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x000000a4, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

