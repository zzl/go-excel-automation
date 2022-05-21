package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020896-0000-0000-C000-000000000046
var IID_Scenarios = syscall.GUID{0x00020896, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Scenarios struct {
	ole.OleClient
}

func NewScenarios(pDisp *win32.IDispatch, addRef bool, scoped bool) *Scenarios {
	p := &Scenarios{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ScenariosFromVar(v ole.Variant) *Scenarios {
	return NewScenarios(v.PdispValVal(), false, false)
}

func (this *Scenarios) IID() *syscall.GUID {
	return &IID_Scenarios
}

func (this *Scenarios) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Scenarios) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Scenarios) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Scenarios) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Scenarios) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Scenarios) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Scenarios) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Scenarios) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Scenarios) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Scenarios) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Scenarios) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Scenarios_Add_OptArgs= []string{
	"Values", "Comment", "Locked", "Hidden", 
}

func (this *Scenarios) Add(name string, changingCells interface{}, optArgs ...interface{}) *Scenario {
	optArgs = ole.ProcessOptArgs(Scenarios_Add_OptArgs, optArgs)
	retVal := this.Call(0x000000b5, []interface{}{name, changingCells}, optArgs...)
	return NewScenario(retVal.PdispValVal(), false, true)
}

func (this *Scenarios) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

var Scenarios_CreateSummary_OptArgs= []string{
	"ResultCells", 
}

func (this *Scenarios) CreateSummary(reportType int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Scenarios_CreateSummary_OptArgs, optArgs)
	retVal := this.Call(0x00000391, []interface{}{reportType}, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Scenarios) Item(index interface{}) *Scenario {
	retVal := this.Call(0x000000aa, []interface{}{index})
	return NewScenario(retVal.PdispValVal(), false, true)
}

func (this *Scenarios) Merge(source interface{}) ole.Variant {
	retVal := this.Call(0x00000234, []interface{}{source})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Scenarios) NewEnum_() *com.UnknownClass {
	retVal := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Scenarios) ForEach(action func(item *Scenario) bool) {
	pEnum := this.NewEnum_()
	var pEnumVar *win32.IEnumVARIANT
	pEnum.QueryInterface(&win32.IID_IEnumVARIANT, unsafe.Pointer(&pEnumVar))
	defer pEnumVar.Release();
	for {
		var c uint32
		var v ole.Variant
		pEnumVar.Next(1, (*win32.VARIANT)(&v), &c)
		if c == 0 {
			break
		}
		pItem := (*Scenario)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

