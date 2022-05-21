package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000208D4-0000-0000-C000-000000000046
var IID_AutoCorrect = syscall.GUID{0x000208D4, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type AutoCorrect struct {
	ole.OleClient
}

func NewAutoCorrect(pDisp *win32.IDispatch, addRef bool, scoped bool) *AutoCorrect {
	p := &AutoCorrect{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func AutoCorrectFromVar(v ole.Variant) *AutoCorrect {
	return NewAutoCorrect(v.PdispValVal(), false, false)
}

func (this *AutoCorrect) IID() *syscall.GUID {
	return &IID_AutoCorrect
}

func (this *AutoCorrect) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *AutoCorrect) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *AutoCorrect) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *AutoCorrect) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *AutoCorrect) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *AutoCorrect) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *AutoCorrect) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *AutoCorrect) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *AutoCorrect) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *AutoCorrect) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *AutoCorrect) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *AutoCorrect) AddReplacement(what string, replacement string) ole.Variant {
	retVal := this.Call(0x0000047a, []interface{}{what, replacement})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *AutoCorrect) CapitalizeNamesOfDays() bool {
	retVal := this.PropGet(0x0000047e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetCapitalizeNamesOfDays(rhs bool)  {
	retVal := this.PropPut(0x0000047e, []interface{}{rhs})
	_= retVal
}

func (this *AutoCorrect) DeleteReplacement(what string) ole.Variant {
	retVal := this.Call(0x0000047b, []interface{}{what})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var AutoCorrect_ReplacementList_OptArgs= []string{
	"Index", 
}

func (this *AutoCorrect) ReplacementList(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(AutoCorrect_ReplacementList_OptArgs, optArgs)
	retVal := this.PropGet(0x0000047f, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var AutoCorrect_SetReplacementList_OptArgs= []string{
	"rhs", 
}

func (this *AutoCorrect) SetReplacementList(index interface{}, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(AutoCorrect_SetReplacementList_OptArgs, optArgs)
	retVal := this.PropPut(0x0000047f, []interface{}{index}, optArgs...)
	_= retVal
}

func (this *AutoCorrect) ReplaceText() bool {
	retVal := this.PropGet(0x0000047c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetReplaceText(rhs bool)  {
	retVal := this.PropPut(0x0000047c, []interface{}{rhs})
	_= retVal
}

func (this *AutoCorrect) TwoInitialCapitals() bool {
	retVal := this.PropGet(0x0000047d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetTwoInitialCapitals(rhs bool)  {
	retVal := this.PropPut(0x0000047d, []interface{}{rhs})
	_= retVal
}

func (this *AutoCorrect) CorrectSentenceCap() bool {
	retVal := this.PropGet(0x00000653, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetCorrectSentenceCap(rhs bool)  {
	retVal := this.PropPut(0x00000653, []interface{}{rhs})
	_= retVal
}

func (this *AutoCorrect) CorrectCapsLock() bool {
	retVal := this.PropGet(0x00000654, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetCorrectCapsLock(rhs bool)  {
	retVal := this.PropPut(0x00000654, []interface{}{rhs})
	_= retVal
}

func (this *AutoCorrect) DisplayAutoCorrectOptions() bool {
	retVal := this.PropGet(0x00000786, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetDisplayAutoCorrectOptions(rhs bool)  {
	retVal := this.PropPut(0x00000786, []interface{}{rhs})
	_= retVal
}

func (this *AutoCorrect) AutoExpandListRange() bool {
	retVal := this.PropGet(0x000008f6, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetAutoExpandListRange(rhs bool)  {
	retVal := this.PropPut(0x000008f6, []interface{}{rhs})
	_= retVal
}

func (this *AutoCorrect) AutoFillFormulasInLists() bool {
	retVal := this.PropGet(0x00000a52, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetAutoFillFormulasInLists(rhs bool)  {
	retVal := this.PropPut(0x00000a52, []interface{}{rhs})
	_= retVal
}

