package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000208B9-0000-0000-C000-000000000046
var IID_Name = syscall.GUID{0x000208B9, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Name struct {
	ole.OleClient
}

func NewName(pDisp *win32.IDispatch, addRef bool, scoped bool) *Name {
	p := &Name{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func NameFromVar(v ole.Variant) *Name {
	return NewName(v.PdispValVal(), false, false)
}

func (this *Name) IID() *syscall.GUID {
	return &IID_Name
}

func (this *Name) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Name) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Name) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Name) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Name) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Name) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Name) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Name) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Name) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Name) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Name) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Name) Default_() string {
	retVal := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Name) Index() int32 {
	retVal := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *Name) Category() string {
	retVal := this.PropGet(0x000003a6, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Name) SetCategory(rhs string)  {
	retVal := this.PropPut(0x000003a6, []interface{}{rhs})
	_= retVal
}

func (this *Name) CategoryLocal() string {
	retVal := this.PropGet(0x000003a7, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Name) SetCategoryLocal(rhs string)  {
	retVal := this.PropPut(0x000003a7, []interface{}{rhs})
	_= retVal
}

func (this *Name) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

func (this *Name) MacroType() int32 {
	retVal := this.PropGet(0x000003a8, nil)
	return retVal.LValVal()
}

func (this *Name) SetMacroType(rhs int32)  {
	retVal := this.PropPut(0x000003a8, []interface{}{rhs})
	_= retVal
}

func (this *Name) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Name) SetName(rhs string)  {
	retVal := this.PropPut(0x0000006e, []interface{}{rhs})
	_= retVal
}

func (this *Name) RefersTo() ole.Variant {
	retVal := this.PropGet(0x000003aa, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Name) SetRefersTo(rhs interface{})  {
	retVal := this.PropPut(0x000003aa, []interface{}{rhs})
	_= retVal
}

func (this *Name) ShortcutKey() string {
	retVal := this.PropGet(0x00000255, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Name) SetShortcutKey(rhs string)  {
	retVal := this.PropPut(0x00000255, []interface{}{rhs})
	_= retVal
}

func (this *Name) Value() string {
	retVal := this.PropGet(0x00000006, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Name) SetValue(rhs string)  {
	retVal := this.PropPut(0x00000006, []interface{}{rhs})
	_= retVal
}

func (this *Name) Visible() bool {
	retVal := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Name) SetVisible(rhs bool)  {
	retVal := this.PropPut(0x0000022e, []interface{}{rhs})
	_= retVal
}

func (this *Name) NameLocal() string {
	retVal := this.PropGet(0x000003a9, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Name) SetNameLocal(rhs string)  {
	retVal := this.PropPut(0x000003a9, []interface{}{rhs})
	_= retVal
}

func (this *Name) RefersToLocal() ole.Variant {
	retVal := this.PropGet(0x000003ab, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Name) SetRefersToLocal(rhs interface{})  {
	retVal := this.PropPut(0x000003ab, []interface{}{rhs})
	_= retVal
}

func (this *Name) RefersToR1C1() ole.Variant {
	retVal := this.PropGet(0x000003ac, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Name) SetRefersToR1C1(rhs interface{})  {
	retVal := this.PropPut(0x000003ac, []interface{}{rhs})
	_= retVal
}

func (this *Name) RefersToR1C1Local() ole.Variant {
	retVal := this.PropGet(0x000003ad, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Name) SetRefersToR1C1Local(rhs interface{})  {
	retVal := this.PropPut(0x000003ad, []interface{}{rhs})
	_= retVal
}

func (this *Name) RefersToRange() *Range {
	retVal := this.PropGet(0x00000488, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Name) Comment() string {
	retVal := this.PropGet(0x0000038e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Name) SetComment(rhs string)  {
	retVal := this.PropPut(0x0000038e, []interface{}{rhs})
	_= retVal
}

func (this *Name) WorkbookParameter() bool {
	retVal := this.PropGet(0x00000a2f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Name) SetWorkbookParameter(rhs bool)  {
	retVal := this.PropPut(0x00000a2f, []interface{}{rhs})
	_= retVal
}

func (this *Name) ValidWorkbookParameter() bool {
	retVal := this.PropGet(0x00000a30, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

