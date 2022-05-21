package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002445E-0000-0000-C000-000000000046
var IID_SmartTagAction = syscall.GUID{0x0002445E, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type SmartTagAction struct {
	ole.OleClient
}

func NewSmartTagAction(pDisp *win32.IDispatch, addRef bool, scoped bool) *SmartTagAction {
	p := &SmartTagAction{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SmartTagActionFromVar(v ole.Variant) *SmartTagAction {
	return NewSmartTagAction(v.PdispValVal(), false, false)
}

func (this *SmartTagAction) IID() *syscall.GUID {
	return &IID_SmartTagAction
}

func (this *SmartTagAction) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SmartTagAction) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *SmartTagAction) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *SmartTagAction) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *SmartTagAction) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *SmartTagAction) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *SmartTagAction) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *SmartTagAction) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *SmartTagAction) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *SmartTagAction) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *SmartTagAction) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *SmartTagAction) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SmartTagAction) Execute()  {
	retVal := this.Call(0x000008a3, nil)
	_= retVal
}

func (this *SmartTagAction) Default_() string {
	retVal := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SmartTagAction) Type() int32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *SmartTagAction) PresentInPane() bool {
	retVal := this.PropGet(0x000008f9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SmartTagAction) ExpandHelp() bool {
	retVal := this.PropGet(0x000008fa, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SmartTagAction) SetExpandHelp(rhs bool)  {
	retVal := this.PropPut(0x000008fa, []interface{}{rhs})
	_= retVal
}

func (this *SmartTagAction) CheckboxState() bool {
	retVal := this.PropGet(0x000008fb, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SmartTagAction) SetCheckboxState(rhs bool)  {
	retVal := this.PropPut(0x000008fb, []interface{}{rhs})
	_= retVal
}

func (this *SmartTagAction) TextboxText() string {
	retVal := this.PropGet(0x000008fc, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SmartTagAction) SetTextboxText(rhs string)  {
	retVal := this.PropPut(0x000008fc, []interface{}{rhs})
	_= retVal
}

func (this *SmartTagAction) ListSelection() int32 {
	retVal := this.PropGet(0x000008fd, nil)
	return retVal.LValVal()
}

func (this *SmartTagAction) SetListSelection(rhs int32)  {
	retVal := this.PropPut(0x000008fd, []interface{}{rhs})
	_= retVal
}

func (this *SmartTagAction) RadioGroupSelection() int32 {
	retVal := this.PropGet(0x000008fe, nil)
	return retVal.LValVal()
}

func (this *SmartTagAction) SetRadioGroupSelection(rhs int32)  {
	retVal := this.PropPut(0x000008fe, []interface{}{rhs})
	_= retVal
}

func (this *SmartTagAction) ActiveXControl() *ole.DispatchClass {
	retVal := this.PropGet(0x000008ff, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

