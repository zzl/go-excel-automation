package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002085E-0000-0000-C000-000000000046
var IID_ToolbarButton = syscall.GUID{0x0002085E, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ToolbarButton struct {
	ole.OleClient
}

func NewToolbarButton(pDisp *win32.IDispatch, addRef bool, scoped bool) *ToolbarButton {
	p := &ToolbarButton{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ToolbarButtonFromVar(v ole.Variant) *ToolbarButton {
	return NewToolbarButton(v.PdispValVal(), false, false)
}

func (this *ToolbarButton) IID() *syscall.GUID {
	return &IID_ToolbarButton
}

func (this *ToolbarButton) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ToolbarButton) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *ToolbarButton) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ToolbarButton) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ToolbarButton) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *ToolbarButton) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *ToolbarButton) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *ToolbarButton) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *ToolbarButton) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *ToolbarButton) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ToolbarButton) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ToolbarButton) BuiltIn() bool {
	retVal := this.PropGet(0x00000229, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ToolbarButton) BuiltInFace() bool {
	retVal := this.PropGet(0x0000022a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ToolbarButton) SetBuiltInFace(rhs bool)  {
	retVal := this.PropPut(0x0000022a, []interface{}{rhs})
	_= retVal
}

func (this *ToolbarButton) Copy(toolbar *Toolbar, before int32)  {
	retVal := this.Call(0x00000227, []interface{}{toolbar, before})
	_= retVal
}

func (this *ToolbarButton) CopyFace()  {
	retVal := this.Call(0x000003c6, nil)
	_= retVal
}

func (this *ToolbarButton) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

func (this *ToolbarButton) Edit()  {
	retVal := this.Call(0x00000232, nil)
	_= retVal
}

func (this *ToolbarButton) Enabled() bool {
	retVal := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ToolbarButton) SetEnabled(rhs bool)  {
	retVal := this.PropPut(0x00000258, []interface{}{rhs})
	_= retVal
}

func (this *ToolbarButton) HelpContextID() int32 {
	retVal := this.PropGet(0x00000163, nil)
	return retVal.LValVal()
}

func (this *ToolbarButton) SetHelpContextID(rhs int32)  {
	retVal := this.PropPut(0x00000163, []interface{}{rhs})
	_= retVal
}

func (this *ToolbarButton) HelpFile() string {
	retVal := this.PropGet(0x00000168, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ToolbarButton) SetHelpFile(rhs string)  {
	retVal := this.PropPut(0x00000168, []interface{}{rhs})
	_= retVal
}

func (this *ToolbarButton) ID() int32 {
	retVal := this.PropGet(0x0000023a, nil)
	return retVal.LValVal()
}

func (this *ToolbarButton) IsGap() bool {
	retVal := this.PropGet(0x00000231, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ToolbarButton) Move(toolbar *Toolbar, before int32)  {
	retVal := this.Call(0x0000027d, []interface{}{toolbar, before})
	_= retVal
}

func (this *ToolbarButton) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ToolbarButton) SetName(rhs string)  {
	retVal := this.PropPut(0x0000006e, []interface{}{rhs})
	_= retVal
}

func (this *ToolbarButton) OnAction() string {
	retVal := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ToolbarButton) SetOnAction(rhs string)  {
	retVal := this.PropPut(0x00000254, []interface{}{rhs})
	_= retVal
}

func (this *ToolbarButton) PasteFace()  {
	retVal := this.Call(0x000003c7, nil)
	_= retVal
}

func (this *ToolbarButton) Pushed() bool {
	retVal := this.PropGet(0x00000230, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ToolbarButton) SetPushed(rhs bool)  {
	retVal := this.PropPut(0x00000230, []interface{}{rhs})
	_= retVal
}

func (this *ToolbarButton) Reset()  {
	retVal := this.Call(0x0000022b, nil)
	_= retVal
}

func (this *ToolbarButton) StatusBar() string {
	retVal := this.PropGet(0x00000182, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ToolbarButton) SetStatusBar(rhs string)  {
	retVal := this.PropPut(0x00000182, []interface{}{rhs})
	_= retVal
}

func (this *ToolbarButton) Width() int32 {
	retVal := this.PropGet(0x0000007a, nil)
	return retVal.LValVal()
}

func (this *ToolbarButton) SetWidth(rhs int32)  {
	retVal := this.PropPut(0x0000007a, []interface{}{rhs})
	_= retVal
}

