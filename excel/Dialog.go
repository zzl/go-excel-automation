package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 0002087A-0000-0000-C000-000000000046
var IID_Dialog = syscall.GUID{0x0002087A, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Dialog struct {
	ole.OleClient
}

func NewDialog(pDisp *win32.IDispatch, addRef bool, scoped bool) *Dialog {
	if pDisp == nil {
		return nil
	}
	p := &Dialog{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DialogFromVar(v ole.Variant) *Dialog {
	return NewDialog(v.IDispatch(), false, false)
}

func (this *Dialog) IID() *syscall.GUID {
	return &IID_Dialog
}

func (this *Dialog) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Dialog) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Dialog) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Dialog) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Dialog) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Dialog) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Dialog) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Dialog) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Dialog) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Dialog) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Dialog) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Dialog_Show_OptArgs = []string{
	"Arg1", "Arg2", "Arg3", "Arg4",
	"Arg5", "Arg6", "Arg7", "Arg8",
	"Arg9", "Arg10", "Arg11", "Arg12",
	"Arg13", "Arg14", "Arg15", "Arg16",
	"Arg17", "Arg18", "Arg19", "Arg20",
	"Arg21", "Arg22", "Arg23", "Arg24",
	"Arg25", "Arg26", "Arg27", "Arg28",
	"Arg29", "Arg30",
}

func (this *Dialog) Show(optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(Dialog_Show_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001f0, nil, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}
