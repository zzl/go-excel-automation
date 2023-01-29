package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 0002085F-0000-0000-C000-000000000046
var IID_ToolbarButtons = syscall.GUID{0x0002085F, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ToolbarButtons struct {
	ole.OleClient
}

func NewToolbarButtons(pDisp *win32.IDispatch, addRef bool, scoped bool) *ToolbarButtons {
	if pDisp == nil {
		return nil
	}
	p := &ToolbarButtons{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ToolbarButtonsFromVar(v ole.Variant) *ToolbarButtons {
	return NewToolbarButtons(v.IDispatch(), false, false)
}

func (this *ToolbarButtons) IID() *syscall.GUID {
	return &IID_ToolbarButtons
}

func (this *ToolbarButtons) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ToolbarButtons) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *ToolbarButtons) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ToolbarButtons) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ToolbarButtons) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *ToolbarButtons) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *ToolbarButtons) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *ToolbarButtons) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *ToolbarButtons) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ToolbarButtons) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ToolbarButtons) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var ToolbarButtons_Add_OptArgs = []string{
	"Button", "Before", "OnAction", "Pushed",
	"Enabled", "StatusBar", "HelpFile", "HelpContextID",
}

func (this *ToolbarButtons) Add(optArgs ...interface{}) *ToolbarButton {
	optArgs = ole.ProcessOptArgs(ToolbarButtons_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b5, nil, optArgs...)
	return NewToolbarButton(retVal.IDispatch(), false, true)
}

func (this *ToolbarButtons) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *ToolbarButtons) Item(index int32) *ToolbarButton {
	retVal, _ := this.PropGet(0x000000aa, []interface{}{index})
	return NewToolbarButton(retVal.IDispatch(), false, true)
}

func (this *ToolbarButtons) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *ToolbarButtons) ForEach(action func(item *ToolbarButton) bool) {
	pEnum := this.NewEnum_()
	var pEnumVar *win32.IEnumVARIANT
	pEnum.QueryInterface(&win32.IID_IEnumVARIANT, unsafe.Pointer(&pEnumVar))
	defer pEnumVar.Release()
	for {
		var c uint32
		var v ole.Variant
		pEnumVar.Next(1, (*win32.VARIANT)(&v), &c)
		if c == 0 {
			break
		}
		pItem := (*ToolbarButton)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *ToolbarButtons) Default_(index int32) *ToolbarButton {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return NewToolbarButton(retVal.IDispatch(), false, true)
}
