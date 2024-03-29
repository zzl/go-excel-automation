package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00020892-0000-0000-C000-000000000046
var IID_Windows = syscall.GUID{0x00020892, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Windows struct {
	ole.OleClient
}

func NewWindows(pDisp *win32.IDispatch, addRef bool, scoped bool) *Windows {
	if pDisp == nil {
		return nil
	}
	p := &Windows{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func WindowsFromVar(v ole.Variant) *Windows {
	return NewWindows(v.IDispatch(), false, false)
}

func (this *Windows) IID() *syscall.GUID {
	return &IID_Windows
}

func (this *Windows) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Windows) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Windows) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Windows) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Windows) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Windows) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Windows) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Windows) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Windows) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Windows) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Windows) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Windows_Arrange_OptArgs = []string{
	"ArrangeStyle", "ActiveWorkbook", "SyncHorizontal", "SyncVertical",
}

func (this *Windows) Arrange(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Windows_Arrange_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000027e, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Windows) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *Windows) Item(index interface{}) *Window {
	retVal, _ := this.PropGet(0x000000aa, []interface{}{index})
	return NewWindow(retVal.IDispatch(), false, true)
}

func (this *Windows) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Windows) ForEach(action func(item *Window) bool) {
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
		pItem := (*Window)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Windows) Default_(index interface{}) *Window {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return NewWindow(retVal.IDispatch(), false, true)
}

func (this *Windows) CompareSideBySideWith(windowName interface{}) bool {
	retVal, _ := this.Call(0x000008c6, []interface{}{windowName})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Windows) BreakSideBySide() bool {
	retVal, _ := this.Call(0x000008c8, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Windows) SyncScrollingSideBySide() bool {
	retVal, _ := this.PropGet(0x000008c9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Windows) SetSyncScrollingSideBySide(rhs bool) {
	_ = this.PropPut(0x000008c9, []interface{}{rhs})
}

func (this *Windows) ResetPositionsSideBySide() {
	retVal, _ := this.Call(0x000008ca, nil)
	_ = retVal
}
