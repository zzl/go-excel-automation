package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244CC-0000-0000-C000-000000000046
var IID_ProtectedViewWindows = syscall.GUID{0x000244CC, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ProtectedViewWindows struct {
	ole.OleClient
}

func NewProtectedViewWindows(pDisp *win32.IDispatch, addRef bool, scoped bool) *ProtectedViewWindows {
	 if pDisp == nil {
		return nil;
	}
	p := &ProtectedViewWindows{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ProtectedViewWindowsFromVar(v ole.Variant) *ProtectedViewWindows {
	return NewProtectedViewWindows(v.IDispatch(), false, false)
}

func (this *ProtectedViewWindows) IID() *syscall.GUID {
	return &IID_ProtectedViewWindows
}

func (this *ProtectedViewWindows) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ProtectedViewWindows) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *ProtectedViewWindows) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ProtectedViewWindows) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ProtectedViewWindows) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *ProtectedViewWindows) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *ProtectedViewWindows) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *ProtectedViewWindows) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *ProtectedViewWindows) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ProtectedViewWindows) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ProtectedViewWindows) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ProtectedViewWindows) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *ProtectedViewWindows) Item(index interface{}) *ProtectedViewWindow {
	retVal, _ := this.PropGet(0x000000aa, []interface{}{index})
	return NewProtectedViewWindow(retVal.IDispatch(), false, true)
}

func (this *ProtectedViewWindows) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *ProtectedViewWindows) ForEach(action func(item *ProtectedViewWindow) bool) {
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
		pItem := (*ProtectedViewWindow)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *ProtectedViewWindows) Default_(index interface{}) *ProtectedViewWindow {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return NewProtectedViewWindow(retVal.IDispatch(), false, true)
}

var ProtectedViewWindows_Open_OptArgs= []string{
	"Password", "AddToMru", "RepairMode", 
}

func (this *ProtectedViewWindows) Open(filename string, optArgs ...interface{}) *ProtectedViewWindow {
	optArgs = ole.ProcessOptArgs(ProtectedViewWindows_Open_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000783, []interface{}{filename}, optArgs...)
	return NewProtectedViewWindow(retVal.IDispatch(), false, true)
}

