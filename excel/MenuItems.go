package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020867-0000-0000-C000-000000000046
var IID_MenuItems = syscall.GUID{0x00020867, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type MenuItems struct {
	ole.OleClient
}

func NewMenuItems(pDisp *win32.IDispatch, addRef bool, scoped bool) *MenuItems {
	p := &MenuItems{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func MenuItemsFromVar(v ole.Variant) *MenuItems {
	return NewMenuItems(v.PdispValVal(), false, false)
}

func (this *MenuItems) IID() *syscall.GUID {
	return &IID_MenuItems
}

func (this *MenuItems) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *MenuItems) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *MenuItems) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *MenuItems) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *MenuItems) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *MenuItems) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *MenuItems) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *MenuItems) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *MenuItems) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *MenuItems) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *MenuItems) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var MenuItems_Add_OptArgs= []string{
	"OnAction", "ShortcutKey", "Before", "Restore", 
	"StatusBar", "HelpFile", "HelpContextID", 
}

func (this *MenuItems) Add(caption string, optArgs ...interface{}) *MenuItem {
	optArgs = ole.ProcessOptArgs(MenuItems_Add_OptArgs, optArgs)
	retVal := this.Call(0x000000b5, []interface{}{caption}, optArgs...)
	return NewMenuItem(retVal.PdispValVal(), false, true)
}

var MenuItems_AddMenu_OptArgs= []string{
	"Before", "Restore", 
}

func (this *MenuItems) AddMenu(caption string, optArgs ...interface{}) *Menu {
	optArgs = ole.ProcessOptArgs(MenuItems_AddMenu_OptArgs, optArgs)
	retVal := this.Call(0x00000256, []interface{}{caption}, optArgs...)
	return NewMenu(retVal.PdispValVal(), false, true)
}

func (this *MenuItems) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *MenuItems) Default_(index interface{}) *ole.DispatchClass {
	retVal := this.PropGet(0x00000000, []interface{}{index})
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *MenuItems) Item(index interface{}) *ole.DispatchClass {
	retVal := this.PropGet(0x000000aa, []interface{}{index})
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *MenuItems) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *MenuItems) ForEach(action func(item *ole.DispatchClass) bool) {
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
		pItem := (*ole.DispatchClass)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

