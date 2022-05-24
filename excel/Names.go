package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000208B8-0000-0000-C000-000000000046
var IID_Names = syscall.GUID{0x000208B8, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Names struct {
	ole.OleClient
}

func NewNames(pDisp *win32.IDispatch, addRef bool, scoped bool) *Names {
	 if pDisp == nil {
		return nil;
	}
	p := &Names{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func NamesFromVar(v ole.Variant) *Names {
	return NewNames(v.IDispatch(), false, false)
}

func (this *Names) IID() *syscall.GUID {
	return &IID_Names
}

func (this *Names) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Names) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Names) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Names) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Names) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Names) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Names) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Names) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Names) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Names) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Names) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Names_Add_OptArgs= []string{
	"Name", "RefersTo", "Visible", "MacroType", 
	"ShortcutKey", "Category", "NameLocal", "RefersToLocal", 
	"CategoryLocal", "RefersToR1C1", "RefersToR1C1Local", 
}

func (this *Names) Add(optArgs ...interface{}) *Name {
	optArgs = ole.ProcessOptArgs(Names_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b5, nil, optArgs...)
	return NewName(retVal.IDispatch(), false, true)
}

var Names_Item_OptArgs= []string{
	"Index", "IndexLocal", "RefersTo", 
}

func (this *Names) Item(optArgs ...interface{}) *Name {
	optArgs = ole.ProcessOptArgs(Names_Item_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000aa, nil, optArgs...)
	return NewName(retVal.IDispatch(), false, true)
}

var Names_Default__OptArgs= []string{
	"Index", "IndexLocal", "RefersTo", 
}

func (this *Names) Default_(optArgs ...interface{}) *Name {
	optArgs = ole.ProcessOptArgs(Names_Default__OptArgs, optArgs)
	retVal, _ := this.Call(0x00000000, nil, optArgs...)
	return NewName(retVal.IDispatch(), false, true)
}

func (this *Names) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *Names) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Names) ForEach(action func(item *Name) bool) {
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
		pItem := (*Name)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

