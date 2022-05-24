package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024454-0000-0000-C000-000000000046
var IID_CalculatedMembers = syscall.GUID{0x00024454, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type CalculatedMembers struct {
	ole.OleClient
}

func NewCalculatedMembers(pDisp *win32.IDispatch, addRef bool, scoped bool) *CalculatedMembers {
	 if pDisp == nil {
		return nil;
	}
	p := &CalculatedMembers{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CalculatedMembersFromVar(v ole.Variant) *CalculatedMembers {
	return NewCalculatedMembers(v.IDispatch(), false, false)
}

func (this *CalculatedMembers) IID() *syscall.GUID {
	return &IID_CalculatedMembers
}

func (this *CalculatedMembers) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CalculatedMembers) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *CalculatedMembers) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *CalculatedMembers) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *CalculatedMembers) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *CalculatedMembers) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *CalculatedMembers) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *CalculatedMembers) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *CalculatedMembers) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *CalculatedMembers) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *CalculatedMembers) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *CalculatedMembers) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *CalculatedMembers) Item(index interface{}) *CalculatedMember {
	retVal, _ := this.PropGet(0x000000aa, []interface{}{index})
	return NewCalculatedMember(retVal.IDispatch(), false, true)
}

func (this *CalculatedMembers) Default_(index interface{}) *CalculatedMember {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return NewCalculatedMember(retVal.IDispatch(), false, true)
}

func (this *CalculatedMembers) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *CalculatedMembers) ForEach(action func(item *CalculatedMember) bool) {
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
		pItem := (*CalculatedMember)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

var CalculatedMembers_Add__OptArgs= []string{
	"SolveOrder", "Type", 
}

func (this *CalculatedMembers) Add_(name string, formula string, optArgs ...interface{}) *CalculatedMember {
	optArgs = ole.ProcessOptArgs(CalculatedMembers_Add__OptArgs, optArgs)
	retVal, _ := this.Call(0x00000825, []interface{}{name, formula}, optArgs...)
	return NewCalculatedMember(retVal.IDispatch(), false, true)
}

var CalculatedMembers_Add_OptArgs= []string{
	"SolveOrder", "Type", "Dynamic", "DisplayFolder", "HierarchizeDistinct", 
}

func (this *CalculatedMembers) Add(name string, formula interface{}, optArgs ...interface{}) *CalculatedMember {
	optArgs = ole.ProcessOptArgs(CalculatedMembers_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b5, []interface{}{name, formula}, optArgs...)
	return NewCalculatedMember(retVal.IDispatch(), false, true)
}

