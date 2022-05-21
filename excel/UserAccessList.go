package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002446C-0000-0000-C000-000000000046
var IID_UserAccessList = syscall.GUID{0x0002446C, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type UserAccessList struct {
	ole.OleClient
}

func NewUserAccessList(pDisp *win32.IDispatch, addRef bool, scoped bool) *UserAccessList {
	p := &UserAccessList{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func UserAccessListFromVar(v ole.Variant) *UserAccessList {
	return NewUserAccessList(v.PdispValVal(), false, false)
}

func (this *UserAccessList) IID() *syscall.GUID {
	return &IID_UserAccessList
}

func (this *UserAccessList) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *UserAccessList) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *UserAccessList) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *UserAccessList) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *UserAccessList) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *UserAccessList) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *UserAccessList) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *UserAccessList) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *UserAccessList) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *UserAccessList) Item(index interface{}) *UserAccess {
	retVal := this.PropGet(0x000000aa, []interface{}{index})
	return NewUserAccess(retVal.PdispValVal(), false, true)
}

func (this *UserAccessList) Add(name string, allowEdit bool) *UserAccess {
	retVal := this.Call(0x000000b5, []interface{}{name, allowEdit})
	return NewUserAccess(retVal.PdispValVal(), false, true)
}

func (this *UserAccessList) DeleteAll()  {
	retVal := this.Call(0x000008bf, nil)
	_= retVal
}

func (this *UserAccessList) Default_(index interface{}) *UserAccess {
	retVal := this.PropGet(0x00000000, []interface{}{index})
	return NewUserAccess(retVal.PdispValVal(), false, true)
}

func (this *UserAccessList) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *UserAccessList) ForEach(action func(item *UserAccess) bool) {
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
		pItem := (*UserAccess)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

