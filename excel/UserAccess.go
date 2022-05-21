package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002446D-0000-0000-C000-000000000046
var IID_UserAccess = syscall.GUID{0x0002446D, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type UserAccess struct {
	ole.OleClient
}

func NewUserAccess(pDisp *win32.IDispatch, addRef bool, scoped bool) *UserAccess {
	p := &UserAccess{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func UserAccessFromVar(v ole.Variant) *UserAccess {
	return NewUserAccess(v.PdispValVal(), false, false)
}

func (this *UserAccess) IID() *syscall.GUID {
	return &IID_UserAccess
}

func (this *UserAccess) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *UserAccess) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *UserAccess) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *UserAccess) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *UserAccess) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *UserAccess) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *UserAccess) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *UserAccess) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *UserAccess) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *UserAccess) AllowEdit() bool {
	retVal := this.PropGet(0x000007e4, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *UserAccess) SetAllowEdit(rhs bool)  {
	retVal := this.PropPut(0x000007e4, []interface{}{rhs})
	_= retVal
}

func (this *UserAccess) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

