package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002446B-0000-0000-C000-000000000046
var IID_AllowEditRange = syscall.GUID{0x0002446B, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type AllowEditRange struct {
	ole.OleClient
}

func NewAllowEditRange(pDisp *win32.IDispatch, addRef bool, scoped bool) *AllowEditRange {
	 if pDisp == nil {
		return nil;
	}
	p := &AllowEditRange{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func AllowEditRangeFromVar(v ole.Variant) *AllowEditRange {
	return NewAllowEditRange(v.IDispatch(), false, false)
}

func (this *AllowEditRange) IID() *syscall.GUID {
	return &IID_AllowEditRange
}

func (this *AllowEditRange) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *AllowEditRange) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *AllowEditRange) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *AllowEditRange) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *AllowEditRange) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *AllowEditRange) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *AllowEditRange) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *AllowEditRange) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *AllowEditRange) Title() string {
	retVal, _ := this.PropGet(0x000000c7, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AllowEditRange) SetTitle(rhs string)  {
	_ = this.PropPut(0x000000c7, []interface{}{rhs})
}

func (this *AllowEditRange) Range() *Range {
	retVal, _ := this.PropGet(0x000000c5, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *AllowEditRange) SetRange(rhs *Range)  {
	_ = this.PropPutRef(0x000000c5, []interface{}{rhs})
}

func (this *AllowEditRange) ChangePassword(password string)  {
	retVal, _ := this.Call(0x000008bd, []interface{}{password})
	_= retVal
}

func (this *AllowEditRange) Delete()  {
	retVal, _ := this.Call(0x00000075, nil)
	_= retVal
}

var AllowEditRange_Unprotect_OptArgs= []string{
	"Password", 
}

func (this *AllowEditRange) Unprotect(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(AllowEditRange_Unprotect_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000011d, nil, optArgs...)
	_= retVal
}

func (this *AllowEditRange) Users() *UserAccessList {
	retVal, _ := this.PropGet(0x000008be, nil)
	return NewUserAccessList(retVal.IDispatch(), false, true)
}

