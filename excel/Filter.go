package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024434-0000-0000-C000-000000000046
var IID_Filter = syscall.GUID{0x00024434, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Filter struct {
	ole.OleClient
}

func NewFilter(pDisp *win32.IDispatch, addRef bool, scoped bool) *Filter {
	 if pDisp == nil {
		return nil;
	}
	p := &Filter{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FilterFromVar(v ole.Variant) *Filter {
	return NewFilter(v.IDispatch(), false, false)
}

func (this *Filter) IID() *syscall.GUID {
	return &IID_Filter
}

func (this *Filter) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Filter) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Filter) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Filter) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Filter) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Filter) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Filter) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Filter) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Filter) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Filter) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Filter) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Filter) On() bool {
	retVal, _ := this.PropGet(0x00000652, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Filter) Criteria1() ole.Variant {
	retVal, _ := this.PropGet(0x0000031c, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Filter) Operator_() int32 {
	retVal, _ := this.PropGet(0x00000a51, nil)
	return retVal.LValVal()
}

func (this *Filter) Criteria2() ole.Variant {
	retVal, _ := this.PropGet(0x0000031e, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Filter) Operator() int32 {
	retVal, _ := this.PropGet(0x0000031d, nil)
	return retVal.LValVal()
}

func (this *Filter) SetOperator(rhs int32)  {
	_ = this.PropPut(0x0000031d, []interface{}{rhs})
}

func (this *Filter) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

