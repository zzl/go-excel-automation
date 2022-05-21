package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244AB-0000-0000-C000-000000000046
var IID_Sort = syscall.GUID{0x000244AB, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Sort struct {
	ole.OleClient
}

func NewSort(pDisp *win32.IDispatch, addRef bool, scoped bool) *Sort {
	p := &Sort{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SortFromVar(v ole.Variant) *Sort {
	return NewSort(v.PdispValVal(), false, false)
}

func (this *Sort) IID() *syscall.GUID {
	return &IID_Sort
}

func (this *Sort) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Sort) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Sort) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Sort) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Sort) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Sort) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Sort) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Sort) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Sort) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Sort) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Sort) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Sort) Rng() *Range {
	retVal := this.PropGet(0x00000abc, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Sort) Header() int32 {
	retVal := this.PropGet(0x0000037f, nil)
	return retVal.LValVal()
}

func (this *Sort) SetHeader(rhs int32)  {
	retVal := this.PropPut(0x0000037f, []interface{}{rhs})
	_= retVal
}

func (this *Sort) MatchCase() bool {
	retVal := this.PropGet(0x000001aa, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Sort) SetMatchCase(rhs bool)  {
	retVal := this.PropPut(0x000001aa, []interface{}{rhs})
	_= retVal
}

func (this *Sort) Orientation() int32 {
	retVal := this.PropGet(0x00000086, nil)
	return retVal.LValVal()
}

func (this *Sort) SetOrientation(rhs int32)  {
	retVal := this.PropPut(0x00000086, []interface{}{rhs})
	_= retVal
}

func (this *Sort) SortMethod() int32 {
	retVal := this.PropGet(0x00000381, nil)
	return retVal.LValVal()
}

func (this *Sort) SetSortMethod(rhs int32)  {
	retVal := this.PropPut(0x00000381, []interface{}{rhs})
	_= retVal
}

func (this *Sort) SortFields() *SortFields {
	retVal := this.PropGet(0x00000abd, nil)
	return NewSortFields(retVal.PdispValVal(), false, true)
}

func (this *Sort) SetRange(rng *Range)  {
	retVal := this.Call(0x00000abe, []interface{}{rng})
	_= retVal
}

func (this *Sort) Apply()  {
	retVal := this.Call(0x0000068b, nil)
	_= retVal
}

