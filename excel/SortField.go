package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 000244A9-0000-0000-C000-000000000046
var IID_SortField = syscall.GUID{0x000244A9, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type SortField struct {
	ole.OleClient
}

func NewSortField(pDisp *win32.IDispatch, addRef bool, scoped bool) *SortField {
	if pDisp == nil {
		return nil
	}
	p := &SortField{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SortFieldFromVar(v ole.Variant) *SortField {
	return NewSortField(v.IDispatch(), false, false)
}

func (this *SortField) IID() *syscall.GUID {
	return &IID_SortField
}

func (this *SortField) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SortField) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *SortField) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *SortField) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *SortField) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *SortField) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *SortField) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *SortField) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *SortField) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *SortField) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *SortField) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *SortField) SortOn() int32 {
	retVal, _ := this.PropGet(0x00000ab5, nil)
	return retVal.LValVal()
}

func (this *SortField) SetSortOn(rhs int32) {
	_ = this.PropPut(0x00000ab5, []interface{}{rhs})
}

func (this *SortField) SortOnValue() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000ab6, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *SortField) Key() *Range {
	retVal, _ := this.PropGet(0x0000009b, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *SortField) Order() int32 {
	retVal, _ := this.PropGet(0x000000c0, nil)
	return retVal.LValVal()
}

func (this *SortField) SetOrder(rhs int32) {
	_ = this.PropPut(0x000000c0, []interface{}{rhs})
}

func (this *SortField) CustomOrder() ole.Variant {
	retVal, _ := this.PropGet(0x00000ab7, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *SortField) SetCustomOrder(rhs interface{}) {
	_ = this.PropPut(0x00000ab7, []interface{}{rhs})
}

func (this *SortField) DataOption() int32 {
	retVal, _ := this.PropGet(0x00000ab8, nil)
	return retVal.LValVal()
}

func (this *SortField) SetDataOption(rhs int32) {
	_ = this.PropPut(0x00000ab8, []interface{}{rhs})
}

func (this *SortField) Priority() int32 {
	retVal, _ := this.PropGet(0x000003d9, nil)
	return retVal.LValVal()
}

func (this *SortField) SetPriority(rhs int32) {
	_ = this.PropPut(0x000003d9, []interface{}{rhs})
}

func (this *SortField) Delete() {
	retVal, _ := this.Call(0x00000075, nil)
	_ = retVal
}

func (this *SortField) ModifyKey(key *Range) {
	retVal, _ := this.Call(0x00000ab9, []interface{}{key})
	_ = retVal
}

func (this *SortField) SetIcon(icon *Icon) {
	retVal, _ := this.Call(0x00000aba, []interface{}{icon})
	_ = retVal
}
