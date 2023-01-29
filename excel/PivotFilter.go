package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00024483-0000-0000-C000-000000000046
var IID_PivotFilter = syscall.GUID{0x00024483, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type PivotFilter struct {
	ole.OleClient
}

func NewPivotFilter(pDisp *win32.IDispatch, addRef bool, scoped bool) *PivotFilter {
	if pDisp == nil {
		return nil
	}
	p := &PivotFilter{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PivotFilterFromVar(v ole.Variant) *PivotFilter {
	return NewPivotFilter(v.IDispatch(), false, false)
}

func (this *PivotFilter) IID() *syscall.GUID {
	return &IID_PivotFilter
}

func (this *PivotFilter) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *PivotFilter) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *PivotFilter) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *PivotFilter) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *PivotFilter) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *PivotFilter) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *PivotFilter) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *PivotFilter) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *PivotFilter) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *PivotFilter) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *PivotFilter) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *PivotFilter) Order() int32 {
	retVal, _ := this.PropGet(0x000000c0, nil)
	return retVal.LValVal()
}

func (this *PivotFilter) SetOrder(rhs int32) {
	_ = this.PropPut(0x000000c0, []interface{}{rhs})
}

func (this *PivotFilter) FilterType() int32 {
	retVal, _ := this.PropGet(0x00000a7e, nil)
	return retVal.LValVal()
}

func (this *PivotFilter) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotFilter) Description() string {
	retVal, _ := this.PropGet(0x000000da, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotFilter) Delete() {
	retVal, _ := this.Call(0x00000075, nil)
	_ = retVal
}

func (this *PivotFilter) Active() bool {
	retVal, _ := this.PropGet(0x00000908, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotFilter) PivotField() *PivotField {
	retVal, _ := this.PropGet(0x000002db, nil)
	return NewPivotField(retVal.IDispatch(), false, true)
}

func (this *PivotFilter) DataField() *PivotField {
	retVal, _ := this.PropGet(0x0000082b, nil)
	return NewPivotField(retVal.IDispatch(), false, true)
}

func (this *PivotFilter) DataCubeField() *CubeField {
	retVal, _ := this.PropGet(0x00000a7f, nil)
	return NewCubeField(retVal.IDispatch(), false, true)
}

func (this *PivotFilter) Value1() ole.Variant {
	retVal, _ := this.PropGet(0x00000a80, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PivotFilter) Value2() ole.Variant {
	retVal, _ := this.PropGet(0x0000056c, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PivotFilter) MemberPropertyField() *PivotField {
	retVal, _ := this.PropGet(0x00000a81, nil)
	return NewPivotField(retVal.IDispatch(), false, true)
}

func (this *PivotFilter) IsMemberPropertyFilter() bool {
	retVal, _ := this.PropGet(0x00000a82, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}
