package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00024455-0000-0000-C000-000000000046
var IID_CalculatedMember = syscall.GUID{0x00024455, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type CalculatedMember struct {
	ole.OleClient
}

func NewCalculatedMember(pDisp *win32.IDispatch, addRef bool, scoped bool) *CalculatedMember {
	if pDisp == nil {
		return nil
	}
	p := &CalculatedMember{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CalculatedMemberFromVar(v ole.Variant) *CalculatedMember {
	return NewCalculatedMember(v.IDispatch(), false, false)
}

func (this *CalculatedMember) IID() *syscall.GUID {
	return &IID_CalculatedMember
}

func (this *CalculatedMember) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CalculatedMember) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *CalculatedMember) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *CalculatedMember) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *CalculatedMember) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *CalculatedMember) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *CalculatedMember) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *CalculatedMember) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *CalculatedMember) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *CalculatedMember) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *CalculatedMember) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *CalculatedMember) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *CalculatedMember) Formula() string {
	retVal, _ := this.PropGet(0x00000105, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *CalculatedMember) SourceName() string {
	retVal, _ := this.PropGet(0x000002d1, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *CalculatedMember) SolveOrder() int32 {
	retVal, _ := this.PropGet(0x0000088b, nil)
	return retVal.LValVal()
}

func (this *CalculatedMember) IsValid() bool {
	retVal, _ := this.PropGet(0x0000088c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CalculatedMember) Default_() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *CalculatedMember) Delete() {
	retVal, _ := this.Call(0x00000075, nil)
	_ = retVal
}

func (this *CalculatedMember) Type() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *CalculatedMember) Dynamic() bool {
	retVal, _ := this.PropGet(0x00000b6e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CalculatedMember) DisplayFolder() string {
	retVal, _ := this.PropGet(0x00000b6f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *CalculatedMember) HierarchizeDistinct() bool {
	retVal, _ := this.PropGet(0x00000b6d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CalculatedMember) SetHierarchizeDistinct(rhs bool) {
	_ = this.PropPut(0x00000b6d, []interface{}{rhs})
}

func (this *CalculatedMember) FlattenHierarchies() bool {
	retVal, _ := this.PropGet(0x00000b6c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CalculatedMember) SetFlattenHierarchies(rhs bool) {
	_ = this.PropPut(0x00000b6c, []interface{}{rhs})
}
