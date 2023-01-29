package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00024499-0000-0000-C000-000000000046
var IID_IconCriterion = syscall.GUID{0x00024499, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IconCriterion struct {
	ole.OleClient
}

func NewIconCriterion(pDisp *win32.IDispatch, addRef bool, scoped bool) *IconCriterion {
	if pDisp == nil {
		return nil
	}
	p := &IconCriterion{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func IconCriterionFromVar(v ole.Variant) *IconCriterion {
	return NewIconCriterion(v.IDispatch(), false, false)
}

func (this *IconCriterion) IID() *syscall.GUID {
	return &IID_IconCriterion
}

func (this *IconCriterion) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *IconCriterion) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *IconCriterion) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *IconCriterion) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *IconCriterion) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *IconCriterion) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *IconCriterion) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *IconCriterion) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *IconCriterion) Index() int32 {
	retVal, _ := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *IconCriterion) Type() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *IconCriterion) SetType(rhs int32) {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

func (this *IconCriterion) Value() ole.Variant {
	retVal, _ := this.PropGet(0x00000006, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *IconCriterion) SetValue(rhs interface{}) {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *IconCriterion) Operator() int32 {
	retVal, _ := this.PropGet(0x0000031d, nil)
	return retVal.LValVal()
}

func (this *IconCriterion) SetOperator(rhs int32) {
	_ = this.PropPut(0x0000031d, []interface{}{rhs})
}

func (this *IconCriterion) Icon() int32 {
	retVal, _ := this.PropGet(0x00000abb, nil)
	return retVal.LValVal()
}

func (this *IconCriterion) SetIcon(rhs int32) {
	_ = this.PropPut(0x00000abb, []interface{}{rhs})
}
