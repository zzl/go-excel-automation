package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00024495-0000-0000-C000-000000000046
var IID_ColorScaleCriterion = syscall.GUID{0x00024495, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ColorScaleCriterion struct {
	ole.OleClient
}

func NewColorScaleCriterion(pDisp *win32.IDispatch, addRef bool, scoped bool) *ColorScaleCriterion {
	if pDisp == nil {
		return nil
	}
	p := &ColorScaleCriterion{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ColorScaleCriterionFromVar(v ole.Variant) *ColorScaleCriterion {
	return NewColorScaleCriterion(v.IDispatch(), false, false)
}

func (this *ColorScaleCriterion) IID() *syscall.GUID {
	return &IID_ColorScaleCriterion
}

func (this *ColorScaleCriterion) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ColorScaleCriterion) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *ColorScaleCriterion) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ColorScaleCriterion) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ColorScaleCriterion) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *ColorScaleCriterion) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *ColorScaleCriterion) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *ColorScaleCriterion) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *ColorScaleCriterion) Index() int32 {
	retVal, _ := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *ColorScaleCriterion) Type() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *ColorScaleCriterion) SetType(rhs int32) {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

func (this *ColorScaleCriterion) Value() ole.Variant {
	retVal, _ := this.PropGet(0x00000006, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ColorScaleCriterion) SetValue(rhs interface{}) {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *ColorScaleCriterion) FormatColor() *FormatColor {
	retVal, _ := this.PropGet(0x00000a9d, nil)
	return NewFormatColor(retVal.IDispatch(), false, true)
}
