package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 0002441E-0000-0000-C000-000000000046
var IID_PivotFormula = syscall.GUID{0x0002441E, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type PivotFormula struct {
	ole.OleClient
}

func NewPivotFormula(pDisp *win32.IDispatch, addRef bool, scoped bool) *PivotFormula {
	if pDisp == nil {
		return nil
	}
	p := &PivotFormula{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PivotFormulaFromVar(v ole.Variant) *PivotFormula {
	return NewPivotFormula(v.IDispatch(), false, false)
}

func (this *PivotFormula) IID() *syscall.GUID {
	return &IID_PivotFormula
}

func (this *PivotFormula) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *PivotFormula) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *PivotFormula) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *PivotFormula) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *PivotFormula) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *PivotFormula) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *PivotFormula) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *PivotFormula) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *PivotFormula) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *PivotFormula) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *PivotFormula) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *PivotFormula) Delete() {
	retVal, _ := this.Call(0x00000075, nil)
	_ = retVal
}

func (this *PivotFormula) Default_() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotFormula) SetDefault_(rhs string) {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *PivotFormula) Formula() string {
	retVal, _ := this.PropGet(0x00000105, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotFormula) SetFormula(rhs string) {
	_ = this.PropPut(0x00000105, []interface{}{rhs})
}

func (this *PivotFormula) Value() string {
	retVal, _ := this.PropGet(0x00000006, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotFormula) SetValue(rhs string) {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *PivotFormula) Index() int32 {
	retVal, _ := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *PivotFormula) SetIndex(rhs int32) {
	_ = this.PropPut(0x000001e6, []interface{}{rhs})
}

func (this *PivotFormula) StandardFormula() string {
	retVal, _ := this.PropGet(0x00000824, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotFormula) SetStandardFormula(rhs string) {
	_ = this.PropPut(0x00000824, []interface{}{rhs})
}
