package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00024432-0000-0000-C000-000000000046
var IID_AutoFilter = syscall.GUID{0x00024432, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type AutoFilter struct {
	ole.OleClient
}

func NewAutoFilter(pDisp *win32.IDispatch, addRef bool, scoped bool) *AutoFilter {
	if pDisp == nil {
		return nil
	}
	p := &AutoFilter{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func AutoFilterFromVar(v ole.Variant) *AutoFilter {
	return NewAutoFilter(v.IDispatch(), false, false)
}

func (this *AutoFilter) IID() *syscall.GUID {
	return &IID_AutoFilter
}

func (this *AutoFilter) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *AutoFilter) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *AutoFilter) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *AutoFilter) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *AutoFilter) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *AutoFilter) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *AutoFilter) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *AutoFilter) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *AutoFilter) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *AutoFilter) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *AutoFilter) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *AutoFilter) Range() *Range {
	retVal, _ := this.PropGet(0x000000c5, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *AutoFilter) Filters() *Filters {
	retVal, _ := this.PropGet(0x00000651, nil)
	return NewFilters(retVal.IDispatch(), false, true)
}

func (this *AutoFilter) FilterMode() bool {
	retVal, _ := this.PropGet(0x00000320, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoFilter) Sort() *Sort {
	retVal, _ := this.PropGet(0x00000370, nil)
	return NewSort(retVal.IDispatch(), false, true)
}

func (this *AutoFilter) ApplyFilter() {
	retVal, _ := this.Call(0x00000a50, nil)
	_ = retVal
}

func (this *AutoFilter) ShowAllData() {
	retVal, _ := this.Call(0x0000031a, nil)
	_ = retVal
}
