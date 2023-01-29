package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00024467-0000-0000-C000-000000000046
var IID_Protection = syscall.GUID{0x00024467, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Protection struct {
	ole.OleClient
}

func NewProtection(pDisp *win32.IDispatch, addRef bool, scoped bool) *Protection {
	if pDisp == nil {
		return nil
	}
	p := &Protection{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ProtectionFromVar(v ole.Variant) *Protection {
	return NewProtection(v.IDispatch(), false, false)
}

func (this *Protection) IID() *syscall.GUID {
	return &IID_Protection
}

func (this *Protection) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Protection) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Protection) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Protection) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Protection) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Protection) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Protection) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Protection) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Protection) AllowFormattingCells() bool {
	retVal, _ := this.PropGet(0x000007f0, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Protection) AllowFormattingColumns() bool {
	retVal, _ := this.PropGet(0x000007f1, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Protection) AllowFormattingRows() bool {
	retVal, _ := this.PropGet(0x000007f2, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Protection) AllowInsertingColumns() bool {
	retVal, _ := this.PropGet(0x000007f3, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Protection) AllowInsertingRows() bool {
	retVal, _ := this.PropGet(0x000007f4, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Protection) AllowInsertingHyperlinks() bool {
	retVal, _ := this.PropGet(0x000007f5, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Protection) AllowDeletingColumns() bool {
	retVal, _ := this.PropGet(0x000007f6, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Protection) AllowDeletingRows() bool {
	retVal, _ := this.PropGet(0x000007f7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Protection) AllowSorting() bool {
	retVal, _ := this.PropGet(0x000007f8, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Protection) AllowFiltering() bool {
	retVal, _ := this.PropGet(0x000007f9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Protection) AllowUsingPivotTables() bool {
	retVal, _ := this.PropGet(0x000007fa, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Protection) AllowEditRanges() *AllowEditRanges {
	retVal, _ := this.PropGet(0x000008bc, nil)
	return NewAllowEditRanges(retVal.IDispatch(), false, true)
}
