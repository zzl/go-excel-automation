package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 000244C4-0000-0000-C000-000000000046
var IID_SlicerCache = syscall.GUID{0x000244C4, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type SlicerCache struct {
	ole.OleClient
}

func NewSlicerCache(pDisp *win32.IDispatch, addRef bool, scoped bool) *SlicerCache {
	if pDisp == nil {
		return nil
	}
	p := &SlicerCache{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SlicerCacheFromVar(v ole.Variant) *SlicerCache {
	return NewSlicerCache(v.IDispatch(), false, false)
}

func (this *SlicerCache) IID() *syscall.GUID {
	return &IID_SlicerCache
}

func (this *SlicerCache) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SlicerCache) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *SlicerCache) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *SlicerCache) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *SlicerCache) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *SlicerCache) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *SlicerCache) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *SlicerCache) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *SlicerCache) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *SlicerCache) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *SlicerCache) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *SlicerCache) Index() int32 {
	retVal, _ := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *SlicerCache) OLAP() bool {
	retVal, _ := this.PropGet(0x0000081d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SlicerCache) SourceType() int32 {
	retVal, _ := this.PropGet(0x000002ad, nil)
	return retVal.LValVal()
}

func (this *SlicerCache) WorkbookConnection() *WorkbookConnection {
	retVal, _ := this.PropGet(0x000009f0, nil)
	return NewWorkbookConnection(retVal.IDispatch(), false, true)
}

func (this *SlicerCache) Slicers() *Slicers {
	retVal, _ := this.PropGet(0x00000b41, nil)
	return NewSlicers(retVal.IDispatch(), false, true)
}

func (this *SlicerCache) PivotTables() *SlicerPivotTables {
	retVal, _ := this.PropGet(0x000002b2, nil)
	return NewSlicerPivotTables(retVal.IDispatch(), false, true)
}

func (this *SlicerCache) SlicerCacheLevels() *SlicerCacheLevels {
	retVal, _ := this.PropGet(0x00000b9e, nil)
	return NewSlicerCacheLevels(retVal.IDispatch(), false, true)
}

func (this *SlicerCache) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SlicerCache) SetName(rhs string) {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *SlicerCache) VisibleSlicerItems() *SlicerItems {
	retVal, _ := this.PropGet(0x00000b9f, nil)
	return NewSlicerItems(retVal.IDispatch(), false, true)
}

func (this *SlicerCache) VisibleSlicerItemsList() ole.Variant {
	retVal, _ := this.PropGet(0x00000ba0, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *SlicerCache) SetVisibleSlicerItemsList(rhs interface{}) {
	_ = this.PropPut(0x00000ba0, []interface{}{rhs})
}

func (this *SlicerCache) SlicerItems() *SlicerItems {
	retVal, _ := this.PropGet(0x00000ba1, nil)
	return NewSlicerItems(retVal.IDispatch(), false, true)
}

func (this *SlicerCache) CrossFilterType() int32 {
	retVal, _ := this.PropGet(0x00000ba2, nil)
	return retVal.LValVal()
}

func (this *SlicerCache) SetCrossFilterType(rhs int32) {
	_ = this.PropPut(0x00000ba2, []interface{}{rhs})
}

func (this *SlicerCache) SortItems() int32 {
	retVal, _ := this.PropGet(0x00000ba3, nil)
	return retVal.LValVal()
}

func (this *SlicerCache) SetSortItems(rhs int32) {
	_ = this.PropPut(0x00000ba3, []interface{}{rhs})
}

func (this *SlicerCache) SourceName() string {
	retVal, _ := this.PropGet(0x000002d1, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SlicerCache) SortUsingCustomLists() bool {
	retVal, _ := this.PropGet(0x00000a0e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SlicerCache) SetSortUsingCustomLists(rhs bool) {
	_ = this.PropPut(0x00000a0e, []interface{}{rhs})
}

func (this *SlicerCache) ShowAllItems() bool {
	retVal, _ := this.PropGet(0x000001c4, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SlicerCache) SetShowAllItems(rhs bool) {
	_ = this.PropPut(0x000001c4, []interface{}{rhs})
}

func (this *SlicerCache) ClearManualFilter() {
	retVal, _ := this.Call(0x00000a22, nil)
	_ = retVal
}

func (this *SlicerCache) Delete() {
	retVal, _ := this.Call(0x00000075, nil)
	_ = retVal
}
