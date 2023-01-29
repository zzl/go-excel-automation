package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 000244C8-0000-0000-C000-000000000046
var IID_Slicer = syscall.GUID{0x000244C8, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Slicer struct {
	ole.OleClient
}

func NewSlicer(pDisp *win32.IDispatch, addRef bool, scoped bool) *Slicer {
	if pDisp == nil {
		return nil
	}
	p := &Slicer{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SlicerFromVar(v ole.Variant) *Slicer {
	return NewSlicer(v.IDispatch(), false, false)
}

func (this *Slicer) IID() *syscall.GUID {
	return &IID_Slicer
}

func (this *Slicer) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Slicer) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Slicer) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Slicer) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Slicer) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Slicer) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Slicer) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Slicer) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Slicer) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Slicer) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Slicer) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Slicer) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Slicer) SetName(rhs string) {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *Slicer) Caption() string {
	retVal, _ := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Slicer) SetCaption(rhs string) {
	_ = this.PropPut(0x0000008b, []interface{}{rhs})
}

func (this *Slicer) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *Slicer) SetTop(rhs float64) {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *Slicer) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *Slicer) SetLeft(rhs float64) {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *Slicer) DisableMoveResizeUI() bool {
	retVal, _ := this.PropGet(0x00000ba7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Slicer) SetDisableMoveResizeUI(rhs bool) {
	_ = this.PropPut(0x00000ba7, []interface{}{rhs})
}

func (this *Slicer) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *Slicer) SetWidth(rhs float64) {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *Slicer) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *Slicer) SetHeight(rhs float64) {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *Slicer) RowHeight() float64 {
	retVal, _ := this.PropGet(0x00000110, nil)
	return retVal.DblValVal()
}

func (this *Slicer) SetRowHeight(rhs float64) {
	_ = this.PropPut(0x00000110, []interface{}{rhs})
}

func (this *Slicer) ColumnWidth() float64 {
	retVal, _ := this.PropGet(0x000000f2, nil)
	return retVal.DblValVal()
}

func (this *Slicer) SetColumnWidth(rhs float64) {
	_ = this.PropPut(0x000000f2, []interface{}{rhs})
}

func (this *Slicer) NumberOfColumns() int32 {
	retVal, _ := this.PropGet(0x00000ba8, nil)
	return retVal.LValVal()
}

func (this *Slicer) SetNumberOfColumns(rhs int32) {
	_ = this.PropPut(0x00000ba8, []interface{}{rhs})
}

func (this *Slicer) DisplayHeader() bool {
	retVal, _ := this.PropGet(0x00000ba9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Slicer) SetDisplayHeader(rhs bool) {
	_ = this.PropPut(0x00000ba9, []interface{}{rhs})
}

func (this *Slicer) Locked() bool {
	retVal, _ := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Slicer) SetLocked(rhs bool) {
	_ = this.PropPut(0x0000010d, []interface{}{rhs})
}

func (this *Slicer) SlicerCache() *SlicerCache {
	retVal, _ := this.PropGet(0x00000baa, nil)
	return NewSlicerCache(retVal.IDispatch(), false, true)
}

func (this *Slicer) SlicerCacheLevel() *SlicerCacheLevel {
	retVal, _ := this.PropGet(0x00000bab, nil)
	return NewSlicerCacheLevel(retVal.IDispatch(), false, true)
}

func (this *Slicer) Shape() *Shape {
	retVal, _ := this.PropGet(0x0000062e, nil)
	return NewShape(retVal.IDispatch(), false, true)
}

func (this *Slicer) Style() ole.Variant {
	retVal, _ := this.PropGet(0x00000104, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Slicer) SetStyle(rhs interface{}) {
	_ = this.PropPut(0x00000104, []interface{}{rhs})
}

func (this *Slicer) Delete() {
	retVal, _ := this.Call(0x00000075, nil)
	_ = retVal
}

func (this *Slicer) Cut() {
	retVal, _ := this.Call(0x00000235, nil)
	_ = retVal
}

func (this *Slicer) Copy() {
	retVal, _ := this.Call(0x00000227, nil)
	_ = retVal
}

func (this *Slicer) ActiveItem() *SlicerItem {
	retVal, _ := this.PropGet(0x00000bac, nil)
	return NewSlicerItem(retVal.IDispatch(), false, true)
}
