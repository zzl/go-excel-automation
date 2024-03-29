package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 0002444A-0000-0000-C000-000000000046
var IID_PivotLayout = syscall.GUID{0x0002444A, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type PivotLayout struct {
	ole.OleClient
}

func NewPivotLayout(pDisp *win32.IDispatch, addRef bool, scoped bool) *PivotLayout {
	if pDisp == nil {
		return nil
	}
	p := &PivotLayout{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PivotLayoutFromVar(v ole.Variant) *PivotLayout {
	return NewPivotLayout(v.IDispatch(), false, false)
}

func (this *PivotLayout) IID() *syscall.GUID {
	return &IID_PivotLayout
}

func (this *PivotLayout) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *PivotLayout) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *PivotLayout) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *PivotLayout) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *PivotLayout) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *PivotLayout) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *PivotLayout) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *PivotLayout) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *PivotLayout) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *PivotLayout) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *PivotLayout) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var PivotLayout_ColumnFields_OptArgs = []string{
	"Index",
}

func (this *PivotLayout) ColumnFields(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(PivotLayout_ColumnFields_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x000002c9, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var PivotLayout_DataFields_OptArgs = []string{
	"Index",
}

func (this *PivotLayout) DataFields(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(PivotLayout_DataFields_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x000002cb, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var PivotLayout_PageFields_OptArgs = []string{
	"Index",
}

func (this *PivotLayout) PageFields(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(PivotLayout_PageFields_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x000002ca, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var PivotLayout_RowFields_OptArgs = []string{
	"Index",
}

func (this *PivotLayout) RowFields(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(PivotLayout_RowFields_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x000002c8, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var PivotLayout_HiddenFields_OptArgs = []string{
	"Index",
}

func (this *PivotLayout) HiddenFields(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(PivotLayout_HiddenFields_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x000002c7, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var PivotLayout_VisibleFields_OptArgs = []string{
	"Index",
}

func (this *PivotLayout) VisibleFields(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(PivotLayout_VisibleFields_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x000002c6, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var PivotLayout_PivotFields_OptArgs = []string{
	"Index",
}

func (this *PivotLayout) PivotFields(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(PivotLayout_PivotFields_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x000002ce, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *PivotLayout) CubeFields() *CubeFields {
	retVal, _ := this.PropGet(0x0000072f, nil)
	return NewCubeFields(retVal.IDispatch(), false, true)
}

func (this *PivotLayout) PivotCache() *PivotCache {
	retVal, _ := this.PropGet(0x000005d8, nil)
	return NewPivotCache(retVal.IDispatch(), false, true)
}

func (this *PivotLayout) PivotTable() *PivotTable {
	retVal, _ := this.PropGet(0x000002cc, nil)
	return NewPivotTable(retVal.IDispatch(), false, true)
}

func (this *PivotLayout) InnerDetail() string {
	retVal, _ := this.PropGet(0x000002ba, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotLayout) SetInnerDetail(rhs string) {
	_ = this.PropPut(0x000002ba, []interface{}{rhs})
}

var PivotLayout_AddFields_OptArgs = []string{
	"RowFields", "ColumnFields", "PageFields", "AppendField",
}

func (this *PivotLayout) AddFields(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(PivotLayout_AddFields_OptArgs, optArgs)
	retVal, _ := this.Call(0x000002c4, nil, optArgs...)
	_ = retVal
}
