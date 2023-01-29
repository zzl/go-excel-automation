package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 000244C2-0000-0000-C000-000000000046
var IID_DisplayFormat = syscall.GUID{0x000244C2, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type DisplayFormat struct {
	ole.OleClient
}

func NewDisplayFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *DisplayFormat {
	if pDisp == nil {
		return nil
	}
	p := &DisplayFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DisplayFormatFromVar(v ole.Variant) *DisplayFormat {
	return NewDisplayFormat(v.IDispatch(), false, false)
}

func (this *DisplayFormat) IID() *syscall.GUID {
	return &IID_DisplayFormat
}

func (this *DisplayFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DisplayFormat) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *DisplayFormat) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *DisplayFormat) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *DisplayFormat) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *DisplayFormat) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *DisplayFormat) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *DisplayFormat) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *DisplayFormat) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *DisplayFormat) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *DisplayFormat) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DisplayFormat) Borders() *Borders {
	retVal, _ := this.PropGet(0x000001b3, nil)
	return NewBorders(retVal.IDispatch(), false, true)
}

var DisplayFormat_Characters_OptArgs = []string{
	"Start", "Length",
}

func (this *DisplayFormat) Characters(optArgs ...interface{}) *Characters {
	optArgs = ole.ProcessOptArgs(DisplayFormat_Characters_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000025b, nil, optArgs...)
	return NewCharacters(retVal.IDispatch(), false, true)
}

func (this *DisplayFormat) Font() *Font {
	retVal, _ := this.PropGet(0x00000092, nil)
	return NewFont(retVal.IDispatch(), false, true)
}

func (this *DisplayFormat) Style() ole.Variant {
	retVal, _ := this.PropGet(0x00000104, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayFormat) AddIndent() ole.Variant {
	retVal, _ := this.PropGet(0x00000427, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayFormat) FormulaHidden() ole.Variant {
	retVal, _ := this.PropGet(0x00000106, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayFormat) HorizontalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x00000088, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayFormat) IndentLevel() ole.Variant {
	retVal, _ := this.PropGet(0x000000c9, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayFormat) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *DisplayFormat) Locked() ole.Variant {
	retVal, _ := this.PropGet(0x0000010d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayFormat) MergeCells() ole.Variant {
	retVal, _ := this.PropGet(0x000000d0, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayFormat) NumberFormat() ole.Variant {
	retVal, _ := this.PropGet(0x000000c1, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayFormat) NumberFormatLocal() ole.Variant {
	retVal, _ := this.PropGet(0x00000449, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayFormat) Orientation() ole.Variant {
	retVal, _ := this.PropGet(0x00000086, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayFormat) ReadingOrder() int32 {
	retVal, _ := this.PropGet(0x000003cf, nil)
	return retVal.LValVal()
}

func (this *DisplayFormat) ShrinkToFit() ole.Variant {
	retVal, _ := this.PropGet(0x000000d1, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayFormat) VerticalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x00000089, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayFormat) WrapText() ole.Variant {
	retVal, _ := this.PropGet(0x00000114, nil)
	com.AddToScope(retVal)
	return *retVal
}
