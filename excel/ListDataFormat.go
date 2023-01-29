package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 0002447D-0000-0000-C000-000000000046
var IID_ListDataFormat = syscall.GUID{0x0002447D, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ListDataFormat struct {
	ole.OleClient
}

func NewListDataFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *ListDataFormat {
	if pDisp == nil {
		return nil
	}
	p := &ListDataFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ListDataFormatFromVar(v ole.Variant) *ListDataFormat {
	return NewListDataFormat(v.IDispatch(), false, false)
}

func (this *ListDataFormat) IID() *syscall.GUID {
	return &IID_ListDataFormat
}

func (this *ListDataFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ListDataFormat) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *ListDataFormat) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ListDataFormat) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ListDataFormat) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *ListDataFormat) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *ListDataFormat) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *ListDataFormat) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *ListDataFormat) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ListDataFormat) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ListDataFormat) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ListDataFormat) Default_() int32 {
	retVal, _ := this.PropGet(0x00000000, nil)
	return retVal.LValVal()
}

func (this *ListDataFormat) Choices() ole.Variant {
	retVal, _ := this.PropGet(0x0000092c, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ListDataFormat) DecimalPlaces() int32 {
	retVal, _ := this.PropGet(0x0000092d, nil)
	return retVal.LValVal()
}

func (this *ListDataFormat) DefaultValue() ole.Variant {
	retVal, _ := this.PropGet(0x0000092e, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ListDataFormat) IsPercent() bool {
	retVal, _ := this.PropGet(0x0000092f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListDataFormat) Lcid() int32 {
	retVal, _ := this.PropGet(0x00000930, nil)
	return retVal.LValVal()
}

func (this *ListDataFormat) MaxCharacters() int32 {
	retVal, _ := this.PropGet(0x00000931, nil)
	return retVal.LValVal()
}

func (this *ListDataFormat) MaxNumber() ole.Variant {
	retVal, _ := this.PropGet(0x00000932, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ListDataFormat) MinNumber() ole.Variant {
	retVal, _ := this.PropGet(0x00000933, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ListDataFormat) Required() bool {
	retVal, _ := this.PropGet(0x00000934, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListDataFormat) Type() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *ListDataFormat) ReadOnly() bool {
	retVal, _ := this.PropGet(0x00000128, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListDataFormat) AllowFillIn() bool {
	retVal, _ := this.PropGet(0x00000935, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}
