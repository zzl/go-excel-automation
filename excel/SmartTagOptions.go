package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00024464-0000-0000-C000-000000000046
var IID_SmartTagOptions = syscall.GUID{0x00024464, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type SmartTagOptions struct {
	ole.OleClient
}

func NewSmartTagOptions(pDisp *win32.IDispatch, addRef bool, scoped bool) *SmartTagOptions {
	if pDisp == nil {
		return nil
	}
	p := &SmartTagOptions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SmartTagOptionsFromVar(v ole.Variant) *SmartTagOptions {
	return NewSmartTagOptions(v.IDispatch(), false, false)
}

func (this *SmartTagOptions) IID() *syscall.GUID {
	return &IID_SmartTagOptions
}

func (this *SmartTagOptions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SmartTagOptions) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *SmartTagOptions) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *SmartTagOptions) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *SmartTagOptions) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *SmartTagOptions) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *SmartTagOptions) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *SmartTagOptions) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *SmartTagOptions) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *SmartTagOptions) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *SmartTagOptions) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *SmartTagOptions) DisplaySmartTags() int32 {
	retVal, _ := this.PropGet(0x000008aa, nil)
	return retVal.LValVal()
}

func (this *SmartTagOptions) SetDisplaySmartTags(rhs int32) {
	_ = this.PropPut(0x000008aa, []interface{}{rhs})
}

func (this *SmartTagOptions) EmbedSmartTags() bool {
	retVal, _ := this.PropGet(0x000008ab, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SmartTagOptions) SetEmbedSmartTags(rhs bool) {
	_ = this.PropPut(0x000008ab, []interface{}{rhs})
}
