package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00024460-0000-0000-C000-000000000046
var IID_SmartTag = syscall.GUID{0x00024460, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type SmartTag struct {
	ole.OleClient
}

func NewSmartTag(pDisp *win32.IDispatch, addRef bool, scoped bool) *SmartTag {
	if pDisp == nil {
		return nil
	}
	p := &SmartTag{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SmartTagFromVar(v ole.Variant) *SmartTag {
	return NewSmartTag(v.IDispatch(), false, false)
}

func (this *SmartTag) IID() *syscall.GUID {
	return &IID_SmartTag
}

func (this *SmartTag) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SmartTag) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *SmartTag) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *SmartTag) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *SmartTag) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *SmartTag) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *SmartTag) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *SmartTag) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *SmartTag) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *SmartTag) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *SmartTag) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *SmartTag) DownloadURL() string {
	retVal, _ := this.PropGet(0x000008a4, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SmartTag) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SmartTag) Default_() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SmartTag) XML() string {
	retVal, _ := this.PropGet(0x000008a5, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SmartTag) Range() *Range {
	retVal, _ := this.PropGet(0x000000c5, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *SmartTag) Delete() {
	retVal, _ := this.Call(0x00000075, nil)
	_ = retVal
}

func (this *SmartTag) SmartTagActions() *SmartTagActions {
	retVal, _ := this.PropGet(0x000008a6, nil)
	return NewSmartTagActions(retVal.IDispatch(), false, true)
}

func (this *SmartTag) Properties() *CustomProperties {
	retVal, _ := this.PropGet(0x00000857, nil)
	return NewCustomProperties(retVal.IDispatch(), false, true)
}
