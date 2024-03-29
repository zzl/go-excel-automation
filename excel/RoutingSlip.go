package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 000208AA-0000-0000-C000-000000000046
var IID_RoutingSlip = syscall.GUID{0x000208AA, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type RoutingSlip struct {
	ole.OleClient
}

func NewRoutingSlip(pDisp *win32.IDispatch, addRef bool, scoped bool) *RoutingSlip {
	if pDisp == nil {
		return nil
	}
	p := &RoutingSlip{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func RoutingSlipFromVar(v ole.Variant) *RoutingSlip {
	return NewRoutingSlip(v.IDispatch(), false, false)
}

func (this *RoutingSlip) IID() *syscall.GUID {
	return &IID_RoutingSlip
}

func (this *RoutingSlip) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *RoutingSlip) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *RoutingSlip) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *RoutingSlip) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *RoutingSlip) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *RoutingSlip) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *RoutingSlip) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *RoutingSlip) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *RoutingSlip) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *RoutingSlip) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *RoutingSlip) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *RoutingSlip) Delivery() int32 {
	retVal, _ := this.PropGet(0x000003bb, nil)
	return retVal.LValVal()
}

func (this *RoutingSlip) SetDelivery(rhs int32) {
	_ = this.PropPut(0x000003bb, []interface{}{rhs})
}

func (this *RoutingSlip) Message() ole.Variant {
	retVal, _ := this.PropGet(0x000003ba, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *RoutingSlip) SetMessage(rhs interface{}) {
	_ = this.PropPut(0x000003ba, []interface{}{rhs})
}

var RoutingSlip_Recipients_OptArgs = []string{
	"Index",
}

func (this *RoutingSlip) Recipients(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(RoutingSlip_Recipients_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x000003b8, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var RoutingSlip_SetRecipients_OptArgs = []string{
	"Index",
}

func (this *RoutingSlip) SetRecipients(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(RoutingSlip_SetRecipients_OptArgs, optArgs)
	_ = this.PropPut(0x000003b8, nil, optArgs...)
}

func (this *RoutingSlip) Reset() ole.Variant {
	retVal, _ := this.Call(0x0000022b, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *RoutingSlip) ReturnWhenDone() bool {
	retVal, _ := this.PropGet(0x000003bc, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *RoutingSlip) SetReturnWhenDone(rhs bool) {
	_ = this.PropPut(0x000003bc, []interface{}{rhs})
}

func (this *RoutingSlip) Status() int32 {
	retVal, _ := this.PropGet(0x000003be, nil)
	return retVal.LValVal()
}

func (this *RoutingSlip) Subject() ole.Variant {
	retVal, _ := this.PropGet(0x000003b9, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *RoutingSlip) SetSubject(rhs interface{}) {
	_ = this.PropPut(0x000003b9, []interface{}{rhs})
}

func (this *RoutingSlip) TrackStatus() bool {
	retVal, _ := this.PropGet(0x000003bd, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *RoutingSlip) SetTrackStatus(rhs bool) {
	_ = this.PropPut(0x000003bd, []interface{}{rhs})
}
