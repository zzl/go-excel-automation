package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"time"
	"unsafe"
)

// 000208D1-0000-0000-C000-000000000046
var IID_Mailer = syscall.GUID{0x000208D1, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Mailer struct {
	ole.OleClient
}

func NewMailer(pDisp *win32.IDispatch, addRef bool, scoped bool) *Mailer {
	if pDisp == nil {
		return nil
	}
	p := &Mailer{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func MailerFromVar(v ole.Variant) *Mailer {
	return NewMailer(v.IDispatch(), false, false)
}

func (this *Mailer) IID() *syscall.GUID {
	return &IID_Mailer
}

func (this *Mailer) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Mailer) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Mailer) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Mailer) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Mailer) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Mailer) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Mailer) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Mailer) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Mailer) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Mailer) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Mailer) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Mailer) BCCRecipients() ole.Variant {
	retVal, _ := this.PropGet(0x000003d7, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Mailer) SetBCCRecipients(rhs interface{}) {
	_ = this.PropPut(0x000003d7, []interface{}{rhs})
}

func (this *Mailer) CCRecipients() ole.Variant {
	retVal, _ := this.PropGet(0x000003d6, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Mailer) SetCCRecipients(rhs interface{}) {
	_ = this.PropPut(0x000003d6, []interface{}{rhs})
}

func (this *Mailer) Enclosures() ole.Variant {
	retVal, _ := this.PropGet(0x000003d8, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Mailer) SetEnclosures(rhs interface{}) {
	_ = this.PropPut(0x000003d8, []interface{}{rhs})
}

func (this *Mailer) Received() bool {
	retVal, _ := this.PropGet(0x000003da, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Mailer) SendDateTime() time.Time {
	retVal, _ := this.PropGet(0x000003db, nil)
	return ole.Date(retVal.DateVal()).ToGoTime()
}

func (this *Mailer) Sender() string {
	retVal, _ := this.PropGet(0x000003dc, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Mailer) Subject() string {
	retVal, _ := this.PropGet(0x000003b9, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Mailer) SetSubject(rhs string) {
	_ = this.PropPut(0x000003b9, []interface{}{rhs})
}

func (this *Mailer) ToRecipients() ole.Variant {
	retVal, _ := this.PropGet(0x000003d5, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Mailer) SetToRecipients(rhs interface{}) {
	_ = this.PropPut(0x000003d5, []interface{}{rhs})
}

func (this *Mailer) WhichAddress() ole.Variant {
	retVal, _ := this.PropGet(0x000003ce, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Mailer) SetWhichAddress(rhs interface{}) {
	_ = this.PropPut(0x000003ce, []interface{}{rhs})
}
