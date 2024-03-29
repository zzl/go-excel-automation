package excel

import (
	"github.com/zzl/go-win32api/v2/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024431-0000-0000-C000-000000000046
var IID_Hyperlink = syscall.GUID{0x00024431, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Hyperlink struct {
	ole.OleClient
}

func NewHyperlink(pDisp *win32.IDispatch, addRef bool, scoped bool) *Hyperlink {
	if pDisp == nil {
		return nil
	}
	p := &Hyperlink{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func HyperlinkFromVar(v ole.Variant) *Hyperlink {
	return NewHyperlink(v.IDispatch(), false, false)
}

func (this *Hyperlink) IID() *syscall.GUID {
	return &IID_Hyperlink
}

func (this *Hyperlink) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Hyperlink) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Hyperlink) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Hyperlink) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Hyperlink) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Hyperlink) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Hyperlink) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Hyperlink) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Hyperlink) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Hyperlink) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Hyperlink) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Hyperlink) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Hyperlink) Range() *Range {
	retVal, _ := this.PropGet(0x000000c5, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Hyperlink) Shape() *Shape {
	retVal, _ := this.PropGet(0x0000062e, nil)
	return NewShape(retVal.IDispatch(), false, true)
}

func (this *Hyperlink) SubAddress() string {
	retVal, _ := this.PropGet(0x000005bf, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Hyperlink) SetSubAddress(rhs string) {
	_ = this.PropPut(0x000005bf, []interface{}{rhs})
}

func (this *Hyperlink) Address() string {
	retVal, _ := this.PropGet(0x000000ec, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Hyperlink) SetAddress(rhs string) {
	_ = this.PropPut(0x000000ec, []interface{}{rhs})
}

func (this *Hyperlink) Type() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *Hyperlink) AddToFavorites() {
	retVal, _ := this.Call(0x000005c4, nil)
	_ = retVal
}

func (this *Hyperlink) Delete() {
	retVal, _ := this.Call(0x00000075, nil)
	_ = retVal
}

var Hyperlink_Follow_OptArgs = []string{
	"NewWindow", "AddHistory", "ExtraInfo", "Method", "HeaderInfo",
}

func (this *Hyperlink) Follow(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Hyperlink_Follow_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000650, nil, optArgs...)
	_ = retVal
}

func (this *Hyperlink) EmailSubject() string {
	retVal, _ := this.PropGet(0x0000075b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Hyperlink) SetEmailSubject(rhs string) {
	_ = this.PropPut(0x0000075b, []interface{}{rhs})
}

func (this *Hyperlink) ScreenTip() string {
	retVal, _ := this.PropGet(0x00000759, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Hyperlink) SetScreenTip(rhs string) {
	_ = this.PropPut(0x00000759, []interface{}{rhs})
}

func (this *Hyperlink) TextToDisplay() string {
	retVal, _ := this.PropGet(0x0000075a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Hyperlink) SetTextToDisplay(rhs string) {
	_ = this.PropPut(0x0000075a, []interface{}{rhs})
}

func (this *Hyperlink) CreateNewDocument(filename string, editNow bool, overwrite bool) {
	retVal, _ := this.Call(0x0000075c, []interface{}{filename, editNow, overwrite})
	_ = retVal
}

