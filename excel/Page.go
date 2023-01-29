package excel

import (
	"github.com/zzl/go-win32api/v2/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244A2-0000-0000-C000-000000000046
var IID_Page = syscall.GUID{0x000244A2, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Page struct {
	ole.OleClient
}

func NewPage(pDisp *win32.IDispatch, addRef bool, scoped bool) *Page {
	if pDisp == nil {
		return nil
	}
	p := &Page{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PageFromVar(v ole.Variant) *Page {
	return NewPage(v.IDispatch(), false, false)
}

func (this *Page) IID() *syscall.GUID {
	return &IID_Page
}

func (this *Page) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Page) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Page) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Page) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Page) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Page) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Page) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Page) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Page) LeftHeader() *HeaderFooter {
	retVal, _ := this.PropGet(0x000003fa, nil)
	return NewHeaderFooter(retVal.IDispatch(), false, true)
}

func (this *Page) CenterHeader() *HeaderFooter {
	retVal, _ := this.PropGet(0x000003f3, nil)
	return NewHeaderFooter(retVal.IDispatch(), false, true)
}

func (this *Page) RightHeader() *HeaderFooter {
	retVal, _ := this.PropGet(0x00000402, nil)
	return NewHeaderFooter(retVal.IDispatch(), false, true)
}

func (this *Page) LeftFooter() *HeaderFooter {
	retVal, _ := this.PropGet(0x000003f9, nil)
	return NewHeaderFooter(retVal.IDispatch(), false, true)
}

func (this *Page) CenterFooter() *HeaderFooter {
	retVal, _ := this.PropGet(0x000003f2, nil)
	return NewHeaderFooter(retVal.IDispatch(), false, true)
}

func (this *Page) RightFooter() *HeaderFooter {
	retVal, _ := this.PropGet(0x00000401, nil)
	return NewHeaderFooter(retVal.IDispatch(), false, true)
}

