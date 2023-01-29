package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

var CLSID_Worksheet = syscall.GUID{0x00020820, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Worksheet struct {
	Worksheet_
}

func NewWorksheet(pDisp *win32.IDispatch, addRef bool, scoped bool) *Worksheet {
	if pDisp == nil {
		return nil
	}
	p := &Worksheet{Worksheet_{ole.OleClient{pDisp}}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func NewWorksheetFromVar(v ole.Variant, addRef bool, scoped bool) *Worksheet {
	return NewWorksheet(v.IDispatch(), addRef, scoped)
}

func NewWorksheetInstance(scoped bool) (*Worksheet, error) {
	var p *win32.IDispatch
	hr := win32.CoCreateInstance(&CLSID_Worksheet, nil,
		win32.CLSCTX_INPROC_SERVER|win32.CLSCTX_LOCAL_SERVER,
		&IID_Worksheet_, unsafe.Pointer(&p))
	if win32.FAILED(hr) {
		return nil, com.NewError(hr)
	}
	return NewWorksheet(p, false, scoped), nil
}

func (this *Worksheet) RegisterEventHandlers(handlers DocEventsHandlers) uint32 {
	var cpc *win32.IConnectionPointContainer
	hr := this.QueryInterface(&win32.IID_IConnectionPointContainer, unsafe.Pointer(&cpc))
	win32.ASSERT_SUCCEEDED(hr)

	var cp *win32.IConnectionPoint
	hr = cpc.FindConnectionPoint(&IID_DocEvents, &cp)
	win32.ASSERT_SUCCEEDED(hr)

	dispImpl := &DocEventsDispImpl{Handlers: handlers}
	disp := NewDocEventsComObj(dispImpl, false)

	var cookie uint32
	hr = cp.Advise(disp.IUnknown(), &cookie)
	win32.ASSERT_SUCCEEDED(hr)

	disp.Release()
	cp.Release()
	cpc.Release()
	return cookie
}

func (this *Worksheet) UnRegisterEventHandlers(cookie uint32) {
	var cpc *win32.IConnectionPointContainer
	hr := this.QueryInterface(&win32.IID_IConnectionPointContainer, unsafe.Pointer(&cpc))
	win32.ASSERT_SUCCEEDED(hr)

	var cp *win32.IConnectionPoint
	hr = cpc.FindConnectionPoint(&IID_DocEvents, &cp)
	win32.ASSERT_SUCCEEDED(hr)

	hr = cp.Unadvise(cookie)
	win32.ASSERT_SUCCEEDED(hr)

	cp.Release()
	cpc.Release()
}
