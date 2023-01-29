package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

var CLSID_Application = syscall.GUID{0x00024500, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Application struct {
	Application_
}

func NewApplication(pDisp *win32.IDispatch, addRef bool, scoped bool) *Application {
	if pDisp == nil {
		return nil
	}
	p := &Application{Application_{ole.OleClient{pDisp}}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func NewApplicationFromVar(v ole.Variant, addRef bool, scoped bool) *Application {
	return NewApplication(v.IDispatch(), addRef, scoped)
}

func NewApplicationInstance(scoped bool) (*Application, error) {
	var p *win32.IDispatch
	hr := win32.CoCreateInstance(&CLSID_Application, nil,
		win32.CLSCTX_INPROC_SERVER|win32.CLSCTX_LOCAL_SERVER,
		&IID_Application_, unsafe.Pointer(&p))
	if win32.FAILED(hr) {
		return nil, com.NewError(hr)
	}
	return NewApplication(p, false, scoped), nil
}

func (this *Application) RegisterEventHandlers(handlers AppEventsHandlers) uint32 {
	var cpc *win32.IConnectionPointContainer
	hr := this.QueryInterface(&win32.IID_IConnectionPointContainer, unsafe.Pointer(&cpc))
	win32.ASSERT_SUCCEEDED(hr)

	var cp *win32.IConnectionPoint
	hr = cpc.FindConnectionPoint(&IID_AppEvents, &cp)
	win32.ASSERT_SUCCEEDED(hr)

	dispImpl := &AppEventsDispImpl{Handlers: handlers}
	disp := NewAppEventsComObj(dispImpl, false)

	var cookie uint32
	hr = cp.Advise(disp.IUnknown(), &cookie)
	win32.ASSERT_SUCCEEDED(hr)

	disp.Release()
	cp.Release()
	cpc.Release()
	return cookie
}

func (this *Application) UnRegisterEventHandlers(cookie uint32) {
	var cpc *win32.IConnectionPointContainer
	hr := this.QueryInterface(&win32.IID_IConnectionPointContainer, unsafe.Pointer(&cpc))
	win32.ASSERT_SUCCEEDED(hr)

	var cp *win32.IConnectionPoint
	hr = cpc.FindConnectionPoint(&IID_AppEvents, &cp)
	win32.ASSERT_SUCCEEDED(hr)

	hr = cp.Unadvise(cookie)
	win32.ASSERT_SUCCEEDED(hr)

	cp.Release()
	cpc.Release()
}
