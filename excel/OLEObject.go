package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

var CLSID_OLEObject = syscall.GUID{0x00020818, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type OLEObject struct {
	OLEObject_
}

func NewOLEObject(pDisp *win32.IDispatch, addRef bool, scoped bool) *OLEObject {
	 if pDisp == nil {
		return nil;
	}
	p := &OLEObject{OLEObject_{ole.OleClient{pDisp}}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func NewOLEObjectFromVar(v ole.Variant, addRef bool, scoped bool) *OLEObject {
	return NewOLEObject(v.IDispatch(), addRef, scoped)
}

func NewOLEObjectInstance(scoped bool) (*OLEObject, error) {
	var p *win32.IDispatch
	hr := win32.CoCreateInstance(&CLSID_OLEObject, nil, 
		win32.CLSCTX_INPROC_SERVER|win32.CLSCTX_LOCAL_SERVER,
		&IID_OLEObject_, unsafe.Pointer(&p))
	if win32.FAILED(hr) {
		return nil, com.NewError(hr)
	}
	return NewOLEObject(p, false, scoped), nil
}

func (this *OLEObject) RegisterEventHandlers(handlers OLEObjectEventsHandlers) uint32 {
	var cpc *win32.IConnectionPointContainer
	hr := this.QueryInterface(&win32.IID_IConnectionPointContainer, unsafe.Pointer(&cpc))
	win32.ASSERT_SUCCEEDED(hr)

	var cp *win32.IConnectionPoint
	hr = cpc.FindConnectionPoint(&IID_OLEObjectEvents, &cp)
	win32.ASSERT_SUCCEEDED(hr)

	dispImpl := &OLEObjectEventsDispImpl{Handlers: handlers}
	disp := NewOLEObjectEventsComObj(dispImpl, false)
	
	var cookie uint32
	hr = cp.Advise(disp.IUnknown(), &cookie)
	win32.ASSERT_SUCCEEDED(hr)

	disp.Release()
	cp.Release()
	cpc.Release()
	return cookie
}

func (this *OLEObject) UnRegisterEventHandlers(cookie uint32) {
	var cpc *win32.IConnectionPointContainer
	hr := this.QueryInterface(&win32.IID_IConnectionPointContainer, unsafe.Pointer(&cpc))
	win32.ASSERT_SUCCEEDED(hr)

	var cp *win32.IConnectionPoint
	hr = cpc.FindConnectionPoint(&IID_OLEObjectEvents, &cp)
	win32.ASSERT_SUCCEEDED(hr)

	hr = cp.Unadvise(cookie)
	win32.ASSERT_SUCCEEDED(hr)

	cp.Release()
	cpc.Release()
}

