package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

var CLSID_Chart = syscall.GUID{0x00020821, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Chart struct {
	Chart_
}

func NewChart(pDisp *win32.IDispatch, addRef bool, scoped bool) *Chart {
	if pDisp == nil {
		return nil
	}
	p := &Chart{Chart_{ole.OleClient{pDisp}}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func NewChartFromVar(v ole.Variant, addRef bool, scoped bool) *Chart {
	return NewChart(v.IDispatch(), addRef, scoped)
}

func NewChartInstance(scoped bool) (*Chart, error) {
	var p *win32.IDispatch
	hr := win32.CoCreateInstance(&CLSID_Chart, nil,
		win32.CLSCTX_INPROC_SERVER|win32.CLSCTX_LOCAL_SERVER,
		&IID_Chart_, unsafe.Pointer(&p))
	if win32.FAILED(hr) {
		return nil, com.NewError(hr)
	}
	return NewChart(p, false, scoped), nil
}

func (this *Chart) RegisterEventHandlers(handlers ChartEventsHandlers) uint32 {
	var cpc *win32.IConnectionPointContainer
	hr := this.QueryInterface(&win32.IID_IConnectionPointContainer, unsafe.Pointer(&cpc))
	win32.ASSERT_SUCCEEDED(hr)

	var cp *win32.IConnectionPoint
	hr = cpc.FindConnectionPoint(&IID_ChartEvents, &cp)
	win32.ASSERT_SUCCEEDED(hr)

	dispImpl := &ChartEventsDispImpl{Handlers: handlers}
	disp := NewChartEventsComObj(dispImpl, false)

	var cookie uint32
	hr = cp.Advise(disp.IUnknown(), &cookie)
	win32.ASSERT_SUCCEEDED(hr)

	disp.Release()
	cp.Release()
	cpc.Release()
	return cookie
}

func (this *Chart) UnRegisterEventHandlers(cookie uint32) {
	var cpc *win32.IConnectionPointContainer
	hr := this.QueryInterface(&win32.IID_IConnectionPointContainer, unsafe.Pointer(&cpc))
	win32.ASSERT_SUCCEEDED(hr)

	var cp *win32.IConnectionPoint
	hr = cpc.FindConnectionPoint(&IID_ChartEvents, &cp)
	win32.ASSERT_SUCCEEDED(hr)

	hr = cp.Unadvise(cookie)
	win32.ASSERT_SUCCEEDED(hr)

	cp.Release()
	cpc.Release()
}
