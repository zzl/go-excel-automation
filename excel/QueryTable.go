package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

var CLSID_QueryTable = syscall.GUID{0x59191DA1, 0xEA47, 0x11CE,
	[8]byte{0xA5, 0x1F, 0x00, 0xAA, 0x00, 0x61, 0x50, 0x7F}}

type QueryTable struct {
	QueryTable_
}

func NewQueryTable(pDisp *win32.IDispatch, addRef bool, scoped bool) *QueryTable {
	if pDisp == nil {
		return nil
	}
	p := &QueryTable{QueryTable_{ole.OleClient{pDisp}}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func NewQueryTableFromVar(v ole.Variant, addRef bool, scoped bool) *QueryTable {
	return NewQueryTable(v.IDispatch(), addRef, scoped)
}

func NewQueryTableInstance(scoped bool) (*QueryTable, error) {
	var p *win32.IDispatch
	hr := win32.CoCreateInstance(&CLSID_QueryTable, nil,
		win32.CLSCTX_INPROC_SERVER|win32.CLSCTX_LOCAL_SERVER,
		&IID_QueryTable_, unsafe.Pointer(&p))
	if win32.FAILED(hr) {
		return nil, com.NewError(hr)
	}
	return NewQueryTable(p, false, scoped), nil
}

func (this *QueryTable) RegisterEventHandlers(handlers RefreshEventsHandlers) uint32 {
	var cpc *win32.IConnectionPointContainer
	hr := this.QueryInterface(&win32.IID_IConnectionPointContainer, unsafe.Pointer(&cpc))
	win32.ASSERT_SUCCEEDED(hr)

	var cp *win32.IConnectionPoint
	hr = cpc.FindConnectionPoint(&IID_RefreshEvents, &cp)
	win32.ASSERT_SUCCEEDED(hr)

	dispImpl := &RefreshEventsDispImpl{Handlers: handlers}
	disp := NewRefreshEventsComObj(dispImpl, false)

	var cookie uint32
	hr = cp.Advise(disp.IUnknown(), &cookie)
	win32.ASSERT_SUCCEEDED(hr)

	disp.Release()
	cp.Release()
	cpc.Release()
	return cookie
}

func (this *QueryTable) UnRegisterEventHandlers(cookie uint32) {
	var cpc *win32.IConnectionPointContainer
	hr := this.QueryInterface(&win32.IID_IConnectionPointContainer, unsafe.Pointer(&cpc))
	win32.ASSERT_SUCCEEDED(hr)

	var cp *win32.IConnectionPoint
	hr = cpc.FindConnectionPoint(&IID_RefreshEvents, &cp)
	win32.ASSERT_SUCCEEDED(hr)

	hr = cp.Unadvise(cookie)
	win32.ASSERT_SUCCEEDED(hr)

	cp.Release()
	cpc.Release()
}
