package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

var CLSID_Global = syscall.GUID{0x00020812, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Global struct {
	Global_
}

func NewGlobal(pDisp *win32.IDispatch, addRef bool, scoped bool) *Global {
	 if pDisp == nil {
		return nil;
	}
	p := &Global{Global_{ole.OleClient{pDisp}}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func NewGlobalFromVar(v ole.Variant, addRef bool, scoped bool) *Global {
	return NewGlobal(v.IDispatch(), addRef, scoped)
}

func NewGlobalInstance(scoped bool) (*Global, error) {
	var p *win32.IDispatch
	hr := win32.CoCreateInstance(&CLSID_Global, nil, 
		win32.CLSCTX_INPROC_SERVER|win32.CLSCTX_LOCAL_SERVER,
		&IID_Global_, unsafe.Pointer(&p))
	if win32.FAILED(hr) {
		return nil, com.NewError(hr)
	}
	return NewGlobal(p, false, scoped), nil
}

