package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 0002444F-0001-0000-C000-000000000046
var IID_ICanvasShapes = syscall.GUID{0x0002444F, 0x0001, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ICanvasShapes struct {
	win32.IDispatch
}

func NewICanvasShapes(pUnk *win32.IUnknown, addRef bool, scoped bool) *ICanvasShapes {
	if pUnk == nil {
		return nil
	}
	p := (*ICanvasShapes)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *ICanvasShapes) IID() *syscall.GUID {
	return &IID_ICanvasShapes
}
