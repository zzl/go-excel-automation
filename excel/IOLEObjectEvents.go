package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"syscall"
	"unsafe"
)

// 00024410-0001-0000-C000-000000000046
var IID_IOLEObjectEvents = syscall.GUID{0x00024410, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IOLEObjectEvents struct {
	win32.IDispatch
}

func NewIOLEObjectEvents(pUnk *win32.IUnknown, addRef bool, scoped bool) *IOLEObjectEvents {
	p := (*IOLEObjectEvents)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IOLEObjectEvents) IID() *syscall.GUID {
	return &IID_IOLEObjectEvents
}

func (this *IOLEObjectEvents) GotFocus() com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IOLEObjectEvents) LostFocus() com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

