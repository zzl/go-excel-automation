package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"syscall"
	"unsafe"
)

// 0002446E-0001-0000-C000-000000000046
var IID_IRTD = syscall.GUID{0x0002446E, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IRTD struct {
	win32.IDispatch
}

func NewIRTD(pUnk *win32.IUnknown, addRef bool, scoped bool) *IRTD {
	p := (*IRTD)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IRTD) IID() *syscall.GUID {
	return &IID_IRTD
}

func (this *IRTD) GetThrottleInterval(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRTD) SetThrottleInterval(rhs int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IRTD) RefreshData() com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IRTD) RestartServers() com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

