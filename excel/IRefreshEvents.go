package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"syscall"
	"unsafe"
)

// 0002441B-0001-0000-C000-000000000046
var IID_IRefreshEvents = syscall.GUID{0x0002441B, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IRefreshEvents struct {
	win32.IDispatch
}

func NewIRefreshEvents(pUnk *win32.IUnknown, addRef bool, scoped bool) *IRefreshEvents {
	 if pUnk == nil {
		return nil;
	}
	p := (*IRefreshEvents)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IRefreshEvents) IID() *syscall.GUID {
	return &IID_IRefreshEvents
}

func (this *IRefreshEvents) BeforeRefresh(cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IRefreshEvents) AfterRefresh(success bool) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&success))))
	return com.Error(ret)
}

