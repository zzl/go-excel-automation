package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"syscall"
	"unsafe"
)

// 0002446D-0001-0000-C000-000000000046
var IID_IUserAccess = syscall.GUID{0x0002446D, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IUserAccess struct {
	win32.IDispatch
}

func NewIUserAccess(pUnk *win32.IUnknown, addRef bool, scoped bool) *IUserAccess {
	 if pUnk == nil {
		return nil;
	}
	p := (*IUserAccess)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IUserAccess) IID() *syscall.GUID {
	return &IID_IUserAccess
}

func (this *IUserAccess) GetName(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IUserAccess) GetAllowEdit(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IUserAccess) SetAllowEdit(rhs bool) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IUserAccess) Delete() com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

