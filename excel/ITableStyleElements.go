package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"syscall"
	"unsafe"
)

// 000244A6-0001-0000-C000-000000000046
var IID_ITableStyleElements = syscall.GUID{0x000244A6, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ITableStyleElements struct {
	win32.IDispatch
}

func NewITableStyleElements(pUnk *win32.IUnknown, addRef bool, scoped bool) *ITableStyleElements {
	 if pUnk == nil {
		return nil;
	}
	p := (*ITableStyleElements)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *ITableStyleElements) IID() *syscall.GUID {
	return &IID_ITableStyleElements
}

func (this *ITableStyleElements) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *ITableStyleElements) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITableStyleElements) GetParent(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *ITableStyleElements) GetCount(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITableStyleElements) Item(index int32, rhs **TableStyleElement) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(index), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *ITableStyleElements) GetDefault_(index int32, rhs **TableStyleElement) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(index), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *ITableStyleElements) GetNewEnum_(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

