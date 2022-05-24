package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"syscall"
	"unsafe"
)

// 000244C5-0001-0000-C000-000000000046
var IID_ISlicerCacheLevels = syscall.GUID{0x000244C5, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ISlicerCacheLevels struct {
	win32.IDispatch
}

func NewISlicerCacheLevels(pUnk *win32.IUnknown, addRef bool, scoped bool) *ISlicerCacheLevels {
	 if pUnk == nil {
		return nil;
	}
	p := (*ISlicerCacheLevels)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *ISlicerCacheLevels) IID() *syscall.GUID {
	return &IID_ISlicerCacheLevels
}

func (this *ISlicerCacheLevels) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *ISlicerCacheLevels) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ISlicerCacheLevels) GetParent(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *ISlicerCacheLevels) GetCount(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ISlicerCacheLevels) GetItem(level interface{}, rhs **SlicerCacheLevel) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&level)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *ISlicerCacheLevels) GetDefault_(level interface{}, rhs **SlicerCacheLevel) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&level)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *ISlicerCacheLevels) GetNewEnum_(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

