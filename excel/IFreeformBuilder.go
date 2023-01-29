package excel

import (
	"github.com/zzl/go-win32api/v2/win32"
	"github.com/zzl/go-com/com"
	"syscall"
	"unsafe"
)

// 0002443F-0001-0000-C000-000000000046
var IID_IFreeformBuilder = syscall.GUID{0x0002443F, 0x0001, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IFreeformBuilder struct {
	win32.IDispatch
}

func NewIFreeformBuilder(pUnk *win32.IUnknown, addRef bool, scoped bool) *IFreeformBuilder {
	if pUnk == nil {
		return nil
	}
	p := (*IFreeformBuilder)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IFreeformBuilder) IID() *syscall.GUID {
	return &IID_IFreeformBuilder
}

func (this *IFreeformBuilder) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IFreeformBuilder) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IFreeformBuilder) GetParent(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IFreeformBuilder) AddNodes(segmentType int32, editingType int32, x1 float32, y1 float32, x2 interface{}, y2 interface{}, x3 interface{}, y3 interface{}) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(segmentType), uintptr(editingType), uintptr(x1), uintptr(y1), (uintptr)(unsafe.Pointer(&x2)), (uintptr)(unsafe.Pointer(&y2)), (uintptr)(unsafe.Pointer(&x3)), (uintptr)(unsafe.Pointer(&y3)))
	return com.Error(ret)
}

func (this *IFreeformBuilder) ConvertToShape(rhs **Shape) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

