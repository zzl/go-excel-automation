package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 000244AE-0001-0000-C000-000000000046
var IID_IColorStops = syscall.GUID{0x000244AE, 0x0001, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IColorStops struct {
	win32.IDispatch
}

func NewIColorStops(pUnk *win32.IUnknown, addRef bool, scoped bool) *IColorStops {
	if pUnk == nil {
		return nil
	}
	p := (*IColorStops)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IColorStops) IID() *syscall.GUID {
	return &IID_IColorStops
}

func (this *IColorStops) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IColorStops) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IColorStops) GetParent(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IColorStops) GetCount(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IColorStops) Add(position float64, rhs **ColorStop) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(position), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IColorStops) Item(index interface{}, rhs **ColorStop) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IColorStops) GetDefault_(index interface{}, rhs **ColorStop) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IColorStops) GetNewEnum_(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[14]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IColorStops) Clear() com.Error {
	addr := (*this.LpVtbl)[15]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}
