package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"syscall"
	"unsafe"
)

// 000244B0-0001-0000-C000-000000000046
var IID_IRectangularGradient = syscall.GUID{0x000244B0, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IRectangularGradient struct {
	win32.IDispatch
}

func NewIRectangularGradient(pUnk *win32.IUnknown, addRef bool, scoped bool) *IRectangularGradient {
	p := (*IRectangularGradient)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IRectangularGradient) IID() *syscall.GUID {
	return &IID_IRectangularGradient
}

func (this *IRectangularGradient) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IRectangularGradient) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRectangularGradient) GetParent(rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IRectangularGradient) GetColorStops(rhs **ColorStops) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IRectangularGradient) GetRectangleTop(rhs *float64) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRectangularGradient) SetRectangleTop(rhs float64) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IRectangularGradient) GetRectangleBottom(rhs *float64) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRectangularGradient) SetRectangleBottom(rhs float64) com.Error {
	addr := (*this.LpVtbl)[14]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IRectangularGradient) GetRectangleLeft(rhs *float64) com.Error {
	addr := (*this.LpVtbl)[15]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRectangularGradient) SetRectangleLeft(rhs float64) com.Error {
	addr := (*this.LpVtbl)[16]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IRectangularGradient) GetRectangleRight(rhs *float64) com.Error {
	addr := (*this.LpVtbl)[17]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IRectangularGradient) SetRectangleRight(rhs float64) com.Error {
	addr := (*this.LpVtbl)[18]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

