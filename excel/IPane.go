package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020895-0001-0000-C000-000000000046
var IID_IPane = syscall.GUID{0x00020895, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IPane struct {
	win32.IDispatch
}

func NewIPane(pUnk *win32.IUnknown, addRef bool, scoped bool) *IPane {
	 if pUnk == nil {
		return nil;
	}
	p := (*IPane)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IPane) IID() *syscall.GUID {
	return &IID_IPane
}

func (this *IPane) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPane) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPane) GetParent(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPane) Activate(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPane) GetIndex(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPane) LargeScroll(down interface{}, up interface{}, toRight interface{}, toLeft interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&down)), (uintptr)(unsafe.Pointer(&up)), (uintptr)(unsafe.Pointer(&toRight)), (uintptr)(unsafe.Pointer(&toLeft)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPane) GetScrollColumn(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPane) SetScrollColumn(rhs int32) com.Error {
	addr := (*this.LpVtbl)[14]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IPane) GetScrollRow(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[15]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPane) SetScrollRow(rhs int32) com.Error {
	addr := (*this.LpVtbl)[16]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IPane) SmallScroll(down interface{}, up interface{}, toRight interface{}, toLeft interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[17]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&down)), (uintptr)(unsafe.Pointer(&up)), (uintptr)(unsafe.Pointer(&toRight)), (uintptr)(unsafe.Pointer(&toLeft)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPane) GetVisibleRange(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[18]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPane) ScrollIntoView(left int32, top int32, width int32, height int32, start interface{}) com.Error {
	addr := (*this.LpVtbl)[19]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(left), uintptr(top), uintptr(width), uintptr(height), (uintptr)(unsafe.Pointer(&start)))
	return com.Error(ret)
}

func (this *IPane) PointsToScreenPixelsX(points int32, rhs *int32) com.Error {
	addr := (*this.LpVtbl)[20]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(points), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPane) PointsToScreenPixelsY(points int32, rhs *int32) com.Error {
	addr := (*this.LpVtbl)[21]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(points), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

