package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"syscall"
	"unsafe"
)

// 00024470-0001-0000-C000-000000000046
var IID_IListObjects = syscall.GUID{0x00024470, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IListObjects struct {
	win32.IDispatch
}

func NewIListObjects(pUnk *win32.IUnknown, addRef bool, scoped bool) *IListObjects {
	 if pUnk == nil {
		return nil;
	}
	p := (*IListObjects)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IListObjects) IID() *syscall.GUID {
	return &IID_IListObjects
}

func (this *IListObjects) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IListObjects) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IListObjects) GetParent(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IListObjects) Add_(sourceType int32, source interface{}, linkSource interface{}, xlListObjectHasHeaders int32, destination interface{}, rhs **ListObject) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(sourceType), (uintptr)(unsafe.Pointer(&source)), (uintptr)(unsafe.Pointer(&linkSource)), uintptr(xlListObjectHasHeaders), (uintptr)(unsafe.Pointer(&destination)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IListObjects) GetDefault_(index interface{}, rhs **ListObject) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IListObjects) GetNewEnum_(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IListObjects) GetItem(index interface{}, rhs **ListObject) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IListObjects) GetCount(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[14]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IListObjects) Add(sourceType int32, source interface{}, linkSource interface{}, xlListObjectHasHeaders int32, destination interface{}, tableStyleName interface{}, rhs **ListObject) com.Error {
	addr := (*this.LpVtbl)[15]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(sourceType), (uintptr)(unsafe.Pointer(&source)), (uintptr)(unsafe.Pointer(&linkSource)), uintptr(xlListObjectHasHeaders), (uintptr)(unsafe.Pointer(&destination)), (uintptr)(unsafe.Pointer(&tableStyleName)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

