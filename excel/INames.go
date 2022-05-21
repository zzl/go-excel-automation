package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"syscall"
	"unsafe"
)

// 000208B8-0001-0000-C000-000000000046
var IID_INames = syscall.GUID{0x000208B8, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type INames struct {
	win32.IDispatch
}

func NewINames(pUnk *win32.IUnknown, addRef bool, scoped bool) *INames {
	p := (*INames)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *INames) IID() *syscall.GUID {
	return &IID_INames
}

func (this *INames) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *INames) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *INames) GetParent(rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *INames) Add(name interface{}, refersTo interface{}, visible interface{}, macroType interface{}, shortcutKey interface{}, category interface{}, nameLocal interface{}, refersToLocal interface{}, categoryLocal interface{}, refersToR1C1 interface{}, refersToR1C1Local interface{}, rhs **Name) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&name)), (uintptr)(unsafe.Pointer(&refersTo)), (uintptr)(unsafe.Pointer(&visible)), (uintptr)(unsafe.Pointer(&macroType)), (uintptr)(unsafe.Pointer(&shortcutKey)), (uintptr)(unsafe.Pointer(&category)), (uintptr)(unsafe.Pointer(&nameLocal)), (uintptr)(unsafe.Pointer(&refersToLocal)), (uintptr)(unsafe.Pointer(&categoryLocal)), (uintptr)(unsafe.Pointer(&refersToR1C1)), (uintptr)(unsafe.Pointer(&refersToR1C1Local)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *INames) Item(index interface{}, indexLocal interface{}, refersTo interface{}, lcid int32, rhs **Name) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), (uintptr)(unsafe.Pointer(&indexLocal)), (uintptr)(unsafe.Pointer(&refersTo)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *INames) Default_(index interface{}, indexLocal interface{}, refersTo interface{}, lcid int32, rhs **Name) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), (uintptr)(unsafe.Pointer(&indexLocal)), (uintptr)(unsafe.Pointer(&refersTo)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *INames) GetCount(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *INames) GetNewEnum_(rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[14]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

