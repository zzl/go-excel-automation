package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00024454-0001-0000-C000-000000000046
var IID_ICalculatedMembers = syscall.GUID{0x00024454, 0x0001, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ICalculatedMembers struct {
	win32.IDispatch
}

func NewICalculatedMembers(pUnk *win32.IUnknown, addRef bool, scoped bool) *ICalculatedMembers {
	if pUnk == nil {
		return nil
	}
	p := (*ICalculatedMembers)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *ICalculatedMembers) IID() *syscall.GUID {
	return &IID_ICalculatedMembers
}

func (this *ICalculatedMembers) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *ICalculatedMembers) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ICalculatedMembers) GetParent(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *ICalculatedMembers) GetCount(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ICalculatedMembers) GetItem(index interface{}, rhs **CalculatedMember) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *ICalculatedMembers) GetDefault_(index interface{}, rhs **CalculatedMember) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *ICalculatedMembers) GetNewEnum_(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *ICalculatedMembers) Add_(name string, formula string, solveOrder interface{}, type_ interface{}, rhs **CalculatedMember) com.Error {
	addr := (*this.LpVtbl)[14]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(name)), uintptr(win32.StrToPointer(formula)), (uintptr)(unsafe.Pointer(&solveOrder)), (uintptr)(unsafe.Pointer(&type_)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *ICalculatedMembers) Add(name string, formula interface{}, solveOrder interface{}, type_ interface{}, dynamic interface{}, displayFolder interface{}, hierarchizeDistinct interface{}, rhs **CalculatedMember) com.Error {
	addr := (*this.LpVtbl)[15]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(name)), (uintptr)(unsafe.Pointer(&formula)), (uintptr)(unsafe.Pointer(&solveOrder)), (uintptr)(unsafe.Pointer(&type_)), (uintptr)(unsafe.Pointer(&dynamic)), (uintptr)(unsafe.Pointer(&displayFolder)), (uintptr)(unsafe.Pointer(&hierarchizeDistinct)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}
