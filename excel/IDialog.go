package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"syscall"
	"unsafe"
)

// 0002087A-0001-0000-C000-000000000046
var IID_IDialog = syscall.GUID{0x0002087A, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IDialog struct {
	win32.IDispatch
}

func NewIDialog(pUnk *win32.IUnknown, addRef bool, scoped bool) *IDialog {
	 if pUnk == nil {
		return nil;
	}
	p := (*IDialog)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IDialog) IID() *syscall.GUID {
	return &IID_IDialog
}

func (this *IDialog) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IDialog) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialog) GetParent(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IDialog) Show(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

