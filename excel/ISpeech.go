package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00024466-0001-0000-C000-000000000046
var IID_ISpeech = syscall.GUID{0x00024466, 0x0001, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ISpeech struct {
	win32.IDispatch
}

func NewISpeech(pUnk *win32.IUnknown, addRef bool, scoped bool) *ISpeech {
	if pUnk == nil {
		return nil
	}
	p := (*ISpeech)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *ISpeech) IID() *syscall.GUID {
	return &IID_ISpeech
}

func (this *ISpeech) Speak(text string, speakAsync interface{}, speakXML interface{}, purge interface{}) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(text)), (uintptr)(unsafe.Pointer(&speakAsync)), (uintptr)(unsafe.Pointer(&speakXML)), (uintptr)(unsafe.Pointer(&purge)))
	return com.Error(ret)
}

func (this *ISpeech) GetDirection(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ISpeech) SetDirection(rhs int32) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *ISpeech) GetSpeakCellOnEnter(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ISpeech) SetSpeakCellOnEnter(rhs bool) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}
