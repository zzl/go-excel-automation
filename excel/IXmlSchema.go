package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00024479-0001-0000-C000-000000000046
var IID_IXmlSchema = syscall.GUID{0x00024479, 0x0001, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IXmlSchema struct {
	win32.IDispatch
}

func NewIXmlSchema(pUnk *win32.IUnknown, addRef bool, scoped bool) *IXmlSchema {
	if pUnk == nil {
		return nil
	}
	p := (*IXmlSchema)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IXmlSchema) IID() *syscall.GUID {
	return &IID_IXmlSchema
}

func (this *IXmlSchema) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IXmlSchema) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IXmlSchema) GetParent(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IXmlSchema) GetNamespace(rhs **XmlNamespace) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IXmlSchema) GetXML(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IXmlSchema) GetName(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}
