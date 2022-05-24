package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244AC-0001-0000-C000-000000000046
var IID_IResearch = syscall.GUID{0x000244AC, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IResearch struct {
	win32.IDispatch
}

func NewIResearch(pUnk *win32.IUnknown, addRef bool, scoped bool) *IResearch {
	 if pUnk == nil {
		return nil;
	}
	p := (*IResearch)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IResearch) IID() *syscall.GUID {
	return &IID_IResearch
}

func (this *IResearch) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IResearch) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IResearch) GetParent(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IResearch) Query(serviceID string, queryString interface{}, queryLanguage interface{}, useSelection interface{}, launchQuery interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(serviceID)), (uintptr)(unsafe.Pointer(&queryString)), (uintptr)(unsafe.Pointer(&queryLanguage)), (uintptr)(unsafe.Pointer(&useSelection)), (uintptr)(unsafe.Pointer(&launchQuery)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IResearch) IsResearchService(serviceID string, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(serviceID)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IResearch) SetLanguagePair(languageFrom int32, languageTo int32, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(languageFrom), uintptr(languageTo), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

