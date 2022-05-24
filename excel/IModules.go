package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000208AE-0001-0000-C000-000000000046
var IID_IModules = syscall.GUID{0x000208AE, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IModules struct {
	win32.IDispatch
}

func NewIModules(pUnk *win32.IUnknown, addRef bool, scoped bool) *IModules {
	 if pUnk == nil {
		return nil;
	}
	p := (*IModules)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IModules) IID() *syscall.GUID {
	return &IID_IModules
}

func (this *IModules) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IModules) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IModules) GetParent(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IModules) Add(before interface{}, after interface{}, count interface{}, rhs **Module) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&before)), (uintptr)(unsafe.Pointer(&after)), (uintptr)(unsafe.Pointer(&count)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IModules) Copy(before interface{}, after interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&before)), (uintptr)(unsafe.Pointer(&after)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IModules) GetCount(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IModules) Delete(lcid int32) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IModules) Dummy7_()  {
	addr := (*this.LpVtbl)[14]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IModules) GetItem(index interface{}, rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[15]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IModules) Move(before interface{}, after interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[16]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&before)), (uintptr)(unsafe.Pointer(&after)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IModules) GetNewEnum_(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[17]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IModules) PrintOut__(from interface{}, to interface{}, copies interface{}, preview interface{}, activePrinter interface{}, printToFile interface{}, collate interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[18]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&from)), (uintptr)(unsafe.Pointer(&to)), (uintptr)(unsafe.Pointer(&copies)), (uintptr)(unsafe.Pointer(&preview)), (uintptr)(unsafe.Pointer(&activePrinter)), (uintptr)(unsafe.Pointer(&printToFile)), (uintptr)(unsafe.Pointer(&collate)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IModules) Dummy12_()  {
	addr := (*this.LpVtbl)[19]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IModules) Select(replace interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[20]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&replace)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IModules) GetHPageBreaks(rhs **HPageBreaks) com.Error {
	addr := (*this.LpVtbl)[21]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IModules) GetVPageBreaks(rhs **VPageBreaks) com.Error {
	addr := (*this.LpVtbl)[22]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IModules) GetVisible(lcid int32, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[23]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IModules) SetVisible(lcid int32, rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[24]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IModules) GetDefault_(index interface{}, rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[25]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IModules) PrintOut_(from interface{}, to interface{}, copies interface{}, preview interface{}, activePrinter interface{}, printToFile interface{}, collate interface{}, prToFileName interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[26]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&from)), (uintptr)(unsafe.Pointer(&to)), (uintptr)(unsafe.Pointer(&copies)), (uintptr)(unsafe.Pointer(&preview)), (uintptr)(unsafe.Pointer(&activePrinter)), (uintptr)(unsafe.Pointer(&printToFile)), (uintptr)(unsafe.Pointer(&collate)), (uintptr)(unsafe.Pointer(&prToFileName)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IModules) PrintOut(from interface{}, to interface{}, copies interface{}, preview interface{}, activePrinter interface{}, printToFile interface{}, collate interface{}, prToFileName interface{}, ignorePrintAreas interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[27]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&from)), (uintptr)(unsafe.Pointer(&to)), (uintptr)(unsafe.Pointer(&copies)), (uintptr)(unsafe.Pointer(&preview)), (uintptr)(unsafe.Pointer(&activePrinter)), (uintptr)(unsafe.Pointer(&printToFile)), (uintptr)(unsafe.Pointer(&collate)), (uintptr)(unsafe.Pointer(&prToFileName)), (uintptr)(unsafe.Pointer(&ignorePrintAreas)), uintptr(lcid))
	return com.Error(ret)
}

