package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020896-0001-0000-C000-000000000046
var IID_IScenarios = syscall.GUID{0x00020896, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IScenarios struct {
	win32.IDispatch
}

func NewIScenarios(pUnk *win32.IUnknown, addRef bool, scoped bool) *IScenarios {
	 if pUnk == nil {
		return nil;
	}
	p := (*IScenarios)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IScenarios) IID() *syscall.GUID {
	return &IID_IScenarios
}

func (this *IScenarios) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IScenarios) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IScenarios) GetParent(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IScenarios) Add(name string, changingCells interface{}, values interface{}, comment interface{}, locked interface{}, hidden interface{}, rhs **Scenario) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(name)), (uintptr)(unsafe.Pointer(&changingCells)), (uintptr)(unsafe.Pointer(&values)), (uintptr)(unsafe.Pointer(&comment)), (uintptr)(unsafe.Pointer(&locked)), (uintptr)(unsafe.Pointer(&hidden)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IScenarios) GetCount(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IScenarios) CreateSummary(reportType int32, resultCells interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(reportType), (uintptr)(unsafe.Pointer(&resultCells)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IScenarios) Item(index interface{}, rhs **Scenario) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IScenarios) Merge(source interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[14]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&source)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IScenarios) NewEnum_(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[15]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

