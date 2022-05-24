package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"syscall"
	"unsafe"
)

// 00024411-0001-0000-C000-000000000046
var IID_IDocEvents = syscall.GUID{0x00024411, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IDocEvents struct {
	win32.IDispatch
}

func NewIDocEvents(pUnk *win32.IUnknown, addRef bool, scoped bool) *IDocEvents {
	 if pUnk == nil {
		return nil;
	}
	p := (*IDocEvents)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IDocEvents) IID() *syscall.GUID {
	return &IID_IDocEvents
}

func (this *IDocEvents) SelectionChange(target *Range) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(target)))
	return com.Error(ret)
}

func (this *IDocEvents) BeforeDoubleClick(target *Range, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(target)), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IDocEvents) BeforeRightClick(target *Range, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(target)), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IDocEvents) Activate() com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IDocEvents) Deactivate() com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IDocEvents) Calculate() com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IDocEvents) Change(target *Range) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(target)))
	return com.Error(ret)
}

func (this *IDocEvents) FollowHyperlink(target *Hyperlink) com.Error {
	addr := (*this.LpVtbl)[14]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(target)))
	return com.Error(ret)
}

func (this *IDocEvents) PivotTableUpdate(target *PivotTable) com.Error {
	addr := (*this.LpVtbl)[15]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(target)))
	return com.Error(ret)
}

func (this *IDocEvents) PivotTableAfterValueChange(targetPivotTable *PivotTable, targetRange *Range) com.Error {
	addr := (*this.LpVtbl)[16]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(targetPivotTable)), uintptr(unsafe.Pointer(targetRange)))
	return com.Error(ret)
}

func (this *IDocEvents) PivotTableBeforeAllocateChanges(targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[17]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(targetPivotTable)), uintptr(valueChangeStart), uintptr(valueChangeEnd), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IDocEvents) PivotTableBeforeCommitChanges(targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[18]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(targetPivotTable)), uintptr(valueChangeStart), uintptr(valueChangeEnd), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IDocEvents) PivotTableBeforeDiscardChanges(targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32) com.Error {
	addr := (*this.LpVtbl)[19]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(targetPivotTable)), uintptr(valueChangeStart), uintptr(valueChangeEnd))
	return com.Error(ret)
}

func (this *IDocEvents) PivotTableChangeSync(target *PivotTable) com.Error {
	addr := (*this.LpVtbl)[20]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(target)))
	return com.Error(ret)
}

