package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 0002440F-0001-0000-C000-000000000046
var IID_IChartEvents = syscall.GUID{0x0002440F, 0x0001, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IChartEvents struct {
	win32.IDispatch
}

func NewIChartEvents(pUnk *win32.IUnknown, addRef bool, scoped bool) *IChartEvents {
	if pUnk == nil {
		return nil
	}
	p := (*IChartEvents)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IChartEvents) IID() *syscall.GUID {
	return &IID_IChartEvents
}

func (this *IChartEvents) Activate() com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IChartEvents) Deactivate() com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IChartEvents) Resize() com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IChartEvents) MouseDown(button int32, shift int32, x int32, y int32) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(button), uintptr(shift), uintptr(x), uintptr(y))
	return com.Error(ret)
}

func (this *IChartEvents) MouseUp(button int32, shift int32, x int32, y int32) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(button), uintptr(shift), uintptr(x), uintptr(y))
	return com.Error(ret)
}

func (this *IChartEvents) MouseMove(button int32, shift int32, x int32, y int32) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(button), uintptr(shift), uintptr(x), uintptr(y))
	return com.Error(ret)
}

func (this *IChartEvents) BeforeRightClick(cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IChartEvents) DragPlot() com.Error {
	addr := (*this.LpVtbl)[14]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IChartEvents) DragOver() com.Error {
	addr := (*this.LpVtbl)[15]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IChartEvents) BeforeDoubleClick(elementID int32, arg1 int32, arg2 int32, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[16]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(elementID), uintptr(arg1), uintptr(arg2), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IChartEvents) Select(elementID int32, arg1 int32, arg2 int32) com.Error {
	addr := (*this.LpVtbl)[17]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(elementID), uintptr(arg1), uintptr(arg2))
	return com.Error(ret)
}

func (this *IChartEvents) SeriesChange(seriesIndex int32, pointIndex int32) com.Error {
	addr := (*this.LpVtbl)[18]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(seriesIndex), uintptr(pointIndex))
	return com.Error(ret)
}

func (this *IChartEvents) Calculate() com.Error {
	addr := (*this.LpVtbl)[19]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}
