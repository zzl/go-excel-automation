package excel

import (
	"github.com/zzl/go-win32api/v2/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002086C-0001-0000-C000-000000000046
var IID_ISeriesCollection = syscall.GUID{0x0002086C, 0x0001, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ISeriesCollection struct {
	win32.IDispatch
}

func NewISeriesCollection(pUnk *win32.IUnknown, addRef bool, scoped bool) *ISeriesCollection {
	if pUnk == nil {
		return nil
	}
	p := (*ISeriesCollection)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *ISeriesCollection) IID() *syscall.GUID {
	return &IID_ISeriesCollection
}

func (this *ISeriesCollection) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *ISeriesCollection) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ISeriesCollection) GetParent(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *ISeriesCollection) Add(source interface{}, rowcol int32, seriesLabels interface{}, categoryLabels interface{}, replace interface{}, rhs **Series) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&source)), uintptr(rowcol), (uintptr)(unsafe.Pointer(&seriesLabels)), (uintptr)(unsafe.Pointer(&categoryLabels)), (uintptr)(unsafe.Pointer(&replace)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *ISeriesCollection) GetCount(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ISeriesCollection) Extend(source interface{}, rowcol interface{}, categoryLabels interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&source)), (uintptr)(unsafe.Pointer(&rowcol)), (uintptr)(unsafe.Pointer(&categoryLabels)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ISeriesCollection) Item(index interface{}, rhs **Series) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *ISeriesCollection) NewEnum_(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[14]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *ISeriesCollection) Paste(rowcol int32, seriesLabels interface{}, categoryLabels interface{}, replace interface{}, newSeries interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[15]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rowcol), (uintptr)(unsafe.Pointer(&seriesLabels)), (uintptr)(unsafe.Pointer(&categoryLabels)), (uintptr)(unsafe.Pointer(&replace)), (uintptr)(unsafe.Pointer(&newSeries)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ISeriesCollection) NewSeries(rhs **Series) com.Error {
	addr := (*this.LpVtbl)[16]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *ISeriesCollection) Default_(index interface{}, rhs **Series) com.Error {
	addr := (*this.LpVtbl)[17]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(rhs)))
	com.AddToScope(rhs)
	return com.Error(ret)
}

