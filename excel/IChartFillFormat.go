package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"syscall"
	"unsafe"
)

// 00024435-0001-0000-C000-000000000046
var IID_IChartFillFormat = syscall.GUID{0x00024435, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IChartFillFormat struct {
	win32.IDispatch
}

func NewIChartFillFormat(pUnk *win32.IUnknown, addRef bool, scoped bool) *IChartFillFormat {
	p := (*IChartFillFormat)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IChartFillFormat) IID() *syscall.GUID {
	return &IID_IChartFillFormat
}

func (this *IChartFillFormat) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IChartFillFormat) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IChartFillFormat) GetParent(rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IChartFillFormat) OneColorGradient(style int32, variant int32, degree float32) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(style), uintptr(variant), uintptr(degree))
	return com.Error(ret)
}

func (this *IChartFillFormat) TwoColorGradient(style int32, variant int32) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(style), uintptr(variant))
	return com.Error(ret)
}

func (this *IChartFillFormat) PresetTextured(presetTexture int32) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(presetTexture))
	return com.Error(ret)
}

func (this *IChartFillFormat) Solid() com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IChartFillFormat) Patterned(pattern int32) com.Error {
	addr := (*this.LpVtbl)[14]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(pattern))
	return com.Error(ret)
}

func (this *IChartFillFormat) UserPicture(pictureFile interface{}, pictureFormat interface{}, pictureStackUnit interface{}, picturePlacement interface{}) com.Error {
	addr := (*this.LpVtbl)[15]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&pictureFile)), (uintptr)(unsafe.Pointer(&pictureFormat)), (uintptr)(unsafe.Pointer(&pictureStackUnit)), (uintptr)(unsafe.Pointer(&picturePlacement)))
	return com.Error(ret)
}

func (this *IChartFillFormat) UserTextured(textureFile string) com.Error {
	addr := (*this.LpVtbl)[16]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(textureFile)))
	return com.Error(ret)
}

func (this *IChartFillFormat) PresetGradient(style int32, variant int32, presetGradientType int32) com.Error {
	addr := (*this.LpVtbl)[17]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(style), uintptr(variant), uintptr(presetGradientType))
	return com.Error(ret)
}

func (this *IChartFillFormat) GetBackColor(rhs **ChartColorFormat) com.Error {
	addr := (*this.LpVtbl)[18]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IChartFillFormat) GetForeColor(rhs **ChartColorFormat) com.Error {
	addr := (*this.LpVtbl)[19]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IChartFillFormat) GetGradientColorType(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[20]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IChartFillFormat) GetGradientDegree(rhs *float32) com.Error {
	addr := (*this.LpVtbl)[21]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IChartFillFormat) GetGradientStyle(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[22]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IChartFillFormat) GetGradientVariant(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[23]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IChartFillFormat) GetPattern(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[24]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IChartFillFormat) GetPresetGradientType(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[25]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IChartFillFormat) GetPresetTexture(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[26]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IChartFillFormat) GetTextureName(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[27]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IChartFillFormat) GetTextureType(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[28]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IChartFillFormat) GetType(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[29]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IChartFillFormat) GetVisible(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[30]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IChartFillFormat) SetVisible(rhs int32) com.Error {
	addr := (*this.LpVtbl)[31]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

