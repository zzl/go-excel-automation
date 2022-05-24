package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"syscall"
	"unsafe"
)

// 0002443A-0001-0000-C000-000000000046
var IID_IShapes = syscall.GUID{0x0002443A, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IShapes struct {
	win32.IDispatch
}

func NewIShapes(pUnk *win32.IUnknown, addRef bool, scoped bool) *IShapes {
	 if pUnk == nil {
		return nil;
	}
	p := (*IShapes)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IShapes) IID() *syscall.GUID {
	return &IID_IShapes
}

func (this *IShapes) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IShapes) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IShapes) GetParent(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IShapes) GetCount(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IShapes) Item(index interface{}, rhs **Shape) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IShapes) Default_(index interface{}, rhs **Shape) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IShapes) GetNewEnum_(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IShapes) AddCallout(type_ int32, left float32, top float32, width float32, height float32, rhs **Shape) com.Error {
	addr := (*this.LpVtbl)[14]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(type_), uintptr(left), uintptr(top), uintptr(width), uintptr(height), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IShapes) AddConnector(type_ int32, beginX float32, beginY float32, endX float32, endY float32, rhs **Shape) com.Error {
	addr := (*this.LpVtbl)[15]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(type_), uintptr(beginX), uintptr(beginY), uintptr(endX), uintptr(endY), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IShapes) AddCurve(safeArrayOfPoints interface{}, rhs **Shape) com.Error {
	addr := (*this.LpVtbl)[16]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&safeArrayOfPoints)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IShapes) AddLabel(orientation int32, left float32, top float32, width float32, height float32, rhs **Shape) com.Error {
	addr := (*this.LpVtbl)[17]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(orientation), uintptr(left), uintptr(top), uintptr(width), uintptr(height), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IShapes) AddLine(beginX float32, beginY float32, endX float32, endY float32, rhs **Shape) com.Error {
	addr := (*this.LpVtbl)[18]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(beginX), uintptr(beginY), uintptr(endX), uintptr(endY), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IShapes) AddPicture(filename string, linkToFile int32, saveWithDocument int32, left float32, top float32, width float32, height float32, rhs **Shape) com.Error {
	addr := (*this.LpVtbl)[19]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(filename)), uintptr(linkToFile), uintptr(saveWithDocument), uintptr(left), uintptr(top), uintptr(width), uintptr(height), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IShapes) AddPolyline(safeArrayOfPoints interface{}, rhs **Shape) com.Error {
	addr := (*this.LpVtbl)[20]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&safeArrayOfPoints)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IShapes) AddShape(type_ int32, left float32, top float32, width float32, height float32, rhs **Shape) com.Error {
	addr := (*this.LpVtbl)[21]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(type_), uintptr(left), uintptr(top), uintptr(width), uintptr(height), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IShapes) AddTextEffect(presetTextEffect int32, text string, fontName string, fontSize float32, fontBold int32, fontItalic int32, left float32, top float32, rhs **Shape) com.Error {
	addr := (*this.LpVtbl)[22]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(presetTextEffect), uintptr(win32.StrToPointer(text)), uintptr(win32.StrToPointer(fontName)), uintptr(fontSize), uintptr(fontBold), uintptr(fontItalic), uintptr(left), uintptr(top), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IShapes) AddTextbox(orientation int32, left float32, top float32, width float32, height float32, rhs **Shape) com.Error {
	addr := (*this.LpVtbl)[23]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(orientation), uintptr(left), uintptr(top), uintptr(width), uintptr(height), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IShapes) BuildFreeform(editingType int32, x1 float32, y1 float32, rhs **FreeformBuilder) com.Error {
	addr := (*this.LpVtbl)[24]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(editingType), uintptr(x1), uintptr(y1), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IShapes) GetRange(index interface{}, rhs **ShapeRange) com.Error {
	addr := (*this.LpVtbl)[25]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IShapes) SelectAll() com.Error {
	addr := (*this.LpVtbl)[26]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IShapes) AddFormControl(type_ int32, left int32, top int32, width int32, height int32, rhs **Shape) com.Error {
	addr := (*this.LpVtbl)[27]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(type_), uintptr(left), uintptr(top), uintptr(width), uintptr(height), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IShapes) AddOLEObject(classType interface{}, filename interface{}, link interface{}, displayAsIcon interface{}, iconFileName interface{}, iconIndex interface{}, iconLabel interface{}, left interface{}, top interface{}, width interface{}, height interface{}, rhs **Shape) com.Error {
	addr := (*this.LpVtbl)[28]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&classType)), (uintptr)(unsafe.Pointer(&filename)), (uintptr)(unsafe.Pointer(&link)), (uintptr)(unsafe.Pointer(&displayAsIcon)), (uintptr)(unsafe.Pointer(&iconFileName)), (uintptr)(unsafe.Pointer(&iconIndex)), (uintptr)(unsafe.Pointer(&iconLabel)), (uintptr)(unsafe.Pointer(&left)), (uintptr)(unsafe.Pointer(&top)), (uintptr)(unsafe.Pointer(&width)), (uintptr)(unsafe.Pointer(&height)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IShapes) AddDiagram(type_ int32, left float32, top float32, width float32, height float32, rhs **Shape) com.Error {
	addr := (*this.LpVtbl)[29]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(type_), uintptr(left), uintptr(top), uintptr(width), uintptr(height), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IShapes) AddCanvas(left float32, top float32, width float32, height float32, rhs **Shape) com.Error {
	addr := (*this.LpVtbl)[30]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(left), uintptr(top), uintptr(width), uintptr(height), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IShapes) AddChart(xlChartType interface{}, left interface{}, top interface{}, width interface{}, height interface{}, rhs **Shape) com.Error {
	addr := (*this.LpVtbl)[31]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&xlChartType)), (uintptr)(unsafe.Pointer(&left)), (uintptr)(unsafe.Pointer(&top)), (uintptr)(unsafe.Pointer(&width)), (uintptr)(unsafe.Pointer(&height)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IShapes) AddSmartArt(layout *win32.IDispatch, left interface{}, top interface{}, width interface{}, height interface{}, rhs **Shape) com.Error {
	addr := (*this.LpVtbl)[32]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(layout)), (uintptr)(unsafe.Pointer(&left)), (uintptr)(unsafe.Pointer(&top)), (uintptr)(unsafe.Pointer(&width)), (uintptr)(unsafe.Pointer(&height)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

