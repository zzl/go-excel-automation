package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000208A5-0001-0000-C000-000000000046
var IID_ITextBoxes = syscall.GUID{0x000208A5, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ITextBoxes struct {
	win32.IDispatch
}

func NewITextBoxes(pUnk *win32.IUnknown, addRef bool, scoped bool) *ITextBoxes {
	p := (*ITextBoxes)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *ITextBoxes) IID() *syscall.GUID {
	return &IID_ITextBoxes
}

func (this *ITextBoxes) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *ITextBoxes) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) GetParent(rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *ITextBoxes) Dummy3_()  {
	addr := (*this.LpVtbl)[10]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *ITextBoxes) BringToFront(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) Copy(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) CopyPicture(appearance int32, format int32, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(appearance), uintptr(format), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) Cut(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[14]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) Delete(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[15]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) Duplicate(rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[16]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *ITextBoxes) GetEnabled(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[17]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SetEnabled(rhs bool) com.Error {
	addr := (*this.LpVtbl)[18]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *ITextBoxes) GetHeight(rhs *float64) com.Error {
	addr := (*this.LpVtbl)[19]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SetHeight(rhs float64) com.Error {
	addr := (*this.LpVtbl)[20]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *ITextBoxes) Dummy12_()  {
	addr := (*this.LpVtbl)[21]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *ITextBoxes) GetLeft(rhs *float64) com.Error {
	addr := (*this.LpVtbl)[22]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SetLeft(rhs float64) com.Error {
	addr := (*this.LpVtbl)[23]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *ITextBoxes) GetLocked(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[24]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SetLocked(rhs bool) com.Error {
	addr := (*this.LpVtbl)[25]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *ITextBoxes) Dummy15_()  {
	addr := (*this.LpVtbl)[26]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *ITextBoxes) GetOnAction(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[27]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SetOnAction(rhs string) com.Error {
	addr := (*this.LpVtbl)[28]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) GetPlacement(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[29]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SetPlacement(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[30]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) GetPrintObject(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[31]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SetPrintObject(rhs bool) com.Error {
	addr := (*this.LpVtbl)[32]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *ITextBoxes) Select(replace interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[33]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&replace)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SendToBack(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[34]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) GetTop(rhs *float64) com.Error {
	addr := (*this.LpVtbl)[35]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SetTop(rhs float64) com.Error {
	addr := (*this.LpVtbl)[36]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *ITextBoxes) Dummy22_()  {
	addr := (*this.LpVtbl)[37]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *ITextBoxes) GetVisible(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[38]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SetVisible(rhs bool) com.Error {
	addr := (*this.LpVtbl)[39]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *ITextBoxes) GetWidth(rhs *float64) com.Error {
	addr := (*this.LpVtbl)[40]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SetWidth(rhs float64) com.Error {
	addr := (*this.LpVtbl)[41]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *ITextBoxes) GetZOrder(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[42]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) GetShapeRange(rhs **ShapeRange) com.Error {
	addr := (*this.LpVtbl)[43]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *ITextBoxes) GetAddIndent(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[44]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SetAddIndent(rhs bool) com.Error {
	addr := (*this.LpVtbl)[45]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *ITextBoxes) GetAutoScaleFont(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[46]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SetAutoScaleFont(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[47]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) GetAutoSize(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[48]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SetAutoSize(rhs bool) com.Error {
	addr := (*this.LpVtbl)[49]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *ITextBoxes) GetCaption(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[50]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SetCaption(rhs string) com.Error {
	addr := (*this.LpVtbl)[51]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) GetCharacters(start interface{}, length interface{}, rhs **Characters) com.Error {
	addr := (*this.LpVtbl)[52]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&start)), (uintptr)(unsafe.Pointer(&length)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *ITextBoxes) CheckSpelling(customDictionary interface{}, ignoreUppercase interface{}, alwaysSuggest interface{}, spellLang interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[53]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&customDictionary)), (uintptr)(unsafe.Pointer(&ignoreUppercase)), (uintptr)(unsafe.Pointer(&alwaysSuggest)), (uintptr)(unsafe.Pointer(&spellLang)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) GetFont(rhs **Font) com.Error {
	addr := (*this.LpVtbl)[54]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *ITextBoxes) GetFormula(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[55]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SetFormula(rhs string) com.Error {
	addr := (*this.LpVtbl)[56]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) GetHorizontalAlignment(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[57]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SetHorizontalAlignment(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[58]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) GetLockedText(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[59]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SetLockedText(rhs bool) com.Error {
	addr := (*this.LpVtbl)[60]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *ITextBoxes) GetOrientation(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[61]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SetOrientation(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[62]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) GetText(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[63]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SetText(rhs string) com.Error {
	addr := (*this.LpVtbl)[64]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) GetVerticalAlignment(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[65]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SetVerticalAlignment(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[66]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) GetReadingOrder(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[67]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SetReadingOrder(rhs int32) com.Error {
	addr := (*this.LpVtbl)[68]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *ITextBoxes) GetBorder(rhs **Border) com.Error {
	addr := (*this.LpVtbl)[69]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *ITextBoxes) GetInterior(rhs **Interior) com.Error {
	addr := (*this.LpVtbl)[70]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *ITextBoxes) GetRoundedCorners(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[71]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SetRoundedCorners(rhs bool) com.Error {
	addr := (*this.LpVtbl)[72]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *ITextBoxes) GetShadow(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[73]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) SetShadow(rhs bool) com.Error {
	addr := (*this.LpVtbl)[74]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *ITextBoxes) Add(left float64, top float64, width float64, height float64, rhs **TextBox) com.Error {
	addr := (*this.LpVtbl)[75]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(left), uintptr(top), uintptr(width), uintptr(height), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *ITextBoxes) GetCount(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[76]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *ITextBoxes) Group(rhs **GroupObject) com.Error {
	addr := (*this.LpVtbl)[77]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *ITextBoxes) Item(index interface{}, rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[78]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *ITextBoxes) NewEnum_(rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[79]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

