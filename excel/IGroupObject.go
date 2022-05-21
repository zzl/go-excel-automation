package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020898-0001-0000-C000-000000000046
var IID_IGroupObject = syscall.GUID{0x00020898, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IGroupObject struct {
	win32.IDispatch
}

func NewIGroupObject(pUnk *win32.IUnknown, addRef bool, scoped bool) *IGroupObject {
	p := (*IGroupObject)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IGroupObject) IID() *syscall.GUID {
	return &IID_IGroupObject
}

func (this *IGroupObject) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IGroupObject) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) GetParent(rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IGroupObject) GetBottomRightCell(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IGroupObject) BringToFront(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) Copy(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) CopyPicture(appearance int32, format int32, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(appearance), uintptr(format), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) Cut(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[14]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) Delete(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[15]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) Duplicate(rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[16]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IGroupObject) GetEnabled(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[17]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SetEnabled(rhs bool) com.Error {
	addr := (*this.LpVtbl)[18]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IGroupObject) GetHeight(rhs *float64) com.Error {
	addr := (*this.LpVtbl)[19]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SetHeight(rhs float64) com.Error {
	addr := (*this.LpVtbl)[20]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IGroupObject) GetIndex(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[21]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) GetLeft(rhs *float64) com.Error {
	addr := (*this.LpVtbl)[22]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SetLeft(rhs float64) com.Error {
	addr := (*this.LpVtbl)[23]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IGroupObject) GetLocked(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[24]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SetLocked(rhs bool) com.Error {
	addr := (*this.LpVtbl)[25]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IGroupObject) GetName(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[26]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SetName(rhs string) com.Error {
	addr := (*this.LpVtbl)[27]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) GetOnAction(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[28]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SetOnAction(rhs string) com.Error {
	addr := (*this.LpVtbl)[29]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) GetPlacement(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[30]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SetPlacement(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[31]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) GetPrintObject(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[32]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SetPrintObject(rhs bool) com.Error {
	addr := (*this.LpVtbl)[33]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IGroupObject) Select(replace interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[34]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&replace)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SendToBack(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[35]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) GetTop(rhs *float64) com.Error {
	addr := (*this.LpVtbl)[36]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SetTop(rhs float64) com.Error {
	addr := (*this.LpVtbl)[37]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IGroupObject) GetTopLeftCell(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[38]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IGroupObject) GetVisible(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[39]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SetVisible(rhs bool) com.Error {
	addr := (*this.LpVtbl)[40]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IGroupObject) GetWidth(rhs *float64) com.Error {
	addr := (*this.LpVtbl)[41]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SetWidth(rhs float64) com.Error {
	addr := (*this.LpVtbl)[42]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IGroupObject) GetZOrder(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[43]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) GetShapeRange(rhs **ShapeRange) com.Error {
	addr := (*this.LpVtbl)[44]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IGroupObject) Dummy27_()  {
	addr := (*this.LpVtbl)[45]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Dummy28_()  {
	addr := (*this.LpVtbl)[46]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) GetAddIndent(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[47]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SetAddIndent(rhs bool) com.Error {
	addr := (*this.LpVtbl)[48]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IGroupObject) Dummy30_()  {
	addr := (*this.LpVtbl)[49]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) GetArrowHeadLength(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[50]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SetArrowHeadLength(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[51]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) GetArrowHeadStyle(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[52]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SetArrowHeadStyle(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[53]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) GetArrowHeadWidth(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[54]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SetArrowHeadWidth(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[55]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) GetAutoSize(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[56]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SetAutoSize(rhs bool) com.Error {
	addr := (*this.LpVtbl)[57]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IGroupObject) GetBorder(rhs **Border) com.Error {
	addr := (*this.LpVtbl)[58]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IGroupObject) Dummy36_()  {
	addr := (*this.LpVtbl)[59]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Dummy37_()  {
	addr := (*this.LpVtbl)[60]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Dummy38_()  {
	addr := (*this.LpVtbl)[61]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) CheckSpelling(customDictionary interface{}, ignoreUppercase interface{}, alwaysSuggest interface{}, spellLang interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[62]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&customDictionary)), (uintptr)(unsafe.Pointer(&ignoreUppercase)), (uintptr)(unsafe.Pointer(&alwaysSuggest)), (uintptr)(unsafe.Pointer(&spellLang)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) GetDefault_(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[63]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SetDefault_(rhs int32) com.Error {
	addr := (*this.LpVtbl)[64]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IGroupObject) Dummy41_()  {
	addr := (*this.LpVtbl)[65]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Dummy42_()  {
	addr := (*this.LpVtbl)[66]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Dummy43_()  {
	addr := (*this.LpVtbl)[67]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Dummy44_()  {
	addr := (*this.LpVtbl)[68]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Dummy45_()  {
	addr := (*this.LpVtbl)[69]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) GetFont(rhs **Font) com.Error {
	addr := (*this.LpVtbl)[70]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IGroupObject) Dummy47_()  {
	addr := (*this.LpVtbl)[71]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Dummy48_()  {
	addr := (*this.LpVtbl)[72]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) GetHorizontalAlignment(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[73]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SetHorizontalAlignment(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[74]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) Dummy50_()  {
	addr := (*this.LpVtbl)[75]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) GetInterior(rhs **Interior) com.Error {
	addr := (*this.LpVtbl)[76]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IGroupObject) Dummy52_()  {
	addr := (*this.LpVtbl)[77]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Dummy53_()  {
	addr := (*this.LpVtbl)[78]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Dummy54_()  {
	addr := (*this.LpVtbl)[79]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Dummy55_()  {
	addr := (*this.LpVtbl)[80]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Dummy56_()  {
	addr := (*this.LpVtbl)[81]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Dummy57_()  {
	addr := (*this.LpVtbl)[82]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Dummy58_()  {
	addr := (*this.LpVtbl)[83]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Dummy59_()  {
	addr := (*this.LpVtbl)[84]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Dummy60_()  {
	addr := (*this.LpVtbl)[85]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Dummy61_()  {
	addr := (*this.LpVtbl)[86]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Dummy62_()  {
	addr := (*this.LpVtbl)[87]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Dummy63_()  {
	addr := (*this.LpVtbl)[88]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) GetOrientation(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[89]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SetOrientation(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[90]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) Dummy65_()  {
	addr := (*this.LpVtbl)[91]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Dummy66_()  {
	addr := (*this.LpVtbl)[92]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Dummy67_()  {
	addr := (*this.LpVtbl)[93]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Dummy68_()  {
	addr := (*this.LpVtbl)[94]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) GetRoundedCorners(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[95]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SetRoundedCorners(rhs bool) com.Error {
	addr := (*this.LpVtbl)[96]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IGroupObject) Dummy70_()  {
	addr := (*this.LpVtbl)[97]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) GetShadow(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[98]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SetShadow(rhs bool) com.Error {
	addr := (*this.LpVtbl)[99]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IGroupObject) Dummy72_()  {
	addr := (*this.LpVtbl)[100]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Dummy73_()  {
	addr := (*this.LpVtbl)[101]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) Ungroup(rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[102]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IGroupObject) Dummy75_()  {
	addr := (*this.LpVtbl)[103]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) GetVerticalAlignment(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[104]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SetVerticalAlignment(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[105]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) Dummy77_()  {
	addr := (*this.LpVtbl)[106]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IGroupObject) GetReadingOrder(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[107]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IGroupObject) SetReadingOrder(rhs int32) com.Error {
	addr := (*this.LpVtbl)[108]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

