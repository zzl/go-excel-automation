package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000208AF-0001-0000-C000-000000000046
var IID_IDialogSheet = syscall.GUID{0x000208AF, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IDialogSheet struct {
	win32.IDispatch
}

func NewIDialogSheet(pUnk *win32.IUnknown, addRef bool, scoped bool) *IDialogSheet {
	p := (*IDialogSheet)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IDialogSheet) IID() *syscall.GUID {
	return &IID_IDialogSheet
}

func (this *IDialogSheet) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) GetParent(rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) Activate(lcid int32) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IDialogSheet) Copy(before interface{}, after interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&before)), (uintptr)(unsafe.Pointer(&after)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IDialogSheet) Delete(lcid int32) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IDialogSheet) GetCodeName(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) GetCodeName_(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[14]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) SetCodeName_(rhs string) com.Error {
	addr := (*this.LpVtbl)[15]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) GetIndex(lcid int32, rhs *int32) com.Error {
	addr := (*this.LpVtbl)[16]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) Move(before interface{}, after interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[17]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&before)), (uintptr)(unsafe.Pointer(&after)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IDialogSheet) GetName(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[18]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) SetName(rhs string) com.Error {
	addr := (*this.LpVtbl)[19]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) GetNext(rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[20]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) GetOnDoubleClick(lcid int32, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[21]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) SetOnDoubleClick(lcid int32, rhs string) com.Error {
	addr := (*this.LpVtbl)[22]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) GetOnSheetActivate(lcid int32, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[23]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) SetOnSheetActivate(lcid int32, rhs string) com.Error {
	addr := (*this.LpVtbl)[24]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) GetOnSheetDeactivate(lcid int32, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[25]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) SetOnSheetDeactivate(lcid int32, rhs string) com.Error {
	addr := (*this.LpVtbl)[26]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) GetPageSetup(rhs **PageSetup) com.Error {
	addr := (*this.LpVtbl)[27]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) GetPrevious(rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[28]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) PrintOut__(from interface{}, to interface{}, copies interface{}, preview interface{}, activePrinter interface{}, printToFile interface{}, collate interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[29]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&from)), (uintptr)(unsafe.Pointer(&to)), (uintptr)(unsafe.Pointer(&copies)), (uintptr)(unsafe.Pointer(&preview)), (uintptr)(unsafe.Pointer(&activePrinter)), (uintptr)(unsafe.Pointer(&printToFile)), (uintptr)(unsafe.Pointer(&collate)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IDialogSheet) PrintPreview(enableChanges interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[30]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&enableChanges)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IDialogSheet) Protect_(password interface{}, drawingObjects interface{}, contents interface{}, scenarios interface{}, userInterfaceOnly interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[31]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&password)), (uintptr)(unsafe.Pointer(&drawingObjects)), (uintptr)(unsafe.Pointer(&contents)), (uintptr)(unsafe.Pointer(&scenarios)), (uintptr)(unsafe.Pointer(&userInterfaceOnly)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IDialogSheet) GetProtectContents(lcid int32, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[32]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) GetProtectDrawingObjects(lcid int32, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[33]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) GetProtectionMode(lcid int32, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[34]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) GetProtectScenarios(lcid int32, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[35]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) SaveAs_(filename string, fileFormat interface{}, password interface{}, writeResPassword interface{}, readOnlyRecommended interface{}, createBackup interface{}, addToMru interface{}, textCodepage interface{}, textVisualLayout interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[36]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(filename)), (uintptr)(unsafe.Pointer(&fileFormat)), (uintptr)(unsafe.Pointer(&password)), (uintptr)(unsafe.Pointer(&writeResPassword)), (uintptr)(unsafe.Pointer(&readOnlyRecommended)), (uintptr)(unsafe.Pointer(&createBackup)), (uintptr)(unsafe.Pointer(&addToMru)), (uintptr)(unsafe.Pointer(&textCodepage)), (uintptr)(unsafe.Pointer(&textVisualLayout)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IDialogSheet) Select(replace interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[37]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&replace)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IDialogSheet) Unprotect(password interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[38]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&password)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IDialogSheet) GetVisible(lcid int32, rhs *int32) com.Error {
	addr := (*this.LpVtbl)[39]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) SetVisible(lcid int32, rhs int32) com.Error {
	addr := (*this.LpVtbl)[40]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(rhs))
	return com.Error(ret)
}

func (this *IDialogSheet) GetShapes(rhs **Shapes) com.Error {
	addr := (*this.LpVtbl)[41]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) Dummy29_()  {
	addr := (*this.LpVtbl)[42]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) Arcs(index interface{}, lcid int32, rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[43]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) Dummy31_()  {
	addr := (*this.LpVtbl)[44]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) Dummy32_()  {
	addr := (*this.LpVtbl)[45]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) Buttons(index interface{}, lcid int32, rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[46]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) Dummy34_()  {
	addr := (*this.LpVtbl)[47]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) GetEnableCalculation(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[48]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) SetEnableCalculation(rhs bool) com.Error {
	addr := (*this.LpVtbl)[49]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IDialogSheet) Dummy36_()  {
	addr := (*this.LpVtbl)[50]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) ChartObjects(index interface{}, lcid int32, rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[51]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) CheckBoxes(index interface{}, lcid int32, rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[52]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) CheckSpelling(customDictionary interface{}, ignoreUppercase interface{}, alwaysSuggest interface{}, spellLang interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[53]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&customDictionary)), (uintptr)(unsafe.Pointer(&ignoreUppercase)), (uintptr)(unsafe.Pointer(&alwaysSuggest)), (uintptr)(unsafe.Pointer(&spellLang)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IDialogSheet) Dummy40_()  {
	addr := (*this.LpVtbl)[54]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) Dummy41_()  {
	addr := (*this.LpVtbl)[55]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) Dummy42_()  {
	addr := (*this.LpVtbl)[56]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) Dummy43_()  {
	addr := (*this.LpVtbl)[57]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) Dummy44_()  {
	addr := (*this.LpVtbl)[58]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) Dummy45_()  {
	addr := (*this.LpVtbl)[59]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) GetDisplayAutomaticPageBreaks(lcid int32, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[60]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) SetDisplayAutomaticPageBreaks(lcid int32, rhs bool) com.Error {
	addr := (*this.LpVtbl)[61]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IDialogSheet) Drawings(index interface{}, lcid int32, rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[62]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) DrawingObjects(index interface{}, lcid int32, rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[63]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) DropDowns(index interface{}, lcid int32, rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[64]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) GetEnableAutoFilter(lcid int32, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[65]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) SetEnableAutoFilter(lcid int32, rhs bool) com.Error {
	addr := (*this.LpVtbl)[66]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IDialogSheet) GetEnableSelection(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[67]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) SetEnableSelection(rhs int32) com.Error {
	addr := (*this.LpVtbl)[68]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IDialogSheet) GetEnableOutlining(lcid int32, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[69]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) SetEnableOutlining(lcid int32, rhs bool) com.Error {
	addr := (*this.LpVtbl)[70]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IDialogSheet) GetEnablePivotTable(lcid int32, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[71]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) SetEnablePivotTable(lcid int32, rhs bool) com.Error {
	addr := (*this.LpVtbl)[72]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IDialogSheet) Evaluate(name interface{}, lcid int32, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[73]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&name)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) Evaluate_(name interface{}, lcid int32, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[74]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&name)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) Dummy56_()  {
	addr := (*this.LpVtbl)[75]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) ResetAllPageBreaks() com.Error {
	addr := (*this.LpVtbl)[76]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IDialogSheet) GroupBoxes(index interface{}, lcid int32, rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[77]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) GroupObjects(index interface{}, lcid int32, rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[78]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) Labels(index interface{}, lcid int32, rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[79]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) Lines(index interface{}, lcid int32, rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[80]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) ListBoxes(index interface{}, lcid int32, rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[81]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) GetNames(rhs **Names) com.Error {
	addr := (*this.LpVtbl)[82]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) OLEObjects(index interface{}, lcid int32, rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[83]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) Dummy65_()  {
	addr := (*this.LpVtbl)[84]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) Dummy66_()  {
	addr := (*this.LpVtbl)[85]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) Dummy67_()  {
	addr := (*this.LpVtbl)[86]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) OptionButtons(index interface{}, lcid int32, rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[87]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) Dummy69_()  {
	addr := (*this.LpVtbl)[88]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) Ovals(index interface{}, lcid int32, rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[89]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) Paste(destination interface{}, link interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[90]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&destination)), (uintptr)(unsafe.Pointer(&link)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IDialogSheet) PasteSpecial_(format interface{}, link interface{}, displayAsIcon interface{}, iconFileName interface{}, iconIndex interface{}, iconLabel interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[91]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&format)), (uintptr)(unsafe.Pointer(&link)), (uintptr)(unsafe.Pointer(&displayAsIcon)), (uintptr)(unsafe.Pointer(&iconFileName)), (uintptr)(unsafe.Pointer(&iconIndex)), (uintptr)(unsafe.Pointer(&iconLabel)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IDialogSheet) Pictures(index interface{}, lcid int32, rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[92]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) Dummy74_()  {
	addr := (*this.LpVtbl)[93]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) Dummy75_()  {
	addr := (*this.LpVtbl)[94]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) Dummy76_()  {
	addr := (*this.LpVtbl)[95]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) Rectangles(index interface{}, lcid int32, rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[96]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) Dummy78_()  {
	addr := (*this.LpVtbl)[97]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) Dummy79_()  {
	addr := (*this.LpVtbl)[98]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) GetScrollArea(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[99]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) SetScrollArea(rhs string) com.Error {
	addr := (*this.LpVtbl)[100]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) ScrollBars(index interface{}, lcid int32, rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[101]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) Dummy82_()  {
	addr := (*this.LpVtbl)[102]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) Dummy83_()  {
	addr := (*this.LpVtbl)[103]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) Spinners(index interface{}, lcid int32, rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[104]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) Dummy85_()  {
	addr := (*this.LpVtbl)[105]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) Dummy86_()  {
	addr := (*this.LpVtbl)[106]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) TextBoxes(index interface{}, lcid int32, rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[107]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) Dummy88_()  {
	addr := (*this.LpVtbl)[108]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) Dummy89_()  {
	addr := (*this.LpVtbl)[109]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) Dummy90_()  {
	addr := (*this.LpVtbl)[110]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) GetHPageBreaks(rhs **HPageBreaks) com.Error {
	addr := (*this.LpVtbl)[111]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) GetVPageBreaks(rhs **VPageBreaks) com.Error {
	addr := (*this.LpVtbl)[112]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) GetQueryTables(rhs **QueryTables) com.Error {
	addr := (*this.LpVtbl)[113]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) GetDisplayPageBreaks(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[114]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) SetDisplayPageBreaks(rhs bool) com.Error {
	addr := (*this.LpVtbl)[115]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IDialogSheet) GetComments(rhs **Comments) com.Error {
	addr := (*this.LpVtbl)[116]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) GetHyperlinks(rhs **Hyperlinks) com.Error {
	addr := (*this.LpVtbl)[117]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) ClearCircles() com.Error {
	addr := (*this.LpVtbl)[118]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IDialogSheet) CircleInvalid() com.Error {
	addr := (*this.LpVtbl)[119]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IDialogSheet) GetDisplayRightToLeft_(lcid int32, rhs *int32) com.Error {
	addr := (*this.LpVtbl)[120]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) SetDisplayRightToLeft_(lcid int32, rhs int32) com.Error {
	addr := (*this.LpVtbl)[121]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(rhs))
	return com.Error(ret)
}

func (this *IDialogSheet) GetAutoFilter(rhs **AutoFilter) com.Error {
	addr := (*this.LpVtbl)[122]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) GetDisplayRightToLeft(lcid int32, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[123]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) SetDisplayRightToLeft(lcid int32, rhs bool) com.Error {
	addr := (*this.LpVtbl)[124]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(lcid), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IDialogSheet) GetScripts(rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[125]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) PrintOut_(from interface{}, to interface{}, copies interface{}, preview interface{}, activePrinter interface{}, printToFile interface{}, collate interface{}, prToFileName interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[126]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&from)), (uintptr)(unsafe.Pointer(&to)), (uintptr)(unsafe.Pointer(&copies)), (uintptr)(unsafe.Pointer(&preview)), (uintptr)(unsafe.Pointer(&activePrinter)), (uintptr)(unsafe.Pointer(&printToFile)), (uintptr)(unsafe.Pointer(&collate)), (uintptr)(unsafe.Pointer(&prToFileName)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IDialogSheet) CheckSpelling_(customDictionary interface{}, ignoreUppercase interface{}, alwaysSuggest interface{}, spellLang interface{}, ignoreFinalYaa interface{}, spellScript interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[127]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&customDictionary)), (uintptr)(unsafe.Pointer(&ignoreUppercase)), (uintptr)(unsafe.Pointer(&alwaysSuggest)), (uintptr)(unsafe.Pointer(&spellLang)), (uintptr)(unsafe.Pointer(&ignoreFinalYaa)), (uintptr)(unsafe.Pointer(&spellScript)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IDialogSheet) GetTab(rhs **Tab) com.Error {
	addr := (*this.LpVtbl)[128]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) GetMailEnvelope(rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[129]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) SaveAs(filename string, fileFormat interface{}, password interface{}, writeResPassword interface{}, readOnlyRecommended interface{}, createBackup interface{}, addToMru interface{}, textCodepage interface{}, textVisualLayout interface{}, local interface{}) com.Error {
	addr := (*this.LpVtbl)[130]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(filename)), (uintptr)(unsafe.Pointer(&fileFormat)), (uintptr)(unsafe.Pointer(&password)), (uintptr)(unsafe.Pointer(&writeResPassword)), (uintptr)(unsafe.Pointer(&readOnlyRecommended)), (uintptr)(unsafe.Pointer(&createBackup)), (uintptr)(unsafe.Pointer(&addToMru)), (uintptr)(unsafe.Pointer(&textCodepage)), (uintptr)(unsafe.Pointer(&textVisualLayout)), (uintptr)(unsafe.Pointer(&local)))
	return com.Error(ret)
}

func (this *IDialogSheet) GetCustomProperties(rhs **CustomProperties) com.Error {
	addr := (*this.LpVtbl)[131]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) GetSmartTags(rhs **SmartTags) com.Error {
	addr := (*this.LpVtbl)[132]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) GetProtection(rhs **Protection) com.Error {
	addr := (*this.LpVtbl)[133]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) PasteSpecial(format interface{}, link interface{}, displayAsIcon interface{}, iconFileName interface{}, iconIndex interface{}, iconLabel interface{}, noHTMLFormatting interface{}, lcid int32) com.Error {
	addr := (*this.LpVtbl)[134]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&format)), (uintptr)(unsafe.Pointer(&link)), (uintptr)(unsafe.Pointer(&displayAsIcon)), (uintptr)(unsafe.Pointer(&iconFileName)), (uintptr)(unsafe.Pointer(&iconIndex)), (uintptr)(unsafe.Pointer(&iconLabel)), (uintptr)(unsafe.Pointer(&noHTMLFormatting)), uintptr(lcid))
	return com.Error(ret)
}

func (this *IDialogSheet) Protect(password interface{}, drawingObjects interface{}, contents interface{}, scenarios interface{}, userInterfaceOnly interface{}, allowFormattingCells interface{}, allowFormattingColumns interface{}, allowFormattingRows interface{}, allowInsertingColumns interface{}, allowInsertingRows interface{}, allowInsertingHyperlinks interface{}, allowDeletingColumns interface{}, allowDeletingRows interface{}, allowSorting interface{}, allowFiltering interface{}, allowUsingPivotTables interface{}) com.Error {
	addr := (*this.LpVtbl)[135]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&password)), (uintptr)(unsafe.Pointer(&drawingObjects)), (uintptr)(unsafe.Pointer(&contents)), (uintptr)(unsafe.Pointer(&scenarios)), (uintptr)(unsafe.Pointer(&userInterfaceOnly)), (uintptr)(unsafe.Pointer(&allowFormattingCells)), (uintptr)(unsafe.Pointer(&allowFormattingColumns)), (uintptr)(unsafe.Pointer(&allowFormattingRows)), (uintptr)(unsafe.Pointer(&allowInsertingColumns)), (uintptr)(unsafe.Pointer(&allowInsertingRows)), (uintptr)(unsafe.Pointer(&allowInsertingHyperlinks)), (uintptr)(unsafe.Pointer(&allowDeletingColumns)), (uintptr)(unsafe.Pointer(&allowDeletingRows)), (uintptr)(unsafe.Pointer(&allowSorting)), (uintptr)(unsafe.Pointer(&allowFiltering)), (uintptr)(unsafe.Pointer(&allowUsingPivotTables)))
	return com.Error(ret)
}

func (this *IDialogSheet) Dummy113_()  {
	addr := (*this.LpVtbl)[136]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) Dummy114_()  {
	addr := (*this.LpVtbl)[137]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) Dummy115_()  {
	addr := (*this.LpVtbl)[138]
	_, _, _ = syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
}

func (this *IDialogSheet) PrintOut(from interface{}, to interface{}, copies interface{}, preview interface{}, activePrinter interface{}, printToFile interface{}, collate interface{}, prToFileName interface{}) com.Error {
	addr := (*this.LpVtbl)[139]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&from)), (uintptr)(unsafe.Pointer(&to)), (uintptr)(unsafe.Pointer(&copies)), (uintptr)(unsafe.Pointer(&preview)), (uintptr)(unsafe.Pointer(&activePrinter)), (uintptr)(unsafe.Pointer(&printToFile)), (uintptr)(unsafe.Pointer(&collate)), (uintptr)(unsafe.Pointer(&prToFileName)))
	return com.Error(ret)
}

func (this *IDialogSheet) GetEnableFormatConditionsCalculation(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[140]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) SetEnableFormatConditionsCalculation(rhs bool) com.Error {
	addr := (*this.LpVtbl)[141]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IDialogSheet) GetSort(rhs **Sort) com.Error {
	addr := (*this.LpVtbl)[142]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) ExportAsFixedFormat(type_ int32, filename interface{}, quality interface{}, includeDocProperties interface{}, ignorePrintAreas interface{}, from interface{}, to interface{}, openAfterPublish interface{}, fixedFormatExtClassPtr interface{}) com.Error {
	addr := (*this.LpVtbl)[143]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(type_), (uintptr)(unsafe.Pointer(&filename)), (uintptr)(unsafe.Pointer(&quality)), (uintptr)(unsafe.Pointer(&includeDocProperties)), (uintptr)(unsafe.Pointer(&ignorePrintAreas)), (uintptr)(unsafe.Pointer(&from)), (uintptr)(unsafe.Pointer(&to)), (uintptr)(unsafe.Pointer(&openAfterPublish)), (uintptr)(unsafe.Pointer(&fixedFormatExtClassPtr)))
	return com.Error(ret)
}

func (this *IDialogSheet) GetPrintedCommentPages(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[144]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) GetDefaultButton(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[145]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) SetDefaultButton(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[146]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) GetDialogFrame(rhs **DialogFrame) com.Error {
	addr := (*this.LpVtbl)[147]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) EditBoxes(index interface{}, rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[148]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IDialogSheet) GetFocus(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[149]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) SetFocus(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[150]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) Hide(cancel interface{}, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[151]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&cancel)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IDialogSheet) Show(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[152]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

