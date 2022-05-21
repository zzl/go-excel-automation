package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020893-0001-0000-C000-000000000046
var IID_IWindow = syscall.GUID{0x00020893, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IWindow struct {
	win32.IDispatch
}

func NewIWindow(pUnk *win32.IUnknown, addRef bool, scoped bool) *IWindow {
	p := (*IWindow)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IWindow) IID() *syscall.GUID {
	return &IID_IWindow
}

func (this *IWindow) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IWindow) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) GetParent(rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IWindow) Activate(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) ActivateNext(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) ActivatePrevious(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) GetActiveCell(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IWindow) GetActiveChart(rhs **Chart) com.Error {
	addr := (*this.LpVtbl)[14]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IWindow) GetActivePane(rhs **Pane) com.Error {
	addr := (*this.LpVtbl)[15]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IWindow) GetActiveSheet(rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[16]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IWindow) GetCaption(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[17]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetCaption(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[18]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IWindow) Close(saveChanges interface{}, filename interface{}, routeWorkbook interface{}, rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[19]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&saveChanges)), (uintptr)(unsafe.Pointer(&filename)), (uintptr)(unsafe.Pointer(&routeWorkbook)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) GetDisplayFormulas(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[20]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetDisplayFormulas(rhs bool) com.Error {
	addr := (*this.LpVtbl)[21]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IWindow) GetDisplayGridlines(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[22]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetDisplayGridlines(rhs bool) com.Error {
	addr := (*this.LpVtbl)[23]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IWindow) GetDisplayHeadings(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[24]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetDisplayHeadings(rhs bool) com.Error {
	addr := (*this.LpVtbl)[25]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IWindow) GetDisplayHorizontalScrollBar(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[26]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetDisplayHorizontalScrollBar(rhs bool) com.Error {
	addr := (*this.LpVtbl)[27]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IWindow) GetDisplayOutline(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[28]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetDisplayOutline(rhs bool) com.Error {
	addr := (*this.LpVtbl)[29]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IWindow) GetDisplayRightToLeft_(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[30]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetDisplayRightToLeft_(rhs bool) com.Error {
	addr := (*this.LpVtbl)[31]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IWindow) GetDisplayVerticalScrollBar(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[32]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetDisplayVerticalScrollBar(rhs bool) com.Error {
	addr := (*this.LpVtbl)[33]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IWindow) GetDisplayWorkbookTabs(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[34]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetDisplayWorkbookTabs(rhs bool) com.Error {
	addr := (*this.LpVtbl)[35]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IWindow) GetDisplayZeros(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[36]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetDisplayZeros(rhs bool) com.Error {
	addr := (*this.LpVtbl)[37]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IWindow) GetEnableResize(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[38]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetEnableResize(rhs bool) com.Error {
	addr := (*this.LpVtbl)[39]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IWindow) GetFreezePanes(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[40]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetFreezePanes(rhs bool) com.Error {
	addr := (*this.LpVtbl)[41]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IWindow) GetGridlineColor(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[42]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetGridlineColor(rhs int32) com.Error {
	addr := (*this.LpVtbl)[43]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IWindow) GetGridlineColorIndex(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[44]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetGridlineColorIndex(rhs int32) com.Error {
	addr := (*this.LpVtbl)[45]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IWindow) GetHeight(rhs *float64) com.Error {
	addr := (*this.LpVtbl)[46]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetHeight(rhs float64) com.Error {
	addr := (*this.LpVtbl)[47]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IWindow) GetIndex(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[48]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) LargeScroll(down interface{}, up interface{}, toRight interface{}, toLeft interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[49]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&down)), (uintptr)(unsafe.Pointer(&up)), (uintptr)(unsafe.Pointer(&toRight)), (uintptr)(unsafe.Pointer(&toLeft)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) GetLeft(rhs *float64) com.Error {
	addr := (*this.LpVtbl)[50]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetLeft(rhs float64) com.Error {
	addr := (*this.LpVtbl)[51]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IWindow) NewWindow(rhs **Window) com.Error {
	addr := (*this.LpVtbl)[52]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IWindow) GetOnWindow(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[53]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetOnWindow(rhs string) com.Error {
	addr := (*this.LpVtbl)[54]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) GetPanes(rhs **Panes) com.Error {
	addr := (*this.LpVtbl)[55]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IWindow) PrintOut_(from interface{}, to interface{}, copies interface{}, preview interface{}, activePrinter interface{}, printToFile interface{}, collate interface{}, prToFileName interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[56]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&from)), (uintptr)(unsafe.Pointer(&to)), (uintptr)(unsafe.Pointer(&copies)), (uintptr)(unsafe.Pointer(&preview)), (uintptr)(unsafe.Pointer(&activePrinter)), (uintptr)(unsafe.Pointer(&printToFile)), (uintptr)(unsafe.Pointer(&collate)), (uintptr)(unsafe.Pointer(&prToFileName)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) PrintPreview(enableChanges interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[57]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&enableChanges)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) GetRangeSelection(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[58]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IWindow) GetScrollColumn(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[59]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetScrollColumn(rhs int32) com.Error {
	addr := (*this.LpVtbl)[60]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IWindow) GetScrollRow(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[61]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetScrollRow(rhs int32) com.Error {
	addr := (*this.LpVtbl)[62]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IWindow) ScrollWorkbookTabs(sheets interface{}, position interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[63]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&sheets)), (uintptr)(unsafe.Pointer(&position)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) GetSelectedSheets(rhs **Sheets) com.Error {
	addr := (*this.LpVtbl)[64]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IWindow) GetSelection(rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[65]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IWindow) SmallScroll(down interface{}, up interface{}, toRight interface{}, toLeft interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[66]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&down)), (uintptr)(unsafe.Pointer(&up)), (uintptr)(unsafe.Pointer(&toRight)), (uintptr)(unsafe.Pointer(&toLeft)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) GetSplit(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[67]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetSplit(rhs bool) com.Error {
	addr := (*this.LpVtbl)[68]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IWindow) GetSplitColumn(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[69]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetSplitColumn(rhs int32) com.Error {
	addr := (*this.LpVtbl)[70]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IWindow) GetSplitHorizontal(rhs *float64) com.Error {
	addr := (*this.LpVtbl)[71]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetSplitHorizontal(rhs float64) com.Error {
	addr := (*this.LpVtbl)[72]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IWindow) GetSplitRow(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[73]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetSplitRow(rhs int32) com.Error {
	addr := (*this.LpVtbl)[74]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IWindow) GetSplitVertical(rhs *float64) com.Error {
	addr := (*this.LpVtbl)[75]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetSplitVertical(rhs float64) com.Error {
	addr := (*this.LpVtbl)[76]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IWindow) GetTabRatio(rhs *float64) com.Error {
	addr := (*this.LpVtbl)[77]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetTabRatio(rhs float64) com.Error {
	addr := (*this.LpVtbl)[78]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IWindow) GetTop(rhs *float64) com.Error {
	addr := (*this.LpVtbl)[79]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetTop(rhs float64) com.Error {
	addr := (*this.LpVtbl)[80]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IWindow) GetType(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[81]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) GetUsableHeight(rhs *float64) com.Error {
	addr := (*this.LpVtbl)[82]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) GetUsableWidth(rhs *float64) com.Error {
	addr := (*this.LpVtbl)[83]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) GetVisible(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[84]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetVisible(rhs bool) com.Error {
	addr := (*this.LpVtbl)[85]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IWindow) GetVisibleRange(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[86]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IWindow) GetWidth(rhs *float64) com.Error {
	addr := (*this.LpVtbl)[87]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetWidth(rhs float64) com.Error {
	addr := (*this.LpVtbl)[88]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IWindow) GetWindowNumber(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[89]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) GetWindowState(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[90]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetWindowState(rhs int32) com.Error {
	addr := (*this.LpVtbl)[91]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IWindow) GetZoom(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[92]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetZoom(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[93]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IWindow) GetView(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[94]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetView(rhs int32) com.Error {
	addr := (*this.LpVtbl)[95]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IWindow) GetDisplayRightToLeft(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[96]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetDisplayRightToLeft(rhs bool) com.Error {
	addr := (*this.LpVtbl)[97]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IWindow) PointsToScreenPixelsX(points int32, rhs *int32) com.Error {
	addr := (*this.LpVtbl)[98]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(points), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) PointsToScreenPixelsY(points int32, rhs *int32) com.Error {
	addr := (*this.LpVtbl)[99]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(points), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) RangeFromPoint(x int32, y int32, rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[100]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(x), uintptr(y), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IWindow) ScrollIntoView(left int32, top int32, width int32, height int32, start interface{}) com.Error {
	addr := (*this.LpVtbl)[101]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(left), uintptr(top), uintptr(width), uintptr(height), (uintptr)(unsafe.Pointer(&start)))
	return com.Error(ret)
}

func (this *IWindow) GetSheetViews(rhs **SheetViews) com.Error {
	addr := (*this.LpVtbl)[102]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IWindow) GetActiveSheetView(rhs **com.UnknownClass) com.Error {
	addr := (*this.LpVtbl)[103]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	if com.CurrentScope != nil {
		com.CurrentScope.Add(unsafe.Pointer(&(*rhs).IUnknown))
	}
	return com.Error(ret)
}

func (this *IWindow) PrintOut(from interface{}, to interface{}, copies interface{}, preview interface{}, activePrinter interface{}, printToFile interface{}, collate interface{}, prToFileName interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[104]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&from)), (uintptr)(unsafe.Pointer(&to)), (uintptr)(unsafe.Pointer(&copies)), (uintptr)(unsafe.Pointer(&preview)), (uintptr)(unsafe.Pointer(&activePrinter)), (uintptr)(unsafe.Pointer(&printToFile)), (uintptr)(unsafe.Pointer(&collate)), (uintptr)(unsafe.Pointer(&prToFileName)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) GetDisplayRuler(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[105]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetDisplayRuler(rhs bool) com.Error {
	addr := (*this.LpVtbl)[106]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IWindow) GetAutoFilterDateGrouping(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[107]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetAutoFilterDateGrouping(rhs bool) com.Error {
	addr := (*this.LpVtbl)[108]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IWindow) GetDisplayWhitespace(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[109]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IWindow) SetDisplayWhitespace(rhs bool) com.Error {
	addr := (*this.LpVtbl)[110]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

