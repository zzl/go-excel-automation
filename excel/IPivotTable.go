package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020872-0001-0000-C000-000000000046
var IID_IPivotTable = syscall.GUID{0x00020872, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IPivotTable struct {
	win32.IDispatch
}

func NewIPivotTable(pUnk *win32.IUnknown, addRef bool, scoped bool) *IPivotTable {
	 if pUnk == nil {
		return nil;
	}
	p := (*IPivotTable)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IPivotTable) IID() *syscall.GUID {
	return &IID_IPivotTable
}

func (this *IPivotTable) GetApplication(rhs **Application) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetCreator(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetParent(rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) AddFields(rowFields interface{}, columnFields interface{}, pageFields interface{}, addToTable interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rowFields)), (uintptr)(unsafe.Pointer(&columnFields)), (uintptr)(unsafe.Pointer(&pageFields)), (uintptr)(unsafe.Pointer(&addToTable)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetColumnFields(index interface{}, rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetColumnGrand(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetColumnGrand(rhs bool) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetColumnRange(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[14]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) ShowPages(pageField interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[15]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&pageField)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetDataBodyRange(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[16]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetDataFields(index interface{}, rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[17]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetDataLabelRange(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[18]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetDefault_(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[19]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetDefault_(rhs string) com.Error {
	addr := (*this.LpVtbl)[20]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetHasAutoFormat(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[21]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetHasAutoFormat(rhs bool) com.Error {
	addr := (*this.LpVtbl)[22]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetHiddenFields(index interface{}, rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[23]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetInnerDetail(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[24]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetInnerDetail(rhs string) com.Error {
	addr := (*this.LpVtbl)[25]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetName(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[26]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetName(rhs string) com.Error {
	addr := (*this.LpVtbl)[27]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetPageFields(index interface{}, rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[28]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetPageRange(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[29]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetPageRangeCells(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[30]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) PivotFields(index interface{}, rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[31]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetRefreshDate(rhs *ole.Date) com.Error {
	addr := (*this.LpVtbl)[32]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetRefreshName(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[33]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) RefreshTable(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[34]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetRowFields(index interface{}, rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[35]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetRowGrand(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[36]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetRowGrand(rhs bool) com.Error {
	addr := (*this.LpVtbl)[37]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetRowRange(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[38]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetSaveData(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[39]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetSaveData(rhs bool) com.Error {
	addr := (*this.LpVtbl)[40]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetSourceData(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[41]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetSourceData(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[42]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetTableRange1(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[43]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetTableRange2(rhs **Range) com.Error {
	addr := (*this.LpVtbl)[44]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetValue(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[45]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetValue(rhs string) com.Error {
	addr := (*this.LpVtbl)[46]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetVisibleFields(index interface{}, rhs **win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[47]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&index)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetCacheIndex(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[48]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetCacheIndex(rhs int32) com.Error {
	addr := (*this.LpVtbl)[49]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IPivotTable) CalculatedFields(rhs **CalculatedFields) com.Error {
	addr := (*this.LpVtbl)[50]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetDisplayErrorString(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[51]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetDisplayErrorString(rhs bool) com.Error {
	addr := (*this.LpVtbl)[52]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetDisplayNullString(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[53]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetDisplayNullString(rhs bool) com.Error {
	addr := (*this.LpVtbl)[54]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetEnableDrilldown(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[55]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetEnableDrilldown(rhs bool) com.Error {
	addr := (*this.LpVtbl)[56]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetEnableFieldDialog(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[57]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetEnableFieldDialog(rhs bool) com.Error {
	addr := (*this.LpVtbl)[58]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetEnableWizard(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[59]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetEnableWizard(rhs bool) com.Error {
	addr := (*this.LpVtbl)[60]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetErrorString(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[61]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetErrorString(rhs string) com.Error {
	addr := (*this.LpVtbl)[62]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetData(name string, rhs *float64) com.Error {
	addr := (*this.LpVtbl)[63]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(name)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) ListFormulas() com.Error {
	addr := (*this.LpVtbl)[64]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IPivotTable) GetManualUpdate(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[65]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetManualUpdate(rhs bool) com.Error {
	addr := (*this.LpVtbl)[66]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetMergeLabels(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[67]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetMergeLabels(rhs bool) com.Error {
	addr := (*this.LpVtbl)[68]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetNullString(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[69]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetNullString(rhs string) com.Error {
	addr := (*this.LpVtbl)[70]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) PivotCache(rhs **PivotCache) com.Error {
	addr := (*this.LpVtbl)[71]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetPivotFormulas(rhs **PivotFormulas) com.Error {
	addr := (*this.LpVtbl)[72]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) PivotTableWizard(sourceType interface{}, sourceData interface{}, tableDestination interface{}, tableName interface{}, rowGrand interface{}, columnGrand interface{}, saveData interface{}, hasAutoFormat interface{}, autoPage interface{}, reserved interface{}, backgroundQuery interface{}, optimizeCache interface{}, pageFieldOrder interface{}, pageFieldWrapCount interface{}, readData interface{}, connection interface{}) com.Error {
	addr := (*this.LpVtbl)[73]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&sourceType)), (uintptr)(unsafe.Pointer(&sourceData)), (uintptr)(unsafe.Pointer(&tableDestination)), (uintptr)(unsafe.Pointer(&tableName)), (uintptr)(unsafe.Pointer(&rowGrand)), (uintptr)(unsafe.Pointer(&columnGrand)), (uintptr)(unsafe.Pointer(&saveData)), (uintptr)(unsafe.Pointer(&hasAutoFormat)), (uintptr)(unsafe.Pointer(&autoPage)), (uintptr)(unsafe.Pointer(&reserved)), (uintptr)(unsafe.Pointer(&backgroundQuery)), (uintptr)(unsafe.Pointer(&optimizeCache)), (uintptr)(unsafe.Pointer(&pageFieldOrder)), (uintptr)(unsafe.Pointer(&pageFieldWrapCount)), (uintptr)(unsafe.Pointer(&readData)), (uintptr)(unsafe.Pointer(&connection)))
	return com.Error(ret)
}

func (this *IPivotTable) GetSubtotalHiddenPageItems(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[74]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetSubtotalHiddenPageItems(rhs bool) com.Error {
	addr := (*this.LpVtbl)[75]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetPageFieldOrder(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[76]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetPageFieldOrder(rhs int32) com.Error {
	addr := (*this.LpVtbl)[77]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IPivotTable) GetPageFieldStyle(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[78]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetPageFieldStyle(rhs string) com.Error {
	addr := (*this.LpVtbl)[79]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetPageFieldWrapCount(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[80]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetPageFieldWrapCount(rhs int32) com.Error {
	addr := (*this.LpVtbl)[81]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IPivotTable) GetPreserveFormatting(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[82]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetPreserveFormatting(rhs bool) com.Error {
	addr := (*this.LpVtbl)[83]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) PivotSelect_(name string, mode int32) com.Error {
	addr := (*this.LpVtbl)[84]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(name)), uintptr(mode))
	return com.Error(ret)
}

func (this *IPivotTable) GetPivotSelection(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[85]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetPivotSelection(rhs string) com.Error {
	addr := (*this.LpVtbl)[86]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetSelectionMode(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[87]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetSelectionMode(rhs int32) com.Error {
	addr := (*this.LpVtbl)[88]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IPivotTable) GetTableStyle(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[89]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetTableStyle(rhs string) com.Error {
	addr := (*this.LpVtbl)[90]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetTag(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[91]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetTag(rhs string) com.Error {
	addr := (*this.LpVtbl)[92]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) Update() com.Error {
	addr := (*this.LpVtbl)[93]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IPivotTable) GetVacatedStyle(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[94]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetVacatedStyle(rhs string) com.Error {
	addr := (*this.LpVtbl)[95]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) Format(format int32) com.Error {
	addr := (*this.LpVtbl)[96]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(format))
	return com.Error(ret)
}

func (this *IPivotTable) GetPrintTitles(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[97]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetPrintTitles(rhs bool) com.Error {
	addr := (*this.LpVtbl)[98]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetCubeFields(rhs **CubeFields) com.Error {
	addr := (*this.LpVtbl)[99]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetGrandTotalName(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[100]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetGrandTotalName(rhs string) com.Error {
	addr := (*this.LpVtbl)[101]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetSmallGrid(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[102]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetSmallGrid(rhs bool) com.Error {
	addr := (*this.LpVtbl)[103]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetRepeatItemsOnEachPrintedPage(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[104]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetRepeatItemsOnEachPrintedPage(rhs bool) com.Error {
	addr := (*this.LpVtbl)[105]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetTotalsAnnotation(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[106]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetTotalsAnnotation(rhs bool) com.Error {
	addr := (*this.LpVtbl)[107]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) PivotSelect(name string, mode int32, useStandardName interface{}) com.Error {
	addr := (*this.LpVtbl)[108]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(name)), uintptr(mode), (uintptr)(unsafe.Pointer(&useStandardName)))
	return com.Error(ret)
}

func (this *IPivotTable) GetPivotSelectionStandard(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[109]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetPivotSelectionStandard(rhs string) com.Error {
	addr := (*this.LpVtbl)[110]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetPivotData(dataField interface{}, field1 interface{}, item1 interface{}, field2 interface{}, item2 interface{}, field3 interface{}, item3 interface{}, field4 interface{}, item4 interface{}, field5 interface{}, item5 interface{}, field6 interface{}, item6 interface{}, field7 interface{}, item7 interface{}, field8 interface{}, item8 interface{}, field9 interface{}, item9 interface{}, field10 interface{}, item10 interface{}, field11 interface{}, item11 interface{}, field12 interface{}, item12 interface{}, field13 interface{}, item13 interface{}, field14 interface{}, item14 interface{}, rhs **Range) com.Error {
	addr := (*this.LpVtbl)[111]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&dataField)), (uintptr)(unsafe.Pointer(&field1)), (uintptr)(unsafe.Pointer(&item1)), (uintptr)(unsafe.Pointer(&field2)), (uintptr)(unsafe.Pointer(&item2)), (uintptr)(unsafe.Pointer(&field3)), (uintptr)(unsafe.Pointer(&item3)), (uintptr)(unsafe.Pointer(&field4)), (uintptr)(unsafe.Pointer(&item4)), (uintptr)(unsafe.Pointer(&field5)), (uintptr)(unsafe.Pointer(&item5)), (uintptr)(unsafe.Pointer(&field6)), (uintptr)(unsafe.Pointer(&item6)), (uintptr)(unsafe.Pointer(&field7)), (uintptr)(unsafe.Pointer(&item7)), (uintptr)(unsafe.Pointer(&field8)), (uintptr)(unsafe.Pointer(&item8)), (uintptr)(unsafe.Pointer(&field9)), (uintptr)(unsafe.Pointer(&item9)), (uintptr)(unsafe.Pointer(&field10)), (uintptr)(unsafe.Pointer(&item10)), (uintptr)(unsafe.Pointer(&field11)), (uintptr)(unsafe.Pointer(&item11)), (uintptr)(unsafe.Pointer(&field12)), (uintptr)(unsafe.Pointer(&item12)), (uintptr)(unsafe.Pointer(&field13)), (uintptr)(unsafe.Pointer(&item13)), (uintptr)(unsafe.Pointer(&field14)), (uintptr)(unsafe.Pointer(&item14)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetDataPivotField(rhs **PivotField) com.Error {
	addr := (*this.LpVtbl)[112]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetEnableDataValueEditing(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[113]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetEnableDataValueEditing(rhs bool) com.Error {
	addr := (*this.LpVtbl)[114]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) AddDataField(field *win32.IUnknown, caption interface{}, function interface{}, rhs **PivotField) com.Error {
	addr := (*this.LpVtbl)[115]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(field)), (uintptr)(unsafe.Pointer(&caption)), (uintptr)(unsafe.Pointer(&function)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetMDX(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[116]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetViewCalculatedMembers(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[117]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetViewCalculatedMembers(rhs bool) com.Error {
	addr := (*this.LpVtbl)[118]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetCalculatedMembers(rhs **CalculatedMembers) com.Error {
	addr := (*this.LpVtbl)[119]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetDisplayImmediateItems(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[120]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetDisplayImmediateItems(rhs bool) com.Error {
	addr := (*this.LpVtbl)[121]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) Dummy15(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, arg9 interface{}, arg10 interface{}, arg11 interface{}, arg12 interface{}, arg13 interface{}, arg14 interface{}, arg15 interface{}, arg16 interface{}, arg17 interface{}, arg18 interface{}, arg19 interface{}, arg20 interface{}, arg21 interface{}, arg22 interface{}, arg23 interface{}, arg24 interface{}, arg25 interface{}, arg26 interface{}, arg27 interface{}, arg28 interface{}, arg29 interface{}, arg30 interface{}, rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[122]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&arg1)), (uintptr)(unsafe.Pointer(&arg2)), (uintptr)(unsafe.Pointer(&arg3)), (uintptr)(unsafe.Pointer(&arg4)), (uintptr)(unsafe.Pointer(&arg5)), (uintptr)(unsafe.Pointer(&arg6)), (uintptr)(unsafe.Pointer(&arg7)), (uintptr)(unsafe.Pointer(&arg8)), (uintptr)(unsafe.Pointer(&arg9)), (uintptr)(unsafe.Pointer(&arg10)), (uintptr)(unsafe.Pointer(&arg11)), (uintptr)(unsafe.Pointer(&arg12)), (uintptr)(unsafe.Pointer(&arg13)), (uintptr)(unsafe.Pointer(&arg14)), (uintptr)(unsafe.Pointer(&arg15)), (uintptr)(unsafe.Pointer(&arg16)), (uintptr)(unsafe.Pointer(&arg17)), (uintptr)(unsafe.Pointer(&arg18)), (uintptr)(unsafe.Pointer(&arg19)), (uintptr)(unsafe.Pointer(&arg20)), (uintptr)(unsafe.Pointer(&arg21)), (uintptr)(unsafe.Pointer(&arg22)), (uintptr)(unsafe.Pointer(&arg23)), (uintptr)(unsafe.Pointer(&arg24)), (uintptr)(unsafe.Pointer(&arg25)), (uintptr)(unsafe.Pointer(&arg26)), (uintptr)(unsafe.Pointer(&arg27)), (uintptr)(unsafe.Pointer(&arg28)), (uintptr)(unsafe.Pointer(&arg29)), (uintptr)(unsafe.Pointer(&arg30)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetEnableFieldList(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[123]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetEnableFieldList(rhs bool) com.Error {
	addr := (*this.LpVtbl)[124]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetVisualTotals(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[125]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetVisualTotals(rhs bool) com.Error {
	addr := (*this.LpVtbl)[126]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetShowPageMultipleItemLabel(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[127]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetShowPageMultipleItemLabel(rhs bool) com.Error {
	addr := (*this.LpVtbl)[128]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetVersion(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[129]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) CreateCubeFile(file string, measures interface{}, levels interface{}, members interface{}, properties interface{}, rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[130]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(file)), (uintptr)(unsafe.Pointer(&measures)), (uintptr)(unsafe.Pointer(&levels)), (uintptr)(unsafe.Pointer(&members)), (uintptr)(unsafe.Pointer(&properties)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetDisplayEmptyRow(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[131]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetDisplayEmptyRow(rhs bool) com.Error {
	addr := (*this.LpVtbl)[132]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetDisplayEmptyColumn(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[133]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetDisplayEmptyColumn(rhs bool) com.Error {
	addr := (*this.LpVtbl)[134]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetShowCellBackgroundFromOLAP(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[135]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetShowCellBackgroundFromOLAP(rhs bool) com.Error {
	addr := (*this.LpVtbl)[136]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetPivotColumnAxis(rhs **PivotAxis) com.Error {
	addr := (*this.LpVtbl)[137]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetPivotRowAxis(rhs **PivotAxis) com.Error {
	addr := (*this.LpVtbl)[138]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetShowDrillIndicators(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[139]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetShowDrillIndicators(rhs bool) com.Error {
	addr := (*this.LpVtbl)[140]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetPrintDrillIndicators(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[141]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetPrintDrillIndicators(rhs bool) com.Error {
	addr := (*this.LpVtbl)[142]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetDisplayMemberPropertyTooltips(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[143]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetDisplayMemberPropertyTooltips(rhs bool) com.Error {
	addr := (*this.LpVtbl)[144]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetDisplayContextTooltips(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[145]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetDisplayContextTooltips(rhs bool) com.Error {
	addr := (*this.LpVtbl)[146]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) ClearTable() com.Error {
	addr := (*this.LpVtbl)[147]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IPivotTable) GetCompactRowIndent(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[148]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetCompactRowIndent(rhs int32) com.Error {
	addr := (*this.LpVtbl)[149]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IPivotTable) GetLayoutRowDefault(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[150]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetLayoutRowDefault(rhs int32) com.Error {
	addr := (*this.LpVtbl)[151]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IPivotTable) GetDisplayFieldCaptions(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[152]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetDisplayFieldCaptions(rhs bool) com.Error {
	addr := (*this.LpVtbl)[153]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) RowAxisLayout(rowLayout int32) com.Error {
	addr := (*this.LpVtbl)[154]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rowLayout))
	return com.Error(ret)
}

func (this *IPivotTable) SubtotalLocation(location int32) com.Error {
	addr := (*this.LpVtbl)[155]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(location))
	return com.Error(ret)
}

func (this *IPivotTable) GetActiveFilters(rhs **PivotFilters) com.Error {
	addr := (*this.LpVtbl)[156]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetInGridDropZones(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[157]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetInGridDropZones(rhs bool) com.Error {
	addr := (*this.LpVtbl)[158]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) ClearAllFilters() com.Error {
	addr := (*this.LpVtbl)[159]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IPivotTable) GetTableStyle2(rhs *ole.Variant) com.Error {
	addr := (*this.LpVtbl)[160]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetTableStyle2(rhs interface{}) com.Error {
	addr := (*this.LpVtbl)[161]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetShowTableStyleLastColumn(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[162]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetShowTableStyleLastColumn(rhs bool) com.Error {
	addr := (*this.LpVtbl)[163]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetShowTableStyleRowStripes(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[164]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetShowTableStyleRowStripes(rhs bool) com.Error {
	addr := (*this.LpVtbl)[165]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetShowTableStyleColumnStripes(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[166]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetShowTableStyleColumnStripes(rhs bool) com.Error {
	addr := (*this.LpVtbl)[167]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetShowTableStyleRowHeaders(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[168]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetShowTableStyleRowHeaders(rhs bool) com.Error {
	addr := (*this.LpVtbl)[169]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetShowTableStyleColumnHeaders(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[170]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetShowTableStyleColumnHeaders(rhs bool) com.Error {
	addr := (*this.LpVtbl)[171]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) ConvertToFormulas(convertFilters bool) com.Error {
	addr := (*this.LpVtbl)[172]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&convertFilters))))
	return com.Error(ret)
}

func (this *IPivotTable) GetAllowMultipleFilters(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[173]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetAllowMultipleFilters(rhs bool) com.Error {
	addr := (*this.LpVtbl)[174]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetCompactLayoutRowHeader(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[175]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetCompactLayoutRowHeader(rhs string) com.Error {
	addr := (*this.LpVtbl)[176]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetCompactLayoutColumnHeader(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[177]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetCompactLayoutColumnHeader(rhs string) com.Error {
	addr := (*this.LpVtbl)[178]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetFieldListSortAscending(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[179]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetFieldListSortAscending(rhs bool) com.Error {
	addr := (*this.LpVtbl)[180]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetSortUsingCustomLists(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[181]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetSortUsingCustomLists(rhs bool) com.Error {
	addr := (*this.LpVtbl)[182]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) ChangeConnection(conn *WorkbookConnection) com.Error {
	addr := (*this.LpVtbl)[183]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(conn)))
	return com.Error(ret)
}

func (this *IPivotTable) ChangePivotCache(pivotCache interface{}) com.Error {
	addr := (*this.LpVtbl)[184]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), (uintptr)(unsafe.Pointer(&pivotCache)))
	return com.Error(ret)
}

func (this *IPivotTable) GetLocation(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[185]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetLocation(rhs string) com.Error {
	addr := (*this.LpVtbl)[186]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetEnableWriteback(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[187]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetEnableWriteback(rhs bool) com.Error {
	addr := (*this.LpVtbl)[188]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetAllocation(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[189]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetAllocation(rhs int32) com.Error {
	addr := (*this.LpVtbl)[190]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IPivotTable) GetAllocationValue(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[191]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetAllocationValue(rhs int32) com.Error {
	addr := (*this.LpVtbl)[192]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IPivotTable) GetAllocationMethod(rhs *int32) com.Error {
	addr := (*this.LpVtbl)[193]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetAllocationMethod(rhs int32) com.Error {
	addr := (*this.LpVtbl)[194]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(rhs))
	return com.Error(ret)
}

func (this *IPivotTable) GetAllocationWeightExpression(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[195]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetAllocationWeightExpression(rhs string) com.Error {
	addr := (*this.LpVtbl)[196]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) AllocateChanges() com.Error {
	addr := (*this.LpVtbl)[197]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IPivotTable) CommitChanges() com.Error {
	addr := (*this.LpVtbl)[198]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IPivotTable) DiscardChanges() com.Error {
	addr := (*this.LpVtbl)[199]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IPivotTable) RefreshDataSourceValues() com.Error {
	addr := (*this.LpVtbl)[200]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IPivotTable) RepeatAllLabels(repeat int32) com.Error {
	addr := (*this.LpVtbl)[201]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(repeat))
	return com.Error(ret)
}

func (this *IPivotTable) GetChangeList(rhs **PivotTableChangeList) com.Error {
	addr := (*this.LpVtbl)[202]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetSlicers(rhs **Slicers) com.Error {
	addr := (*this.LpVtbl)[203]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
		com.AddToScope(rhs)
	return com.Error(ret)
}

func (this *IPivotTable) GetAlternativeText(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[204]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetAlternativeText(rhs string) com.Error {
	addr := (*this.LpVtbl)[205]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetSummary(rhs *win32.BSTR) com.Error {
	addr := (*this.LpVtbl)[206]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetSummary(rhs string) com.Error {
	addr := (*this.LpVtbl)[207]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) GetVisualTotalsForSets(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[208]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetVisualTotalsForSets(rhs bool) com.Error {
	addr := (*this.LpVtbl)[209]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetShowValuesRow(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[210]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetShowValuesRow(rhs bool) com.Error {
	addr := (*this.LpVtbl)[211]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

func (this *IPivotTable) GetCalculatedMembersInFilters(rhs *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[212]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(rhs)))
	return com.Error(ret)
}

func (this *IPivotTable) SetCalculatedMembersInFilters(rhs bool) com.Error {
	addr := (*this.LpVtbl)[213]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&rhs))))
	return com.Error(ret)
}

