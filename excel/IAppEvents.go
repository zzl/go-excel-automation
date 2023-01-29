package excel

import (
	"github.com/zzl/go-win32api/v2/win32"
	"github.com/zzl/go-com/com"
	"syscall"
	"unsafe"
)

// 00024413-0001-0000-C000-000000000046
var IID_IAppEvents = syscall.GUID{0x00024413, 0x0001, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IAppEvents struct {
	win32.IDispatch
}

func NewIAppEvents(pUnk *win32.IUnknown, addRef bool, scoped bool) *IAppEvents {
	if pUnk == nil {
		return nil
	}
	p := (*IAppEvents)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IAppEvents) IID() *syscall.GUID {
	return &IID_IAppEvents
}

func (this *IAppEvents) NewWorkbook(wb *Workbook) com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wb)))
	return com.Error(ret)
}

func (this *IAppEvents) SheetSelectionChange(sh *win32.IUnknown, target *Range) com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)), uintptr(unsafe.Pointer(target)))
	return com.Error(ret)
}

func (this *IAppEvents) SheetBeforeDoubleClick(sh *win32.IUnknown, target *Range, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)), uintptr(unsafe.Pointer(target)), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IAppEvents) SheetBeforeRightClick(sh *win32.IUnknown, target *Range, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)), uintptr(unsafe.Pointer(target)), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IAppEvents) SheetActivate(sh *win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)))
	return com.Error(ret)
}

func (this *IAppEvents) SheetDeactivate(sh *win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)))
	return com.Error(ret)
}

func (this *IAppEvents) SheetCalculate(sh *win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)))
	return com.Error(ret)
}

func (this *IAppEvents) SheetChange(sh *win32.IUnknown, target *Range) com.Error {
	addr := (*this.LpVtbl)[14]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)), uintptr(unsafe.Pointer(target)))
	return com.Error(ret)
}

func (this *IAppEvents) WorkbookOpen(wb *Workbook) com.Error {
	addr := (*this.LpVtbl)[15]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wb)))
	return com.Error(ret)
}

func (this *IAppEvents) WorkbookActivate(wb *Workbook) com.Error {
	addr := (*this.LpVtbl)[16]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wb)))
	return com.Error(ret)
}

func (this *IAppEvents) WorkbookDeactivate(wb *Workbook) com.Error {
	addr := (*this.LpVtbl)[17]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wb)))
	return com.Error(ret)
}

func (this *IAppEvents) WorkbookBeforeClose(wb *Workbook, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[18]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wb)), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IAppEvents) WorkbookBeforeSave(wb *Workbook, saveAsUI bool, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[19]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wb)), uintptr(*(*uint8)(unsafe.Pointer(&saveAsUI))), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IAppEvents) WorkbookBeforePrint(wb *Workbook, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[20]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wb)), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IAppEvents) WorkbookNewSheet(wb *Workbook, sh *win32.IUnknown) com.Error {
	addr := (*this.LpVtbl)[21]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wb)), uintptr(unsafe.Pointer(sh)))
	return com.Error(ret)
}

func (this *IAppEvents) WorkbookAddinInstall(wb *Workbook) com.Error {
	addr := (*this.LpVtbl)[22]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wb)))
	return com.Error(ret)
}

func (this *IAppEvents) WorkbookAddinUninstall(wb *Workbook) com.Error {
	addr := (*this.LpVtbl)[23]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wb)))
	return com.Error(ret)
}

func (this *IAppEvents) WindowResize(wb *Workbook, wn *Window) com.Error {
	addr := (*this.LpVtbl)[24]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wb)), uintptr(unsafe.Pointer(wn)))
	return com.Error(ret)
}

func (this *IAppEvents) WindowActivate(wb *Workbook, wn *Window) com.Error {
	addr := (*this.LpVtbl)[25]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wb)), uintptr(unsafe.Pointer(wn)))
	return com.Error(ret)
}

func (this *IAppEvents) WindowDeactivate(wb *Workbook, wn *Window) com.Error {
	addr := (*this.LpVtbl)[26]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wb)), uintptr(unsafe.Pointer(wn)))
	return com.Error(ret)
}

func (this *IAppEvents) SheetFollowHyperlink(sh *win32.IUnknown, target *Hyperlink) com.Error {
	addr := (*this.LpVtbl)[27]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)), uintptr(unsafe.Pointer(target)))
	return com.Error(ret)
}

func (this *IAppEvents) SheetPivotTableUpdate(sh *win32.IUnknown, target *PivotTable) com.Error {
	addr := (*this.LpVtbl)[28]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)), uintptr(unsafe.Pointer(target)))
	return com.Error(ret)
}

func (this *IAppEvents) WorkbookPivotTableCloseConnection(wb *Workbook, target *PivotTable) com.Error {
	addr := (*this.LpVtbl)[29]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wb)), uintptr(unsafe.Pointer(target)))
	return com.Error(ret)
}

func (this *IAppEvents) WorkbookPivotTableOpenConnection(wb *Workbook, target *PivotTable) com.Error {
	addr := (*this.LpVtbl)[30]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wb)), uintptr(unsafe.Pointer(target)))
	return com.Error(ret)
}

func (this *IAppEvents) WorkbookSync(wb *Workbook, syncEventType int32) com.Error {
	addr := (*this.LpVtbl)[31]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wb)), uintptr(syncEventType))
	return com.Error(ret)
}

func (this *IAppEvents) WorkbookBeforeXmlImport(wb *Workbook, map_ *XmlMap, url string, isRefresh bool, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[32]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wb)), uintptr(unsafe.Pointer(map_)), uintptr(win32.StrToPointer(url)), uintptr(*(*uint8)(unsafe.Pointer(&isRefresh))), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IAppEvents) WorkbookAfterXmlImport(wb *Workbook, map_ *XmlMap, isRefresh bool, result int32) com.Error {
	addr := (*this.LpVtbl)[33]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wb)), uintptr(unsafe.Pointer(map_)), uintptr(*(*uint8)(unsafe.Pointer(&isRefresh))), uintptr(result))
	return com.Error(ret)
}

func (this *IAppEvents) WorkbookBeforeXmlExport(wb *Workbook, map_ *XmlMap, url string, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[34]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wb)), uintptr(unsafe.Pointer(map_)), uintptr(win32.StrToPointer(url)), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IAppEvents) WorkbookAfterXmlExport(wb *Workbook, map_ *XmlMap, url string, result int32) com.Error {
	addr := (*this.LpVtbl)[35]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wb)), uintptr(unsafe.Pointer(map_)), uintptr(win32.StrToPointer(url)), uintptr(result))
	return com.Error(ret)
}

func (this *IAppEvents) WorkbookRowsetComplete(wb *Workbook, description string, sheet string, success bool) com.Error {
	addr := (*this.LpVtbl)[36]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wb)), uintptr(win32.StrToPointer(description)), uintptr(win32.StrToPointer(sheet)), uintptr(*(*uint8)(unsafe.Pointer(&success))))
	return com.Error(ret)
}

func (this *IAppEvents) AfterCalculate() com.Error {
	addr := (*this.LpVtbl)[37]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IAppEvents) SheetPivotTableAfterValueChange(sh *win32.IUnknown, targetPivotTable *PivotTable, targetRange *Range) com.Error {
	addr := (*this.LpVtbl)[38]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)), uintptr(unsafe.Pointer(targetPivotTable)), uintptr(unsafe.Pointer(targetRange)))
	return com.Error(ret)
}

func (this *IAppEvents) SheetPivotTableBeforeAllocateChanges(sh *win32.IUnknown, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[39]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)), uintptr(unsafe.Pointer(targetPivotTable)), uintptr(valueChangeStart), uintptr(valueChangeEnd), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IAppEvents) SheetPivotTableBeforeCommitChanges(sh *win32.IUnknown, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[40]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)), uintptr(unsafe.Pointer(targetPivotTable)), uintptr(valueChangeStart), uintptr(valueChangeEnd), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IAppEvents) SheetPivotTableBeforeDiscardChanges(sh *win32.IUnknown, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32) com.Error {
	addr := (*this.LpVtbl)[41]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)), uintptr(unsafe.Pointer(targetPivotTable)), uintptr(valueChangeStart), uintptr(valueChangeEnd))
	return com.Error(ret)
}

func (this *IAppEvents) ProtectedViewWindowOpen(pvw *ProtectedViewWindow) com.Error {
	addr := (*this.LpVtbl)[42]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(pvw)))
	return com.Error(ret)
}

func (this *IAppEvents) ProtectedViewWindowBeforeEdit(pvw *ProtectedViewWindow, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[43]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(pvw)), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IAppEvents) ProtectedViewWindowBeforeClose(pvw *ProtectedViewWindow, reason int32, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[44]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(pvw)), uintptr(reason), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IAppEvents) ProtectedViewWindowResize(pvw *ProtectedViewWindow) com.Error {
	addr := (*this.LpVtbl)[45]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(pvw)))
	return com.Error(ret)
}

func (this *IAppEvents) ProtectedViewWindowActivate(pvw *ProtectedViewWindow) com.Error {
	addr := (*this.LpVtbl)[46]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(pvw)))
	return com.Error(ret)
}

func (this *IAppEvents) ProtectedViewWindowDeactivate(pvw *ProtectedViewWindow) com.Error {
	addr := (*this.LpVtbl)[47]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(pvw)))
	return com.Error(ret)
}

func (this *IAppEvents) WorkbookAfterSave(wb *Workbook, success bool) com.Error {
	addr := (*this.LpVtbl)[48]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wb)), uintptr(*(*uint8)(unsafe.Pointer(&success))))
	return com.Error(ret)
}

func (this *IAppEvents) WorkbookNewChart(wb *Workbook, ch *Chart) com.Error {
	addr := (*this.LpVtbl)[49]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wb)), uintptr(unsafe.Pointer(ch)))
	return com.Error(ret)
}

