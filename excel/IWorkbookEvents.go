package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024412-0001-0000-C000-000000000046
var IID_IWorkbookEvents = syscall.GUID{0x00024412, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IWorkbookEvents struct {
	win32.IDispatch
}

func NewIWorkbookEvents(pUnk *win32.IUnknown, addRef bool, scoped bool) *IWorkbookEvents {
	p := (*IWorkbookEvents)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IWorkbookEvents) IID() *syscall.GUID {
	return &IID_IWorkbookEvents
}

func (this *IWorkbookEvents) Open() com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) Activate() com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) Deactivate() com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) BeforeClose(cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) BeforeSave(saveAsUI bool, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&saveAsUI))), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) BeforePrint(cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) NewSheet(sh *ole.DispatchClass) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) AddinInstall() com.Error {
	addr := (*this.LpVtbl)[14]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) AddinUninstall() com.Error {
	addr := (*this.LpVtbl)[15]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) WindowResize(wn *Window) com.Error {
	addr := (*this.LpVtbl)[16]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wn)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) WindowActivate(wn *Window) com.Error {
	addr := (*this.LpVtbl)[17]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wn)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) WindowDeactivate(wn *Window) com.Error {
	addr := (*this.LpVtbl)[18]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(wn)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) SheetSelectionChange(sh *ole.DispatchClass, target *Range) com.Error {
	addr := (*this.LpVtbl)[19]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)), uintptr(unsafe.Pointer(target)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) SheetBeforeDoubleClick(sh *ole.DispatchClass, target *Range, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[20]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)), uintptr(unsafe.Pointer(target)), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) SheetBeforeRightClick(sh *ole.DispatchClass, target *Range, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[21]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)), uintptr(unsafe.Pointer(target)), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) SheetActivate(sh *ole.DispatchClass) com.Error {
	addr := (*this.LpVtbl)[22]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) SheetDeactivate(sh *ole.DispatchClass) com.Error {
	addr := (*this.LpVtbl)[23]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) SheetCalculate(sh *ole.DispatchClass) com.Error {
	addr := (*this.LpVtbl)[24]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) SheetChange(sh *ole.DispatchClass, target *Range) com.Error {
	addr := (*this.LpVtbl)[25]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)), uintptr(unsafe.Pointer(target)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) SheetFollowHyperlink(sh *ole.DispatchClass, target *Hyperlink) com.Error {
	addr := (*this.LpVtbl)[26]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)), uintptr(unsafe.Pointer(target)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) SheetPivotTableUpdate(sh *ole.DispatchClass, target *PivotTable) com.Error {
	addr := (*this.LpVtbl)[27]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)), uintptr(unsafe.Pointer(target)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) PivotTableCloseConnection(target *PivotTable) com.Error {
	addr := (*this.LpVtbl)[28]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(target)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) PivotTableOpenConnection(target *PivotTable) com.Error {
	addr := (*this.LpVtbl)[29]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(target)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) Sync(syncEventType int32) com.Error {
	addr := (*this.LpVtbl)[30]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(syncEventType))
	return com.Error(ret)
}

func (this *IWorkbookEvents) BeforeXmlImport(map_ *XmlMap, url string, isRefresh bool, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[31]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(map_)), uintptr(win32.StrToPointer(url)), uintptr(*(*uint8)(unsafe.Pointer(&isRefresh))), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) AfterXmlImport(map_ *XmlMap, isRefresh bool, result int32) com.Error {
	addr := (*this.LpVtbl)[32]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(map_)), uintptr(*(*uint8)(unsafe.Pointer(&isRefresh))), uintptr(result))
	return com.Error(ret)
}

func (this *IWorkbookEvents) BeforeXmlExport(map_ *XmlMap, url string, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[33]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(map_)), uintptr(win32.StrToPointer(url)), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) AfterXmlExport(map_ *XmlMap, url string, result int32) com.Error {
	addr := (*this.LpVtbl)[34]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(map_)), uintptr(win32.StrToPointer(url)), uintptr(result))
	return com.Error(ret)
}

func (this *IWorkbookEvents) RowsetComplete(description string, sheet string, success bool) com.Error {
	addr := (*this.LpVtbl)[35]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(win32.StrToPointer(description)), uintptr(win32.StrToPointer(sheet)), uintptr(*(*uint8)(unsafe.Pointer(&success))))
	return com.Error(ret)
}

func (this *IWorkbookEvents) SheetPivotTableAfterValueChange(sh *ole.DispatchClass, targetPivotTable *PivotTable, targetRange *Range) com.Error {
	addr := (*this.LpVtbl)[36]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)), uintptr(unsafe.Pointer(targetPivotTable)), uintptr(unsafe.Pointer(targetRange)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) SheetPivotTableBeforeAllocateChanges(sh *ole.DispatchClass, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[37]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)), uintptr(unsafe.Pointer(targetPivotTable)), uintptr(valueChangeStart), uintptr(valueChangeEnd), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) SheetPivotTableBeforeCommitChanges(sh *ole.DispatchClass, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[38]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)), uintptr(unsafe.Pointer(targetPivotTable)), uintptr(valueChangeStart), uintptr(valueChangeEnd), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) SheetPivotTableBeforeDiscardChanges(sh *ole.DispatchClass, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32) com.Error {
	addr := (*this.LpVtbl)[39]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)), uintptr(unsafe.Pointer(targetPivotTable)), uintptr(valueChangeStart), uintptr(valueChangeEnd))
	return com.Error(ret)
}

func (this *IWorkbookEvents) SheetPivotTableChangeSync(sh *ole.DispatchClass, target *PivotTable) com.Error {
	addr := (*this.LpVtbl)[40]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sh)), uintptr(unsafe.Pointer(target)))
	return com.Error(ret)
}

func (this *IWorkbookEvents) AfterSave(success bool) com.Error {
	addr := (*this.LpVtbl)[41]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(*(*uint8)(unsafe.Pointer(&success))))
	return com.Error(ret)
}

func (this *IWorkbookEvents) NewChart(ch *Chart) com.Error {
	addr := (*this.LpVtbl)[42]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(ch)))
	return com.Error(ret)
}

