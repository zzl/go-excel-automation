package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024413-0000-0000-C000-000000000046
var IID_AppEvents = syscall.GUID{0x00024413, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type AppEventsDispInterface interface {
	QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) 
	AddRef_() uint32
	Release_() uint32
	GetTypeInfoCount_(pctinfo *uint32) 
	GetTypeInfo_(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) 
	GetIDsOfNames_(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) 
	Invoke_(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) 
	NewWorkbook(wb *Workbook) 
	SheetSelectionChange(sh *win32.IUnknown, target *Range) 
	SheetBeforeDoubleClick(sh *win32.IUnknown, target *Range, cancel *win32.VARIANT_BOOL) 
	SheetBeforeRightClick(sh *win32.IUnknown, target *Range, cancel *win32.VARIANT_BOOL) 
	SheetActivate(sh *win32.IUnknown) 
	SheetDeactivate(sh *win32.IUnknown) 
	SheetCalculate(sh *win32.IUnknown) 
	SheetChange(sh *win32.IUnknown, target *Range) 
	WorkbookOpen(wb *Workbook) 
	WorkbookActivate(wb *Workbook) 
	WorkbookDeactivate(wb *Workbook) 
	WorkbookBeforeClose(wb *Workbook, cancel *win32.VARIANT_BOOL) 
	WorkbookBeforeSave(wb *Workbook, saveAsUI bool, cancel *win32.VARIANT_BOOL) 
	WorkbookBeforePrint(wb *Workbook, cancel *win32.VARIANT_BOOL) 
	WorkbookNewSheet(wb *Workbook, sh *win32.IUnknown) 
	WorkbookAddinInstall(wb *Workbook) 
	WorkbookAddinUninstall(wb *Workbook) 
	WindowResize(wb *Workbook, wn *Window) 
	WindowActivate(wb *Workbook, wn *Window) 
	WindowDeactivate(wb *Workbook, wn *Window) 
	SheetFollowHyperlink(sh *win32.IUnknown, target *Hyperlink) 
	SheetPivotTableUpdate(sh *win32.IUnknown, target *PivotTable) 
	WorkbookPivotTableCloseConnection(wb *Workbook, target *PivotTable) 
	WorkbookPivotTableOpenConnection(wb *Workbook, target *PivotTable) 
	WorkbookSync(wb *Workbook, syncEventType int32) 
	WorkbookBeforeXmlImport(wb *Workbook, map_ *XmlMap, url string, isRefresh bool, cancel *win32.VARIANT_BOOL) 
	WorkbookAfterXmlImport(wb *Workbook, map_ *XmlMap, isRefresh bool, result int32) 
	WorkbookBeforeXmlExport(wb *Workbook, map_ *XmlMap, url string, cancel *win32.VARIANT_BOOL) 
	WorkbookAfterXmlExport(wb *Workbook, map_ *XmlMap, url string, result int32) 
	WorkbookRowsetComplete(wb *Workbook, description string, sheet string, success bool) 
	AfterCalculate() 
	SheetPivotTableAfterValueChange(sh *win32.IUnknown, targetPivotTable *PivotTable, targetRange *Range) 
	SheetPivotTableBeforeAllocateChanges(sh *win32.IUnknown, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL) 
	SheetPivotTableBeforeCommitChanges(sh *win32.IUnknown, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL) 
	SheetPivotTableBeforeDiscardChanges(sh *win32.IUnknown, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32) 
	ProtectedViewWindowOpen(pvw *ProtectedViewWindow) 
	ProtectedViewWindowBeforeEdit(pvw *ProtectedViewWindow, cancel *win32.VARIANT_BOOL) 
	ProtectedViewWindowBeforeClose(pvw *ProtectedViewWindow, reason int32, cancel *win32.VARIANT_BOOL) 
	ProtectedViewWindowResize(pvw *ProtectedViewWindow) 
	ProtectedViewWindowActivate(pvw *ProtectedViewWindow) 
	ProtectedViewWindowDeactivate(pvw *ProtectedViewWindow) 
	WorkbookAfterSave(wb *Workbook, success bool) 
	WorkbookNewChart(wb *Workbook, ch *Chart) 
}

type AppEventsHandlers struct {
	QueryInterface_ func(riid *syscall.GUID, ppvObj unsafe.Pointer) 
	AddRef_ func() uint32
	Release_ func() uint32
	GetTypeInfoCount_ func(pctinfo *uint32) 
	GetTypeInfo_ func(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) 
	GetIDsOfNames_ func(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) 
	Invoke_ func(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) 
	NewWorkbook func(wb *Workbook) 
	SheetSelectionChange func(sh *win32.IUnknown, target *Range) 
	SheetBeforeDoubleClick func(sh *win32.IUnknown, target *Range, cancel *win32.VARIANT_BOOL) 
	SheetBeforeRightClick func(sh *win32.IUnknown, target *Range, cancel *win32.VARIANT_BOOL) 
	SheetActivate func(sh *win32.IUnknown) 
	SheetDeactivate func(sh *win32.IUnknown) 
	SheetCalculate func(sh *win32.IUnknown) 
	SheetChange func(sh *win32.IUnknown, target *Range) 
	WorkbookOpen func(wb *Workbook) 
	WorkbookActivate func(wb *Workbook) 
	WorkbookDeactivate func(wb *Workbook) 
	WorkbookBeforeClose func(wb *Workbook, cancel *win32.VARIANT_BOOL) 
	WorkbookBeforeSave func(wb *Workbook, saveAsUI bool, cancel *win32.VARIANT_BOOL) 
	WorkbookBeforePrint func(wb *Workbook, cancel *win32.VARIANT_BOOL) 
	WorkbookNewSheet func(wb *Workbook, sh *win32.IUnknown) 
	WorkbookAddinInstall func(wb *Workbook) 
	WorkbookAddinUninstall func(wb *Workbook) 
	WindowResize func(wb *Workbook, wn *Window) 
	WindowActivate func(wb *Workbook, wn *Window) 
	WindowDeactivate func(wb *Workbook, wn *Window) 
	SheetFollowHyperlink func(sh *win32.IUnknown, target *Hyperlink) 
	SheetPivotTableUpdate func(sh *win32.IUnknown, target *PivotTable) 
	WorkbookPivotTableCloseConnection func(wb *Workbook, target *PivotTable) 
	WorkbookPivotTableOpenConnection func(wb *Workbook, target *PivotTable) 
	WorkbookSync func(wb *Workbook, syncEventType int32) 
	WorkbookBeforeXmlImport func(wb *Workbook, map_ *XmlMap, url string, isRefresh bool, cancel *win32.VARIANT_BOOL) 
	WorkbookAfterXmlImport func(wb *Workbook, map_ *XmlMap, isRefresh bool, result int32) 
	WorkbookBeforeXmlExport func(wb *Workbook, map_ *XmlMap, url string, cancel *win32.VARIANT_BOOL) 
	WorkbookAfterXmlExport func(wb *Workbook, map_ *XmlMap, url string, result int32) 
	WorkbookRowsetComplete func(wb *Workbook, description string, sheet string, success bool) 
	AfterCalculate func() 
	SheetPivotTableAfterValueChange func(sh *win32.IUnknown, targetPivotTable *PivotTable, targetRange *Range) 
	SheetPivotTableBeforeAllocateChanges func(sh *win32.IUnknown, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL) 
	SheetPivotTableBeforeCommitChanges func(sh *win32.IUnknown, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL) 
	SheetPivotTableBeforeDiscardChanges func(sh *win32.IUnknown, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32) 
	ProtectedViewWindowOpen func(pvw *ProtectedViewWindow) 
	ProtectedViewWindowBeforeEdit func(pvw *ProtectedViewWindow, cancel *win32.VARIANT_BOOL) 
	ProtectedViewWindowBeforeClose func(pvw *ProtectedViewWindow, reason int32, cancel *win32.VARIANT_BOOL) 
	ProtectedViewWindowResize func(pvw *ProtectedViewWindow) 
	ProtectedViewWindowActivate func(pvw *ProtectedViewWindow) 
	ProtectedViewWindowDeactivate func(pvw *ProtectedViewWindow) 
	WorkbookAfterSave func(wb *Workbook, success bool) 
	WorkbookNewChart func(wb *Workbook, ch *Chart) 
}

type AppEventsDispImpl struct {
	Handlers AppEventsHandlers
}

func (this *AppEventsDispImpl) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	if this.Handlers.QueryInterface_ != nil {
		this.Handlers.QueryInterface_(riid, ppvObj)
	}
}

func (this *AppEventsDispImpl) AddRef_() uint32 {
	if this.Handlers.AddRef_ != nil {
		return this.Handlers.AddRef_()
	}
	var ret uint32
	return ret
}

func (this *AppEventsDispImpl) Release_() uint32 {
	if this.Handlers.Release_ != nil {
		return this.Handlers.Release_()
	}
	var ret uint32
	return ret
}

func (this *AppEventsDispImpl) GetTypeInfoCount_(pctinfo *uint32) {
	if this.Handlers.GetTypeInfoCount_ != nil {
		this.Handlers.GetTypeInfoCount_(pctinfo)
	}
}

func (this *AppEventsDispImpl) GetTypeInfo_(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	if this.Handlers.GetTypeInfo_ != nil {
		this.Handlers.GetTypeInfo_(itinfo, lcid, pptinfo)
	}
}

func (this *AppEventsDispImpl) GetIDsOfNames_(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	if this.Handlers.GetIDsOfNames_ != nil {
		this.Handlers.GetIDsOfNames_(riid, rgszNames, cNames, lcid, rgdispid)
	}
}

func (this *AppEventsDispImpl) Invoke_(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	if this.Handlers.Invoke_ != nil {
		this.Handlers.Invoke_(dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr)
	}
}

func (this *AppEventsDispImpl) NewWorkbook(wb *Workbook) {
	if this.Handlers.NewWorkbook != nil {
		this.Handlers.NewWorkbook(wb)
	}
}

func (this *AppEventsDispImpl) SheetSelectionChange(sh *win32.IUnknown, target *Range) {
	if this.Handlers.SheetSelectionChange != nil {
		this.Handlers.SheetSelectionChange(sh, target)
	}
}

func (this *AppEventsDispImpl) SheetBeforeDoubleClick(sh *win32.IUnknown, target *Range, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.SheetBeforeDoubleClick != nil {
		this.Handlers.SheetBeforeDoubleClick(sh, target, cancel)
	}
}

func (this *AppEventsDispImpl) SheetBeforeRightClick(sh *win32.IUnknown, target *Range, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.SheetBeforeRightClick != nil {
		this.Handlers.SheetBeforeRightClick(sh, target, cancel)
	}
}

func (this *AppEventsDispImpl) SheetActivate(sh *win32.IUnknown) {
	if this.Handlers.SheetActivate != nil {
		this.Handlers.SheetActivate(sh)
	}
}

func (this *AppEventsDispImpl) SheetDeactivate(sh *win32.IUnknown) {
	if this.Handlers.SheetDeactivate != nil {
		this.Handlers.SheetDeactivate(sh)
	}
}

func (this *AppEventsDispImpl) SheetCalculate(sh *win32.IUnknown) {
	if this.Handlers.SheetCalculate != nil {
		this.Handlers.SheetCalculate(sh)
	}
}

func (this *AppEventsDispImpl) SheetChange(sh *win32.IUnknown, target *Range) {
	if this.Handlers.SheetChange != nil {
		this.Handlers.SheetChange(sh, target)
	}
}

func (this *AppEventsDispImpl) WorkbookOpen(wb *Workbook) {
	if this.Handlers.WorkbookOpen != nil {
		this.Handlers.WorkbookOpen(wb)
	}
}

func (this *AppEventsDispImpl) WorkbookActivate(wb *Workbook) {
	if this.Handlers.WorkbookActivate != nil {
		this.Handlers.WorkbookActivate(wb)
	}
}

func (this *AppEventsDispImpl) WorkbookDeactivate(wb *Workbook) {
	if this.Handlers.WorkbookDeactivate != nil {
		this.Handlers.WorkbookDeactivate(wb)
	}
}

func (this *AppEventsDispImpl) WorkbookBeforeClose(wb *Workbook, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.WorkbookBeforeClose != nil {
		this.Handlers.WorkbookBeforeClose(wb, cancel)
	}
}

func (this *AppEventsDispImpl) WorkbookBeforeSave(wb *Workbook, saveAsUI bool, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.WorkbookBeforeSave != nil {
		this.Handlers.WorkbookBeforeSave(wb, saveAsUI, cancel)
	}
}

func (this *AppEventsDispImpl) WorkbookBeforePrint(wb *Workbook, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.WorkbookBeforePrint != nil {
		this.Handlers.WorkbookBeforePrint(wb, cancel)
	}
}

func (this *AppEventsDispImpl) WorkbookNewSheet(wb *Workbook, sh *win32.IUnknown) {
	if this.Handlers.WorkbookNewSheet != nil {
		this.Handlers.WorkbookNewSheet(wb, sh)
	}
}

func (this *AppEventsDispImpl) WorkbookAddinInstall(wb *Workbook) {
	if this.Handlers.WorkbookAddinInstall != nil {
		this.Handlers.WorkbookAddinInstall(wb)
	}
}

func (this *AppEventsDispImpl) WorkbookAddinUninstall(wb *Workbook) {
	if this.Handlers.WorkbookAddinUninstall != nil {
		this.Handlers.WorkbookAddinUninstall(wb)
	}
}

func (this *AppEventsDispImpl) WindowResize(wb *Workbook, wn *Window) {
	if this.Handlers.WindowResize != nil {
		this.Handlers.WindowResize(wb, wn)
	}
}

func (this *AppEventsDispImpl) WindowActivate(wb *Workbook, wn *Window) {
	if this.Handlers.WindowActivate != nil {
		this.Handlers.WindowActivate(wb, wn)
	}
}

func (this *AppEventsDispImpl) WindowDeactivate(wb *Workbook, wn *Window) {
	if this.Handlers.WindowDeactivate != nil {
		this.Handlers.WindowDeactivate(wb, wn)
	}
}

func (this *AppEventsDispImpl) SheetFollowHyperlink(sh *win32.IUnknown, target *Hyperlink) {
	if this.Handlers.SheetFollowHyperlink != nil {
		this.Handlers.SheetFollowHyperlink(sh, target)
	}
}

func (this *AppEventsDispImpl) SheetPivotTableUpdate(sh *win32.IUnknown, target *PivotTable) {
	if this.Handlers.SheetPivotTableUpdate != nil {
		this.Handlers.SheetPivotTableUpdate(sh, target)
	}
}

func (this *AppEventsDispImpl) WorkbookPivotTableCloseConnection(wb *Workbook, target *PivotTable) {
	if this.Handlers.WorkbookPivotTableCloseConnection != nil {
		this.Handlers.WorkbookPivotTableCloseConnection(wb, target)
	}
}

func (this *AppEventsDispImpl) WorkbookPivotTableOpenConnection(wb *Workbook, target *PivotTable) {
	if this.Handlers.WorkbookPivotTableOpenConnection != nil {
		this.Handlers.WorkbookPivotTableOpenConnection(wb, target)
	}
}

func (this *AppEventsDispImpl) WorkbookSync(wb *Workbook, syncEventType int32) {
	if this.Handlers.WorkbookSync != nil {
		this.Handlers.WorkbookSync(wb, syncEventType)
	}
}

func (this *AppEventsDispImpl) WorkbookBeforeXmlImport(wb *Workbook, map_ *XmlMap, url string, isRefresh bool, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.WorkbookBeforeXmlImport != nil {
		this.Handlers.WorkbookBeforeXmlImport(wb, map_, url, isRefresh, cancel)
	}
}

func (this *AppEventsDispImpl) WorkbookAfterXmlImport(wb *Workbook, map_ *XmlMap, isRefresh bool, result int32) {
	if this.Handlers.WorkbookAfterXmlImport != nil {
		this.Handlers.WorkbookAfterXmlImport(wb, map_, isRefresh, result)
	}
}

func (this *AppEventsDispImpl) WorkbookBeforeXmlExport(wb *Workbook, map_ *XmlMap, url string, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.WorkbookBeforeXmlExport != nil {
		this.Handlers.WorkbookBeforeXmlExport(wb, map_, url, cancel)
	}
}

func (this *AppEventsDispImpl) WorkbookAfterXmlExport(wb *Workbook, map_ *XmlMap, url string, result int32) {
	if this.Handlers.WorkbookAfterXmlExport != nil {
		this.Handlers.WorkbookAfterXmlExport(wb, map_, url, result)
	}
}

func (this *AppEventsDispImpl) WorkbookRowsetComplete(wb *Workbook, description string, sheet string, success bool) {
	if this.Handlers.WorkbookRowsetComplete != nil {
		this.Handlers.WorkbookRowsetComplete(wb, description, sheet, success)
	}
}

func (this *AppEventsDispImpl) AfterCalculate() {
	if this.Handlers.AfterCalculate != nil {
		this.Handlers.AfterCalculate()
	}
}

func (this *AppEventsDispImpl) SheetPivotTableAfterValueChange(sh *win32.IUnknown, targetPivotTable *PivotTable, targetRange *Range) {
	if this.Handlers.SheetPivotTableAfterValueChange != nil {
		this.Handlers.SheetPivotTableAfterValueChange(sh, targetPivotTable, targetRange)
	}
}

func (this *AppEventsDispImpl) SheetPivotTableBeforeAllocateChanges(sh *win32.IUnknown, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.SheetPivotTableBeforeAllocateChanges != nil {
		this.Handlers.SheetPivotTableBeforeAllocateChanges(sh, targetPivotTable, valueChangeStart, valueChangeEnd, cancel)
	}
}

func (this *AppEventsDispImpl) SheetPivotTableBeforeCommitChanges(sh *win32.IUnknown, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.SheetPivotTableBeforeCommitChanges != nil {
		this.Handlers.SheetPivotTableBeforeCommitChanges(sh, targetPivotTable, valueChangeStart, valueChangeEnd, cancel)
	}
}

func (this *AppEventsDispImpl) SheetPivotTableBeforeDiscardChanges(sh *win32.IUnknown, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32) {
	if this.Handlers.SheetPivotTableBeforeDiscardChanges != nil {
		this.Handlers.SheetPivotTableBeforeDiscardChanges(sh, targetPivotTable, valueChangeStart, valueChangeEnd)
	}
}

func (this *AppEventsDispImpl) ProtectedViewWindowOpen(pvw *ProtectedViewWindow) {
	if this.Handlers.ProtectedViewWindowOpen != nil {
		this.Handlers.ProtectedViewWindowOpen(pvw)
	}
}

func (this *AppEventsDispImpl) ProtectedViewWindowBeforeEdit(pvw *ProtectedViewWindow, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.ProtectedViewWindowBeforeEdit != nil {
		this.Handlers.ProtectedViewWindowBeforeEdit(pvw, cancel)
	}
}

func (this *AppEventsDispImpl) ProtectedViewWindowBeforeClose(pvw *ProtectedViewWindow, reason int32, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.ProtectedViewWindowBeforeClose != nil {
		this.Handlers.ProtectedViewWindowBeforeClose(pvw, reason, cancel)
	}
}

func (this *AppEventsDispImpl) ProtectedViewWindowResize(pvw *ProtectedViewWindow) {
	if this.Handlers.ProtectedViewWindowResize != nil {
		this.Handlers.ProtectedViewWindowResize(pvw)
	}
}

func (this *AppEventsDispImpl) ProtectedViewWindowActivate(pvw *ProtectedViewWindow) {
	if this.Handlers.ProtectedViewWindowActivate != nil {
		this.Handlers.ProtectedViewWindowActivate(pvw)
	}
}

func (this *AppEventsDispImpl) ProtectedViewWindowDeactivate(pvw *ProtectedViewWindow) {
	if this.Handlers.ProtectedViewWindowDeactivate != nil {
		this.Handlers.ProtectedViewWindowDeactivate(pvw)
	}
}

func (this *AppEventsDispImpl) WorkbookAfterSave(wb *Workbook, success bool) {
	if this.Handlers.WorkbookAfterSave != nil {
		this.Handlers.WorkbookAfterSave(wb, success)
	}
}

func (this *AppEventsDispImpl) WorkbookNewChart(wb *Workbook, ch *Chart) {
	if this.Handlers.WorkbookNewChart != nil {
		this.Handlers.WorkbookNewChart(wb, ch)
	}
}

type AppEventsImpl struct {
	ole.IDispatchImpl
	DispImpl AppEventsDispInterface
}

func (this *AppEventsImpl) QueryInterface(riid *syscall.GUID, ppvObject unsafe.Pointer) win32.HRESULT {
	if *riid == IID_AppEvents {
		this.AssignPpvObject(ppvObject)
		this.AddRef()
		return win32.S_OK
	}
	return this.IDispatchImpl.QueryInterface(riid, ppvObject)
}

func (this *AppEventsImpl) Invoke(dispIdMember int32, riid *syscall.GUID, lcid uint32,
	wFlags uint16, pDispParams *win32.DISPPARAMS, pVarResult *win32.VARIANT,
	pExcepInfo *win32.EXCEPINFO, puArgErr *uint32) win32.HRESULT {
	var unwrapActions ole.Actions
	defer unwrapActions.Execute()
	switch dispIdMember {
	case 1610612736:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*syscall.GUID)(vArgs[0].ToPointer())
		p2 := (unsafe.Pointer)(vArgs[1].ToPointer())
		this.DispImpl.QueryInterface_(p1, p2)
		return win32.S_OK
	case 1610612737:
		ret := this.DispImpl.AddRef_()
		ole.SetVariantParam((*ole.Variant)(pVarResult), ret, &unwrapActions)
		return win32.S_OK
	case 1610612738:
		ret := this.DispImpl.Release_()
		ole.SetVariantParam((*ole.Variant)(pVarResult), ret, &unwrapActions)
		return win32.S_OK
	case 1610678272:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*uint32)(vArgs[0].ToPointer())
		this.DispImpl.GetTypeInfoCount_(p1)
		return win32.S_OK
	case 1610678273:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 3)
		p1, _ := vArgs[0].ToUint32()
		p2, _ := vArgs[1].ToUint32()
		p3 := (unsafe.Pointer)(vArgs[2].ToPointer())
		this.DispImpl.GetTypeInfo_(p1, p2, p3)
		return win32.S_OK
	case 1610678274:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 5)
		p1 := (*syscall.GUID)(vArgs[0].ToPointer())
		p2 := (**int8)(vArgs[1].ToPointer())
		p3, _ := vArgs[2].ToUint32()
		p4, _ := vArgs[3].ToUint32()
		p5 := (*int32)(vArgs[4].ToPointer())
		this.DispImpl.GetIDsOfNames_(p1, p2, p3, p4, p5)
		return win32.S_OK
	case 1610678275:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 8)
		p1, _ := vArgs[0].ToInt32()
		p2 := (*syscall.GUID)(vArgs[1].ToPointer())
		p3, _ := vArgs[2].ToUint32()
		p4, _ := vArgs[3].ToUint16()
		p5 := (*win32.DISPPARAMS)(vArgs[4].ToPointer())
		p6 := (*ole.Variant)(vArgs[5].ToPointer())
		p7 := (*win32.EXCEPINFO)(vArgs[6].ToPointer())
		p8 := (*uint32)(vArgs[7].ToPointer())
		this.DispImpl.Invoke_(p1, p2, p3, p4, p5, p6, p7, p8)
		return win32.S_OK
	case 1565:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*Workbook)(vArgs[0].ToPointer())
		this.DispImpl.NewWorkbook(p1)
		return win32.S_OK
	case 1558:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*win32.IUnknown)(vArgs[0].ToPointer())
		p2 := (*Range)(vArgs[1].ToPointer())
		this.DispImpl.SheetSelectionChange(p1, p2)
		return win32.S_OK
	case 1559:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 3)
		p1 := (*win32.IUnknown)(vArgs[0].ToPointer())
		p2 := (*Range)(vArgs[1].ToPointer())
		p3 := (*win32.VARIANT_BOOL)(vArgs[2].ToPointer())
		this.DispImpl.SheetBeforeDoubleClick(p1, p2, p3)
		return win32.S_OK
	case 1560:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 3)
		p1 := (*win32.IUnknown)(vArgs[0].ToPointer())
		p2 := (*Range)(vArgs[1].ToPointer())
		p3 := (*win32.VARIANT_BOOL)(vArgs[2].ToPointer())
		this.DispImpl.SheetBeforeRightClick(p1, p2, p3)
		return win32.S_OK
	case 1561:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*win32.IUnknown)(vArgs[0].ToPointer())
		this.DispImpl.SheetActivate(p1)
		return win32.S_OK
	case 1562:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*win32.IUnknown)(vArgs[0].ToPointer())
		this.DispImpl.SheetDeactivate(p1)
		return win32.S_OK
	case 1563:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*win32.IUnknown)(vArgs[0].ToPointer())
		this.DispImpl.SheetCalculate(p1)
		return win32.S_OK
	case 1564:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*win32.IUnknown)(vArgs[0].ToPointer())
		p2 := (*Range)(vArgs[1].ToPointer())
		this.DispImpl.SheetChange(p1, p2)
		return win32.S_OK
	case 1567:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*Workbook)(vArgs[0].ToPointer())
		this.DispImpl.WorkbookOpen(p1)
		return win32.S_OK
	case 1568:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*Workbook)(vArgs[0].ToPointer())
		this.DispImpl.WorkbookActivate(p1)
		return win32.S_OK
	case 1569:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*Workbook)(vArgs[0].ToPointer())
		this.DispImpl.WorkbookDeactivate(p1)
		return win32.S_OK
	case 1570:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*Workbook)(vArgs[0].ToPointer())
		p2 := (*win32.VARIANT_BOOL)(vArgs[1].ToPointer())
		this.DispImpl.WorkbookBeforeClose(p1, p2)
		return win32.S_OK
	case 1571:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 3)
		p1 := (*Workbook)(vArgs[0].ToPointer())
		p2, _ := vArgs[1].ToBool()
		p3 := (*win32.VARIANT_BOOL)(vArgs[2].ToPointer())
		this.DispImpl.WorkbookBeforeSave(p1, p2, p3)
		return win32.S_OK
	case 1572:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*Workbook)(vArgs[0].ToPointer())
		p2 := (*win32.VARIANT_BOOL)(vArgs[1].ToPointer())
		this.DispImpl.WorkbookBeforePrint(p1, p2)
		return win32.S_OK
	case 1573:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*Workbook)(vArgs[0].ToPointer())
		p2 := (*win32.IUnknown)(vArgs[1].ToPointer())
		this.DispImpl.WorkbookNewSheet(p1, p2)
		return win32.S_OK
	case 1574:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*Workbook)(vArgs[0].ToPointer())
		this.DispImpl.WorkbookAddinInstall(p1)
		return win32.S_OK
	case 1575:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*Workbook)(vArgs[0].ToPointer())
		this.DispImpl.WorkbookAddinUninstall(p1)
		return win32.S_OK
	case 1554:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*Workbook)(vArgs[0].ToPointer())
		p2 := (*Window)(vArgs[1].ToPointer())
		this.DispImpl.WindowResize(p1, p2)
		return win32.S_OK
	case 1556:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*Workbook)(vArgs[0].ToPointer())
		p2 := (*Window)(vArgs[1].ToPointer())
		this.DispImpl.WindowActivate(p1, p2)
		return win32.S_OK
	case 1557:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*Workbook)(vArgs[0].ToPointer())
		p2 := (*Window)(vArgs[1].ToPointer())
		this.DispImpl.WindowDeactivate(p1, p2)
		return win32.S_OK
	case 1854:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*win32.IUnknown)(vArgs[0].ToPointer())
		p2 := (*Hyperlink)(vArgs[1].ToPointer())
		this.DispImpl.SheetFollowHyperlink(p1, p2)
		return win32.S_OK
	case 2157:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*win32.IUnknown)(vArgs[0].ToPointer())
		p2 := (*PivotTable)(vArgs[1].ToPointer())
		this.DispImpl.SheetPivotTableUpdate(p1, p2)
		return win32.S_OK
	case 2160:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*Workbook)(vArgs[0].ToPointer())
		p2 := (*PivotTable)(vArgs[1].ToPointer())
		this.DispImpl.WorkbookPivotTableCloseConnection(p1, p2)
		return win32.S_OK
	case 2161:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*Workbook)(vArgs[0].ToPointer())
		p2 := (*PivotTable)(vArgs[1].ToPointer())
		this.DispImpl.WorkbookPivotTableOpenConnection(p1, p2)
		return win32.S_OK
	case 2289:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*Workbook)(vArgs[0].ToPointer())
		p2, _ := vArgs[1].ToInt32()
		this.DispImpl.WorkbookSync(p1, p2)
		return win32.S_OK
	case 2290:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 5)
		p1 := (*Workbook)(vArgs[0].ToPointer())
		p2 := (*XmlMap)(vArgs[1].ToPointer())
		p3, _ := vArgs[2].ToString()
		p4, _ := vArgs[3].ToBool()
		p5 := (*win32.VARIANT_BOOL)(vArgs[4].ToPointer())
		this.DispImpl.WorkbookBeforeXmlImport(p1, p2, p3, p4, p5)
		return win32.S_OK
	case 2291:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 4)
		p1 := (*Workbook)(vArgs[0].ToPointer())
		p2 := (*XmlMap)(vArgs[1].ToPointer())
		p3, _ := vArgs[2].ToBool()
		p4, _ := vArgs[3].ToInt32()
		this.DispImpl.WorkbookAfterXmlImport(p1, p2, p3, p4)
		return win32.S_OK
	case 2292:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 4)
		p1 := (*Workbook)(vArgs[0].ToPointer())
		p2 := (*XmlMap)(vArgs[1].ToPointer())
		p3, _ := vArgs[2].ToString()
		p4 := (*win32.VARIANT_BOOL)(vArgs[3].ToPointer())
		this.DispImpl.WorkbookBeforeXmlExport(p1, p2, p3, p4)
		return win32.S_OK
	case 2293:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 4)
		p1 := (*Workbook)(vArgs[0].ToPointer())
		p2 := (*XmlMap)(vArgs[1].ToPointer())
		p3, _ := vArgs[2].ToString()
		p4, _ := vArgs[3].ToInt32()
		this.DispImpl.WorkbookAfterXmlExport(p1, p2, p3, p4)
		return win32.S_OK
	case 2611:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 4)
		p1 := (*Workbook)(vArgs[0].ToPointer())
		p2, _ := vArgs[1].ToString()
		p3, _ := vArgs[2].ToString()
		p4, _ := vArgs[3].ToBool()
		this.DispImpl.WorkbookRowsetComplete(p1, p2, p3, p4)
		return win32.S_OK
	case 2612:
		this.DispImpl.AfterCalculate()
		return win32.S_OK
	case 2895:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 3)
		p1 := (*win32.IUnknown)(vArgs[0].ToPointer())
		p2 := (*PivotTable)(vArgs[1].ToPointer())
		p3 := (*Range)(vArgs[2].ToPointer())
		this.DispImpl.SheetPivotTableAfterValueChange(p1, p2, p3)
		return win32.S_OK
	case 2896:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 5)
		p1 := (*win32.IUnknown)(vArgs[0].ToPointer())
		p2 := (*PivotTable)(vArgs[1].ToPointer())
		p3, _ := vArgs[2].ToInt32()
		p4, _ := vArgs[3].ToInt32()
		p5 := (*win32.VARIANT_BOOL)(vArgs[4].ToPointer())
		this.DispImpl.SheetPivotTableBeforeAllocateChanges(p1, p2, p3, p4, p5)
		return win32.S_OK
	case 2897:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 5)
		p1 := (*win32.IUnknown)(vArgs[0].ToPointer())
		p2 := (*PivotTable)(vArgs[1].ToPointer())
		p3, _ := vArgs[2].ToInt32()
		p4, _ := vArgs[3].ToInt32()
		p5 := (*win32.VARIANT_BOOL)(vArgs[4].ToPointer())
		this.DispImpl.SheetPivotTableBeforeCommitChanges(p1, p2, p3, p4, p5)
		return win32.S_OK
	case 2898:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 4)
		p1 := (*win32.IUnknown)(vArgs[0].ToPointer())
		p2 := (*PivotTable)(vArgs[1].ToPointer())
		p3, _ := vArgs[2].ToInt32()
		p4, _ := vArgs[3].ToInt32()
		this.DispImpl.SheetPivotTableBeforeDiscardChanges(p1, p2, p3, p4)
		return win32.S_OK
	case 2903:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*ProtectedViewWindow)(vArgs[0].ToPointer())
		this.DispImpl.ProtectedViewWindowOpen(p1)
		return win32.S_OK
	case 2905:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*ProtectedViewWindow)(vArgs[0].ToPointer())
		p2 := (*win32.VARIANT_BOOL)(vArgs[1].ToPointer())
		this.DispImpl.ProtectedViewWindowBeforeEdit(p1, p2)
		return win32.S_OK
	case 2906:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 3)
		p1 := (*ProtectedViewWindow)(vArgs[0].ToPointer())
		p2, _ := vArgs[1].ToInt32()
		p3 := (*win32.VARIANT_BOOL)(vArgs[2].ToPointer())
		this.DispImpl.ProtectedViewWindowBeforeClose(p1, p2, p3)
		return win32.S_OK
	case 2908:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*ProtectedViewWindow)(vArgs[0].ToPointer())
		this.DispImpl.ProtectedViewWindowResize(p1)
		return win32.S_OK
	case 2909:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*ProtectedViewWindow)(vArgs[0].ToPointer())
		this.DispImpl.ProtectedViewWindowActivate(p1)
		return win32.S_OK
	case 2910:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*ProtectedViewWindow)(vArgs[0].ToPointer())
		this.DispImpl.ProtectedViewWindowDeactivate(p1)
		return win32.S_OK
	case 2911:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*Workbook)(vArgs[0].ToPointer())
		p2, _ := vArgs[1].ToBool()
		this.DispImpl.WorkbookAfterSave(p1, p2)
		return win32.S_OK
	case 2912:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*Workbook)(vArgs[0].ToPointer())
		p2 := (*Chart)(vArgs[1].ToPointer())
		this.DispImpl.WorkbookNewChart(p1, p2)
		return win32.S_OK
	}
	return win32.E_NOTIMPL
}

type AppEventsComObj struct {
	ole.IDispatchComObj
}

func NewAppEventsComObj(dispImpl AppEventsDispInterface, scoped bool) *AppEventsComObj {
	comObj := com.NewComObj[AppEventsComObj](
		&AppEventsImpl {DispImpl: dispImpl})
	if scoped {
		com.AddToScope(comObj)
	}
	return comObj
}

