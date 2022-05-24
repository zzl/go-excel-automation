package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024412-0000-0000-C000-000000000046
var IID_WorkbookEvents = syscall.GUID{0x00024412, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type WorkbookEventsDispInterface interface {
	QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) 
	AddRef_() uint32
	Release_() uint32
	GetTypeInfoCount_(pctinfo *uint32) 
	GetTypeInfo_(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) 
	GetIDsOfNames_(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) 
	Invoke_(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) 
	Open() 
	Activate() 
	Deactivate() 
	BeforeClose(cancel *win32.VARIANT_BOOL) 
	BeforeSave(saveAsUI bool, cancel *win32.VARIANT_BOOL) 
	BeforePrint(cancel *win32.VARIANT_BOOL) 
	NewSheet(sh *win32.IUnknown) 
	AddinInstall() 
	AddinUninstall() 
	WindowResize(wn *Window) 
	WindowActivate(wn *Window) 
	WindowDeactivate(wn *Window) 
	SheetSelectionChange(sh *win32.IUnknown, target *Range) 
	SheetBeforeDoubleClick(sh *win32.IUnknown, target *Range, cancel *win32.VARIANT_BOOL) 
	SheetBeforeRightClick(sh *win32.IUnknown, target *Range, cancel *win32.VARIANT_BOOL) 
	SheetActivate(sh *win32.IUnknown) 
	SheetDeactivate(sh *win32.IUnknown) 
	SheetCalculate(sh *win32.IUnknown) 
	SheetChange(sh *win32.IUnknown, target *Range) 
	SheetFollowHyperlink(sh *win32.IUnknown, target *Hyperlink) 
	SheetPivotTableUpdate(sh *win32.IUnknown, target *PivotTable) 
	PivotTableCloseConnection(target *PivotTable) 
	PivotTableOpenConnection(target *PivotTable) 
	Sync(syncEventType int32) 
	BeforeXmlImport(map_ *XmlMap, url string, isRefresh bool, cancel *win32.VARIANT_BOOL) 
	AfterXmlImport(map_ *XmlMap, isRefresh bool, result int32) 
	BeforeXmlExport(map_ *XmlMap, url string, cancel *win32.VARIANT_BOOL) 
	AfterXmlExport(map_ *XmlMap, url string, result int32) 
	RowsetComplete(description string, sheet string, success bool) 
	SheetPivotTableAfterValueChange(sh *win32.IUnknown, targetPivotTable *PivotTable, targetRange *Range) 
	SheetPivotTableBeforeAllocateChanges(sh *win32.IUnknown, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL) 
	SheetPivotTableBeforeCommitChanges(sh *win32.IUnknown, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL) 
	SheetPivotTableBeforeDiscardChanges(sh *win32.IUnknown, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32) 
	SheetPivotTableChangeSync(sh *win32.IUnknown, target *PivotTable) 
	AfterSave(success bool) 
	NewChart(ch *Chart) 
}

type WorkbookEventsHandlers struct {
	QueryInterface_ func(riid *syscall.GUID, ppvObj unsafe.Pointer) 
	AddRef_ func() uint32
	Release_ func() uint32
	GetTypeInfoCount_ func(pctinfo *uint32) 
	GetTypeInfo_ func(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) 
	GetIDsOfNames_ func(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) 
	Invoke_ func(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) 
	Open func() 
	Activate func() 
	Deactivate func() 
	BeforeClose func(cancel *win32.VARIANT_BOOL) 
	BeforeSave func(saveAsUI bool, cancel *win32.VARIANT_BOOL) 
	BeforePrint func(cancel *win32.VARIANT_BOOL) 
	NewSheet func(sh *win32.IUnknown) 
	AddinInstall func() 
	AddinUninstall func() 
	WindowResize func(wn *Window) 
	WindowActivate func(wn *Window) 
	WindowDeactivate func(wn *Window) 
	SheetSelectionChange func(sh *win32.IUnknown, target *Range) 
	SheetBeforeDoubleClick func(sh *win32.IUnknown, target *Range, cancel *win32.VARIANT_BOOL) 
	SheetBeforeRightClick func(sh *win32.IUnknown, target *Range, cancel *win32.VARIANT_BOOL) 
	SheetActivate func(sh *win32.IUnknown) 
	SheetDeactivate func(sh *win32.IUnknown) 
	SheetCalculate func(sh *win32.IUnknown) 
	SheetChange func(sh *win32.IUnknown, target *Range) 
	SheetFollowHyperlink func(sh *win32.IUnknown, target *Hyperlink) 
	SheetPivotTableUpdate func(sh *win32.IUnknown, target *PivotTable) 
	PivotTableCloseConnection func(target *PivotTable) 
	PivotTableOpenConnection func(target *PivotTable) 
	Sync func(syncEventType int32) 
	BeforeXmlImport func(map_ *XmlMap, url string, isRefresh bool, cancel *win32.VARIANT_BOOL) 
	AfterXmlImport func(map_ *XmlMap, isRefresh bool, result int32) 
	BeforeXmlExport func(map_ *XmlMap, url string, cancel *win32.VARIANT_BOOL) 
	AfterXmlExport func(map_ *XmlMap, url string, result int32) 
	RowsetComplete func(description string, sheet string, success bool) 
	SheetPivotTableAfterValueChange func(sh *win32.IUnknown, targetPivotTable *PivotTable, targetRange *Range) 
	SheetPivotTableBeforeAllocateChanges func(sh *win32.IUnknown, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL) 
	SheetPivotTableBeforeCommitChanges func(sh *win32.IUnknown, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL) 
	SheetPivotTableBeforeDiscardChanges func(sh *win32.IUnknown, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32) 
	SheetPivotTableChangeSync func(sh *win32.IUnknown, target *PivotTable) 
	AfterSave func(success bool) 
	NewChart func(ch *Chart) 
}

type WorkbookEventsDispImpl struct {
	Handlers WorkbookEventsHandlers
}

func (this *WorkbookEventsDispImpl) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	if this.Handlers.QueryInterface_ != nil {
		this.Handlers.QueryInterface_(riid, ppvObj)
	}
}

func (this *WorkbookEventsDispImpl) AddRef_() uint32 {
	if this.Handlers.AddRef_ != nil {
		return this.Handlers.AddRef_()
	}
	var ret uint32
	return ret
}

func (this *WorkbookEventsDispImpl) Release_() uint32 {
	if this.Handlers.Release_ != nil {
		return this.Handlers.Release_()
	}
	var ret uint32
	return ret
}

func (this *WorkbookEventsDispImpl) GetTypeInfoCount_(pctinfo *uint32) {
	if this.Handlers.GetTypeInfoCount_ != nil {
		this.Handlers.GetTypeInfoCount_(pctinfo)
	}
}

func (this *WorkbookEventsDispImpl) GetTypeInfo_(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	if this.Handlers.GetTypeInfo_ != nil {
		this.Handlers.GetTypeInfo_(itinfo, lcid, pptinfo)
	}
}

func (this *WorkbookEventsDispImpl) GetIDsOfNames_(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	if this.Handlers.GetIDsOfNames_ != nil {
		this.Handlers.GetIDsOfNames_(riid, rgszNames, cNames, lcid, rgdispid)
	}
}

func (this *WorkbookEventsDispImpl) Invoke_(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	if this.Handlers.Invoke_ != nil {
		this.Handlers.Invoke_(dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr)
	}
}

func (this *WorkbookEventsDispImpl) Open() {
	if this.Handlers.Open != nil {
		this.Handlers.Open()
	}
}

func (this *WorkbookEventsDispImpl) Activate() {
	if this.Handlers.Activate != nil {
		this.Handlers.Activate()
	}
}

func (this *WorkbookEventsDispImpl) Deactivate() {
	if this.Handlers.Deactivate != nil {
		this.Handlers.Deactivate()
	}
}

func (this *WorkbookEventsDispImpl) BeforeClose(cancel *win32.VARIANT_BOOL) {
	if this.Handlers.BeforeClose != nil {
		this.Handlers.BeforeClose(cancel)
	}
}

func (this *WorkbookEventsDispImpl) BeforeSave(saveAsUI bool, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.BeforeSave != nil {
		this.Handlers.BeforeSave(saveAsUI, cancel)
	}
}

func (this *WorkbookEventsDispImpl) BeforePrint(cancel *win32.VARIANT_BOOL) {
	if this.Handlers.BeforePrint != nil {
		this.Handlers.BeforePrint(cancel)
	}
}

func (this *WorkbookEventsDispImpl) NewSheet(sh *win32.IUnknown) {
	if this.Handlers.NewSheet != nil {
		this.Handlers.NewSheet(sh)
	}
}

func (this *WorkbookEventsDispImpl) AddinInstall() {
	if this.Handlers.AddinInstall != nil {
		this.Handlers.AddinInstall()
	}
}

func (this *WorkbookEventsDispImpl) AddinUninstall() {
	if this.Handlers.AddinUninstall != nil {
		this.Handlers.AddinUninstall()
	}
}

func (this *WorkbookEventsDispImpl) WindowResize(wn *Window) {
	if this.Handlers.WindowResize != nil {
		this.Handlers.WindowResize(wn)
	}
}

func (this *WorkbookEventsDispImpl) WindowActivate(wn *Window) {
	if this.Handlers.WindowActivate != nil {
		this.Handlers.WindowActivate(wn)
	}
}

func (this *WorkbookEventsDispImpl) WindowDeactivate(wn *Window) {
	if this.Handlers.WindowDeactivate != nil {
		this.Handlers.WindowDeactivate(wn)
	}
}

func (this *WorkbookEventsDispImpl) SheetSelectionChange(sh *win32.IUnknown, target *Range) {
	if this.Handlers.SheetSelectionChange != nil {
		this.Handlers.SheetSelectionChange(sh, target)
	}
}

func (this *WorkbookEventsDispImpl) SheetBeforeDoubleClick(sh *win32.IUnknown, target *Range, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.SheetBeforeDoubleClick != nil {
		this.Handlers.SheetBeforeDoubleClick(sh, target, cancel)
	}
}

func (this *WorkbookEventsDispImpl) SheetBeforeRightClick(sh *win32.IUnknown, target *Range, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.SheetBeforeRightClick != nil {
		this.Handlers.SheetBeforeRightClick(sh, target, cancel)
	}
}

func (this *WorkbookEventsDispImpl) SheetActivate(sh *win32.IUnknown) {
	if this.Handlers.SheetActivate != nil {
		this.Handlers.SheetActivate(sh)
	}
}

func (this *WorkbookEventsDispImpl) SheetDeactivate(sh *win32.IUnknown) {
	if this.Handlers.SheetDeactivate != nil {
		this.Handlers.SheetDeactivate(sh)
	}
}

func (this *WorkbookEventsDispImpl) SheetCalculate(sh *win32.IUnknown) {
	if this.Handlers.SheetCalculate != nil {
		this.Handlers.SheetCalculate(sh)
	}
}

func (this *WorkbookEventsDispImpl) SheetChange(sh *win32.IUnknown, target *Range) {
	if this.Handlers.SheetChange != nil {
		this.Handlers.SheetChange(sh, target)
	}
}

func (this *WorkbookEventsDispImpl) SheetFollowHyperlink(sh *win32.IUnknown, target *Hyperlink) {
	if this.Handlers.SheetFollowHyperlink != nil {
		this.Handlers.SheetFollowHyperlink(sh, target)
	}
}

func (this *WorkbookEventsDispImpl) SheetPivotTableUpdate(sh *win32.IUnknown, target *PivotTable) {
	if this.Handlers.SheetPivotTableUpdate != nil {
		this.Handlers.SheetPivotTableUpdate(sh, target)
	}
}

func (this *WorkbookEventsDispImpl) PivotTableCloseConnection(target *PivotTable) {
	if this.Handlers.PivotTableCloseConnection != nil {
		this.Handlers.PivotTableCloseConnection(target)
	}
}

func (this *WorkbookEventsDispImpl) PivotTableOpenConnection(target *PivotTable) {
	if this.Handlers.PivotTableOpenConnection != nil {
		this.Handlers.PivotTableOpenConnection(target)
	}
}

func (this *WorkbookEventsDispImpl) Sync(syncEventType int32) {
	if this.Handlers.Sync != nil {
		this.Handlers.Sync(syncEventType)
	}
}

func (this *WorkbookEventsDispImpl) BeforeXmlImport(map_ *XmlMap, url string, isRefresh bool, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.BeforeXmlImport != nil {
		this.Handlers.BeforeXmlImport(map_, url, isRefresh, cancel)
	}
}

func (this *WorkbookEventsDispImpl) AfterXmlImport(map_ *XmlMap, isRefresh bool, result int32) {
	if this.Handlers.AfterXmlImport != nil {
		this.Handlers.AfterXmlImport(map_, isRefresh, result)
	}
}

func (this *WorkbookEventsDispImpl) BeforeXmlExport(map_ *XmlMap, url string, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.BeforeXmlExport != nil {
		this.Handlers.BeforeXmlExport(map_, url, cancel)
	}
}

func (this *WorkbookEventsDispImpl) AfterXmlExport(map_ *XmlMap, url string, result int32) {
	if this.Handlers.AfterXmlExport != nil {
		this.Handlers.AfterXmlExport(map_, url, result)
	}
}

func (this *WorkbookEventsDispImpl) RowsetComplete(description string, sheet string, success bool) {
	if this.Handlers.RowsetComplete != nil {
		this.Handlers.RowsetComplete(description, sheet, success)
	}
}

func (this *WorkbookEventsDispImpl) SheetPivotTableAfterValueChange(sh *win32.IUnknown, targetPivotTable *PivotTable, targetRange *Range) {
	if this.Handlers.SheetPivotTableAfterValueChange != nil {
		this.Handlers.SheetPivotTableAfterValueChange(sh, targetPivotTable, targetRange)
	}
}

func (this *WorkbookEventsDispImpl) SheetPivotTableBeforeAllocateChanges(sh *win32.IUnknown, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.SheetPivotTableBeforeAllocateChanges != nil {
		this.Handlers.SheetPivotTableBeforeAllocateChanges(sh, targetPivotTable, valueChangeStart, valueChangeEnd, cancel)
	}
}

func (this *WorkbookEventsDispImpl) SheetPivotTableBeforeCommitChanges(sh *win32.IUnknown, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.SheetPivotTableBeforeCommitChanges != nil {
		this.Handlers.SheetPivotTableBeforeCommitChanges(sh, targetPivotTable, valueChangeStart, valueChangeEnd, cancel)
	}
}

func (this *WorkbookEventsDispImpl) SheetPivotTableBeforeDiscardChanges(sh *win32.IUnknown, targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32) {
	if this.Handlers.SheetPivotTableBeforeDiscardChanges != nil {
		this.Handlers.SheetPivotTableBeforeDiscardChanges(sh, targetPivotTable, valueChangeStart, valueChangeEnd)
	}
}

func (this *WorkbookEventsDispImpl) SheetPivotTableChangeSync(sh *win32.IUnknown, target *PivotTable) {
	if this.Handlers.SheetPivotTableChangeSync != nil {
		this.Handlers.SheetPivotTableChangeSync(sh, target)
	}
}

func (this *WorkbookEventsDispImpl) AfterSave(success bool) {
	if this.Handlers.AfterSave != nil {
		this.Handlers.AfterSave(success)
	}
}

func (this *WorkbookEventsDispImpl) NewChart(ch *Chart) {
	if this.Handlers.NewChart != nil {
		this.Handlers.NewChart(ch)
	}
}

type WorkbookEventsImpl struct {
	ole.IDispatchImpl
	DispImpl WorkbookEventsDispInterface
}

func (this *WorkbookEventsImpl) QueryInterface(riid *syscall.GUID, ppvObject unsafe.Pointer) win32.HRESULT {
	if *riid == IID_WorkbookEvents {
		this.AssignPpvObject(ppvObject)
		this.AddRef()
		return win32.S_OK
	}
	return this.IDispatchImpl.QueryInterface(riid, ppvObject)
}

func (this *WorkbookEventsImpl) Invoke(dispIdMember int32, riid *syscall.GUID, lcid uint32,
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
	case 1923:
		this.DispImpl.Open()
		return win32.S_OK
	case 304:
		this.DispImpl.Activate()
		return win32.S_OK
	case 1530:
		this.DispImpl.Deactivate()
		return win32.S_OK
	case 1546:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*win32.VARIANT_BOOL)(vArgs[0].ToPointer())
		this.DispImpl.BeforeClose(p1)
		return win32.S_OK
	case 1547:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1, _ := vArgs[0].ToBool()
		p2 := (*win32.VARIANT_BOOL)(vArgs[1].ToPointer())
		this.DispImpl.BeforeSave(p1, p2)
		return win32.S_OK
	case 1549:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*win32.VARIANT_BOOL)(vArgs[0].ToPointer())
		this.DispImpl.BeforePrint(p1)
		return win32.S_OK
	case 1550:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*win32.IUnknown)(vArgs[0].ToPointer())
		this.DispImpl.NewSheet(p1)
		return win32.S_OK
	case 1552:
		this.DispImpl.AddinInstall()
		return win32.S_OK
	case 1553:
		this.DispImpl.AddinUninstall()
		return win32.S_OK
	case 1554:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*Window)(vArgs[0].ToPointer())
		this.DispImpl.WindowResize(p1)
		return win32.S_OK
	case 1556:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*Window)(vArgs[0].ToPointer())
		this.DispImpl.WindowActivate(p1)
		return win32.S_OK
	case 1557:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*Window)(vArgs[0].ToPointer())
		this.DispImpl.WindowDeactivate(p1)
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
	case 2158:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*PivotTable)(vArgs[0].ToPointer())
		this.DispImpl.PivotTableCloseConnection(p1)
		return win32.S_OK
	case 2159:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*PivotTable)(vArgs[0].ToPointer())
		this.DispImpl.PivotTableOpenConnection(p1)
		return win32.S_OK
	case 2266:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1, _ := vArgs[0].ToInt32()
		this.DispImpl.Sync(p1)
		return win32.S_OK
	case 2283:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 4)
		p1 := (*XmlMap)(vArgs[0].ToPointer())
		p2, _ := vArgs[1].ToString()
		p3, _ := vArgs[2].ToBool()
		p4 := (*win32.VARIANT_BOOL)(vArgs[3].ToPointer())
		this.DispImpl.BeforeXmlImport(p1, p2, p3, p4)
		return win32.S_OK
	case 2285:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 3)
		p1 := (*XmlMap)(vArgs[0].ToPointer())
		p2, _ := vArgs[1].ToBool()
		p3, _ := vArgs[2].ToInt32()
		this.DispImpl.AfterXmlImport(p1, p2, p3)
		return win32.S_OK
	case 2287:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 3)
		p1 := (*XmlMap)(vArgs[0].ToPointer())
		p2, _ := vArgs[1].ToString()
		p3 := (*win32.VARIANT_BOOL)(vArgs[2].ToPointer())
		this.DispImpl.BeforeXmlExport(p1, p2, p3)
		return win32.S_OK
	case 2288:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 3)
		p1 := (*XmlMap)(vArgs[0].ToPointer())
		p2, _ := vArgs[1].ToString()
		p3, _ := vArgs[2].ToInt32()
		this.DispImpl.AfterXmlExport(p1, p2, p3)
		return win32.S_OK
	case 2610:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 3)
		p1, _ := vArgs[0].ToString()
		p2, _ := vArgs[1].ToString()
		p3, _ := vArgs[2].ToBool()
		this.DispImpl.RowsetComplete(p1, p2, p3)
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
	case 2899:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*win32.IUnknown)(vArgs[0].ToPointer())
		p2 := (*PivotTable)(vArgs[1].ToPointer())
		this.DispImpl.SheetPivotTableChangeSync(p1, p2)
		return win32.S_OK
	case 2900:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1, _ := vArgs[0].ToBool()
		this.DispImpl.AfterSave(p1)
		return win32.S_OK
	case 2901:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*Chart)(vArgs[0].ToPointer())
		this.DispImpl.NewChart(p1)
		return win32.S_OK
	}
	return win32.E_NOTIMPL
}

type WorkbookEventsComObj struct {
	ole.IDispatchComObj
}

func NewWorkbookEventsComObj(dispImpl WorkbookEventsDispInterface, scoped bool) *WorkbookEventsComObj {
	comObj := com.NewComObj[WorkbookEventsComObj](
		&WorkbookEventsImpl {DispImpl: dispImpl})
	if scoped {
		com.AddToScope(comObj)
	}
	return comObj
}

