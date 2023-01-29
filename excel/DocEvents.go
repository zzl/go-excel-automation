package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00024411-0000-0000-C000-000000000046
var IID_DocEvents = syscall.GUID{0x00024411, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type DocEventsDispInterface interface {
	QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)
	AddRef_() uint32
	Release_() uint32
	GetTypeInfoCount_(pctinfo *uint32)
	GetTypeInfo_(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)
	GetIDsOfNames_(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)
	Invoke_(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)
	SelectionChange(target *Range)
	BeforeDoubleClick(target *Range, cancel *win32.VARIANT_BOOL)
	BeforeRightClick(target *Range, cancel *win32.VARIANT_BOOL)
	Activate()
	Deactivate()
	Calculate()
	Change(target *Range)
	FollowHyperlink(target *Hyperlink)
	PivotTableUpdate(target *PivotTable)
	PivotTableAfterValueChange(targetPivotTable *PivotTable, targetRange *Range)
	PivotTableBeforeAllocateChanges(targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL)
	PivotTableBeforeCommitChanges(targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL)
	PivotTableBeforeDiscardChanges(targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32)
	PivotTableChangeSync(target *PivotTable)
}

type DocEventsHandlers struct {
	QueryInterface_                 func(riid *syscall.GUID, ppvObj unsafe.Pointer)
	AddRef_                         func() uint32
	Release_                        func() uint32
	GetTypeInfoCount_               func(pctinfo *uint32)
	GetTypeInfo_                    func(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)
	GetIDsOfNames_                  func(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)
	Invoke_                         func(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)
	SelectionChange                 func(target *Range)
	BeforeDoubleClick               func(target *Range, cancel *win32.VARIANT_BOOL)
	BeforeRightClick                func(target *Range, cancel *win32.VARIANT_BOOL)
	Activate                        func()
	Deactivate                      func()
	Calculate                       func()
	Change                          func(target *Range)
	FollowHyperlink                 func(target *Hyperlink)
	PivotTableUpdate                func(target *PivotTable)
	PivotTableAfterValueChange      func(targetPivotTable *PivotTable, targetRange *Range)
	PivotTableBeforeAllocateChanges func(targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL)
	PivotTableBeforeCommitChanges   func(targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL)
	PivotTableBeforeDiscardChanges  func(targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32)
	PivotTableChangeSync            func(target *PivotTable)
}

type DocEventsDispImpl struct {
	Handlers DocEventsHandlers
}

func (this *DocEventsDispImpl) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	if this.Handlers.QueryInterface_ != nil {
		this.Handlers.QueryInterface_(riid, ppvObj)
	}
}

func (this *DocEventsDispImpl) AddRef_() uint32 {
	if this.Handlers.AddRef_ != nil {
		return this.Handlers.AddRef_()
	}
	var ret uint32
	return ret
}

func (this *DocEventsDispImpl) Release_() uint32 {
	if this.Handlers.Release_ != nil {
		return this.Handlers.Release_()
	}
	var ret uint32
	return ret
}

func (this *DocEventsDispImpl) GetTypeInfoCount_(pctinfo *uint32) {
	if this.Handlers.GetTypeInfoCount_ != nil {
		this.Handlers.GetTypeInfoCount_(pctinfo)
	}
}

func (this *DocEventsDispImpl) GetTypeInfo_(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	if this.Handlers.GetTypeInfo_ != nil {
		this.Handlers.GetTypeInfo_(itinfo, lcid, pptinfo)
	}
}

func (this *DocEventsDispImpl) GetIDsOfNames_(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	if this.Handlers.GetIDsOfNames_ != nil {
		this.Handlers.GetIDsOfNames_(riid, rgszNames, cNames, lcid, rgdispid)
	}
}

func (this *DocEventsDispImpl) Invoke_(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	if this.Handlers.Invoke_ != nil {
		this.Handlers.Invoke_(dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr)
	}
}

func (this *DocEventsDispImpl) SelectionChange(target *Range) {
	if this.Handlers.SelectionChange != nil {
		this.Handlers.SelectionChange(target)
	}
}

func (this *DocEventsDispImpl) BeforeDoubleClick(target *Range, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.BeforeDoubleClick != nil {
		this.Handlers.BeforeDoubleClick(target, cancel)
	}
}

func (this *DocEventsDispImpl) BeforeRightClick(target *Range, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.BeforeRightClick != nil {
		this.Handlers.BeforeRightClick(target, cancel)
	}
}

func (this *DocEventsDispImpl) Activate() {
	if this.Handlers.Activate != nil {
		this.Handlers.Activate()
	}
}

func (this *DocEventsDispImpl) Deactivate() {
	if this.Handlers.Deactivate != nil {
		this.Handlers.Deactivate()
	}
}

func (this *DocEventsDispImpl) Calculate() {
	if this.Handlers.Calculate != nil {
		this.Handlers.Calculate()
	}
}

func (this *DocEventsDispImpl) Change(target *Range) {
	if this.Handlers.Change != nil {
		this.Handlers.Change(target)
	}
}

func (this *DocEventsDispImpl) FollowHyperlink(target *Hyperlink) {
	if this.Handlers.FollowHyperlink != nil {
		this.Handlers.FollowHyperlink(target)
	}
}

func (this *DocEventsDispImpl) PivotTableUpdate(target *PivotTable) {
	if this.Handlers.PivotTableUpdate != nil {
		this.Handlers.PivotTableUpdate(target)
	}
}

func (this *DocEventsDispImpl) PivotTableAfterValueChange(targetPivotTable *PivotTable, targetRange *Range) {
	if this.Handlers.PivotTableAfterValueChange != nil {
		this.Handlers.PivotTableAfterValueChange(targetPivotTable, targetRange)
	}
}

func (this *DocEventsDispImpl) PivotTableBeforeAllocateChanges(targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.PivotTableBeforeAllocateChanges != nil {
		this.Handlers.PivotTableBeforeAllocateChanges(targetPivotTable, valueChangeStart, valueChangeEnd, cancel)
	}
}

func (this *DocEventsDispImpl) PivotTableBeforeCommitChanges(targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.PivotTableBeforeCommitChanges != nil {
		this.Handlers.PivotTableBeforeCommitChanges(targetPivotTable, valueChangeStart, valueChangeEnd, cancel)
	}
}

func (this *DocEventsDispImpl) PivotTableBeforeDiscardChanges(targetPivotTable *PivotTable, valueChangeStart int32, valueChangeEnd int32) {
	if this.Handlers.PivotTableBeforeDiscardChanges != nil {
		this.Handlers.PivotTableBeforeDiscardChanges(targetPivotTable, valueChangeStart, valueChangeEnd)
	}
}

func (this *DocEventsDispImpl) PivotTableChangeSync(target *PivotTable) {
	if this.Handlers.PivotTableChangeSync != nil {
		this.Handlers.PivotTableChangeSync(target)
	}
}

type DocEventsImpl struct {
	ole.IDispatchImpl
	DispImpl DocEventsDispInterface
}

func (this *DocEventsImpl) QueryInterface(riid *syscall.GUID, ppvObject unsafe.Pointer) win32.HRESULT {
	if *riid == IID_DocEvents {
		this.AssignPpvObject(ppvObject)
		this.AddRef()
		return win32.S_OK
	}
	return this.IDispatchImpl.QueryInterface(riid, ppvObject)
}

func (this *DocEventsImpl) Invoke(dispIdMember int32, riid *syscall.GUID, lcid uint32,
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
	case 1543:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*Range)(vArgs[0].ToPointer())
		this.DispImpl.SelectionChange(p1)
		return win32.S_OK
	case 1537:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*Range)(vArgs[0].ToPointer())
		p2 := (*win32.VARIANT_BOOL)(vArgs[1].ToPointer())
		this.DispImpl.BeforeDoubleClick(p1, p2)
		return win32.S_OK
	case 1534:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*Range)(vArgs[0].ToPointer())
		p2 := (*win32.VARIANT_BOOL)(vArgs[1].ToPointer())
		this.DispImpl.BeforeRightClick(p1, p2)
		return win32.S_OK
	case 304:
		this.DispImpl.Activate()
		return win32.S_OK
	case 1530:
		this.DispImpl.Deactivate()
		return win32.S_OK
	case 279:
		this.DispImpl.Calculate()
		return win32.S_OK
	case 1545:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*Range)(vArgs[0].ToPointer())
		this.DispImpl.Change(p1)
		return win32.S_OK
	case 1470:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*Hyperlink)(vArgs[0].ToPointer())
		this.DispImpl.FollowHyperlink(p1)
		return win32.S_OK
	case 2156:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*PivotTable)(vArgs[0].ToPointer())
		this.DispImpl.PivotTableUpdate(p1)
		return win32.S_OK
	case 2886:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*PivotTable)(vArgs[0].ToPointer())
		p2 := (*Range)(vArgs[1].ToPointer())
		this.DispImpl.PivotTableAfterValueChange(p1, p2)
		return win32.S_OK
	case 2889:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 4)
		p1 := (*PivotTable)(vArgs[0].ToPointer())
		p2, _ := vArgs[1].ToInt32()
		p3, _ := vArgs[2].ToInt32()
		p4 := (*win32.VARIANT_BOOL)(vArgs[3].ToPointer())
		this.DispImpl.PivotTableBeforeAllocateChanges(p1, p2, p3, p4)
		return win32.S_OK
	case 2892:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 4)
		p1 := (*PivotTable)(vArgs[0].ToPointer())
		p2, _ := vArgs[1].ToInt32()
		p3, _ := vArgs[2].ToInt32()
		p4 := (*win32.VARIANT_BOOL)(vArgs[3].ToPointer())
		this.DispImpl.PivotTableBeforeCommitChanges(p1, p2, p3, p4)
		return win32.S_OK
	case 2893:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 3)
		p1 := (*PivotTable)(vArgs[0].ToPointer())
		p2, _ := vArgs[1].ToInt32()
		p3, _ := vArgs[2].ToInt32()
		this.DispImpl.PivotTableBeforeDiscardChanges(p1, p2, p3)
		return win32.S_OK
	case 2894:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*PivotTable)(vArgs[0].ToPointer())
		this.DispImpl.PivotTableChangeSync(p1)
		return win32.S_OK
	}
	return win32.E_NOTIMPL
}

type DocEventsComObj struct {
	ole.IDispatchComObj
}

func NewDocEventsComObj(dispImpl DocEventsDispInterface, scoped bool) *DocEventsComObj {
	comObj := com.NewComObj[DocEventsComObj](
		&DocEventsImpl{DispImpl: dispImpl})
	if scoped {
		com.AddToScope(comObj)
	}
	return comObj
}
