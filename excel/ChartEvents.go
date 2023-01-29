package excel

import (
	"github.com/zzl/go-win32api/v2/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002440F-0000-0000-C000-000000000046
var IID_ChartEvents = syscall.GUID{0x0002440F, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ChartEventsDispInterface interface {
	QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)
	AddRef_() uint32
	Release_() uint32
	GetTypeInfoCount_(pctinfo *uint32)
	GetTypeInfo_(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)
	GetIDsOfNames_(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)
	Invoke_(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)
	Activate()
	Deactivate()
	Resize()
	MouseDown(button int32, shift int32, x int32, y int32)
	MouseUp(button int32, shift int32, x int32, y int32)
	MouseMove(button int32, shift int32, x int32, y int32)
	BeforeRightClick(cancel *win32.VARIANT_BOOL)
	DragPlot()
	DragOver()
	BeforeDoubleClick(elementID int32, arg1 int32, arg2 int32, cancel *win32.VARIANT_BOOL)
	Select(elementID int32, arg1 int32, arg2 int32)
	SeriesChange(seriesIndex int32, pointIndex int32)
	Calculate()
}

type ChartEventsHandlers struct {
	QueryInterface_   func(riid *syscall.GUID, ppvObj unsafe.Pointer)
	AddRef_           func() uint32
	Release_          func() uint32
	GetTypeInfoCount_ func(pctinfo *uint32)
	GetTypeInfo_      func(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)
	GetIDsOfNames_    func(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)
	Invoke_           func(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)
	Activate          func()
	Deactivate        func()
	Resize            func()
	MouseDown         func(button int32, shift int32, x int32, y int32)
	MouseUp           func(button int32, shift int32, x int32, y int32)
	MouseMove         func(button int32, shift int32, x int32, y int32)
	BeforeRightClick  func(cancel *win32.VARIANT_BOOL)
	DragPlot          func()
	DragOver          func()
	BeforeDoubleClick func(elementID int32, arg1 int32, arg2 int32, cancel *win32.VARIANT_BOOL)
	Select            func(elementID int32, arg1 int32, arg2 int32)
	SeriesChange      func(seriesIndex int32, pointIndex int32)
	Calculate         func()
}

type ChartEventsDispImpl struct {
	Handlers ChartEventsHandlers
}

func (this *ChartEventsDispImpl) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	if this.Handlers.QueryInterface_ != nil {
		this.Handlers.QueryInterface_(riid, ppvObj)
	}
}

func (this *ChartEventsDispImpl) AddRef_() uint32 {
	if this.Handlers.AddRef_ != nil {
		return this.Handlers.AddRef_()
	}
	var ret uint32
	return ret
}

func (this *ChartEventsDispImpl) Release_() uint32 {
	if this.Handlers.Release_ != nil {
		return this.Handlers.Release_()
	}
	var ret uint32
	return ret
}

func (this *ChartEventsDispImpl) GetTypeInfoCount_(pctinfo *uint32) {
	if this.Handlers.GetTypeInfoCount_ != nil {
		this.Handlers.GetTypeInfoCount_(pctinfo)
	}
}

func (this *ChartEventsDispImpl) GetTypeInfo_(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	if this.Handlers.GetTypeInfo_ != nil {
		this.Handlers.GetTypeInfo_(itinfo, lcid, pptinfo)
	}
}

func (this *ChartEventsDispImpl) GetIDsOfNames_(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	if this.Handlers.GetIDsOfNames_ != nil {
		this.Handlers.GetIDsOfNames_(riid, rgszNames, cNames, lcid, rgdispid)
	}
}

func (this *ChartEventsDispImpl) Invoke_(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	if this.Handlers.Invoke_ != nil {
		this.Handlers.Invoke_(dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr)
	}
}

func (this *ChartEventsDispImpl) Activate() {
	if this.Handlers.Activate != nil {
		this.Handlers.Activate()
	}
}

func (this *ChartEventsDispImpl) Deactivate() {
	if this.Handlers.Deactivate != nil {
		this.Handlers.Deactivate()
	}
}

func (this *ChartEventsDispImpl) Resize() {
	if this.Handlers.Resize != nil {
		this.Handlers.Resize()
	}
}

func (this *ChartEventsDispImpl) MouseDown(button int32, shift int32, x int32, y int32) {
	if this.Handlers.MouseDown != nil {
		this.Handlers.MouseDown(button, shift, x, y)
	}
}

func (this *ChartEventsDispImpl) MouseUp(button int32, shift int32, x int32, y int32) {
	if this.Handlers.MouseUp != nil {
		this.Handlers.MouseUp(button, shift, x, y)
	}
}

func (this *ChartEventsDispImpl) MouseMove(button int32, shift int32, x int32, y int32) {
	if this.Handlers.MouseMove != nil {
		this.Handlers.MouseMove(button, shift, x, y)
	}
}

func (this *ChartEventsDispImpl) BeforeRightClick(cancel *win32.VARIANT_BOOL) {
	if this.Handlers.BeforeRightClick != nil {
		this.Handlers.BeforeRightClick(cancel)
	}
}

func (this *ChartEventsDispImpl) DragPlot() {
	if this.Handlers.DragPlot != nil {
		this.Handlers.DragPlot()
	}
}

func (this *ChartEventsDispImpl) DragOver() {
	if this.Handlers.DragOver != nil {
		this.Handlers.DragOver()
	}
}

func (this *ChartEventsDispImpl) BeforeDoubleClick(elementID int32, arg1 int32, arg2 int32, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.BeforeDoubleClick != nil {
		this.Handlers.BeforeDoubleClick(elementID, arg1, arg2, cancel)
	}
}

func (this *ChartEventsDispImpl) Select(elementID int32, arg1 int32, arg2 int32) {
	if this.Handlers.Select != nil {
		this.Handlers.Select(elementID, arg1, arg2)
	}
}

func (this *ChartEventsDispImpl) SeriesChange(seriesIndex int32, pointIndex int32) {
	if this.Handlers.SeriesChange != nil {
		this.Handlers.SeriesChange(seriesIndex, pointIndex)
	}
}

func (this *ChartEventsDispImpl) Calculate() {
	if this.Handlers.Calculate != nil {
		this.Handlers.Calculate()
	}
}

type ChartEventsImpl struct {
	ole.IDispatchImpl
	DispImpl ChartEventsDispInterface
}

func (this *ChartEventsImpl) QueryInterface(riid *syscall.GUID, ppvObject unsafe.Pointer) win32.HRESULT {
	if *riid == IID_ChartEvents {
		this.AssignPpvObject(ppvObject)
		this.AddRef()
		return win32.S_OK
	}
	return this.IDispatchImpl.QueryInterface(riid, ppvObject)
}

func (this *ChartEventsImpl) Invoke(dispIdMember int32, riid *syscall.GUID, lcid uint32,
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
	case 304:
		this.DispImpl.Activate()
		return win32.S_OK
	case 1530:
		this.DispImpl.Deactivate()
		return win32.S_OK
	case 256:
		this.DispImpl.Resize()
		return win32.S_OK
	case 1531:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 4)
		p1, _ := vArgs[0].ToInt32()
		p2, _ := vArgs[1].ToInt32()
		p3, _ := vArgs[2].ToInt32()
		p4, _ := vArgs[3].ToInt32()
		this.DispImpl.MouseDown(p1, p2, p3, p4)
		return win32.S_OK
	case 1532:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 4)
		p1, _ := vArgs[0].ToInt32()
		p2, _ := vArgs[1].ToInt32()
		p3, _ := vArgs[2].ToInt32()
		p4, _ := vArgs[3].ToInt32()
		this.DispImpl.MouseUp(p1, p2, p3, p4)
		return win32.S_OK
	case 1533:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 4)
		p1, _ := vArgs[0].ToInt32()
		p2, _ := vArgs[1].ToInt32()
		p3, _ := vArgs[2].ToInt32()
		p4, _ := vArgs[3].ToInt32()
		this.DispImpl.MouseMove(p1, p2, p3, p4)
		return win32.S_OK
	case 1534:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*win32.VARIANT_BOOL)(vArgs[0].ToPointer())
		this.DispImpl.BeforeRightClick(p1)
		return win32.S_OK
	case 1535:
		this.DispImpl.DragPlot()
		return win32.S_OK
	case 1536:
		this.DispImpl.DragOver()
		return win32.S_OK
	case 1537:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 4)
		p1, _ := vArgs[0].ToInt32()
		p2, _ := vArgs[1].ToInt32()
		p3, _ := vArgs[2].ToInt32()
		p4 := (*win32.VARIANT_BOOL)(vArgs[3].ToPointer())
		this.DispImpl.BeforeDoubleClick(p1, p2, p3, p4)
		return win32.S_OK
	case 235:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 3)
		p1, _ := vArgs[0].ToInt32()
		p2, _ := vArgs[1].ToInt32()
		p3, _ := vArgs[2].ToInt32()
		this.DispImpl.Select(p1, p2, p3)
		return win32.S_OK
	case 1538:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1, _ := vArgs[0].ToInt32()
		p2, _ := vArgs[1].ToInt32()
		this.DispImpl.SeriesChange(p1, p2)
		return win32.S_OK
	case 279:
		this.DispImpl.Calculate()
		return win32.S_OK
	}
	return win32.E_NOTIMPL
}

type ChartEventsComObj struct {
	ole.IDispatchComObj
}

func NewChartEventsComObj(dispImpl ChartEventsDispInterface, scoped bool) *ChartEventsComObj {
	comObj := com.NewComObj[ChartEventsComObj](
		&ChartEventsImpl{DispImpl: dispImpl})
	if scoped {
		com.AddToScope(comObj)
	}
	return comObj
}

