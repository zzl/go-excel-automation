package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024410-0000-0000-C000-000000000046
var IID_OLEObjectEvents = syscall.GUID{0x00024410, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type OLEObjectEventsDispInterface interface {
	QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) 
	AddRef_() uint32
	Release_() uint32
	GetTypeInfoCount_(pctinfo *uint32) 
	GetTypeInfo_(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) 
	GetIDsOfNames_(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) 
	Invoke_(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) 
	GotFocus() 
	LostFocus() 
}

type OLEObjectEventsHandlers struct {
	QueryInterface_ func(riid *syscall.GUID, ppvObj unsafe.Pointer) 
	AddRef_ func() uint32
	Release_ func() uint32
	GetTypeInfoCount_ func(pctinfo *uint32) 
	GetTypeInfo_ func(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) 
	GetIDsOfNames_ func(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) 
	Invoke_ func(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) 
	GotFocus func() 
	LostFocus func() 
}

type OLEObjectEventsDispImpl struct {
	Handlers OLEObjectEventsHandlers
}

func (this *OLEObjectEventsDispImpl) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	if this.Handlers.QueryInterface_ != nil {
		this.Handlers.QueryInterface_(riid, ppvObj)
	}
}

func (this *OLEObjectEventsDispImpl) AddRef_() uint32 {
	if this.Handlers.AddRef_ != nil {
		return this.Handlers.AddRef_()
	}
	var ret uint32
	return ret
}

func (this *OLEObjectEventsDispImpl) Release_() uint32 {
	if this.Handlers.Release_ != nil {
		return this.Handlers.Release_()
	}
	var ret uint32
	return ret
}

func (this *OLEObjectEventsDispImpl) GetTypeInfoCount_(pctinfo *uint32) {
	if this.Handlers.GetTypeInfoCount_ != nil {
		this.Handlers.GetTypeInfoCount_(pctinfo)
	}
}

func (this *OLEObjectEventsDispImpl) GetTypeInfo_(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	if this.Handlers.GetTypeInfo_ != nil {
		this.Handlers.GetTypeInfo_(itinfo, lcid, pptinfo)
	}
}

func (this *OLEObjectEventsDispImpl) GetIDsOfNames_(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	if this.Handlers.GetIDsOfNames_ != nil {
		this.Handlers.GetIDsOfNames_(riid, rgszNames, cNames, lcid, rgdispid)
	}
}

func (this *OLEObjectEventsDispImpl) Invoke_(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	if this.Handlers.Invoke_ != nil {
		this.Handlers.Invoke_(dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr)
	}
}

func (this *OLEObjectEventsDispImpl) GotFocus() {
	if this.Handlers.GotFocus != nil {
		this.Handlers.GotFocus()
	}
}

func (this *OLEObjectEventsDispImpl) LostFocus() {
	if this.Handlers.LostFocus != nil {
		this.Handlers.LostFocus()
	}
}

type OLEObjectEventsImpl struct {
	ole.IDispatchImpl
	DispImpl OLEObjectEventsDispInterface
}

func (this *OLEObjectEventsImpl) QueryInterface(riid *syscall.GUID, ppvObject unsafe.Pointer) win32.HRESULT {
	if *riid == IID_OLEObjectEvents {
		this.AssignPpvObject(ppvObject)
		this.AddRef()
		return win32.S_OK
	}
	return this.IDispatchImpl.QueryInterface(riid, ppvObject)
}

func (this *OLEObjectEventsImpl) Invoke(dispIdMember int32, riid *syscall.GUID, lcid uint32,
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
	case 1541:
		this.DispImpl.GotFocus()
		return win32.S_OK
	case 1542:
		this.DispImpl.LostFocus()
		return win32.S_OK
	}
	return win32.E_NOTIMPL
}

type OLEObjectEventsComObj struct {
	ole.IDispatchComObj
}

func NewOLEObjectEventsComObj(dispImpl OLEObjectEventsDispInterface, scoped bool) *OLEObjectEventsComObj {
	comObj := com.NewComObj[OLEObjectEventsComObj](
		&OLEObjectEventsImpl {DispImpl: dispImpl})
	if scoped {
		com.AddToScope(comObj)
	}
	return comObj
}

