package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002443E-0000-0000-C000-000000000046
var IID_ConnectorFormat = syscall.GUID{0x0002443E, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ConnectorFormat struct {
	ole.OleClient
}

func NewConnectorFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *ConnectorFormat {
	 if pDisp == nil {
		return nil;
	}
	p := &ConnectorFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ConnectorFormatFromVar(v ole.Variant) *ConnectorFormat {
	return NewConnectorFormat(v.IDispatch(), false, false)
}

func (this *ConnectorFormat) IID() *syscall.GUID {
	return &IID_ConnectorFormat
}

func (this *ConnectorFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ConnectorFormat) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *ConnectorFormat) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ConnectorFormat) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ConnectorFormat) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *ConnectorFormat) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *ConnectorFormat) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *ConnectorFormat) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *ConnectorFormat) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ConnectorFormat) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ConnectorFormat) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ConnectorFormat) BeginConnect(connectedShape *Shape, connectionSite int32)  {
	retVal, _ := this.Call(0x000006d6, []interface{}{connectedShape, connectionSite})
	_= retVal
}

func (this *ConnectorFormat) BeginDisconnect()  {
	retVal, _ := this.Call(0x000006d9, nil)
	_= retVal
}

func (this *ConnectorFormat) EndConnect(connectedShape *Shape, connectionSite int32)  {
	retVal, _ := this.Call(0x000006da, []interface{}{connectedShape, connectionSite})
	_= retVal
}

func (this *ConnectorFormat) EndDisconnect()  {
	retVal, _ := this.Call(0x000006db, nil)
	_= retVal
}

func (this *ConnectorFormat) BeginConnected() int32 {
	retVal, _ := this.PropGet(0x000006dc, nil)
	return retVal.LValVal()
}

func (this *ConnectorFormat) BeginConnectedShape() *Shape {
	retVal, _ := this.PropGet(0x000006dd, nil)
	return NewShape(retVal.IDispatch(), false, true)
}

func (this *ConnectorFormat) BeginConnectionSite() int32 {
	retVal, _ := this.PropGet(0x000006de, nil)
	return retVal.LValVal()
}

func (this *ConnectorFormat) EndConnected() int32 {
	retVal, _ := this.PropGet(0x000006df, nil)
	return retVal.LValVal()
}

func (this *ConnectorFormat) EndConnectedShape() *Shape {
	retVal, _ := this.PropGet(0x000006e0, nil)
	return NewShape(retVal.IDispatch(), false, true)
}

func (this *ConnectorFormat) EndConnectionSite() int32 {
	retVal, _ := this.PropGet(0x000006e1, nil)
	return retVal.LValVal()
}

func (this *ConnectorFormat) Type() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *ConnectorFormat) SetType(rhs int32)  {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

