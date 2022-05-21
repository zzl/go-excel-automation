package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024485-0000-0000-C000-000000000046
var IID_WorkbookConnection = syscall.GUID{0x00024485, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type WorkbookConnection struct {
	ole.OleClient
}

func NewWorkbookConnection(pDisp *win32.IDispatch, addRef bool, scoped bool) *WorkbookConnection {
	p := &WorkbookConnection{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func WorkbookConnectionFromVar(v ole.Variant) *WorkbookConnection {
	return NewWorkbookConnection(v.PdispValVal(), false, false)
}

func (this *WorkbookConnection) IID() *syscall.GUID {
	return &IID_WorkbookConnection
}

func (this *WorkbookConnection) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *WorkbookConnection) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *WorkbookConnection) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *WorkbookConnection) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *WorkbookConnection) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *WorkbookConnection) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *WorkbookConnection) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *WorkbookConnection) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *WorkbookConnection) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *WorkbookConnection) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *WorkbookConnection) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *WorkbookConnection) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorkbookConnection) SetName(rhs string)  {
	retVal := this.PropPut(0x0000006e, []interface{}{rhs})
	_= retVal
}

func (this *WorkbookConnection) Description() string {
	retVal := this.PropGet(0x000000da, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorkbookConnection) SetDescription(rhs string)  {
	retVal := this.PropPut(0x000000da, []interface{}{rhs})
	_= retVal
}

func (this *WorkbookConnection) Default_() string {
	retVal := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorkbookConnection) SetDefault_(rhs string)  {
	retVal := this.PropPut(0x00000000, []interface{}{rhs})
	_= retVal
}

func (this *WorkbookConnection) Type() int32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *WorkbookConnection) OLEDBConnection() *OLEDBConnection {
	retVal := this.PropGet(0x00000a89, nil)
	return NewOLEDBConnection(retVal.PdispValVal(), false, true)
}

func (this *WorkbookConnection) ODBCConnection() *ODBCConnection {
	retVal := this.PropGet(0x00000a8a, nil)
	return NewODBCConnection(retVal.PdispValVal(), false, true)
}

func (this *WorkbookConnection) Ranges() *Ranges {
	retVal := this.PropGet(0x00000a8b, nil)
	return NewRanges(retVal.PdispValVal(), false, true)
}

func (this *WorkbookConnection) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

func (this *WorkbookConnection) Refresh()  {
	retVal := this.Call(0x00000589, nil)
	_= retVal
}

