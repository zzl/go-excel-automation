package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
	"time"
)

// 0002448E-0000-0000-C000-000000000046
var IID_ODBCConnection = syscall.GUID{0x0002448E, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ODBCConnection struct {
	ole.OleClient
}

func NewODBCConnection(pDisp *win32.IDispatch, addRef bool, scoped bool) *ODBCConnection {
	p := &ODBCConnection{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ODBCConnectionFromVar(v ole.Variant) *ODBCConnection {
	return NewODBCConnection(v.PdispValVal(), false, false)
}

func (this *ODBCConnection) IID() *syscall.GUID {
	return &IID_ODBCConnection
}

func (this *ODBCConnection) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ODBCConnection) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *ODBCConnection) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ODBCConnection) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ODBCConnection) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *ODBCConnection) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *ODBCConnection) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *ODBCConnection) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *ODBCConnection) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *ODBCConnection) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ODBCConnection) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ODBCConnection) BackgroundQuery() bool {
	retVal := this.PropGet(0x00000593, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ODBCConnection) SetBackgroundQuery(rhs bool)  {
	retVal := this.PropPut(0x00000593, []interface{}{rhs})
	_= retVal
}

func (this *ODBCConnection) CancelRefresh()  {
	retVal := this.Call(0x00000635, nil)
	_= retVal
}

func (this *ODBCConnection) CommandText() ole.Variant {
	retVal := this.PropGet(0x00000725, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ODBCConnection) SetCommandText(rhs interface{})  {
	retVal := this.PropPut(0x00000725, []interface{}{rhs})
	_= retVal
}

func (this *ODBCConnection) CommandType() int32 {
	retVal := this.PropGet(0x00000726, nil)
	return retVal.LValVal()
}

func (this *ODBCConnection) SetCommandType(rhs int32)  {
	retVal := this.PropPut(0x00000726, []interface{}{rhs})
	_= retVal
}

func (this *ODBCConnection) Connection() ole.Variant {
	retVal := this.PropGet(0x00000598, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ODBCConnection) SetConnection(rhs interface{})  {
	retVal := this.PropPut(0x00000598, []interface{}{rhs})
	_= retVal
}

func (this *ODBCConnection) EnableRefresh() bool {
	retVal := this.PropGet(0x000005c5, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ODBCConnection) SetEnableRefresh(rhs bool)  {
	retVal := this.PropPut(0x000005c5, []interface{}{rhs})
	_= retVal
}

func (this *ODBCConnection) Refresh()  {
	retVal := this.Call(0x00000589, nil)
	_= retVal
}

func (this *ODBCConnection) RefreshDate() time.Time {
	retVal := this.PropGet(0x000002b8, nil)
	return ole.Date(retVal.DateVal()).ToGoTime()
}

func (this *ODBCConnection) Refreshing() bool {
	retVal := this.PropGet(0x00000633, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ODBCConnection) RefreshOnFileOpen() bool {
	retVal := this.PropGet(0x000005c7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ODBCConnection) SetRefreshOnFileOpen(rhs bool)  {
	retVal := this.PropPut(0x000005c7, []interface{}{rhs})
	_= retVal
}

func (this *ODBCConnection) RefreshPeriod() int32 {
	retVal := this.PropGet(0x00000729, nil)
	return retVal.LValVal()
}

func (this *ODBCConnection) SetRefreshPeriod(rhs int32)  {
	retVal := this.PropPut(0x00000729, []interface{}{rhs})
	_= retVal
}

func (this *ODBCConnection) RobustConnect() int32 {
	retVal := this.PropGet(0x00000821, nil)
	return retVal.LValVal()
}

func (this *ODBCConnection) SetRobustConnect(rhs int32)  {
	retVal := this.PropPut(0x00000821, []interface{}{rhs})
	_= retVal
}

var ODBCConnection_SaveAsODC_OptArgs= []string{
	"Description", "Keywords", 
}

func (this *ODBCConnection) SaveAsODC(odcfileName string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ODBCConnection_SaveAsODC_OptArgs, optArgs)
	retVal := this.Call(0x00000822, []interface{}{odcfileName}, optArgs...)
	_= retVal
}

func (this *ODBCConnection) SavePassword() bool {
	retVal := this.PropGet(0x000005c9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ODBCConnection) SetSavePassword(rhs bool)  {
	retVal := this.PropPut(0x000005c9, []interface{}{rhs})
	_= retVal
}

func (this *ODBCConnection) SourceConnectionFile() string {
	retVal := this.PropGet(0x0000081f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ODBCConnection) SetSourceConnectionFile(rhs string)  {
	retVal := this.PropPut(0x0000081f, []interface{}{rhs})
	_= retVal
}

func (this *ODBCConnection) SourceData() ole.Variant {
	retVal := this.PropGet(0x000002ae, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ODBCConnection) SetSourceData(rhs interface{})  {
	retVal := this.PropPut(0x000002ae, []interface{}{rhs})
	_= retVal
}

func (this *ODBCConnection) SourceDataFile() string {
	retVal := this.PropGet(0x00000820, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ODBCConnection) SetSourceDataFile(rhs string)  {
	retVal := this.PropPut(0x00000820, []interface{}{rhs})
	_= retVal
}

func (this *ODBCConnection) ServerCredentialsMethod() int32 {
	retVal := this.PropGet(0x00000a90, nil)
	return retVal.LValVal()
}

func (this *ODBCConnection) SetServerCredentialsMethod(rhs int32)  {
	retVal := this.PropPut(0x00000a90, []interface{}{rhs})
	_= retVal
}

func (this *ODBCConnection) ServerSSOApplicationID() string {
	retVal := this.PropGet(0x00000a91, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ODBCConnection) SetServerSSOApplicationID(rhs string)  {
	retVal := this.PropPut(0x00000a91, []interface{}{rhs})
	_= retVal
}

func (this *ODBCConnection) AlwaysUseConnectionFile() bool {
	retVal := this.PropGet(0x00000a92, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ODBCConnection) SetAlwaysUseConnectionFile(rhs bool)  {
	retVal := this.PropPut(0x00000a92, []interface{}{rhs})
	_= retVal
}

