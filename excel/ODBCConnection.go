package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"time"
	"unsafe"
)

// 0002448E-0000-0000-C000-000000000046
var IID_ODBCConnection = syscall.GUID{0x0002448E, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ODBCConnection struct {
	ole.OleClient
}

func NewODBCConnection(pDisp *win32.IDispatch, addRef bool, scoped bool) *ODBCConnection {
	if pDisp == nil {
		return nil
	}
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
	return NewODBCConnection(v.IDispatch(), false, false)
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

func (this *ODBCConnection) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *ODBCConnection) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ODBCConnection) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ODBCConnection) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *ODBCConnection) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *ODBCConnection) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *ODBCConnection) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *ODBCConnection) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ODBCConnection) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ODBCConnection) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ODBCConnection) BackgroundQuery() bool {
	retVal, _ := this.PropGet(0x00000593, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ODBCConnection) SetBackgroundQuery(rhs bool) {
	_ = this.PropPut(0x00000593, []interface{}{rhs})
}

func (this *ODBCConnection) CancelRefresh() {
	retVal, _ := this.Call(0x00000635, nil)
	_ = retVal
}

func (this *ODBCConnection) CommandText() ole.Variant {
	retVal, _ := this.PropGet(0x00000725, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ODBCConnection) SetCommandText(rhs interface{}) {
	_ = this.PropPut(0x00000725, []interface{}{rhs})
}

func (this *ODBCConnection) CommandType() int32 {
	retVal, _ := this.PropGet(0x00000726, nil)
	return retVal.LValVal()
}

func (this *ODBCConnection) SetCommandType(rhs int32) {
	_ = this.PropPut(0x00000726, []interface{}{rhs})
}

func (this *ODBCConnection) Connection() ole.Variant {
	retVal, _ := this.PropGet(0x00000598, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ODBCConnection) SetConnection(rhs interface{}) {
	_ = this.PropPut(0x00000598, []interface{}{rhs})
}

func (this *ODBCConnection) EnableRefresh() bool {
	retVal, _ := this.PropGet(0x000005c5, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ODBCConnection) SetEnableRefresh(rhs bool) {
	_ = this.PropPut(0x000005c5, []interface{}{rhs})
}

func (this *ODBCConnection) Refresh() {
	retVal, _ := this.Call(0x00000589, nil)
	_ = retVal
}

func (this *ODBCConnection) RefreshDate() time.Time {
	retVal, _ := this.PropGet(0x000002b8, nil)
	return ole.Date(retVal.DateVal()).ToGoTime()
}

func (this *ODBCConnection) Refreshing() bool {
	retVal, _ := this.PropGet(0x00000633, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ODBCConnection) RefreshOnFileOpen() bool {
	retVal, _ := this.PropGet(0x000005c7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ODBCConnection) SetRefreshOnFileOpen(rhs bool) {
	_ = this.PropPut(0x000005c7, []interface{}{rhs})
}

func (this *ODBCConnection) RefreshPeriod() int32 {
	retVal, _ := this.PropGet(0x00000729, nil)
	return retVal.LValVal()
}

func (this *ODBCConnection) SetRefreshPeriod(rhs int32) {
	_ = this.PropPut(0x00000729, []interface{}{rhs})
}

func (this *ODBCConnection) RobustConnect() int32 {
	retVal, _ := this.PropGet(0x00000821, nil)
	return retVal.LValVal()
}

func (this *ODBCConnection) SetRobustConnect(rhs int32) {
	_ = this.PropPut(0x00000821, []interface{}{rhs})
}

var ODBCConnection_SaveAsODC_OptArgs = []string{
	"Description", "Keywords",
}

func (this *ODBCConnection) SaveAsODC(odcfileName string, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(ODBCConnection_SaveAsODC_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000822, []interface{}{odcfileName}, optArgs...)
	_ = retVal
}

func (this *ODBCConnection) SavePassword() bool {
	retVal, _ := this.PropGet(0x000005c9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ODBCConnection) SetSavePassword(rhs bool) {
	_ = this.PropPut(0x000005c9, []interface{}{rhs})
}

func (this *ODBCConnection) SourceConnectionFile() string {
	retVal, _ := this.PropGet(0x0000081f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ODBCConnection) SetSourceConnectionFile(rhs string) {
	_ = this.PropPut(0x0000081f, []interface{}{rhs})
}

func (this *ODBCConnection) SourceData() ole.Variant {
	retVal, _ := this.PropGet(0x000002ae, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ODBCConnection) SetSourceData(rhs interface{}) {
	_ = this.PropPut(0x000002ae, []interface{}{rhs})
}

func (this *ODBCConnection) SourceDataFile() string {
	retVal, _ := this.PropGet(0x00000820, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ODBCConnection) SetSourceDataFile(rhs string) {
	_ = this.PropPut(0x00000820, []interface{}{rhs})
}

func (this *ODBCConnection) ServerCredentialsMethod() int32 {
	retVal, _ := this.PropGet(0x00000a90, nil)
	return retVal.LValVal()
}

func (this *ODBCConnection) SetServerCredentialsMethod(rhs int32) {
	_ = this.PropPut(0x00000a90, []interface{}{rhs})
}

func (this *ODBCConnection) ServerSSOApplicationID() string {
	retVal, _ := this.PropGet(0x00000a91, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ODBCConnection) SetServerSSOApplicationID(rhs string) {
	_ = this.PropPut(0x00000a91, []interface{}{rhs})
}

func (this *ODBCConnection) AlwaysUseConnectionFile() bool {
	retVal, _ := this.PropGet(0x00000a92, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ODBCConnection) SetAlwaysUseConnectionFile(rhs bool) {
	_ = this.PropPut(0x00000a92, []interface{}{rhs})
}
