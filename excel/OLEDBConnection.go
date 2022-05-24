package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
	"time"
)

// 0002448D-0000-0000-C000-000000000046
var IID_OLEDBConnection = syscall.GUID{0x0002448D, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type OLEDBConnection struct {
	ole.OleClient
}

func NewOLEDBConnection(pDisp *win32.IDispatch, addRef bool, scoped bool) *OLEDBConnection {
	 if pDisp == nil {
		return nil;
	}
	p := &OLEDBConnection{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OLEDBConnectionFromVar(v ole.Variant) *OLEDBConnection {
	return NewOLEDBConnection(v.IDispatch(), false, false)
}

func (this *OLEDBConnection) IID() *syscall.GUID {
	return &IID_OLEDBConnection
}

func (this *OLEDBConnection) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OLEDBConnection) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *OLEDBConnection) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *OLEDBConnection) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *OLEDBConnection) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *OLEDBConnection) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *OLEDBConnection) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *OLEDBConnection) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *OLEDBConnection) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OLEDBConnection) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *OLEDBConnection) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OLEDBConnection) ADOConnection() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000081a, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OLEDBConnection) BackgroundQuery() bool {
	retVal, _ := this.PropGet(0x00000593, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEDBConnection) SetBackgroundQuery(rhs bool)  {
	_ = this.PropPut(0x00000593, []interface{}{rhs})
}

func (this *OLEDBConnection) CancelRefresh()  {
	retVal, _ := this.Call(0x00000635, nil)
	_= retVal
}

func (this *OLEDBConnection) CommandText() ole.Variant {
	retVal, _ := this.PropGet(0x00000725, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *OLEDBConnection) SetCommandText(rhs interface{})  {
	_ = this.PropPut(0x00000725, []interface{}{rhs})
}

func (this *OLEDBConnection) CommandType() int32 {
	retVal, _ := this.PropGet(0x00000726, nil)
	return retVal.LValVal()
}

func (this *OLEDBConnection) SetCommandType(rhs int32)  {
	_ = this.PropPut(0x00000726, []interface{}{rhs})
}

func (this *OLEDBConnection) Connection() ole.Variant {
	retVal, _ := this.PropGet(0x00000598, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *OLEDBConnection) SetConnection(rhs interface{})  {
	_ = this.PropPut(0x00000598, []interface{}{rhs})
}

func (this *OLEDBConnection) EnableRefresh() bool {
	retVal, _ := this.PropGet(0x000005c5, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEDBConnection) SetEnableRefresh(rhs bool)  {
	_ = this.PropPut(0x000005c5, []interface{}{rhs})
}

func (this *OLEDBConnection) LocalConnection() ole.Variant {
	retVal, _ := this.PropGet(0x0000072b, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *OLEDBConnection) SetLocalConnection(rhs interface{})  {
	_ = this.PropPut(0x0000072b, []interface{}{rhs})
}

func (this *OLEDBConnection) MaintainConnection() bool {
	retVal, _ := this.PropGet(0x00000728, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEDBConnection) SetMaintainConnection(rhs bool)  {
	_ = this.PropPut(0x00000728, []interface{}{rhs})
}

func (this *OLEDBConnection) MakeConnection()  {
	retVal, _ := this.Call(0x0000081c, nil)
	_= retVal
}

func (this *OLEDBConnection) Refresh()  {
	retVal, _ := this.Call(0x00000589, nil)
	_= retVal
}

func (this *OLEDBConnection) RefreshDate() time.Time {
	retVal, _ := this.PropGet(0x000002b8, nil)
	return ole.Date(retVal.DateVal()).ToGoTime()
}

func (this *OLEDBConnection) Refreshing() bool {
	retVal, _ := this.PropGet(0x00000633, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEDBConnection) RefreshOnFileOpen() bool {
	retVal, _ := this.PropGet(0x000005c7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEDBConnection) SetRefreshOnFileOpen(rhs bool)  {
	_ = this.PropPut(0x000005c7, []interface{}{rhs})
}

func (this *OLEDBConnection) RefreshPeriod() int32 {
	retVal, _ := this.PropGet(0x00000729, nil)
	return retVal.LValVal()
}

func (this *OLEDBConnection) SetRefreshPeriod(rhs int32)  {
	_ = this.PropPut(0x00000729, []interface{}{rhs})
}

func (this *OLEDBConnection) RobustConnect() int32 {
	retVal, _ := this.PropGet(0x00000821, nil)
	return retVal.LValVal()
}

func (this *OLEDBConnection) SetRobustConnect(rhs int32)  {
	_ = this.PropPut(0x00000821, []interface{}{rhs})
}

var OLEDBConnection_SaveAsODC_OptArgs= []string{
	"Description", "Keywords", 
}

func (this *OLEDBConnection) SaveAsODC(odcfileName string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(OLEDBConnection_SaveAsODC_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000822, []interface{}{odcfileName}, optArgs...)
	_= retVal
}

func (this *OLEDBConnection) SavePassword() bool {
	retVal, _ := this.PropGet(0x000005c9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEDBConnection) SetSavePassword(rhs bool)  {
	_ = this.PropPut(0x000005c9, []interface{}{rhs})
}

func (this *OLEDBConnection) SourceConnectionFile() string {
	retVal, _ := this.PropGet(0x0000081f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OLEDBConnection) SetSourceConnectionFile(rhs string)  {
	_ = this.PropPut(0x0000081f, []interface{}{rhs})
}

func (this *OLEDBConnection) SourceDataFile() string {
	retVal, _ := this.PropGet(0x00000820, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OLEDBConnection) SetSourceDataFile(rhs string)  {
	_ = this.PropPut(0x00000820, []interface{}{rhs})
}

func (this *OLEDBConnection) OLAP() bool {
	retVal, _ := this.PropGet(0x0000081d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEDBConnection) UseLocalConnection() bool {
	retVal, _ := this.PropGet(0x0000072d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEDBConnection) SetUseLocalConnection(rhs bool)  {
	_ = this.PropPut(0x0000072d, []interface{}{rhs})
}

func (this *OLEDBConnection) MaxDrillthroughRecords() int32 {
	retVal, _ := this.PropGet(0x00000a8f, nil)
	return retVal.LValVal()
}

func (this *OLEDBConnection) SetMaxDrillthroughRecords(rhs int32)  {
	_ = this.PropPut(0x00000a8f, []interface{}{rhs})
}

func (this *OLEDBConnection) IsConnected() bool {
	retVal, _ := this.PropGet(0x0000081b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEDBConnection) ServerCredentialsMethod() int32 {
	retVal, _ := this.PropGet(0x00000a90, nil)
	return retVal.LValVal()
}

func (this *OLEDBConnection) SetServerCredentialsMethod(rhs int32)  {
	_ = this.PropPut(0x00000a90, []interface{}{rhs})
}

func (this *OLEDBConnection) ServerSSOApplicationID() string {
	retVal, _ := this.PropGet(0x00000a91, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OLEDBConnection) SetServerSSOApplicationID(rhs string)  {
	_ = this.PropPut(0x00000a91, []interface{}{rhs})
}

func (this *OLEDBConnection) AlwaysUseConnectionFile() bool {
	retVal, _ := this.PropGet(0x00000a92, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEDBConnection) SetAlwaysUseConnectionFile(rhs bool)  {
	_ = this.PropPut(0x00000a92, []interface{}{rhs})
}

func (this *OLEDBConnection) ServerFillColor() bool {
	retVal, _ := this.PropGet(0x00000a93, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEDBConnection) SetServerFillColor(rhs bool)  {
	_ = this.PropPut(0x00000a93, []interface{}{rhs})
}

func (this *OLEDBConnection) ServerFontStyle() bool {
	retVal, _ := this.PropGet(0x00000a94, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEDBConnection) SetServerFontStyle(rhs bool)  {
	_ = this.PropPut(0x00000a94, []interface{}{rhs})
}

func (this *OLEDBConnection) ServerNumberFormat() bool {
	retVal, _ := this.PropGet(0x00000a95, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEDBConnection) SetServerNumberFormat(rhs bool)  {
	_ = this.PropPut(0x00000a95, []interface{}{rhs})
}

func (this *OLEDBConnection) ServerTextColor() bool {
	retVal, _ := this.PropGet(0x00000a96, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEDBConnection) SetServerTextColor(rhs bool)  {
	_ = this.PropPut(0x00000a96, []interface{}{rhs})
}

func (this *OLEDBConnection) RetrieveInOfficeUILang() bool {
	retVal, _ := this.PropGet(0x00000a97, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEDBConnection) SetRetrieveInOfficeUILang(rhs bool)  {
	_ = this.PropPut(0x00000a97, []interface{}{rhs})
}

func (this *OLEDBConnection) Reconnect()  {
	retVal, _ := this.Call(0x00000b7b, nil)
	_= retVal
}

func (this *OLEDBConnection) CalculatedMembers() *CalculatedMembers {
	retVal, _ := this.PropGet(0x0000084d, nil)
	return NewCalculatedMembers(retVal.IDispatch(), false, true)
}

func (this *OLEDBConnection) LocaleID() int32 {
	retVal, _ := this.PropGet(0x00000b7c, nil)
	return retVal.LValVal()
}

func (this *OLEDBConnection) SetLocaleID(rhs int32)  {
	_ = this.PropPut(0x00000b7c, []interface{}{rhs})
}

