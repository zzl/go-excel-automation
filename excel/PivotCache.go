package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
	"time"
)

// 0002441C-0000-0000-C000-000000000046
var IID_PivotCache = syscall.GUID{0x0002441C, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type PivotCache struct {
	ole.OleClient
}

func NewPivotCache(pDisp *win32.IDispatch, addRef bool, scoped bool) *PivotCache {
	p := &PivotCache{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PivotCacheFromVar(v ole.Variant) *PivotCache {
	return NewPivotCache(v.PdispValVal(), false, false)
}

func (this *PivotCache) IID() *syscall.GUID {
	return &IID_PivotCache
}

func (this *PivotCache) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *PivotCache) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *PivotCache) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *PivotCache) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *PivotCache) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *PivotCache) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *PivotCache) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *PivotCache) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *PivotCache) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *PivotCache) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *PivotCache) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *PivotCache) BackgroundQuery() bool {
	retVal := this.PropGet(0x00000593, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotCache) SetBackgroundQuery(rhs bool)  {
	retVal := this.PropPut(0x00000593, []interface{}{rhs})
	_= retVal
}

func (this *PivotCache) Connection() ole.Variant {
	retVal := this.PropGet(0x00000598, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *PivotCache) SetConnection(rhs interface{})  {
	retVal := this.PropPut(0x00000598, []interface{}{rhs})
	_= retVal
}

func (this *PivotCache) EnableRefresh() bool {
	retVal := this.PropGet(0x000005c5, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotCache) SetEnableRefresh(rhs bool)  {
	retVal := this.PropPut(0x000005c5, []interface{}{rhs})
	_= retVal
}

func (this *PivotCache) Index() int32 {
	retVal := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *PivotCache) MemoryUsed() int32 {
	retVal := this.PropGet(0x00000174, nil)
	return retVal.LValVal()
}

func (this *PivotCache) OptimizeCache() bool {
	retVal := this.PropGet(0x00000594, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotCache) SetOptimizeCache(rhs bool)  {
	retVal := this.PropPut(0x00000594, []interface{}{rhs})
	_= retVal
}

func (this *PivotCache) RecordCount() int32 {
	retVal := this.PropGet(0x000005c6, nil)
	return retVal.LValVal()
}

func (this *PivotCache) Refresh()  {
	retVal := this.Call(0x00000589, nil)
	_= retVal
}

func (this *PivotCache) RefreshDate() time.Time {
	retVal := this.PropGet(0x000002b8, nil)
	return ole.Date(retVal.DateVal()).ToGoTime()
}

func (this *PivotCache) RefreshName() string {
	retVal := this.PropGet(0x000002b9, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotCache) RefreshOnFileOpen() bool {
	retVal := this.PropGet(0x000005c7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotCache) SetRefreshOnFileOpen(rhs bool)  {
	retVal := this.PropPut(0x000005c7, []interface{}{rhs})
	_= retVal
}

func (this *PivotCache) Sql() ole.Variant {
	retVal := this.PropGet(0x000005c8, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *PivotCache) SetSql(rhs interface{})  {
	retVal := this.PropPut(0x000005c8, []interface{}{rhs})
	_= retVal
}

func (this *PivotCache) SavePassword() bool {
	retVal := this.PropGet(0x000005c9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotCache) SetSavePassword(rhs bool)  {
	retVal := this.PropPut(0x000005c9, []interface{}{rhs})
	_= retVal
}

func (this *PivotCache) SourceData() ole.Variant {
	retVal := this.PropGet(0x000002ae, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *PivotCache) SetSourceData(rhs interface{})  {
	retVal := this.PropPut(0x000002ae, []interface{}{rhs})
	_= retVal
}

func (this *PivotCache) CommandText() ole.Variant {
	retVal := this.PropGet(0x00000725, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *PivotCache) SetCommandText(rhs interface{})  {
	retVal := this.PropPut(0x00000725, []interface{}{rhs})
	_= retVal
}

func (this *PivotCache) CommandType() int32 {
	retVal := this.PropGet(0x00000726, nil)
	return retVal.LValVal()
}

func (this *PivotCache) SetCommandType(rhs int32)  {
	retVal := this.PropPut(0x00000726, []interface{}{rhs})
	_= retVal
}

func (this *PivotCache) QueryType() int32 {
	retVal := this.PropGet(0x00000727, nil)
	return retVal.LValVal()
}

func (this *PivotCache) MaintainConnection() bool {
	retVal := this.PropGet(0x00000728, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotCache) SetMaintainConnection(rhs bool)  {
	retVal := this.PropPut(0x00000728, []interface{}{rhs})
	_= retVal
}

func (this *PivotCache) RefreshPeriod() int32 {
	retVal := this.PropGet(0x00000729, nil)
	return retVal.LValVal()
}

func (this *PivotCache) SetRefreshPeriod(rhs int32)  {
	retVal := this.PropPut(0x00000729, []interface{}{rhs})
	_= retVal
}

func (this *PivotCache) Recordset() *ole.DispatchClass {
	retVal := this.PropGet(0x0000048d, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *PivotCache) SetRecordset(rhs *ole.DispatchClass)  {
	retVal := this.PropPutRef(0x0000048d, []interface{}{rhs})
	_= retVal
}

func (this *PivotCache) ResetTimer()  {
	retVal := this.Call(0x0000072a, nil)
	_= retVal
}

func (this *PivotCache) LocalConnection() ole.Variant {
	retVal := this.PropGet(0x0000072b, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *PivotCache) SetLocalConnection(rhs interface{})  {
	retVal := this.PropPut(0x0000072b, []interface{}{rhs})
	_= retVal
}

var PivotCache_CreatePivotTable_OptArgs= []string{
	"TableName", "ReadData", "DefaultVersion", 
}

func (this *PivotCache) CreatePivotTable(tableDestination interface{}, optArgs ...interface{}) *PivotTable {
	optArgs = ole.ProcessOptArgs(PivotCache_CreatePivotTable_OptArgs, optArgs)
	retVal := this.Call(0x0000072c, []interface{}{tableDestination}, optArgs...)
	return NewPivotTable(retVal.PdispValVal(), false, true)
}

func (this *PivotCache) UseLocalConnection() bool {
	retVal := this.PropGet(0x0000072d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotCache) SetUseLocalConnection(rhs bool)  {
	retVal := this.PropPut(0x0000072d, []interface{}{rhs})
	_= retVal
}

func (this *PivotCache) ADOConnection() *ole.DispatchClass {
	retVal := this.PropGet(0x0000081a, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *PivotCache) IsConnected() bool {
	retVal := this.PropGet(0x0000081b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotCache) MakeConnection()  {
	retVal := this.Call(0x0000081c, nil)
	_= retVal
}

func (this *PivotCache) OLAP() bool {
	retVal := this.PropGet(0x0000081d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotCache) SourceType() int32 {
	retVal := this.PropGet(0x000002ad, nil)
	return retVal.LValVal()
}

func (this *PivotCache) MissingItemsLimit() int32 {
	retVal := this.PropGet(0x0000081e, nil)
	return retVal.LValVal()
}

func (this *PivotCache) SetMissingItemsLimit(rhs int32)  {
	retVal := this.PropPut(0x0000081e, []interface{}{rhs})
	_= retVal
}

func (this *PivotCache) SourceConnectionFile() string {
	retVal := this.PropGet(0x0000081f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotCache) SetSourceConnectionFile(rhs string)  {
	retVal := this.PropPut(0x0000081f, []interface{}{rhs})
	_= retVal
}

func (this *PivotCache) SourceDataFile() string {
	retVal := this.PropGet(0x00000820, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotCache) RobustConnect() int32 {
	retVal := this.PropGet(0x00000821, nil)
	return retVal.LValVal()
}

func (this *PivotCache) SetRobustConnect(rhs int32)  {
	retVal := this.PropPut(0x00000821, []interface{}{rhs})
	_= retVal
}

var PivotCache_SaveAsODC_OptArgs= []string{
	"Description", "Keywords", 
}

func (this *PivotCache) SaveAsODC(odcfileName string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(PivotCache_SaveAsODC_OptArgs, optArgs)
	retVal := this.Call(0x00000822, []interface{}{odcfileName}, optArgs...)
	_= retVal
}

func (this *PivotCache) WorkbookConnection() *WorkbookConnection {
	retVal := this.PropGet(0x000009f0, nil)
	return NewWorkbookConnection(retVal.PdispValVal(), false, true)
}

func (this *PivotCache) Version() int32 {
	retVal := this.PropGet(0x00000188, nil)
	return retVal.LValVal()
}

func (this *PivotCache) UpgradeOnRefresh() bool {
	retVal := this.PropGet(0x000009f1, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotCache) SetUpgradeOnRefresh(rhs bool)  {
	retVal := this.PropPut(0x000009f1, []interface{}{rhs})
	_= retVal
}

