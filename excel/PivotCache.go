package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"time"
	"unsafe"
)

// 0002441C-0000-0000-C000-000000000046
var IID_PivotCache = syscall.GUID{0x0002441C, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type PivotCache struct {
	ole.OleClient
}

func NewPivotCache(pDisp *win32.IDispatch, addRef bool, scoped bool) *PivotCache {
	if pDisp == nil {
		return nil
	}
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
	return NewPivotCache(v.IDispatch(), false, false)
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

func (this *PivotCache) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *PivotCache) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *PivotCache) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *PivotCache) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *PivotCache) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *PivotCache) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *PivotCache) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *PivotCache) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *PivotCache) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *PivotCache) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *PivotCache) BackgroundQuery() bool {
	retVal, _ := this.PropGet(0x00000593, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotCache) SetBackgroundQuery(rhs bool) {
	_ = this.PropPut(0x00000593, []interface{}{rhs})
}

func (this *PivotCache) Connection() ole.Variant {
	retVal, _ := this.PropGet(0x00000598, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PivotCache) SetConnection(rhs interface{}) {
	_ = this.PropPut(0x00000598, []interface{}{rhs})
}

func (this *PivotCache) EnableRefresh() bool {
	retVal, _ := this.PropGet(0x000005c5, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotCache) SetEnableRefresh(rhs bool) {
	_ = this.PropPut(0x000005c5, []interface{}{rhs})
}

func (this *PivotCache) Index() int32 {
	retVal, _ := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *PivotCache) MemoryUsed() int32 {
	retVal, _ := this.PropGet(0x00000174, nil)
	return retVal.LValVal()
}

func (this *PivotCache) OptimizeCache() bool {
	retVal, _ := this.PropGet(0x00000594, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotCache) SetOptimizeCache(rhs bool) {
	_ = this.PropPut(0x00000594, []interface{}{rhs})
}

func (this *PivotCache) RecordCount() int32 {
	retVal, _ := this.PropGet(0x000005c6, nil)
	return retVal.LValVal()
}

func (this *PivotCache) Refresh() {
	retVal, _ := this.Call(0x00000589, nil)
	_ = retVal
}

func (this *PivotCache) RefreshDate() time.Time {
	retVal, _ := this.PropGet(0x000002b8, nil)
	return ole.Date(retVal.DateVal()).ToGoTime()
}

func (this *PivotCache) RefreshName() string {
	retVal, _ := this.PropGet(0x000002b9, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotCache) RefreshOnFileOpen() bool {
	retVal, _ := this.PropGet(0x000005c7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotCache) SetRefreshOnFileOpen(rhs bool) {
	_ = this.PropPut(0x000005c7, []interface{}{rhs})
}

func (this *PivotCache) Sql() ole.Variant {
	retVal, _ := this.PropGet(0x000005c8, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PivotCache) SetSql(rhs interface{}) {
	_ = this.PropPut(0x000005c8, []interface{}{rhs})
}

func (this *PivotCache) SavePassword() bool {
	retVal, _ := this.PropGet(0x000005c9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotCache) SetSavePassword(rhs bool) {
	_ = this.PropPut(0x000005c9, []interface{}{rhs})
}

func (this *PivotCache) SourceData() ole.Variant {
	retVal, _ := this.PropGet(0x000002ae, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PivotCache) SetSourceData(rhs interface{}) {
	_ = this.PropPut(0x000002ae, []interface{}{rhs})
}

func (this *PivotCache) CommandText() ole.Variant {
	retVal, _ := this.PropGet(0x00000725, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PivotCache) SetCommandText(rhs interface{}) {
	_ = this.PropPut(0x00000725, []interface{}{rhs})
}

func (this *PivotCache) CommandType() int32 {
	retVal, _ := this.PropGet(0x00000726, nil)
	return retVal.LValVal()
}

func (this *PivotCache) SetCommandType(rhs int32) {
	_ = this.PropPut(0x00000726, []interface{}{rhs})
}

func (this *PivotCache) QueryType() int32 {
	retVal, _ := this.PropGet(0x00000727, nil)
	return retVal.LValVal()
}

func (this *PivotCache) MaintainConnection() bool {
	retVal, _ := this.PropGet(0x00000728, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotCache) SetMaintainConnection(rhs bool) {
	_ = this.PropPut(0x00000728, []interface{}{rhs})
}

func (this *PivotCache) RefreshPeriod() int32 {
	retVal, _ := this.PropGet(0x00000729, nil)
	return retVal.LValVal()
}

func (this *PivotCache) SetRefreshPeriod(rhs int32) {
	_ = this.PropPut(0x00000729, []interface{}{rhs})
}

func (this *PivotCache) Recordset() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000048d, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *PivotCache) SetRecordset(rhs *win32.IUnknown) {
	_ = this.PropPutRef(0x0000048d, []interface{}{rhs})
}

func (this *PivotCache) ResetTimer() {
	retVal, _ := this.Call(0x0000072a, nil)
	_ = retVal
}

func (this *PivotCache) LocalConnection() ole.Variant {
	retVal, _ := this.PropGet(0x0000072b, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PivotCache) SetLocalConnection(rhs interface{}) {
	_ = this.PropPut(0x0000072b, []interface{}{rhs})
}

var PivotCache_CreatePivotTable_OptArgs = []string{
	"TableName", "ReadData", "DefaultVersion",
}

func (this *PivotCache) CreatePivotTable(tableDestination interface{}, optArgs ...interface{}) *PivotTable {
	optArgs = ole.ProcessOptArgs(PivotCache_CreatePivotTable_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000072c, []interface{}{tableDestination}, optArgs...)
	return NewPivotTable(retVal.IDispatch(), false, true)
}

func (this *PivotCache) UseLocalConnection() bool {
	retVal, _ := this.PropGet(0x0000072d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotCache) SetUseLocalConnection(rhs bool) {
	_ = this.PropPut(0x0000072d, []interface{}{rhs})
}

func (this *PivotCache) ADOConnection() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000081a, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *PivotCache) IsConnected() bool {
	retVal, _ := this.PropGet(0x0000081b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotCache) MakeConnection() {
	retVal, _ := this.Call(0x0000081c, nil)
	_ = retVal
}

func (this *PivotCache) OLAP() bool {
	retVal, _ := this.PropGet(0x0000081d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotCache) SourceType() int32 {
	retVal, _ := this.PropGet(0x000002ad, nil)
	return retVal.LValVal()
}

func (this *PivotCache) MissingItemsLimit() int32 {
	retVal, _ := this.PropGet(0x0000081e, nil)
	return retVal.LValVal()
}

func (this *PivotCache) SetMissingItemsLimit(rhs int32) {
	_ = this.PropPut(0x0000081e, []interface{}{rhs})
}

func (this *PivotCache) SourceConnectionFile() string {
	retVal, _ := this.PropGet(0x0000081f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotCache) SetSourceConnectionFile(rhs string) {
	_ = this.PropPut(0x0000081f, []interface{}{rhs})
}

func (this *PivotCache) SourceDataFile() string {
	retVal, _ := this.PropGet(0x00000820, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PivotCache) RobustConnect() int32 {
	retVal, _ := this.PropGet(0x00000821, nil)
	return retVal.LValVal()
}

func (this *PivotCache) SetRobustConnect(rhs int32) {
	_ = this.PropPut(0x00000821, []interface{}{rhs})
}

var PivotCache_SaveAsODC_OptArgs = []string{
	"Description", "Keywords",
}

func (this *PivotCache) SaveAsODC(odcfileName string, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(PivotCache_SaveAsODC_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000822, []interface{}{odcfileName}, optArgs...)
	_ = retVal
}

func (this *PivotCache) WorkbookConnection() *WorkbookConnection {
	retVal, _ := this.PropGet(0x000009f0, nil)
	return NewWorkbookConnection(retVal.IDispatch(), false, true)
}

func (this *PivotCache) Version() int32 {
	retVal, _ := this.PropGet(0x00000188, nil)
	return retVal.LValVal()
}

func (this *PivotCache) UpgradeOnRefresh() bool {
	retVal, _ := this.PropGet(0x000009f1, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PivotCache) SetUpgradeOnRefresh(rhs bool) {
	_ = this.PropPut(0x000009f1, []interface{}{rhs})
}
