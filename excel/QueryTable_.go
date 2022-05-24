package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024428-0000-0000-C000-000000000046
var IID_QueryTable_ = syscall.GUID{0x00024428, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type QueryTable_ struct {
	ole.OleClient
}

func NewQueryTable_(pDisp *win32.IDispatch, addRef bool, scoped bool) *QueryTable_ {
	 if pDisp == nil {
		return nil;
	}
	p := &QueryTable_{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func QueryTable_FromVar(v ole.Variant) *QueryTable_ {
	return NewQueryTable_(v.IDispatch(), false, false)
}

func (this *QueryTable_) IID() *syscall.GUID {
	return &IID_QueryTable_
}

func (this *QueryTable_) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *QueryTable_) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *QueryTable_) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *QueryTable_) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *QueryTable_) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *QueryTable_) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *QueryTable_) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *QueryTable_) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *QueryTable_) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *QueryTable_) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *QueryTable_) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *QueryTable_) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *QueryTable_) SetName(rhs string)  {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *QueryTable_) FieldNames() bool {
	retVal, _ := this.PropGet(0x00000630, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetFieldNames(rhs bool)  {
	_ = this.PropPut(0x00000630, []interface{}{rhs})
}

func (this *QueryTable_) RowNumbers() bool {
	retVal, _ := this.PropGet(0x00000631, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetRowNumbers(rhs bool)  {
	_ = this.PropPut(0x00000631, []interface{}{rhs})
}

func (this *QueryTable_) FillAdjacentFormulas() bool {
	retVal, _ := this.PropGet(0x00000632, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetFillAdjacentFormulas(rhs bool)  {
	_ = this.PropPut(0x00000632, []interface{}{rhs})
}

func (this *QueryTable_) HasAutoFormat() bool {
	retVal, _ := this.PropGet(0x000002b7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetHasAutoFormat(rhs bool)  {
	_ = this.PropPut(0x000002b7, []interface{}{rhs})
}

func (this *QueryTable_) RefreshOnFileOpen() bool {
	retVal, _ := this.PropGet(0x000005c7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetRefreshOnFileOpen(rhs bool)  {
	_ = this.PropPut(0x000005c7, []interface{}{rhs})
}

func (this *QueryTable_) Refreshing() bool {
	retVal, _ := this.PropGet(0x00000633, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) FetchedRowOverflow() bool {
	retVal, _ := this.PropGet(0x00000634, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) BackgroundQuery() bool {
	retVal, _ := this.PropGet(0x00000593, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetBackgroundQuery(rhs bool)  {
	_ = this.PropPut(0x00000593, []interface{}{rhs})
}

func (this *QueryTable_) CancelRefresh()  {
	retVal, _ := this.Call(0x00000635, nil)
	_= retVal
}

func (this *QueryTable_) RefreshStyle() int32 {
	retVal, _ := this.PropGet(0x00000636, nil)
	return retVal.LValVal()
}

func (this *QueryTable_) SetRefreshStyle(rhs int32)  {
	_ = this.PropPut(0x00000636, []interface{}{rhs})
}

func (this *QueryTable_) EnableRefresh() bool {
	retVal, _ := this.PropGet(0x000005c5, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetEnableRefresh(rhs bool)  {
	_ = this.PropPut(0x000005c5, []interface{}{rhs})
}

func (this *QueryTable_) SavePassword() bool {
	retVal, _ := this.PropGet(0x000005c9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetSavePassword(rhs bool)  {
	_ = this.PropPut(0x000005c9, []interface{}{rhs})
}

func (this *QueryTable_) Destination() *Range {
	retVal, _ := this.PropGet(0x000002a9, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *QueryTable_) Connection() ole.Variant {
	retVal, _ := this.PropGet(0x00000598, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *QueryTable_) SetConnection(rhs interface{})  {
	_ = this.PropPut(0x00000598, []interface{}{rhs})
}

func (this *QueryTable_) Sql() ole.Variant {
	retVal, _ := this.PropGet(0x000005c8, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *QueryTable_) SetSql(rhs interface{})  {
	_ = this.PropPut(0x000005c8, []interface{}{rhs})
}

func (this *QueryTable_) PostText() string {
	retVal, _ := this.PropGet(0x00000637, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *QueryTable_) SetPostText(rhs string)  {
	_ = this.PropPut(0x00000637, []interface{}{rhs})
}

func (this *QueryTable_) ResultRange() *Range {
	retVal, _ := this.PropGet(0x00000638, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *QueryTable_) Delete()  {
	retVal, _ := this.Call(0x00000075, nil)
	_= retVal
}

var QueryTable__Refresh_OptArgs= []string{
	"BackgroundQuery", 
}

func (this *QueryTable_) Refresh(optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(QueryTable__Refresh_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000589, nil, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) Parameters() *Parameters {
	retVal, _ := this.PropGet(0x00000639, nil)
	return NewParameters(retVal.IDispatch(), false, true)
}

func (this *QueryTable_) Recordset() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000048d, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *QueryTable_) SetRecordset(rhs *win32.IUnknown)  {
	_ = this.PropPutRef(0x0000048d, []interface{}{rhs})
}

func (this *QueryTable_) SaveData() bool {
	retVal, _ := this.PropGet(0x000002b4, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetSaveData(rhs bool)  {
	_ = this.PropPut(0x000002b4, []interface{}{rhs})
}

func (this *QueryTable_) TablesOnlyFromHTML() bool {
	retVal, _ := this.PropGet(0x0000063a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetTablesOnlyFromHTML(rhs bool)  {
	_ = this.PropPut(0x0000063a, []interface{}{rhs})
}

func (this *QueryTable_) EnableEditing() bool {
	retVal, _ := this.PropGet(0x0000063b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetEnableEditing(rhs bool)  {
	_ = this.PropPut(0x0000063b, []interface{}{rhs})
}

func (this *QueryTable_) TextFilePlatform() int32 {
	retVal, _ := this.PropGet(0x0000073f, nil)
	return retVal.LValVal()
}

func (this *QueryTable_) SetTextFilePlatform(rhs int32)  {
	_ = this.PropPut(0x0000073f, []interface{}{rhs})
}

func (this *QueryTable_) TextFileStartRow() int32 {
	retVal, _ := this.PropGet(0x00000740, nil)
	return retVal.LValVal()
}

func (this *QueryTable_) SetTextFileStartRow(rhs int32)  {
	_ = this.PropPut(0x00000740, []interface{}{rhs})
}

func (this *QueryTable_) TextFileParseType() int32 {
	retVal, _ := this.PropGet(0x00000741, nil)
	return retVal.LValVal()
}

func (this *QueryTable_) SetTextFileParseType(rhs int32)  {
	_ = this.PropPut(0x00000741, []interface{}{rhs})
}

func (this *QueryTable_) TextFileTextQualifier() int32 {
	retVal, _ := this.PropGet(0x00000742, nil)
	return retVal.LValVal()
}

func (this *QueryTable_) SetTextFileTextQualifier(rhs int32)  {
	_ = this.PropPut(0x00000742, []interface{}{rhs})
}

func (this *QueryTable_) TextFileConsecutiveDelimiter() bool {
	retVal, _ := this.PropGet(0x00000743, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetTextFileConsecutiveDelimiter(rhs bool)  {
	_ = this.PropPut(0x00000743, []interface{}{rhs})
}

func (this *QueryTable_) TextFileTabDelimiter() bool {
	retVal, _ := this.PropGet(0x00000744, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetTextFileTabDelimiter(rhs bool)  {
	_ = this.PropPut(0x00000744, []interface{}{rhs})
}

func (this *QueryTable_) TextFileSemicolonDelimiter() bool {
	retVal, _ := this.PropGet(0x00000745, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetTextFileSemicolonDelimiter(rhs bool)  {
	_ = this.PropPut(0x00000745, []interface{}{rhs})
}

func (this *QueryTable_) TextFileCommaDelimiter() bool {
	retVal, _ := this.PropGet(0x00000746, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetTextFileCommaDelimiter(rhs bool)  {
	_ = this.PropPut(0x00000746, []interface{}{rhs})
}

func (this *QueryTable_) TextFileSpaceDelimiter() bool {
	retVal, _ := this.PropGet(0x00000747, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetTextFileSpaceDelimiter(rhs bool)  {
	_ = this.PropPut(0x00000747, []interface{}{rhs})
}

func (this *QueryTable_) TextFileOtherDelimiter() string {
	retVal, _ := this.PropGet(0x00000748, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *QueryTable_) SetTextFileOtherDelimiter(rhs string)  {
	_ = this.PropPut(0x00000748, []interface{}{rhs})
}

func (this *QueryTable_) TextFileColumnDataTypes() ole.Variant {
	retVal, _ := this.PropGet(0x00000749, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *QueryTable_) SetTextFileColumnDataTypes(rhs interface{})  {
	_ = this.PropPut(0x00000749, []interface{}{rhs})
}

func (this *QueryTable_) TextFileFixedColumnWidths() ole.Variant {
	retVal, _ := this.PropGet(0x0000074a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *QueryTable_) SetTextFileFixedColumnWidths(rhs interface{})  {
	_ = this.PropPut(0x0000074a, []interface{}{rhs})
}

func (this *QueryTable_) PreserveColumnInfo() bool {
	retVal, _ := this.PropGet(0x0000074b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetPreserveColumnInfo(rhs bool)  {
	_ = this.PropPut(0x0000074b, []interface{}{rhs})
}

func (this *QueryTable_) PreserveFormatting() bool {
	retVal, _ := this.PropGet(0x000005dc, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetPreserveFormatting(rhs bool)  {
	_ = this.PropPut(0x000005dc, []interface{}{rhs})
}

func (this *QueryTable_) AdjustColumnWidth() bool {
	retVal, _ := this.PropGet(0x0000074c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetAdjustColumnWidth(rhs bool)  {
	_ = this.PropPut(0x0000074c, []interface{}{rhs})
}

func (this *QueryTable_) CommandText() ole.Variant {
	retVal, _ := this.PropGet(0x00000725, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *QueryTable_) SetCommandText(rhs interface{})  {
	_ = this.PropPut(0x00000725, []interface{}{rhs})
}

func (this *QueryTable_) CommandType() int32 {
	retVal, _ := this.PropGet(0x00000726, nil)
	return retVal.LValVal()
}

func (this *QueryTable_) SetCommandType(rhs int32)  {
	_ = this.PropPut(0x00000726, []interface{}{rhs})
}

func (this *QueryTable_) TextFilePromptOnRefresh() bool {
	retVal, _ := this.PropGet(0x0000074d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetTextFilePromptOnRefresh(rhs bool)  {
	_ = this.PropPut(0x0000074d, []interface{}{rhs})
}

func (this *QueryTable_) QueryType() int32 {
	retVal, _ := this.PropGet(0x00000727, nil)
	return retVal.LValVal()
}

func (this *QueryTable_) MaintainConnection() bool {
	retVal, _ := this.PropGet(0x00000728, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetMaintainConnection(rhs bool)  {
	_ = this.PropPut(0x00000728, []interface{}{rhs})
}

func (this *QueryTable_) TextFileDecimalSeparator() string {
	retVal, _ := this.PropGet(0x0000074e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *QueryTable_) SetTextFileDecimalSeparator(rhs string)  {
	_ = this.PropPut(0x0000074e, []interface{}{rhs})
}

func (this *QueryTable_) TextFileThousandsSeparator() string {
	retVal, _ := this.PropGet(0x0000074f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *QueryTable_) SetTextFileThousandsSeparator(rhs string)  {
	_ = this.PropPut(0x0000074f, []interface{}{rhs})
}

func (this *QueryTable_) RefreshPeriod() int32 {
	retVal, _ := this.PropGet(0x00000729, nil)
	return retVal.LValVal()
}

func (this *QueryTable_) SetRefreshPeriod(rhs int32)  {
	_ = this.PropPut(0x00000729, []interface{}{rhs})
}

func (this *QueryTable_) ResetTimer()  {
	retVal, _ := this.Call(0x0000072a, nil)
	_= retVal
}

func (this *QueryTable_) WebSelectionType() int32 {
	retVal, _ := this.PropGet(0x00000750, nil)
	return retVal.LValVal()
}

func (this *QueryTable_) SetWebSelectionType(rhs int32)  {
	_ = this.PropPut(0x00000750, []interface{}{rhs})
}

func (this *QueryTable_) WebFormatting() int32 {
	retVal, _ := this.PropGet(0x00000751, nil)
	return retVal.LValVal()
}

func (this *QueryTable_) SetWebFormatting(rhs int32)  {
	_ = this.PropPut(0x00000751, []interface{}{rhs})
}

func (this *QueryTable_) WebTables() string {
	retVal, _ := this.PropGet(0x00000752, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *QueryTable_) SetWebTables(rhs string)  {
	_ = this.PropPut(0x00000752, []interface{}{rhs})
}

func (this *QueryTable_) WebPreFormattedTextToColumns() bool {
	retVal, _ := this.PropGet(0x00000753, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetWebPreFormattedTextToColumns(rhs bool)  {
	_ = this.PropPut(0x00000753, []interface{}{rhs})
}

func (this *QueryTable_) WebSingleBlockTextImport() bool {
	retVal, _ := this.PropGet(0x00000754, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetWebSingleBlockTextImport(rhs bool)  {
	_ = this.PropPut(0x00000754, []interface{}{rhs})
}

func (this *QueryTable_) WebDisableDateRecognition() bool {
	retVal, _ := this.PropGet(0x00000755, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetWebDisableDateRecognition(rhs bool)  {
	_ = this.PropPut(0x00000755, []interface{}{rhs})
}

func (this *QueryTable_) WebConsecutiveDelimitersAsOne() bool {
	retVal, _ := this.PropGet(0x00000756, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetWebConsecutiveDelimitersAsOne(rhs bool)  {
	_ = this.PropPut(0x00000756, []interface{}{rhs})
}

func (this *QueryTable_) WebDisableRedirections() bool {
	retVal, _ := this.PropGet(0x00000872, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetWebDisableRedirections(rhs bool)  {
	_ = this.PropPut(0x00000872, []interface{}{rhs})
}

func (this *QueryTable_) EditWebPage() ole.Variant {
	retVal, _ := this.PropGet(0x00000873, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *QueryTable_) SetEditWebPage(rhs interface{})  {
	_ = this.PropPut(0x00000873, []interface{}{rhs})
}

func (this *QueryTable_) SourceConnectionFile() string {
	retVal, _ := this.PropGet(0x0000081f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *QueryTable_) SetSourceConnectionFile(rhs string)  {
	_ = this.PropPut(0x0000081f, []interface{}{rhs})
}

func (this *QueryTable_) SourceDataFile() string {
	retVal, _ := this.PropGet(0x00000820, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *QueryTable_) SetSourceDataFile(rhs string)  {
	_ = this.PropPut(0x00000820, []interface{}{rhs})
}

func (this *QueryTable_) RobustConnect() int32 {
	retVal, _ := this.PropGet(0x00000821, nil)
	return retVal.LValVal()
}

func (this *QueryTable_) SetRobustConnect(rhs int32)  {
	_ = this.PropPut(0x00000821, []interface{}{rhs})
}

func (this *QueryTable_) TextFileTrailingMinusNumbers() bool {
	retVal, _ := this.PropGet(0x00000874, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *QueryTable_) SetTextFileTrailingMinusNumbers(rhs bool)  {
	_ = this.PropPut(0x00000874, []interface{}{rhs})
}

var QueryTable__SaveAsODC_OptArgs= []string{
	"Description", "Keywords", 
}

func (this *QueryTable_) SaveAsODC(odcfileName string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(QueryTable__SaveAsODC_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000822, []interface{}{odcfileName}, optArgs...)
	_= retVal
}

func (this *QueryTable_) ListObject() *ListObject {
	retVal, _ := this.PropGet(0x000008d1, nil)
	return NewListObject(retVal.IDispatch(), false, true)
}

func (this *QueryTable_) TextFileVisualLayout() int32 {
	retVal, _ := this.PropGet(0x000008c5, nil)
	return retVal.LValVal()
}

func (this *QueryTable_) SetTextFileVisualLayout(rhs int32)  {
	_ = this.PropPut(0x000008c5, []interface{}{rhs})
}

func (this *QueryTable_) WorkbookConnection() *WorkbookConnection {
	retVal, _ := this.PropGet(0x000009f0, nil)
	return NewWorkbookConnection(retVal.IDispatch(), false, true)
}

func (this *QueryTable_) Sort() *Sort {
	retVal, _ := this.PropGet(0x00000370, nil)
	return NewSort(retVal.IDispatch(), false, true)
}

