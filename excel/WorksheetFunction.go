package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00020845-0000-0000-C000-000000000046
var IID_WorksheetFunction = syscall.GUID{0x00020845, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type WorksheetFunction struct {
	ole.OleClient
}

func NewWorksheetFunction(pDisp *win32.IDispatch, addRef bool, scoped bool) *WorksheetFunction {
	if pDisp == nil {
		return nil
	}
	p := &WorksheetFunction{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func WorksheetFunctionFromVar(v ole.Variant) *WorksheetFunction {
	return NewWorksheetFunction(v.IDispatch(), false, false)
}

func (this *WorksheetFunction) IID() *syscall.GUID {
	return &IID_WorksheetFunction
}

func (this *WorksheetFunction) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *WorksheetFunction) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *WorksheetFunction) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *WorksheetFunction) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *WorksheetFunction) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *WorksheetFunction) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *WorksheetFunction) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *WorksheetFunction) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *WorksheetFunction) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *WorksheetFunction) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *WorksheetFunction) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var WorksheetFunction_WSFunction__OptArgs = []string{
	"Arg1", "Arg2", "Arg3", "Arg4",
	"Arg5", "Arg6", "Arg7", "Arg8",
	"Arg9", "Arg10", "Arg11", "Arg12",
	"Arg13", "Arg14", "Arg15", "Arg16",
	"Arg17", "Arg18", "Arg19", "Arg20",
	"Arg21", "Arg22", "Arg23", "Arg24",
	"Arg25", "Arg26", "Arg27", "Arg28",
	"Arg29", "Arg30",
}

func (this *WorksheetFunction) WSFunction_(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_WSFunction__OptArgs, optArgs)
	retVal, _ := this.Call(0x000000a9, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var WorksheetFunction_Count_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) Count(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Count_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004000, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) IsNA(arg1 interface{}) bool {
	retVal, _ := this.Call(0x00004002, []interface{}{arg1})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WorksheetFunction) IsError(arg1 interface{}) bool {
	retVal, _ := this.Call(0x00004003, []interface{}{arg1})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var WorksheetFunction_Sum_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) Sum(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Sum_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004004, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Average_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) Average(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Average_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004005, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Min_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) Min(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Min_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004006, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Max_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) Max(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Max_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004007, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Npv_OptArgs = []string{
	"Arg3", "Arg4", "Arg5", "Arg6",
	"Arg7", "Arg8", "Arg9", "Arg10",
	"Arg11", "Arg12", "Arg13", "Arg14",
	"Arg15", "Arg16", "Arg17", "Arg18",
	"Arg19", "Arg20", "Arg21", "Arg22",
	"Arg23", "Arg24", "Arg25", "Arg26",
	"Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) Npv(arg1 float64, arg2 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Npv_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000400b, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_StDev_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) StDev(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_StDev_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000400c, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Dollar_OptArgs = []string{
	"Arg2",
}

func (this *WorksheetFunction) Dollar(arg1 float64, optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Dollar_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000400d, []interface{}{arg1}, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var WorksheetFunction_Fixed_OptArgs = []string{
	"Arg2", "Arg3",
}

func (this *WorksheetFunction) Fixed(arg1 float64, optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Fixed_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000400e, []interface{}{arg1}, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) Pi() float64 {
	retVal, _ := this.Call(0x00004013, nil)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Ln(arg1 float64) float64 {
	retVal, _ := this.Call(0x00004016, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Log10(arg1 float64) float64 {
	retVal, _ := this.Call(0x00004017, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Round(arg1 float64, arg2 float64) float64 {
	retVal, _ := this.Call(0x0000401b, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

var WorksheetFunction_Lookup_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) Lookup(arg1 interface{}, arg2 interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Lookup_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000401c, []interface{}{arg1, arg2}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var WorksheetFunction_Index_OptArgs = []string{
	"Arg3", "Arg4",
}

func (this *WorksheetFunction) Index(arg1 interface{}, arg2 float64, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Index_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000401d, []interface{}{arg1, arg2}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *WorksheetFunction) Rept(arg1 string, arg2 float64) string {
	retVal, _ := this.Call(0x0000401e, []interface{}{arg1, arg2})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var WorksheetFunction_And_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) And(arg1 interface{}, optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_And_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004024, []interface{}{arg1}, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var WorksheetFunction_Or_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) Or(arg1 interface{}, optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Or_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004025, []interface{}{arg1}, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WorksheetFunction) DCount(arg1 *Range, arg2 interface{}, arg3 interface{}) float64 {
	retVal, _ := this.Call(0x00004028, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) DSum(arg1 *Range, arg2 interface{}, arg3 interface{}) float64 {
	retVal, _ := this.Call(0x00004029, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) DAverage(arg1 *Range, arg2 interface{}, arg3 interface{}) float64 {
	retVal, _ := this.Call(0x0000402a, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) DMin(arg1 *Range, arg2 interface{}, arg3 interface{}) float64 {
	retVal, _ := this.Call(0x0000402b, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) DMax(arg1 *Range, arg2 interface{}, arg3 interface{}) float64 {
	retVal, _ := this.Call(0x0000402c, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) DStDev(arg1 *Range, arg2 interface{}, arg3 interface{}) float64 {
	retVal, _ := this.Call(0x0000402d, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

var WorksheetFunction_Var_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) Var(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Var_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000402e, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) DVar(arg1 *Range, arg2 interface{}, arg3 interface{}) float64 {
	retVal, _ := this.Call(0x0000402f, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Text(arg1 interface{}, arg2 string) string {
	retVal, _ := this.Call(0x00004030, []interface{}{arg1, arg2})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var WorksheetFunction_LinEst_OptArgs = []string{
	"Arg2", "Arg3", "Arg4",
}

func (this *WorksheetFunction) LinEst(arg1 interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_LinEst_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004031, []interface{}{arg1}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var WorksheetFunction_Trend_OptArgs = []string{
	"Arg2", "Arg3", "Arg4",
}

func (this *WorksheetFunction) Trend(arg1 interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Trend_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004032, []interface{}{arg1}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var WorksheetFunction_LogEst_OptArgs = []string{
	"Arg2", "Arg3", "Arg4",
}

func (this *WorksheetFunction) LogEst(arg1 interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_LogEst_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004033, []interface{}{arg1}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var WorksheetFunction_Growth_OptArgs = []string{
	"Arg2", "Arg3", "Arg4",
}

func (this *WorksheetFunction) Growth(arg1 interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Growth_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004034, []interface{}{arg1}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var WorksheetFunction_Pv_OptArgs = []string{
	"Arg4", "Arg5",
}

func (this *WorksheetFunction) Pv(arg1 float64, arg2 float64, arg3 float64, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Pv_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004038, []interface{}{arg1, arg2, arg3}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Fv_OptArgs = []string{
	"Arg4", "Arg5",
}

func (this *WorksheetFunction) Fv(arg1 float64, arg2 float64, arg3 float64, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Fv_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004039, []interface{}{arg1, arg2, arg3}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_NPer_OptArgs = []string{
	"Arg4", "Arg5",
}

func (this *WorksheetFunction) NPer(arg1 float64, arg2 float64, arg3 float64, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_NPer_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000403a, []interface{}{arg1, arg2, arg3}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Pmt_OptArgs = []string{
	"Arg4", "Arg5",
}

func (this *WorksheetFunction) Pmt(arg1 float64, arg2 float64, arg3 float64, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Pmt_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000403b, []interface{}{arg1, arg2, arg3}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Rate_OptArgs = []string{
	"Arg4", "Arg5", "Arg6",
}

func (this *WorksheetFunction) Rate(arg1 float64, arg2 float64, arg3 float64, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Rate_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000403c, []interface{}{arg1, arg2, arg3}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) MIrr(arg1 interface{}, arg2 float64, arg3 float64) float64 {
	retVal, _ := this.Call(0x0000403d, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

var WorksheetFunction_Irr_OptArgs = []string{
	"Arg2",
}

func (this *WorksheetFunction) Irr(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Irr_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000403e, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Match_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) Match(arg1 interface{}, arg2 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Match_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004040, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Weekday_OptArgs = []string{
	"Arg2",
}

func (this *WorksheetFunction) Weekday(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Weekday_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004046, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Search_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) Search(arg1 string, arg2 string, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Search_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004052, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Transpose(arg1 interface{}) ole.Variant {
	retVal, _ := this.Call(0x00004053, []interface{}{arg1})
	com.AddToScope(retVal)
	return *retVal
}

func (this *WorksheetFunction) Atan2(arg1 float64, arg2 float64) float64 {
	retVal, _ := this.Call(0x00004061, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Asin(arg1 float64) float64 {
	retVal, _ := this.Call(0x00004062, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Acos(arg1 float64) float64 {
	retVal, _ := this.Call(0x00004063, []interface{}{arg1})
	return retVal.DblValVal()
}

var WorksheetFunction_Choose_OptArgs = []string{
	"Arg3", "Arg4", "Arg5", "Arg6",
	"Arg7", "Arg8", "Arg9", "Arg10",
	"Arg11", "Arg12", "Arg13", "Arg14",
	"Arg15", "Arg16", "Arg17", "Arg18",
	"Arg19", "Arg20", "Arg21", "Arg22",
	"Arg23", "Arg24", "Arg25", "Arg26",
	"Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) Choose(arg1 interface{}, arg2 interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Choose_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004064, []interface{}{arg1, arg2}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var WorksheetFunction_HLookup_OptArgs = []string{
	"Arg4",
}

func (this *WorksheetFunction) HLookup(arg1 interface{}, arg2 interface{}, arg3 interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_HLookup_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004065, []interface{}{arg1, arg2, arg3}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var WorksheetFunction_VLookup_OptArgs = []string{
	"Arg4",
}

func (this *WorksheetFunction) VLookup(arg1 interface{}, arg2 interface{}, arg3 interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_VLookup_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004066, []interface{}{arg1, arg2, arg3}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var WorksheetFunction_Log_OptArgs = []string{
	"Arg2",
}

func (this *WorksheetFunction) Log(arg1 float64, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Log_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000406d, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Proper(arg1 string) string {
	retVal, _ := this.Call(0x00004072, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) Trim(arg1 string) string {
	retVal, _ := this.Call(0x00004076, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) Replace(arg1 string, arg2 float64, arg3 float64, arg4 string) string {
	retVal, _ := this.Call(0x00004077, []interface{}{arg1, arg2, arg3, arg4})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var WorksheetFunction_Substitute_OptArgs = []string{
	"Arg4",
}

func (this *WorksheetFunction) Substitute(arg1 string, arg2 string, arg3 string, optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Substitute_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004078, []interface{}{arg1, arg2, arg3}, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var WorksheetFunction_Find_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) Find(arg1 string, arg2 string, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Find_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000407c, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) IsErr(arg1 interface{}) bool {
	retVal, _ := this.Call(0x0000407e, []interface{}{arg1})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WorksheetFunction) IsText(arg1 interface{}) bool {
	retVal, _ := this.Call(0x0000407f, []interface{}{arg1})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WorksheetFunction) IsNumber(arg1 interface{}) bool {
	retVal, _ := this.Call(0x00004080, []interface{}{arg1})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WorksheetFunction) Sln(arg1 float64, arg2 float64, arg3 float64) float64 {
	retVal, _ := this.Call(0x0000408e, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Syd(arg1 float64, arg2 float64, arg3 float64, arg4 float64) float64 {
	retVal, _ := this.Call(0x0000408f, []interface{}{arg1, arg2, arg3, arg4})
	return retVal.DblValVal()
}

var WorksheetFunction_Ddb_OptArgs = []string{
	"Arg5",
}

func (this *WorksheetFunction) Ddb(arg1 float64, arg2 float64, arg3 float64, arg4 float64, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Ddb_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004090, []interface{}{arg1, arg2, arg3, arg4}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Clean(arg1 string) string {
	retVal, _ := this.Call(0x000040a2, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) MDeterm(arg1 interface{}) float64 {
	retVal, _ := this.Call(0x000040a3, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) MInverse(arg1 interface{}) ole.Variant {
	retVal, _ := this.Call(0x000040a4, []interface{}{arg1})
	com.AddToScope(retVal)
	return *retVal
}

func (this *WorksheetFunction) MMult(arg1 interface{}, arg2 interface{}) ole.Variant {
	retVal, _ := this.Call(0x000040a5, []interface{}{arg1, arg2})
	com.AddToScope(retVal)
	return *retVal
}

var WorksheetFunction_Ipmt_OptArgs = []string{
	"Arg5", "Arg6",
}

func (this *WorksheetFunction) Ipmt(arg1 float64, arg2 float64, arg3 float64, arg4 float64, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Ipmt_OptArgs, optArgs)
	retVal, _ := this.Call(0x000040a7, []interface{}{arg1, arg2, arg3, arg4}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Ppmt_OptArgs = []string{
	"Arg5", "Arg6",
}

func (this *WorksheetFunction) Ppmt(arg1 float64, arg2 float64, arg3 float64, arg4 float64, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Ppmt_OptArgs, optArgs)
	retVal, _ := this.Call(0x000040a8, []interface{}{arg1, arg2, arg3, arg4}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_CountA_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) CountA(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_CountA_OptArgs, optArgs)
	retVal, _ := this.Call(0x000040a9, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Product_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) Product(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Product_OptArgs, optArgs)
	retVal, _ := this.Call(0x000040b7, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Fact(arg1 float64) float64 {
	retVal, _ := this.Call(0x000040b8, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) DProduct(arg1 *Range, arg2 interface{}, arg3 interface{}) float64 {
	retVal, _ := this.Call(0x000040bd, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) IsNonText(arg1 interface{}) bool {
	retVal, _ := this.Call(0x000040be, []interface{}{arg1})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var WorksheetFunction_StDevP_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) StDevP(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_StDevP_OptArgs, optArgs)
	retVal, _ := this.Call(0x000040c1, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_VarP_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) VarP(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_VarP_OptArgs, optArgs)
	retVal, _ := this.Call(0x000040c2, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) DStDevP(arg1 *Range, arg2 interface{}, arg3 interface{}) float64 {
	retVal, _ := this.Call(0x000040c3, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) DVarP(arg1 *Range, arg2 interface{}, arg3 interface{}) float64 {
	retVal, _ := this.Call(0x000040c4, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) IsLogical(arg1 interface{}) bool {
	retVal, _ := this.Call(0x000040c6, []interface{}{arg1})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WorksheetFunction) DCountA(arg1 *Range, arg2 interface{}, arg3 interface{}) float64 {
	retVal, _ := this.Call(0x000040c7, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) USDollar(arg1 float64, arg2 float64) string {
	retVal, _ := this.Call(0x000040cc, []interface{}{arg1, arg2})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var WorksheetFunction_FindB_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) FindB(arg1 string, arg2 string, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_FindB_OptArgs, optArgs)
	retVal, _ := this.Call(0x000040cd, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_SearchB_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) SearchB(arg1 string, arg2 string, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_SearchB_OptArgs, optArgs)
	retVal, _ := this.Call(0x000040ce, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) ReplaceB(arg1 string, arg2 float64, arg3 float64, arg4 string) string {
	retVal, _ := this.Call(0x000040cf, []interface{}{arg1, arg2, arg3, arg4})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) RoundUp(arg1 float64, arg2 float64) float64 {
	retVal, _ := this.Call(0x000040d4, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) RoundDown(arg1 float64, arg2 float64) float64 {
	retVal, _ := this.Call(0x000040d5, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

var WorksheetFunction_Rank_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) Rank(arg1 float64, arg2 *Range, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Rank_OptArgs, optArgs)
	retVal, _ := this.Call(0x000040d8, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Days360_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) Days360(arg1 interface{}, arg2 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Days360_OptArgs, optArgs)
	retVal, _ := this.Call(0x000040dc, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Vdb_OptArgs = []string{
	"Arg6", "Arg7",
}

func (this *WorksheetFunction) Vdb(arg1 float64, arg2 float64, arg3 float64, arg4 float64, arg5 float64, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Vdb_OptArgs, optArgs)
	retVal, _ := this.Call(0x000040de, []interface{}{arg1, arg2, arg3, arg4, arg5}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Median_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) Median(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Median_OptArgs, optArgs)
	retVal, _ := this.Call(0x000040e3, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_SumProduct_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) SumProduct(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_SumProduct_OptArgs, optArgs)
	retVal, _ := this.Call(0x000040e4, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Sinh(arg1 float64) float64 {
	retVal, _ := this.Call(0x000040e5, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Cosh(arg1 float64) float64 {
	retVal, _ := this.Call(0x000040e6, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Tanh(arg1 float64) float64 {
	retVal, _ := this.Call(0x000040e7, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Asinh(arg1 float64) float64 {
	retVal, _ := this.Call(0x000040e8, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Acosh(arg1 float64) float64 {
	retVal, _ := this.Call(0x000040e9, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Atanh(arg1 float64) float64 {
	retVal, _ := this.Call(0x000040ea, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) DGet(arg1 *Range, arg2 interface{}, arg3 interface{}) ole.Variant {
	retVal, _ := this.Call(0x000040eb, []interface{}{arg1, arg2, arg3})
	com.AddToScope(retVal)
	return *retVal
}

var WorksheetFunction_Db_OptArgs = []string{
	"Arg5",
}

func (this *WorksheetFunction) Db(arg1 float64, arg2 float64, arg3 float64, arg4 float64, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Db_OptArgs, optArgs)
	retVal, _ := this.Call(0x000040f7, []interface{}{arg1, arg2, arg3, arg4}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Frequency(arg1 interface{}, arg2 interface{}) ole.Variant {
	retVal, _ := this.Call(0x000040fc, []interface{}{arg1, arg2})
	com.AddToScope(retVal)
	return *retVal
}

var WorksheetFunction_AveDev_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) AveDev(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_AveDev_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000410d, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_BetaDist_OptArgs = []string{
	"Arg4", "Arg5",
}

func (this *WorksheetFunction) BetaDist(arg1 float64, arg2 float64, arg3 float64, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_BetaDist_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000410e, []interface{}{arg1, arg2, arg3}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) GammaLn(arg1 float64) float64 {
	retVal, _ := this.Call(0x0000410f, []interface{}{arg1})
	return retVal.DblValVal()
}

var WorksheetFunction_BetaInv_OptArgs = []string{
	"Arg4", "Arg5",
}

func (this *WorksheetFunction) BetaInv(arg1 float64, arg2 float64, arg3 float64, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_BetaInv_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004110, []interface{}{arg1, arg2, arg3}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) BinomDist(arg1 float64, arg2 float64, arg3 float64, arg4 bool) float64 {
	retVal, _ := this.Call(0x00004111, []interface{}{arg1, arg2, arg3, arg4})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) ChiDist(arg1 float64, arg2 float64) float64 {
	retVal, _ := this.Call(0x00004112, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) ChiInv(arg1 float64, arg2 float64) float64 {
	retVal, _ := this.Call(0x00004113, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Combin(arg1 float64, arg2 float64) float64 {
	retVal, _ := this.Call(0x00004114, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Confidence(arg1 float64, arg2 float64, arg3 float64) float64 {
	retVal, _ := this.Call(0x00004115, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) CritBinom(arg1 float64, arg2 float64, arg3 float64) float64 {
	retVal, _ := this.Call(0x00004116, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Even(arg1 float64) float64 {
	retVal, _ := this.Call(0x00004117, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) ExponDist(arg1 float64, arg2 float64, arg3 bool) float64 {
	retVal, _ := this.Call(0x00004118, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) FDist(arg1 float64, arg2 float64, arg3 float64) float64 {
	retVal, _ := this.Call(0x00004119, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) FInv(arg1 float64, arg2 float64, arg3 float64) float64 {
	retVal, _ := this.Call(0x0000411a, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Fisher(arg1 float64) float64 {
	retVal, _ := this.Call(0x0000411b, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) FisherInv(arg1 float64) float64 {
	retVal, _ := this.Call(0x0000411c, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Floor(arg1 float64, arg2 float64) float64 {
	retVal, _ := this.Call(0x0000411d, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) GammaDist(arg1 float64, arg2 float64, arg3 float64, arg4 bool) float64 {
	retVal, _ := this.Call(0x0000411e, []interface{}{arg1, arg2, arg3, arg4})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) GammaInv(arg1 float64, arg2 float64, arg3 float64) float64 {
	retVal, _ := this.Call(0x0000411f, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Ceiling(arg1 float64, arg2 float64) float64 {
	retVal, _ := this.Call(0x00004120, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) HypGeomDist(arg1 float64, arg2 float64, arg3 float64, arg4 float64) float64 {
	retVal, _ := this.Call(0x00004121, []interface{}{arg1, arg2, arg3, arg4})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) LogNormDist(arg1 float64, arg2 float64, arg3 float64) float64 {
	retVal, _ := this.Call(0x00004122, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) LogInv(arg1 float64, arg2 float64, arg3 float64) float64 {
	retVal, _ := this.Call(0x00004123, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) NegBinomDist(arg1 float64, arg2 float64, arg3 float64) float64 {
	retVal, _ := this.Call(0x00004124, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) NormDist(arg1 float64, arg2 float64, arg3 float64, arg4 bool) float64 {
	retVal, _ := this.Call(0x00004125, []interface{}{arg1, arg2, arg3, arg4})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) NormSDist(arg1 float64) float64 {
	retVal, _ := this.Call(0x00004126, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) NormInv(arg1 float64, arg2 float64, arg3 float64) float64 {
	retVal, _ := this.Call(0x00004127, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) NormSInv(arg1 float64) float64 {
	retVal, _ := this.Call(0x00004128, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Standardize(arg1 float64, arg2 float64, arg3 float64) float64 {
	retVal, _ := this.Call(0x00004129, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Odd(arg1 float64) float64 {
	retVal, _ := this.Call(0x0000412a, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Permut(arg1 float64, arg2 float64) float64 {
	retVal, _ := this.Call(0x0000412b, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Poisson(arg1 float64, arg2 float64, arg3 bool) float64 {
	retVal, _ := this.Call(0x0000412c, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) TDist(arg1 float64, arg2 float64, arg3 float64) float64 {
	retVal, _ := this.Call(0x0000412d, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Weibull(arg1 float64, arg2 float64, arg3 float64, arg4 bool) float64 {
	retVal, _ := this.Call(0x0000412e, []interface{}{arg1, arg2, arg3, arg4})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) SumXMY2(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x0000412f, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) SumX2MY2(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x00004130, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) SumX2PY2(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x00004131, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) ChiTest(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x00004132, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Correl(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x00004133, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Covar(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x00004134, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Forecast(arg1 float64, arg2 interface{}, arg3 interface{}) float64 {
	retVal, _ := this.Call(0x00004135, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) FTest(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x00004136, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Intercept(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x00004137, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Pearson(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x00004138, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) RSq(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x00004139, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) StEyx(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x0000413a, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Slope(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x0000413b, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) TTest(arg1 interface{}, arg2 interface{}, arg3 float64, arg4 float64) float64 {
	retVal, _ := this.Call(0x0000413c, []interface{}{arg1, arg2, arg3, arg4})
	return retVal.DblValVal()
}

var WorksheetFunction_Prob_OptArgs = []string{
	"Arg4",
}

func (this *WorksheetFunction) Prob(arg1 interface{}, arg2 interface{}, arg3 float64, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Prob_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000413d, []interface{}{arg1, arg2, arg3}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_DevSq_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) DevSq(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_DevSq_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000413e, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_GeoMean_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) GeoMean(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_GeoMean_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000413f, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_HarMean_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) HarMean(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_HarMean_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004140, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_SumSq_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) SumSq(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_SumSq_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004141, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Kurt_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) Kurt(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Kurt_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004142, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Skew_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) Skew(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Skew_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004143, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_ZTest_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) ZTest(arg1 interface{}, arg2 float64, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_ZTest_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004144, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Large(arg1 interface{}, arg2 float64) float64 {
	retVal, _ := this.Call(0x00004145, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Small(arg1 interface{}, arg2 float64) float64 {
	retVal, _ := this.Call(0x00004146, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Quartile(arg1 interface{}, arg2 float64) float64 {
	retVal, _ := this.Call(0x00004147, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Percentile(arg1 interface{}, arg2 float64) float64 {
	retVal, _ := this.Call(0x00004148, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

var WorksheetFunction_PercentRank_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) PercentRank(arg1 interface{}, arg2 float64, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_PercentRank_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004149, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Mode_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) Mode(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Mode_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000414a, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) TrimMean(arg1 interface{}, arg2 float64) float64 {
	retVal, _ := this.Call(0x0000414b, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) TInv(arg1 float64, arg2 float64) float64 {
	retVal, _ := this.Call(0x0000414c, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Power(arg1 float64, arg2 float64) float64 {
	retVal, _ := this.Call(0x00004151, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Radians(arg1 float64) float64 {
	retVal, _ := this.Call(0x00004156, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Degrees(arg1 float64) float64 {
	retVal, _ := this.Call(0x00004157, []interface{}{arg1})
	return retVal.DblValVal()
}

var WorksheetFunction_Subtotal_OptArgs = []string{
	"Arg3", "Arg4", "Arg5", "Arg6",
	"Arg7", "Arg8", "Arg9", "Arg10",
	"Arg11", "Arg12", "Arg13", "Arg14",
	"Arg15", "Arg16", "Arg17", "Arg18",
	"Arg19", "Arg20", "Arg21", "Arg22",
	"Arg23", "Arg24", "Arg25", "Arg26",
	"Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) Subtotal(arg1 float64, arg2 *Range, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Subtotal_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004158, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_SumIf_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) SumIf(arg1 *Range, arg2 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_SumIf_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004159, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) CountIf(arg1 *Range, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x0000415a, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) CountBlank(arg1 *Range) float64 {
	retVal, _ := this.Call(0x0000415b, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Ispmt(arg1 float64, arg2 float64, arg3 float64, arg4 float64) float64 {
	retVal, _ := this.Call(0x0000415e, []interface{}{arg1, arg2, arg3, arg4})
	return retVal.DblValVal()
}

var WorksheetFunction_Roman_OptArgs = []string{
	"Arg2",
}

func (this *WorksheetFunction) Roman(arg1 float64, optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Roman_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004162, []interface{}{arg1}, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) Asc(arg1 string) string {
	retVal, _ := this.Call(0x000040d6, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) Dbcs(arg1 string) string {
	retVal, _ := this.Call(0x000040d7, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) Phonetic(arg1 *Range) string {
	retVal, _ := this.Call(0x00004168, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) BahtText(arg1 float64) string {
	retVal, _ := this.Call(0x00004170, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) ThaiDayOfWeek(arg1 float64) string {
	retVal, _ := this.Call(0x00004171, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) ThaiDigit(arg1 string) string {
	retVal, _ := this.Call(0x00004172, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) ThaiMonthOfYear(arg1 float64) string {
	retVal, _ := this.Call(0x00004173, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) ThaiNumSound(arg1 float64) string {
	retVal, _ := this.Call(0x00004174, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) ThaiNumString(arg1 float64) string {
	retVal, _ := this.Call(0x00004175, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) ThaiStringLength(arg1 string) float64 {
	retVal, _ := this.Call(0x00004176, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) IsThaiDigit(arg1 string) bool {
	retVal, _ := this.Call(0x00004177, []interface{}{arg1})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WorksheetFunction) RoundBahtDown(arg1 float64) float64 {
	retVal, _ := this.Call(0x00004178, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) RoundBahtUp(arg1 float64) float64 {
	retVal, _ := this.Call(0x00004179, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) ThaiYear(arg1 float64) float64 {
	retVal, _ := this.Call(0x0000417a, []interface{}{arg1})
	return retVal.DblValVal()
}

var WorksheetFunction_RTD_OptArgs = []string{
	"topic2", "topic3", "topic4", "topic5",
	"topic6", "topic7", "topic8", "topic9",
	"topic10", "topic11", "topic12", "topic13",
	"topic14", "topic15", "topic16", "topic17",
	"topic18", "topic19", "topic20", "topic21",
	"topic22", "topic23", "topic24", "topic25",
	"topic26", "topic27", "topic28",
}

func (this *WorksheetFunction) RTD(progID interface{}, server interface{}, topic1 interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_RTD_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000417b, []interface{}{progID, server, topic1}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var WorksheetFunction_Hex2Bin_OptArgs = []string{
	"Arg2",
}

func (this *WorksheetFunction) Hex2Bin(arg1 interface{}, optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Hex2Bin_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004180, []interface{}{arg1}, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) Hex2Dec(arg1 interface{}) string {
	retVal, _ := this.Call(0x00004181, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var WorksheetFunction_Hex2Oct_OptArgs = []string{
	"Arg2",
}

func (this *WorksheetFunction) Hex2Oct(arg1 interface{}, optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Hex2Oct_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004182, []interface{}{arg1}, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var WorksheetFunction_Dec2Bin_OptArgs = []string{
	"Arg2",
}

func (this *WorksheetFunction) Dec2Bin(arg1 interface{}, optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Dec2Bin_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004183, []interface{}{arg1}, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var WorksheetFunction_Dec2Hex_OptArgs = []string{
	"Arg2",
}

func (this *WorksheetFunction) Dec2Hex(arg1 interface{}, optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Dec2Hex_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004184, []interface{}{arg1}, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var WorksheetFunction_Dec2Oct_OptArgs = []string{
	"Arg2",
}

func (this *WorksheetFunction) Dec2Oct(arg1 interface{}, optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Dec2Oct_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004185, []interface{}{arg1}, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var WorksheetFunction_Oct2Bin_OptArgs = []string{
	"Arg2",
}

func (this *WorksheetFunction) Oct2Bin(arg1 interface{}, optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Oct2Bin_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004186, []interface{}{arg1}, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var WorksheetFunction_Oct2Hex_OptArgs = []string{
	"Arg2",
}

func (this *WorksheetFunction) Oct2Hex(arg1 interface{}, optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Oct2Hex_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004187, []interface{}{arg1}, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) Oct2Dec(arg1 interface{}) string {
	retVal, _ := this.Call(0x00004188, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) Bin2Dec(arg1 interface{}) string {
	retVal, _ := this.Call(0x00004189, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var WorksheetFunction_Bin2Oct_OptArgs = []string{
	"Arg2",
}

func (this *WorksheetFunction) Bin2Oct(arg1 interface{}, optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Bin2Oct_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000418a, []interface{}{arg1}, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var WorksheetFunction_Bin2Hex_OptArgs = []string{
	"Arg2",
}

func (this *WorksheetFunction) Bin2Hex(arg1 interface{}, optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Bin2Hex_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000418b, []interface{}{arg1}, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) ImSub(arg1 interface{}, arg2 interface{}) string {
	retVal, _ := this.Call(0x0000418c, []interface{}{arg1, arg2})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) ImDiv(arg1 interface{}, arg2 interface{}) string {
	retVal, _ := this.Call(0x0000418d, []interface{}{arg1, arg2})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) ImPower(arg1 interface{}, arg2 interface{}) string {
	retVal, _ := this.Call(0x0000418e, []interface{}{arg1, arg2})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) ImAbs(arg1 interface{}) string {
	retVal, _ := this.Call(0x0000418f, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) ImSqrt(arg1 interface{}) string {
	retVal, _ := this.Call(0x00004190, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) ImLn(arg1 interface{}) string {
	retVal, _ := this.Call(0x00004191, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) ImLog2(arg1 interface{}) string {
	retVal, _ := this.Call(0x00004192, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) ImLog10(arg1 interface{}) string {
	retVal, _ := this.Call(0x00004193, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) ImSin(arg1 interface{}) string {
	retVal, _ := this.Call(0x00004194, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) ImCos(arg1 interface{}) string {
	retVal, _ := this.Call(0x00004195, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) ImExp(arg1 interface{}) string {
	retVal, _ := this.Call(0x00004196, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) ImArgument(arg1 interface{}) string {
	retVal, _ := this.Call(0x00004197, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) ImConjugate(arg1 interface{}) string {
	retVal, _ := this.Call(0x00004198, []interface{}{arg1})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) Imaginary(arg1 interface{}) float64 {
	retVal, _ := this.Call(0x00004199, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) ImReal(arg1 interface{}) float64 {
	retVal, _ := this.Call(0x0000419a, []interface{}{arg1})
	return retVal.DblValVal()
}

var WorksheetFunction_Complex_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) Complex(arg1 interface{}, arg2 interface{}, optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Complex_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000419b, []interface{}{arg1, arg2}, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var WorksheetFunction_ImSum_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) ImSum(arg1 interface{}, optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_ImSum_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000419c, []interface{}{arg1}, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var WorksheetFunction_ImProduct_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) ImProduct(arg1 interface{}, optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_ImProduct_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000419d, []interface{}{arg1}, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WorksheetFunction) SeriesSum(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}) float64 {
	retVal, _ := this.Call(0x0000419e, []interface{}{arg1, arg2, arg3, arg4})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) FactDouble(arg1 interface{}) float64 {
	retVal, _ := this.Call(0x0000419f, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) SqrtPi(arg1 interface{}) float64 {
	retVal, _ := this.Call(0x000041a0, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Quotient(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x000041a1, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

var WorksheetFunction_Delta_OptArgs = []string{
	"Arg2",
}

func (this *WorksheetFunction) Delta(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Delta_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041a2, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_GeStep_OptArgs = []string{
	"Arg2",
}

func (this *WorksheetFunction) GeStep(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_GeStep_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041a3, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) IsEven(arg1 interface{}) bool {
	retVal, _ := this.Call(0x000041a4, []interface{}{arg1})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WorksheetFunction) IsOdd(arg1 interface{}) bool {
	retVal, _ := this.Call(0x000041a5, []interface{}{arg1})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WorksheetFunction) MRound(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x000041a6, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

var WorksheetFunction_Erf_OptArgs = []string{
	"Arg2",
}

func (this *WorksheetFunction) Erf(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Erf_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041a7, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) ErfC(arg1 interface{}) float64 {
	retVal, _ := this.Call(0x000041a8, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) BesselJ(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x000041a9, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) BesselK(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x000041aa, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) BesselY(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x000041ab, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) BesselI(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x000041ac, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

var WorksheetFunction_Xirr_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) Xirr(arg1 interface{}, arg2 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Xirr_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041ad, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Xnpv(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x000041ae, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

var WorksheetFunction_PriceMat_OptArgs = []string{
	"Arg6",
}

func (this *WorksheetFunction) PriceMat(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_PriceMat_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041af, []interface{}{arg1, arg2, arg3, arg4, arg5}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_YieldMat_OptArgs = []string{
	"Arg6",
}

func (this *WorksheetFunction) YieldMat(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_YieldMat_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041b0, []interface{}{arg1, arg2, arg3, arg4, arg5}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_IntRate_OptArgs = []string{
	"Arg5",
}

func (this *WorksheetFunction) IntRate(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_IntRate_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041b1, []interface{}{arg1, arg2, arg3, arg4}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Received_OptArgs = []string{
	"Arg5",
}

func (this *WorksheetFunction) Received(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Received_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041b2, []interface{}{arg1, arg2, arg3, arg4}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Disc_OptArgs = []string{
	"Arg5",
}

func (this *WorksheetFunction) Disc(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Disc_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041b3, []interface{}{arg1, arg2, arg3, arg4}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_PriceDisc_OptArgs = []string{
	"Arg5",
}

func (this *WorksheetFunction) PriceDisc(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_PriceDisc_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041b4, []interface{}{arg1, arg2, arg3, arg4}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_YieldDisc_OptArgs = []string{
	"Arg5",
}

func (this *WorksheetFunction) YieldDisc(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_YieldDisc_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041b5, []interface{}{arg1, arg2, arg3, arg4}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_TBillEq_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) TBillEq(arg1 interface{}, arg2 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_TBillEq_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041b6, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_TBillPrice_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) TBillPrice(arg1 interface{}, arg2 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_TBillPrice_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041b7, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_TBillYield_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) TBillYield(arg1 interface{}, arg2 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_TBillYield_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041b8, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Price_OptArgs = []string{
	"Arg7",
}

func (this *WorksheetFunction) Price(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Price_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041b9, []interface{}{arg1, arg2, arg3, arg4, arg5, arg6}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) DollarDe(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x000041bb, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) DollarFr(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x000041bc, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Nominal(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x000041bd, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Effect(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x000041be, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) CumPrinc(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}) float64 {
	retVal, _ := this.Call(0x000041bf, []interface{}{arg1, arg2, arg3, arg4, arg5, arg6})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) CumIPmt(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}) float64 {
	retVal, _ := this.Call(0x000041c0, []interface{}{arg1, arg2, arg3, arg4, arg5, arg6})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) EDate(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x000041c1, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) EoMonth(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x000041c2, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

var WorksheetFunction_YearFrac_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) YearFrac(arg1 interface{}, arg2 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_YearFrac_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041c3, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_CoupDayBs_OptArgs = []string{
	"Arg4",
}

func (this *WorksheetFunction) CoupDayBs(arg1 interface{}, arg2 interface{}, arg3 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_CoupDayBs_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041c4, []interface{}{arg1, arg2, arg3}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_CoupDays_OptArgs = []string{
	"Arg4",
}

func (this *WorksheetFunction) CoupDays(arg1 interface{}, arg2 interface{}, arg3 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_CoupDays_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041c5, []interface{}{arg1, arg2, arg3}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_CoupDaysNc_OptArgs = []string{
	"Arg4",
}

func (this *WorksheetFunction) CoupDaysNc(arg1 interface{}, arg2 interface{}, arg3 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_CoupDaysNc_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041c6, []interface{}{arg1, arg2, arg3}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_CoupNcd_OptArgs = []string{
	"Arg4",
}

func (this *WorksheetFunction) CoupNcd(arg1 interface{}, arg2 interface{}, arg3 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_CoupNcd_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041c7, []interface{}{arg1, arg2, arg3}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_CoupNum_OptArgs = []string{
	"Arg4",
}

func (this *WorksheetFunction) CoupNum(arg1 interface{}, arg2 interface{}, arg3 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_CoupNum_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041c8, []interface{}{arg1, arg2, arg3}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_CoupPcd_OptArgs = []string{
	"Arg4",
}

func (this *WorksheetFunction) CoupPcd(arg1 interface{}, arg2 interface{}, arg3 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_CoupPcd_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041c9, []interface{}{arg1, arg2, arg3}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Duration_OptArgs = []string{
	"Arg6",
}

func (this *WorksheetFunction) Duration(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Duration_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041ca, []interface{}{arg1, arg2, arg3, arg4, arg5}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_MDuration_OptArgs = []string{
	"Arg6",
}

func (this *WorksheetFunction) MDuration(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_MDuration_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041cb, []interface{}{arg1, arg2, arg3, arg4, arg5}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_OddLPrice_OptArgs = []string{
	"Arg8",
}

func (this *WorksheetFunction) OddLPrice(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_OddLPrice_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041cc, []interface{}{arg1, arg2, arg3, arg4, arg5, arg6, arg7}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_OddLYield_OptArgs = []string{
	"Arg8",
}

func (this *WorksheetFunction) OddLYield(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_OddLYield_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041cd, []interface{}{arg1, arg2, arg3, arg4, arg5, arg6, arg7}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_OddFPrice_OptArgs = []string{
	"Arg9",
}

func (this *WorksheetFunction) OddFPrice(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_OddFPrice_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041ce, []interface{}{arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_OddFYield_OptArgs = []string{
	"Arg9",
}

func (this *WorksheetFunction) OddFYield(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, arg7 interface{}, arg8 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_OddFYield_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041cf, []interface{}{arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) RandBetween(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x000041d0, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

var WorksheetFunction_WeekNum_OptArgs = []string{
	"Arg2",
}

func (this *WorksheetFunction) WeekNum(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_WeekNum_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041d1, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_AmorDegrc_OptArgs = []string{
	"Arg7",
}

func (this *WorksheetFunction) AmorDegrc(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_AmorDegrc_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041d2, []interface{}{arg1, arg2, arg3, arg4, arg5, arg6}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_AmorLinc_OptArgs = []string{
	"Arg7",
}

func (this *WorksheetFunction) AmorLinc(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_AmorLinc_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041d3, []interface{}{arg1, arg2, arg3, arg4, arg5, arg6}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Convert(arg1 interface{}, arg2 interface{}, arg3 interface{}) float64 {
	retVal, _ := this.Call(0x000041d4, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

var WorksheetFunction_AccrInt_OptArgs = []string{
	"Arg7",
}

func (this *WorksheetFunction) AccrInt(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, arg5 interface{}, arg6 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_AccrInt_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041d5, []interface{}{arg1, arg2, arg3, arg4, arg5, arg6}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_AccrIntM_OptArgs = []string{
	"Arg5",
}

func (this *WorksheetFunction) AccrIntM(arg1 interface{}, arg2 interface{}, arg3 interface{}, arg4 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_AccrIntM_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041d6, []interface{}{arg1, arg2, arg3, arg4}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_WorkDay_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) WorkDay(arg1 interface{}, arg2 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_WorkDay_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041d7, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_NetworkDays_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) NetworkDays(arg1 interface{}, arg2 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_NetworkDays_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041d8, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Gcd_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) Gcd(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Gcd_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041d9, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_MultiNomial_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) MultiNomial(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_MultiNomial_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041da, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Lcm_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) Lcm(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Lcm_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041db, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) FVSchedule(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x000041dc, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

var WorksheetFunction_SumIfs_OptArgs = []string{
	"Arg4", "Arg5", "Arg6", "Arg7",
	"Arg8", "Arg9", "Arg10", "Arg11",
	"Arg12", "Arg13", "Arg14", "Arg15",
	"Arg16", "Arg17", "Arg18", "Arg19",
	"Arg20", "Arg21", "Arg22", "Arg23",
	"Arg24", "Arg25", "Arg26", "Arg27",
	"Arg28", "Arg29",
}

func (this *WorksheetFunction) SumIfs(arg1 *Range, arg2 *Range, arg3 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_SumIfs_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041e2, []interface{}{arg1, arg2, arg3}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_CountIfs_OptArgs = []string{
	"Arg3", "Arg4", "Arg5", "Arg6",
	"Arg7", "Arg8", "Arg9", "Arg10",
	"Arg11", "Arg12", "Arg13", "Arg14",
	"Arg15", "Arg16", "Arg17", "Arg18",
	"Arg19", "Arg20", "Arg21", "Arg22",
	"Arg23", "Arg24", "Arg25", "Arg26",
	"Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) CountIfs(arg1 *Range, arg2 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_CountIfs_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041e1, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_AverageIf_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) AverageIf(arg1 *Range, arg2 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_AverageIf_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041e3, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_AverageIfs_OptArgs = []string{
	"Arg4", "Arg5", "Arg6", "Arg7",
	"Arg8", "Arg9", "Arg10", "Arg11",
	"Arg12", "Arg13", "Arg14", "Arg15",
	"Arg16", "Arg17", "Arg18", "Arg19",
	"Arg20", "Arg21", "Arg22", "Arg23",
	"Arg24", "Arg25", "Arg26", "Arg27",
	"Arg28", "Arg29",
}

func (this *WorksheetFunction) AverageIfs(arg1 *Range, arg2 *Range, arg3 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_AverageIfs_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041e4, []interface{}{arg1, arg2, arg3}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) IfError(arg1 interface{}, arg2 interface{}) ole.Variant {
	retVal, _ := this.Call(0x000041e0, []interface{}{arg1, arg2})
	com.AddToScope(retVal)
	return *retVal
}

var WorksheetFunction_Aggregate_OptArgs = []string{
	"Arg4", "Arg5", "Arg6", "Arg7",
	"Arg8", "Arg9", "Arg10", "Arg11",
	"Arg12", "Arg13", "Arg14", "Arg15",
	"Arg16", "Arg17", "Arg18", "Arg19",
	"Arg20", "Arg21", "Arg22", "Arg23",
	"Arg24", "Arg25", "Arg26", "Arg27",
	"Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) Aggregate(arg1 float64, arg2 float64, arg3 *Range, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Aggregate_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041e5, []interface{}{arg1, arg2, arg3}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Confidence_Norm(arg1 float64, arg2 float64, arg3 float64) float64 {
	retVal, _ := this.Call(0x000041e8, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Confidence_T(arg1 float64, arg2 float64, arg3 float64) float64 {
	retVal, _ := this.Call(0x000041e9, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) ChiSq_Test(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x000041ea, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) F_Test(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x000041eb, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Covariance_P(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x000041ec, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Covariance_S(arg1 interface{}, arg2 interface{}) float64 {
	retVal, _ := this.Call(0x000041ed, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Expon_Dist(arg1 float64, arg2 float64, arg3 bool) float64 {
	retVal, _ := this.Call(0x000041ee, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Gamma_Dist(arg1 float64, arg2 float64, arg3 float64, arg4 bool) float64 {
	retVal, _ := this.Call(0x000041ef, []interface{}{arg1, arg2, arg3, arg4})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Gamma_Inv(arg1 float64, arg2 float64, arg3 float64) float64 {
	retVal, _ := this.Call(0x000041f0, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

var WorksheetFunction_Mode_Mult_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) Mode_Mult(arg1 interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Mode_Mult_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041f1, []interface{}{arg1}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var WorksheetFunction_Mode_Sngl_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) Mode_Sngl(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Mode_Sngl_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041f2, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Norm_Dist(arg1 float64, arg2 float64, arg3 float64, arg4 bool) float64 {
	retVal, _ := this.Call(0x000041f3, []interface{}{arg1, arg2, arg3, arg4})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Norm_Inv(arg1 float64, arg2 float64, arg3 float64) float64 {
	retVal, _ := this.Call(0x000041f4, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Percentile_Exc(arg1 interface{}, arg2 float64) float64 {
	retVal, _ := this.Call(0x000041f5, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Percentile_Inc(arg1 interface{}, arg2 float64) float64 {
	retVal, _ := this.Call(0x000041f6, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

var WorksheetFunction_PercentRank_Exc_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) PercentRank_Exc(arg1 interface{}, arg2 float64, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_PercentRank_Exc_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041f7, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_PercentRank_Inc_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) PercentRank_Inc(arg1 interface{}, arg2 float64, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_PercentRank_Inc_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041f8, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Poisson_Dist(arg1 float64, arg2 float64, arg3 bool) float64 {
	retVal, _ := this.Call(0x000041f9, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Quartile_Exc(arg1 interface{}, arg2 float64) float64 {
	retVal, _ := this.Call(0x000041fa, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Quartile_Inc(arg1 interface{}, arg2 float64) float64 {
	retVal, _ := this.Call(0x000041fb, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

var WorksheetFunction_Rank_Avg_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) Rank_Avg(arg1 float64, arg2 *Range, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Rank_Avg_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041fc, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Rank_Eq_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) Rank_Eq(arg1 float64, arg2 *Range, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Rank_Eq_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041fd, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_StDev_S_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) StDev_S(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_StDev_S_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041fe, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_StDev_P_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) StDev_P(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_StDev_P_OptArgs, optArgs)
	retVal, _ := this.Call(0x000041ff, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) T_Dist(arg1 float64, arg2 float64, arg3 bool) float64 {
	retVal, _ := this.Call(0x00004200, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) T_Dist_2T(arg1 float64, arg2 float64) float64 {
	retVal, _ := this.Call(0x00004201, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) T_Dist_RT(arg1 float64, arg2 float64) float64 {
	retVal, _ := this.Call(0x00004202, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) T_Inv(arg1 float64, arg2 float64) float64 {
	retVal, _ := this.Call(0x00004203, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) T_Inv_2T(arg1 float64, arg2 float64) float64 {
	retVal, _ := this.Call(0x00004204, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

var WorksheetFunction_Var_S_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) Var_S(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Var_S_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004205, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Var_P_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) Var_P(arg1 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Var_P_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004206, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Weibull_Dist(arg1 float64, arg2 float64, arg3 float64, arg4 bool) float64 {
	retVal, _ := this.Call(0x00004207, []interface{}{arg1, arg2, arg3, arg4})
	return retVal.DblValVal()
}

var WorksheetFunction_NetworkDays_Intl_OptArgs = []string{
	"Arg3", "Arg4",
}

func (this *WorksheetFunction) NetworkDays_Intl(arg1 interface{}, arg2 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_NetworkDays_Intl_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004208, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_WorkDay_Intl_OptArgs = []string{
	"Arg3", "Arg4",
}

func (this *WorksheetFunction) WorkDay_Intl(arg1 interface{}, arg2 interface{}, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_WorkDay_Intl_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004209, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_ISO_Ceiling_OptArgs = []string{
	"Arg2",
}

func (this *WorksheetFunction) ISO_Ceiling(arg1 float64, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_ISO_Ceiling_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000420b, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Dummy21(arg1 float64, arg2 float64) float64 {
	retVal, _ := this.Call(0x00000b0a, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

var WorksheetFunction_Dummy19_OptArgs = []string{
	"Arg2", "Arg3", "Arg4", "Arg5",
	"Arg6", "Arg7", "Arg8", "Arg9",
	"Arg10", "Arg11", "Arg12", "Arg13",
	"Arg14", "Arg15", "Arg16", "Arg17",
	"Arg18", "Arg19", "Arg20", "Arg21",
	"Arg22", "Arg23", "Arg24", "Arg25",
	"Arg26", "Arg27", "Arg28", "Arg29", "Arg30",
}

func (this *WorksheetFunction) Dummy19(arg1 interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Dummy19_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000b0b, []interface{}{arg1}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var WorksheetFunction_Beta_Dist_OptArgs = []string{
	"Arg5", "Arg6",
}

func (this *WorksheetFunction) Beta_Dist(arg1 float64, arg2 float64, arg3 float64, arg4 bool, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Beta_Dist_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000420d, []interface{}{arg1, arg2, arg3, arg4}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Beta_Inv_OptArgs = []string{
	"Arg4", "Arg5",
}

func (this *WorksheetFunction) Beta_Inv(arg1 float64, arg2 float64, arg3 float64, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Beta_Inv_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000420e, []interface{}{arg1, arg2, arg3}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) ChiSq_Dist(arg1 float64, arg2 float64, arg3 bool) float64 {
	retVal, _ := this.Call(0x0000420f, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) ChiSq_Dist_RT(arg1 float64, arg2 float64) float64 {
	retVal, _ := this.Call(0x00004210, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) ChiSq_Inv(arg1 float64, arg2 float64) float64 {
	retVal, _ := this.Call(0x00004211, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) ChiSq_Inv_RT(arg1 float64, arg2 float64) float64 {
	retVal, _ := this.Call(0x00004212, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) F_Dist(arg1 float64, arg2 float64, arg3 float64, arg4 bool) float64 {
	retVal, _ := this.Call(0x00004213, []interface{}{arg1, arg2, arg3, arg4})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) F_Dist_RT(arg1 float64, arg2 float64, arg3 float64) float64 {
	retVal, _ := this.Call(0x00004214, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) F_Inv(arg1 float64, arg2 float64, arg3 float64) float64 {
	retVal, _ := this.Call(0x00004215, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) F_Inv_RT(arg1 float64, arg2 float64, arg3 float64) float64 {
	retVal, _ := this.Call(0x00004216, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) HypGeom_Dist(arg1 float64, arg2 float64, arg3 float64, arg4 float64, arg5 bool) float64 {
	retVal, _ := this.Call(0x00004217, []interface{}{arg1, arg2, arg3, arg4, arg5})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) LogNorm_Dist(arg1 float64, arg2 float64, arg3 float64, arg4 bool) float64 {
	retVal, _ := this.Call(0x00004218, []interface{}{arg1, arg2, arg3, arg4})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) LogNorm_Inv(arg1 float64, arg2 float64, arg3 float64) float64 {
	retVal, _ := this.Call(0x00004219, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) NegBinom_Dist(arg1 float64, arg2 float64, arg3 float64, arg4 bool) float64 {
	retVal, _ := this.Call(0x0000421a, []interface{}{arg1, arg2, arg3, arg4})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Norm_S_Dist(arg1 float64, arg2 bool) float64 {
	retVal, _ := this.Call(0x0000421b, []interface{}{arg1, arg2})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Norm_S_Inv(arg1 float64) float64 {
	retVal, _ := this.Call(0x0000421c, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) T_Test(arg1 interface{}, arg2 interface{}, arg3 float64, arg4 float64) float64 {
	retVal, _ := this.Call(0x0000421d, []interface{}{arg1, arg2, arg3, arg4})
	return retVal.DblValVal()
}

var WorksheetFunction_Z_Test_OptArgs = []string{
	"Arg3",
}

func (this *WorksheetFunction) Z_Test(arg1 interface{}, arg2 float64, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Z_Test_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000421e, []interface{}{arg1, arg2}, optArgs...)
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Binom_Dist(arg1 float64, arg2 float64, arg3 float64, arg4 bool) float64 {
	retVal, _ := this.Call(0x000041e6, []interface{}{arg1, arg2, arg3, arg4})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Binom_Inv(arg1 float64, arg2 float64, arg3 float64) float64 {
	retVal, _ := this.Call(0x000041e7, []interface{}{arg1, arg2, arg3})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) Erf_Precise(arg1 interface{}) float64 {
	retVal, _ := this.Call(0x0000421f, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) ErfC_Precise(arg1 interface{}) float64 {
	retVal, _ := this.Call(0x00004220, []interface{}{arg1})
	return retVal.DblValVal()
}

func (this *WorksheetFunction) GammaLn_Precise(arg1 float64) float64 {
	retVal, _ := this.Call(0x00004221, []interface{}{arg1})
	return retVal.DblValVal()
}

var WorksheetFunction_Ceiling_Precise_OptArgs = []string{
	"Arg2",
}

func (this *WorksheetFunction) Ceiling_Precise(arg1 float64, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Ceiling_Precise_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004222, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}

var WorksheetFunction_Floor_Precise_OptArgs = []string{
	"Arg2",
}

func (this *WorksheetFunction) Floor_Precise(arg1 float64, optArgs ...interface{}) float64 {
	optArgs = ole.ProcessOptArgs(WorksheetFunction_Floor_Precise_OptArgs, optArgs)
	retVal, _ := this.Call(0x00004223, []interface{}{arg1}, optArgs...)
	return retVal.DblValVal()
}
