package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020859-0000-0000-C000-000000000046
var IID_ChartGroup = syscall.GUID{0x00020859, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ChartGroup struct {
	ole.OleClient
}

func NewChartGroup(pDisp *win32.IDispatch, addRef bool, scoped bool) *ChartGroup {
	p := &ChartGroup{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ChartGroupFromVar(v ole.Variant) *ChartGroup {
	return NewChartGroup(v.PdispValVal(), false, false)
}

func (this *ChartGroup) IID() *syscall.GUID {
	return &IID_ChartGroup
}

func (this *ChartGroup) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ChartGroup) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *ChartGroup) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ChartGroup) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ChartGroup) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *ChartGroup) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *ChartGroup) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *ChartGroup) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *ChartGroup) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *ChartGroup) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ChartGroup) AxisGroup() int32 {
	retVal := this.PropGet(0x0000002f, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetAxisGroup(rhs int32)  {
	retVal := this.PropPut(0x0000002f, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) DoughnutHoleSize() int32 {
	retVal := this.PropGet(0x00000466, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetDoughnutHoleSize(rhs int32)  {
	retVal := this.PropPut(0x00000466, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) DownBars() *DownBars {
	retVal := this.PropGet(0x0000008d, nil)
	return NewDownBars(retVal.PdispValVal(), false, true)
}

func (this *ChartGroup) DropLines() *DropLines {
	retVal := this.PropGet(0x0000008e, nil)
	return NewDropLines(retVal.PdispValVal(), false, true)
}

func (this *ChartGroup) FirstSliceAngle() int32 {
	retVal := this.PropGet(0x0000003f, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetFirstSliceAngle(rhs int32)  {
	retVal := this.PropPut(0x0000003f, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) GapWidth() int32 {
	retVal := this.PropGet(0x00000033, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetGapWidth(rhs int32)  {
	retVal := this.PropPut(0x00000033, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) HasDropLines() bool {
	retVal := this.PropGet(0x0000003d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetHasDropLines(rhs bool)  {
	retVal := this.PropPut(0x0000003d, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) HasHiLoLines() bool {
	retVal := this.PropGet(0x0000003e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetHasHiLoLines(rhs bool)  {
	retVal := this.PropPut(0x0000003e, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) HasRadarAxisLabels() bool {
	retVal := this.PropGet(0x00000040, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetHasRadarAxisLabels(rhs bool)  {
	retVal := this.PropPut(0x00000040, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) HasSeriesLines() bool {
	retVal := this.PropGet(0x00000041, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetHasSeriesLines(rhs bool)  {
	retVal := this.PropPut(0x00000041, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) HasUpDownBars() bool {
	retVal := this.PropGet(0x00000042, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetHasUpDownBars(rhs bool)  {
	retVal := this.PropPut(0x00000042, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) HiLoLines() *HiLoLines {
	retVal := this.PropGet(0x0000008f, nil)
	return NewHiLoLines(retVal.PdispValVal(), false, true)
}

func (this *ChartGroup) Index() int32 {
	retVal := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) Overlap() int32 {
	retVal := this.PropGet(0x00000038, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetOverlap(rhs int32)  {
	retVal := this.PropPut(0x00000038, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) RadarAxisLabels() *TickLabels {
	retVal := this.PropGet(0x00000090, nil)
	return NewTickLabels(retVal.PdispValVal(), false, true)
}

var ChartGroup_SeriesCollection_OptArgs= []string{
	"Index", 
}

func (this *ChartGroup) SeriesCollection(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(ChartGroup_SeriesCollection_OptArgs, optArgs)
	retVal := this.Call(0x00000044, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ChartGroup) SeriesLines() *SeriesLines {
	retVal := this.PropGet(0x00000091, nil)
	return NewSeriesLines(retVal.PdispValVal(), false, true)
}

func (this *ChartGroup) SubType() int32 {
	retVal := this.PropGet(0x0000006d, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetSubType(rhs int32)  {
	retVal := this.PropPut(0x0000006d, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) Type() int32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetType(rhs int32)  {
	retVal := this.PropPut(0x0000006c, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) UpBars() *UpBars {
	retVal := this.PropGet(0x0000008c, nil)
	return NewUpBars(retVal.PdispValVal(), false, true)
}

func (this *ChartGroup) VaryByCategories() bool {
	retVal := this.PropGet(0x0000003c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetVaryByCategories(rhs bool)  {
	retVal := this.PropPut(0x0000003c, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) SizeRepresents() int32 {
	retVal := this.PropGet(0x00000674, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetSizeRepresents(rhs int32)  {
	retVal := this.PropPut(0x00000674, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) BubbleScale() int32 {
	retVal := this.PropGet(0x00000675, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetBubbleScale(rhs int32)  {
	retVal := this.PropPut(0x00000675, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) ShowNegativeBubbles() bool {
	retVal := this.PropGet(0x00000676, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetShowNegativeBubbles(rhs bool)  {
	retVal := this.PropPut(0x00000676, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) SplitType() int32 {
	retVal := this.PropGet(0x00000677, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetSplitType(rhs int32)  {
	retVal := this.PropPut(0x00000677, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) SplitValue() ole.Variant {
	retVal := this.PropGet(0x00000678, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartGroup) SetSplitValue(rhs interface{})  {
	retVal := this.PropPut(0x00000678, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) SecondPlotSize() int32 {
	retVal := this.PropGet(0x00000679, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetSecondPlotSize(rhs int32)  {
	retVal := this.PropPut(0x00000679, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) Has3DShading() bool {
	retVal := this.PropGet(0x0000067a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetHas3DShading(rhs bool)  {
	retVal := this.PropPut(0x0000067a, []interface{}{rhs})
	_= retVal
}

