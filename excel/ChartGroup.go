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
	 if pDisp == nil {
		return nil;
	}
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
	return NewChartGroup(v.IDispatch(), false, false)
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
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *ChartGroup) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ChartGroup) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ChartGroup) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *ChartGroup) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *ChartGroup) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *ChartGroup) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *ChartGroup) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ChartGroup) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ChartGroup) AxisGroup() int32 {
	retVal, _ := this.PropGet(0x0000002f, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetAxisGroup(rhs int32)  {
	_ = this.PropPut(0x0000002f, []interface{}{rhs})
}

func (this *ChartGroup) DoughnutHoleSize() int32 {
	retVal, _ := this.PropGet(0x00000466, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetDoughnutHoleSize(rhs int32)  {
	_ = this.PropPut(0x00000466, []interface{}{rhs})
}

func (this *ChartGroup) DownBars() *DownBars {
	retVal, _ := this.PropGet(0x0000008d, nil)
	return NewDownBars(retVal.IDispatch(), false, true)
}

func (this *ChartGroup) DropLines() *DropLines {
	retVal, _ := this.PropGet(0x0000008e, nil)
	return NewDropLines(retVal.IDispatch(), false, true)
}

func (this *ChartGroup) FirstSliceAngle() int32 {
	retVal, _ := this.PropGet(0x0000003f, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetFirstSliceAngle(rhs int32)  {
	_ = this.PropPut(0x0000003f, []interface{}{rhs})
}

func (this *ChartGroup) GapWidth() int32 {
	retVal, _ := this.PropGet(0x00000033, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetGapWidth(rhs int32)  {
	_ = this.PropPut(0x00000033, []interface{}{rhs})
}

func (this *ChartGroup) HasDropLines() bool {
	retVal, _ := this.PropGet(0x0000003d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetHasDropLines(rhs bool)  {
	_ = this.PropPut(0x0000003d, []interface{}{rhs})
}

func (this *ChartGroup) HasHiLoLines() bool {
	retVal, _ := this.PropGet(0x0000003e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetHasHiLoLines(rhs bool)  {
	_ = this.PropPut(0x0000003e, []interface{}{rhs})
}

func (this *ChartGroup) HasRadarAxisLabels() bool {
	retVal, _ := this.PropGet(0x00000040, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetHasRadarAxisLabels(rhs bool)  {
	_ = this.PropPut(0x00000040, []interface{}{rhs})
}

func (this *ChartGroup) HasSeriesLines() bool {
	retVal, _ := this.PropGet(0x00000041, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetHasSeriesLines(rhs bool)  {
	_ = this.PropPut(0x00000041, []interface{}{rhs})
}

func (this *ChartGroup) HasUpDownBars() bool {
	retVal, _ := this.PropGet(0x00000042, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetHasUpDownBars(rhs bool)  {
	_ = this.PropPut(0x00000042, []interface{}{rhs})
}

func (this *ChartGroup) HiLoLines() *HiLoLines {
	retVal, _ := this.PropGet(0x0000008f, nil)
	return NewHiLoLines(retVal.IDispatch(), false, true)
}

func (this *ChartGroup) Index() int32 {
	retVal, _ := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) Overlap() int32 {
	retVal, _ := this.PropGet(0x00000038, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetOverlap(rhs int32)  {
	_ = this.PropPut(0x00000038, []interface{}{rhs})
}

func (this *ChartGroup) RadarAxisLabels() *TickLabels {
	retVal, _ := this.PropGet(0x00000090, nil)
	return NewTickLabels(retVal.IDispatch(), false, true)
}

var ChartGroup_SeriesCollection_OptArgs= []string{
	"Index", 
}

func (this *ChartGroup) SeriesCollection(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(ChartGroup_SeriesCollection_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000044, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ChartGroup) SeriesLines() *SeriesLines {
	retVal, _ := this.PropGet(0x00000091, nil)
	return NewSeriesLines(retVal.IDispatch(), false, true)
}

func (this *ChartGroup) SubType() int32 {
	retVal, _ := this.PropGet(0x0000006d, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetSubType(rhs int32)  {
	_ = this.PropPut(0x0000006d, []interface{}{rhs})
}

func (this *ChartGroup) Type() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetType(rhs int32)  {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

func (this *ChartGroup) UpBars() *UpBars {
	retVal, _ := this.PropGet(0x0000008c, nil)
	return NewUpBars(retVal.IDispatch(), false, true)
}

func (this *ChartGroup) VaryByCategories() bool {
	retVal, _ := this.PropGet(0x0000003c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetVaryByCategories(rhs bool)  {
	_ = this.PropPut(0x0000003c, []interface{}{rhs})
}

func (this *ChartGroup) SizeRepresents() int32 {
	retVal, _ := this.PropGet(0x00000674, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetSizeRepresents(rhs int32)  {
	_ = this.PropPut(0x00000674, []interface{}{rhs})
}

func (this *ChartGroup) BubbleScale() int32 {
	retVal, _ := this.PropGet(0x00000675, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetBubbleScale(rhs int32)  {
	_ = this.PropPut(0x00000675, []interface{}{rhs})
}

func (this *ChartGroup) ShowNegativeBubbles() bool {
	retVal, _ := this.PropGet(0x00000676, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetShowNegativeBubbles(rhs bool)  {
	_ = this.PropPut(0x00000676, []interface{}{rhs})
}

func (this *ChartGroup) SplitType() int32 {
	retVal, _ := this.PropGet(0x00000677, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetSplitType(rhs int32)  {
	_ = this.PropPut(0x00000677, []interface{}{rhs})
}

func (this *ChartGroup) SplitValue() ole.Variant {
	retVal, _ := this.PropGet(0x00000678, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartGroup) SetSplitValue(rhs interface{})  {
	_ = this.PropPut(0x00000678, []interface{}{rhs})
}

func (this *ChartGroup) SecondPlotSize() int32 {
	retVal, _ := this.PropGet(0x00000679, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetSecondPlotSize(rhs int32)  {
	_ = this.PropPut(0x00000679, []interface{}{rhs})
}

func (this *ChartGroup) Has3DShading() bool {
	retVal, _ := this.PropGet(0x0000067a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetHas3DShading(rhs bool)  {
	_ = this.PropPut(0x0000067a, []interface{}{rhs})
}

