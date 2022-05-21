package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244B7-0000-0000-C000-000000000046
var IID_SparklineGroup = syscall.GUID{0x000244B7, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type SparklineGroup struct {
	ole.OleClient
}

func NewSparklineGroup(pDisp *win32.IDispatch, addRef bool, scoped bool) *SparklineGroup {
	p := &SparklineGroup{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SparklineGroupFromVar(v ole.Variant) *SparklineGroup {
	return NewSparklineGroup(v.PdispValVal(), false, false)
}

func (this *SparklineGroup) IID() *syscall.GUID {
	return &IID_SparklineGroup
}

func (this *SparklineGroup) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SparklineGroup) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *SparklineGroup) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *SparklineGroup) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *SparklineGroup) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *SparklineGroup) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *SparklineGroup) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *SparklineGroup) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *SparklineGroup) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *SparklineGroup) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *SparklineGroup) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *SparklineGroup) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *SparklineGroup) Item(index interface{}) *Sparkline {
	retVal := this.PropGet(0x000000aa, []interface{}{index})
	return NewSparkline(retVal.PdispValVal(), false, true)
}

func (this *SparklineGroup) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *SparklineGroup) ForEach(action func(item *Sparkline) bool) {
	pEnum := this.NewEnum_()
	var pEnumVar *win32.IEnumVARIANT
	pEnum.QueryInterface(&win32.IID_IEnumVARIANT, unsafe.Pointer(&pEnumVar))
	defer pEnumVar.Release();
	for {
		var c uint32
		var v ole.Variant
		pEnumVar.Next(1, (*win32.VARIANT)(&v), &c)
		if c == 0 {
			break
		}
		pItem := (*Sparkline)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *SparklineGroup) Location() *Range {
	retVal := this.PropGet(0x00000575, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *SparklineGroup) SetLocation(rhs *Range)  {
	retVal := this.PropPutRef(0x00000575, []interface{}{rhs})
	_= retVal
}

func (this *SparklineGroup) SourceData() string {
	retVal := this.PropGet(0x000002ae, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SparklineGroup) SetSourceData(rhs string)  {
	retVal := this.PropPut(0x000002ae, []interface{}{rhs})
	_= retVal
}

func (this *SparklineGroup) DateRange() string {
	retVal := this.PropGet(0x00000b84, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SparklineGroup) SetDateRange(rhs string)  {
	retVal := this.PropPut(0x00000b84, []interface{}{rhs})
	_= retVal
}

func (this *SparklineGroup) ModifyLocation(location *Range)  {
	retVal := this.Call(0x00000b85, []interface{}{location})
	_= retVal
}

func (this *SparklineGroup) ModifySourceData(sourceData string)  {
	retVal := this.Call(0x00000b86, []interface{}{sourceData})
	_= retVal
}

func (this *SparklineGroup) Modify(location *Range, sourceData string)  {
	retVal := this.Call(0x0000062d, []interface{}{location, sourceData})
	_= retVal
}

func (this *SparklineGroup) ModifyDateRange(dateRange string)  {
	retVal := this.Call(0x00000b87, []interface{}{dateRange})
	_= retVal
}

func (this *SparklineGroup) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

func (this *SparklineGroup) Type() int32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *SparklineGroup) SetType(rhs int32)  {
	retVal := this.PropPut(0x0000006c, []interface{}{rhs})
	_= retVal
}

func (this *SparklineGroup) SeriesColor() *FormatColor {
	retVal := this.PropGet(0x00000b88, nil)
	return NewFormatColor(retVal.PdispValVal(), false, true)
}

func (this *SparklineGroup) Points() *SparkPoints {
	retVal := this.PropGet(0x00000046, nil)
	return NewSparkPoints(retVal.PdispValVal(), false, true)
}

func (this *SparklineGroup) Axes() *SparkAxes {
	retVal := this.PropGet(0x00000017, nil)
	return NewSparkAxes(retVal.PdispValVal(), false, true)
}

func (this *SparklineGroup) DisplayBlanksAs() int32 {
	retVal := this.PropGet(0x0000005d, nil)
	return retVal.LValVal()
}

func (this *SparklineGroup) SetDisplayBlanksAs(rhs int32)  {
	retVal := this.PropPut(0x0000005d, []interface{}{rhs})
	_= retVal
}

func (this *SparklineGroup) DisplayHidden() bool {
	retVal := this.PropGet(0x00000b89, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SparklineGroup) SetDisplayHidden(rhs bool)  {
	retVal := this.PropPut(0x00000b89, []interface{}{rhs})
	_= retVal
}

func (this *SparklineGroup) LineWeight() ole.Variant {
	retVal := this.PropGet(0x00000b8a, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *SparklineGroup) SetLineWeight(rhs interface{})  {
	retVal := this.PropPut(0x00000b8a, []interface{}{rhs})
	_= retVal
}

func (this *SparklineGroup) PlotBy() int32 {
	retVal := this.PropGet(0x000000ca, nil)
	return retVal.LValVal()
}

func (this *SparklineGroup) SetPlotBy(rhs int32)  {
	retVal := this.PropPut(0x000000ca, []interface{}{rhs})
	_= retVal
}

