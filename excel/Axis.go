package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020848-0000-0000-C000-000000000046
var IID_Axis = syscall.GUID{0x00020848, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Axis struct {
	ole.OleClient
}

func NewAxis(pDisp *win32.IDispatch, addRef bool, scoped bool) *Axis {
	 if pDisp == nil {
		return nil;
	}
	p := &Axis{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func AxisFromVar(v ole.Variant) *Axis {
	return NewAxis(v.IDispatch(), false, false)
}

func (this *Axis) IID() *syscall.GUID {
	return &IID_Axis
}

func (this *Axis) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Axis) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Axis) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Axis) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Axis) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Axis) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Axis) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Axis) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Axis) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Axis) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Axis) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Axis) AxisBetweenCategories() bool {
	retVal, _ := this.PropGet(0x0000002d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetAxisBetweenCategories(rhs bool)  {
	_ = this.PropPut(0x0000002d, []interface{}{rhs})
}

func (this *Axis) AxisGroup() int32 {
	retVal, _ := this.PropGet(0x0000002f, nil)
	return retVal.LValVal()
}

func (this *Axis) AxisTitle() *AxisTitle {
	retVal, _ := this.PropGet(0x00000052, nil)
	return NewAxisTitle(retVal.IDispatch(), false, true)
}

func (this *Axis) Border() *Border {
	retVal, _ := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.IDispatch(), false, true)
}

func (this *Axis) CategoryNames() ole.Variant {
	retVal, _ := this.PropGet(0x0000009c, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Axis) SetCategoryNames(rhs interface{})  {
	_ = this.PropPut(0x0000009c, []interface{}{rhs})
}

func (this *Axis) Crosses() int32 {
	retVal, _ := this.PropGet(0x0000002a, nil)
	return retVal.LValVal()
}

func (this *Axis) SetCrosses(rhs int32)  {
	_ = this.PropPut(0x0000002a, []interface{}{rhs})
}

func (this *Axis) CrossesAt() float64 {
	retVal, _ := this.PropGet(0x0000002b, nil)
	return retVal.DblValVal()
}

func (this *Axis) SetCrossesAt(rhs float64)  {
	_ = this.PropPut(0x0000002b, []interface{}{rhs})
}

func (this *Axis) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Axis) HasMajorGridlines() bool {
	retVal, _ := this.PropGet(0x00000018, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetHasMajorGridlines(rhs bool)  {
	_ = this.PropPut(0x00000018, []interface{}{rhs})
}

func (this *Axis) HasMinorGridlines() bool {
	retVal, _ := this.PropGet(0x00000019, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetHasMinorGridlines(rhs bool)  {
	_ = this.PropPut(0x00000019, []interface{}{rhs})
}

func (this *Axis) HasTitle() bool {
	retVal, _ := this.PropGet(0x00000036, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetHasTitle(rhs bool)  {
	_ = this.PropPut(0x00000036, []interface{}{rhs})
}

func (this *Axis) MajorGridlines() *Gridlines {
	retVal, _ := this.PropGet(0x00000059, nil)
	return NewGridlines(retVal.IDispatch(), false, true)
}

func (this *Axis) MajorTickMark() int32 {
	retVal, _ := this.PropGet(0x0000001a, nil)
	return retVal.LValVal()
}

func (this *Axis) SetMajorTickMark(rhs int32)  {
	_ = this.PropPut(0x0000001a, []interface{}{rhs})
}

func (this *Axis) MajorUnit() float64 {
	retVal, _ := this.PropGet(0x00000025, nil)
	return retVal.DblValVal()
}

func (this *Axis) SetMajorUnit(rhs float64)  {
	_ = this.PropPut(0x00000025, []interface{}{rhs})
}

func (this *Axis) MajorUnitIsAuto() bool {
	retVal, _ := this.PropGet(0x00000026, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetMajorUnitIsAuto(rhs bool)  {
	_ = this.PropPut(0x00000026, []interface{}{rhs})
}

func (this *Axis) MaximumScale() float64 {
	retVal, _ := this.PropGet(0x00000023, nil)
	return retVal.DblValVal()
}

func (this *Axis) SetMaximumScale(rhs float64)  {
	_ = this.PropPut(0x00000023, []interface{}{rhs})
}

func (this *Axis) MaximumScaleIsAuto() bool {
	retVal, _ := this.PropGet(0x00000024, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetMaximumScaleIsAuto(rhs bool)  {
	_ = this.PropPut(0x00000024, []interface{}{rhs})
}

func (this *Axis) MinimumScale() float64 {
	retVal, _ := this.PropGet(0x00000021, nil)
	return retVal.DblValVal()
}

func (this *Axis) SetMinimumScale(rhs float64)  {
	_ = this.PropPut(0x00000021, []interface{}{rhs})
}

func (this *Axis) MinimumScaleIsAuto() bool {
	retVal, _ := this.PropGet(0x00000022, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetMinimumScaleIsAuto(rhs bool)  {
	_ = this.PropPut(0x00000022, []interface{}{rhs})
}

func (this *Axis) MinorGridlines() *Gridlines {
	retVal, _ := this.PropGet(0x0000005a, nil)
	return NewGridlines(retVal.IDispatch(), false, true)
}

func (this *Axis) MinorTickMark() int32 {
	retVal, _ := this.PropGet(0x0000001b, nil)
	return retVal.LValVal()
}

func (this *Axis) SetMinorTickMark(rhs int32)  {
	_ = this.PropPut(0x0000001b, []interface{}{rhs})
}

func (this *Axis) MinorUnit() float64 {
	retVal, _ := this.PropGet(0x00000027, nil)
	return retVal.DblValVal()
}

func (this *Axis) SetMinorUnit(rhs float64)  {
	_ = this.PropPut(0x00000027, []interface{}{rhs})
}

func (this *Axis) MinorUnitIsAuto() bool {
	retVal, _ := this.PropGet(0x00000028, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetMinorUnitIsAuto(rhs bool)  {
	_ = this.PropPut(0x00000028, []interface{}{rhs})
}

func (this *Axis) ReversePlotOrder() bool {
	retVal, _ := this.PropGet(0x0000002c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetReversePlotOrder(rhs bool)  {
	_ = this.PropPut(0x0000002c, []interface{}{rhs})
}

func (this *Axis) ScaleType() int32 {
	retVal, _ := this.PropGet(0x00000029, nil)
	return retVal.LValVal()
}

func (this *Axis) SetScaleType(rhs int32)  {
	_ = this.PropPut(0x00000029, []interface{}{rhs})
}

func (this *Axis) Select() ole.Variant {
	retVal, _ := this.Call(0x000000eb, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Axis) TickLabelPosition() int32 {
	retVal, _ := this.PropGet(0x0000001c, nil)
	return retVal.LValVal()
}

func (this *Axis) SetTickLabelPosition(rhs int32)  {
	_ = this.PropPut(0x0000001c, []interface{}{rhs})
}

func (this *Axis) TickLabels() *TickLabels {
	retVal, _ := this.PropGet(0x0000005b, nil)
	return NewTickLabels(retVal.IDispatch(), false, true)
}

func (this *Axis) TickLabelSpacing() int32 {
	retVal, _ := this.PropGet(0x0000001d, nil)
	return retVal.LValVal()
}

func (this *Axis) SetTickLabelSpacing(rhs int32)  {
	_ = this.PropPut(0x0000001d, []interface{}{rhs})
}

func (this *Axis) TickMarkSpacing() int32 {
	retVal, _ := this.PropGet(0x0000001f, nil)
	return retVal.LValVal()
}

func (this *Axis) SetTickMarkSpacing(rhs int32)  {
	_ = this.PropPut(0x0000001f, []interface{}{rhs})
}

func (this *Axis) Type() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *Axis) SetType(rhs int32)  {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

func (this *Axis) BaseUnit() int32 {
	retVal, _ := this.PropGet(0x0000066f, nil)
	return retVal.LValVal()
}

func (this *Axis) SetBaseUnit(rhs int32)  {
	_ = this.PropPut(0x0000066f, []interface{}{rhs})
}

func (this *Axis) BaseUnitIsAuto() bool {
	retVal, _ := this.PropGet(0x00000670, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetBaseUnitIsAuto(rhs bool)  {
	_ = this.PropPut(0x00000670, []interface{}{rhs})
}

func (this *Axis) MajorUnitScale() int32 {
	retVal, _ := this.PropGet(0x00000671, nil)
	return retVal.LValVal()
}

func (this *Axis) SetMajorUnitScale(rhs int32)  {
	_ = this.PropPut(0x00000671, []interface{}{rhs})
}

func (this *Axis) MinorUnitScale() int32 {
	retVal, _ := this.PropGet(0x00000672, nil)
	return retVal.LValVal()
}

func (this *Axis) SetMinorUnitScale(rhs int32)  {
	_ = this.PropPut(0x00000672, []interface{}{rhs})
}

func (this *Axis) CategoryType() int32 {
	retVal, _ := this.PropGet(0x00000673, nil)
	return retVal.LValVal()
}

func (this *Axis) SetCategoryType(rhs int32)  {
	_ = this.PropPut(0x00000673, []interface{}{rhs})
}

func (this *Axis) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *Axis) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *Axis) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *Axis) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *Axis) DisplayUnit() int32 {
	retVal, _ := this.PropGet(0x0000075e, nil)
	return retVal.LValVal()
}

func (this *Axis) SetDisplayUnit(rhs int32)  {
	_ = this.PropPut(0x0000075e, []interface{}{rhs})
}

func (this *Axis) DisplayUnitCustom() float64 {
	retVal, _ := this.PropGet(0x0000075f, nil)
	return retVal.DblValVal()
}

func (this *Axis) SetDisplayUnitCustom(rhs float64)  {
	_ = this.PropPut(0x0000075f, []interface{}{rhs})
}

func (this *Axis) HasDisplayUnitLabel() bool {
	retVal, _ := this.PropGet(0x00000760, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetHasDisplayUnitLabel(rhs bool)  {
	_ = this.PropPut(0x00000760, []interface{}{rhs})
}

func (this *Axis) DisplayUnitLabel() *DisplayUnitLabel {
	retVal, _ := this.PropGet(0x00000761, nil)
	return NewDisplayUnitLabel(retVal.IDispatch(), false, true)
}

func (this *Axis) LogBase() float64 {
	retVal, _ := this.PropGet(0x00000a56, nil)
	return retVal.DblValVal()
}

func (this *Axis) SetLogBase(rhs float64)  {
	_ = this.PropPut(0x00000a56, []interface{}{rhs})
}

func (this *Axis) TickLabelSpacingIsAuto() bool {
	retVal, _ := this.PropGet(0x00000a57, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetTickLabelSpacingIsAuto(rhs bool)  {
	_ = this.PropPut(0x00000a57, []interface{}{rhs})
}

func (this *Axis) Format() *ChartFormat {
	retVal, _ := this.PropGet(0x00000074, nil)
	return NewChartFormat(retVal.IDispatch(), false, true)
}

