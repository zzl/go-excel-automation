package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002084A-0000-0000-C000-000000000046
var IID_AxisTitle = syscall.GUID{0x0002084A, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type AxisTitle struct {
	ole.OleClient
}

func NewAxisTitle(pDisp *win32.IDispatch, addRef bool, scoped bool) *AxisTitle {
	p := &AxisTitle{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func AxisTitleFromVar(v ole.Variant) *AxisTitle {
	return NewAxisTitle(v.PdispValVal(), false, false)
}

func (this *AxisTitle) IID() *syscall.GUID {
	return &IID_AxisTitle
}

func (this *AxisTitle) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *AxisTitle) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *AxisTitle) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *AxisTitle) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *AxisTitle) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *AxisTitle) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *AxisTitle) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *AxisTitle) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *AxisTitle) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *AxisTitle) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *AxisTitle) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *AxisTitle) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AxisTitle) Select() ole.Variant {
	retVal := this.Call(0x000000eb, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *AxisTitle) Border() *Border {
	retVal := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.PdispValVal(), false, true)
}

func (this *AxisTitle) Delete() ole.Variant {
	retVal := this.Call(0x00000075, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *AxisTitle) Interior() *Interior {
	retVal := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.PdispValVal(), false, true)
}

func (this *AxisTitle) Fill() *ChartFillFormat {
	retVal := this.PropGet(0x0000067f, nil)
	return NewChartFillFormat(retVal.PdispValVal(), false, true)
}

func (this *AxisTitle) Caption() string {
	retVal := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AxisTitle) SetCaption(rhs string)  {
	retVal := this.PropPut(0x0000008b, []interface{}{rhs})
	_= retVal
}

var AxisTitle_Characters_OptArgs= []string{
	"Start", "Length", 
}

func (this *AxisTitle) Characters(optArgs ...interface{}) *Characters {
	optArgs = ole.ProcessOptArgs(AxisTitle_Characters_OptArgs, optArgs)
	retVal := this.PropGet(0x0000025b, nil, optArgs...)
	return NewCharacters(retVal.PdispValVal(), false, true)
}

func (this *AxisTitle) Font() *Font {
	retVal := this.PropGet(0x00000092, nil)
	return NewFont(retVal.PdispValVal(), false, true)
}

func (this *AxisTitle) HorizontalAlignment() ole.Variant {
	retVal := this.PropGet(0x00000088, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *AxisTitle) SetHorizontalAlignment(rhs interface{})  {
	retVal := this.PropPut(0x00000088, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) Left() float64 {
	retVal := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *AxisTitle) SetLeft(rhs float64)  {
	retVal := this.PropPut(0x0000007f, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) Orientation() ole.Variant {
	retVal := this.PropGet(0x00000086, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *AxisTitle) SetOrientation(rhs interface{})  {
	retVal := this.PropPut(0x00000086, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) Shadow() bool {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AxisTitle) SetShadow(rhs bool)  {
	retVal := this.PropPut(0x00000067, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) Text() string {
	retVal := this.PropGet(0x0000008a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AxisTitle) SetText(rhs string)  {
	retVal := this.PropPut(0x0000008a, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) Top() float64 {
	retVal := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *AxisTitle) SetTop(rhs float64)  {
	retVal := this.PropPut(0x0000007e, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) VerticalAlignment() ole.Variant {
	retVal := this.PropGet(0x00000089, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *AxisTitle) SetVerticalAlignment(rhs interface{})  {
	retVal := this.PropPut(0x00000089, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) ReadingOrder() int32 {
	retVal := this.PropGet(0x000003cf, nil)
	return retVal.LValVal()
}

func (this *AxisTitle) SetReadingOrder(rhs int32)  {
	retVal := this.PropPut(0x000003cf, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) AutoScaleFont() ole.Variant {
	retVal := this.PropGet(0x000005f5, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *AxisTitle) SetAutoScaleFont(rhs interface{})  {
	retVal := this.PropPut(0x000005f5, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) IncludeInLayout() bool {
	retVal := this.PropGet(0x00000a58, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AxisTitle) SetIncludeInLayout(rhs bool)  {
	retVal := this.PropPut(0x00000a58, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) Position() int32 {
	retVal := this.PropGet(0x00000085, nil)
	return retVal.LValVal()
}

func (this *AxisTitle) SetPosition(rhs int32)  {
	retVal := this.PropPut(0x00000085, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) Format() *ChartFormat {
	retVal := this.PropGet(0x00000074, nil)
	return NewChartFormat(retVal.PdispValVal(), false, true)
}

func (this *AxisTitle) Height() float64 {
	retVal := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *AxisTitle) Width() float64 {
	retVal := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *AxisTitle) Formula() string {
	retVal := this.PropGet(0x00000105, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AxisTitle) SetFormula(rhs string)  {
	retVal := this.PropPut(0x00000105, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) FormulaR1C1() string {
	retVal := this.PropGet(0x00000108, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AxisTitle) SetFormulaR1C1(rhs string)  {
	retVal := this.PropPut(0x00000108, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) FormulaLocal() string {
	retVal := this.PropGet(0x00000107, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AxisTitle) SetFormulaLocal(rhs string)  {
	retVal := this.PropPut(0x00000107, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) FormulaR1C1Local() string {
	retVal := this.PropGet(0x00000109, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AxisTitle) SetFormulaR1C1Local(rhs string)  {
	retVal := this.PropPut(0x00000109, []interface{}{rhs})
	_= retVal
}

