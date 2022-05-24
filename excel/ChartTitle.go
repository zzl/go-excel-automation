package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020849-0000-0000-C000-000000000046
var IID_ChartTitle = syscall.GUID{0x00020849, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ChartTitle struct {
	ole.OleClient
}

func NewChartTitle(pDisp *win32.IDispatch, addRef bool, scoped bool) *ChartTitle {
	 if pDisp == nil {
		return nil;
	}
	p := &ChartTitle{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ChartTitleFromVar(v ole.Variant) *ChartTitle {
	return NewChartTitle(v.IDispatch(), false, false)
}

func (this *ChartTitle) IID() *syscall.GUID {
	return &IID_ChartTitle
}

func (this *ChartTitle) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ChartTitle) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *ChartTitle) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ChartTitle) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ChartTitle) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *ChartTitle) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *ChartTitle) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *ChartTitle) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *ChartTitle) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ChartTitle) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ChartTitle) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ChartTitle) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ChartTitle) Select() ole.Variant {
	retVal, _ := this.Call(0x000000eb, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartTitle) Border() *Border {
	retVal, _ := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.IDispatch(), false, true)
}

func (this *ChartTitle) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartTitle) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *ChartTitle) Fill() *ChartFillFormat {
	retVal, _ := this.PropGet(0x0000067f, nil)
	return NewChartFillFormat(retVal.IDispatch(), false, true)
}

func (this *ChartTitle) Caption() string {
	retVal, _ := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ChartTitle) SetCaption(rhs string)  {
	_ = this.PropPut(0x0000008b, []interface{}{rhs})
}

var ChartTitle_Characters_OptArgs= []string{
	"Start", "Length", 
}

func (this *ChartTitle) Characters(optArgs ...interface{}) *Characters {
	optArgs = ole.ProcessOptArgs(ChartTitle_Characters_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000025b, nil, optArgs...)
	return NewCharacters(retVal.IDispatch(), false, true)
}

func (this *ChartTitle) Font() *Font {
	retVal, _ := this.PropGet(0x00000092, nil)
	return NewFont(retVal.IDispatch(), false, true)
}

func (this *ChartTitle) HorizontalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x00000088, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartTitle) SetHorizontalAlignment(rhs interface{})  {
	_ = this.PropPut(0x00000088, []interface{}{rhs})
}

func (this *ChartTitle) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *ChartTitle) SetLeft(rhs float64)  {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *ChartTitle) Orientation() ole.Variant {
	retVal, _ := this.PropGet(0x00000086, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartTitle) SetOrientation(rhs interface{})  {
	_ = this.PropPut(0x00000086, []interface{}{rhs})
}

func (this *ChartTitle) Shadow() bool {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartTitle) SetShadow(rhs bool)  {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *ChartTitle) Text() string {
	retVal, _ := this.PropGet(0x0000008a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ChartTitle) SetText(rhs string)  {
	_ = this.PropPut(0x0000008a, []interface{}{rhs})
}

func (this *ChartTitle) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *ChartTitle) SetTop(rhs float64)  {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *ChartTitle) VerticalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x00000089, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartTitle) SetVerticalAlignment(rhs interface{})  {
	_ = this.PropPut(0x00000089, []interface{}{rhs})
}

func (this *ChartTitle) ReadingOrder() int32 {
	retVal, _ := this.PropGet(0x000003cf, nil)
	return retVal.LValVal()
}

func (this *ChartTitle) SetReadingOrder(rhs int32)  {
	_ = this.PropPut(0x000003cf, []interface{}{rhs})
}

func (this *ChartTitle) AutoScaleFont() ole.Variant {
	retVal, _ := this.PropGet(0x000005f5, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartTitle) SetAutoScaleFont(rhs interface{})  {
	_ = this.PropPut(0x000005f5, []interface{}{rhs})
}

func (this *ChartTitle) IncludeInLayout() bool {
	retVal, _ := this.PropGet(0x00000a58, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartTitle) SetIncludeInLayout(rhs bool)  {
	_ = this.PropPut(0x00000a58, []interface{}{rhs})
}

func (this *ChartTitle) Position() int32 {
	retVal, _ := this.PropGet(0x00000085, nil)
	return retVal.LValVal()
}

func (this *ChartTitle) SetPosition(rhs int32)  {
	_ = this.PropPut(0x00000085, []interface{}{rhs})
}

func (this *ChartTitle) Format() *ChartFormat {
	retVal, _ := this.PropGet(0x00000074, nil)
	return NewChartFormat(retVal.IDispatch(), false, true)
}

func (this *ChartTitle) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *ChartTitle) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *ChartTitle) Formula() string {
	retVal, _ := this.PropGet(0x00000105, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ChartTitle) SetFormula(rhs string)  {
	_ = this.PropPut(0x00000105, []interface{}{rhs})
}

func (this *ChartTitle) FormulaR1C1() string {
	retVal, _ := this.PropGet(0x00000108, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ChartTitle) SetFormulaR1C1(rhs string)  {
	_ = this.PropPut(0x00000108, []interface{}{rhs})
}

func (this *ChartTitle) FormulaLocal() string {
	retVal, _ := this.PropGet(0x00000107, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ChartTitle) SetFormulaLocal(rhs string)  {
	_ = this.PropPut(0x00000107, []interface{}{rhs})
}

func (this *ChartTitle) FormulaR1C1Local() string {
	retVal, _ := this.PropGet(0x00000109, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ChartTitle) SetFormulaR1C1Local(rhs string)  {
	_ = this.PropPut(0x00000109, []interface{}{rhs})
}

