package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000208BE-0000-0000-C000-000000000046
var IID_Trendline = syscall.GUID{0x000208BE, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Trendline struct {
	ole.OleClient
}

func NewTrendline(pDisp *win32.IDispatch, addRef bool, scoped bool) *Trendline {
	 if pDisp == nil {
		return nil;
	}
	p := &Trendline{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TrendlineFromVar(v ole.Variant) *Trendline {
	return NewTrendline(v.IDispatch(), false, false)
}

func (this *Trendline) IID() *syscall.GUID {
	return &IID_Trendline
}

func (this *Trendline) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Trendline) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Trendline) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Trendline) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Trendline) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Trendline) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Trendline) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Trendline) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Trendline) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Trendline) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Trendline) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Trendline) Backward() int32 {
	retVal, _ := this.PropGet(0x000000b9, nil)
	return retVal.LValVal()
}

func (this *Trendline) SetBackward(rhs int32)  {
	_ = this.PropPut(0x000000b9, []interface{}{rhs})
}

func (this *Trendline) Border() *Border {
	retVal, _ := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.IDispatch(), false, true)
}

func (this *Trendline) ClearFormats() ole.Variant {
	retVal, _ := this.Call(0x00000070, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Trendline) DataLabel() *DataLabel {
	retVal, _ := this.PropGet(0x0000009e, nil)
	return NewDataLabel(retVal.IDispatch(), false, true)
}

func (this *Trendline) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Trendline) DisplayEquation() bool {
	retVal, _ := this.PropGet(0x000000be, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Trendline) SetDisplayEquation(rhs bool)  {
	_ = this.PropPut(0x000000be, []interface{}{rhs})
}

func (this *Trendline) DisplayRSquared() bool {
	retVal, _ := this.PropGet(0x000000bd, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Trendline) SetDisplayRSquared(rhs bool)  {
	_ = this.PropPut(0x000000bd, []interface{}{rhs})
}

func (this *Trendline) Forward() int32 {
	retVal, _ := this.PropGet(0x000000bf, nil)
	return retVal.LValVal()
}

func (this *Trendline) SetForward(rhs int32)  {
	_ = this.PropPut(0x000000bf, []interface{}{rhs})
}

func (this *Trendline) Index() int32 {
	retVal, _ := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *Trendline) Intercept() float64 {
	retVal, _ := this.PropGet(0x000000ba, nil)
	return retVal.DblValVal()
}

func (this *Trendline) SetIntercept(rhs float64)  {
	_ = this.PropPut(0x000000ba, []interface{}{rhs})
}

func (this *Trendline) InterceptIsAuto() bool {
	retVal, _ := this.PropGet(0x000000bb, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Trendline) SetInterceptIsAuto(rhs bool)  {
	_ = this.PropPut(0x000000bb, []interface{}{rhs})
}

func (this *Trendline) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Trendline) SetName(rhs string)  {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *Trendline) NameIsAuto() bool {
	retVal, _ := this.PropGet(0x000000bc, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Trendline) SetNameIsAuto(rhs bool)  {
	_ = this.PropPut(0x000000bc, []interface{}{rhs})
}

func (this *Trendline) Order() int32 {
	retVal, _ := this.PropGet(0x000000c0, nil)
	return retVal.LValVal()
}

func (this *Trendline) SetOrder(rhs int32)  {
	_ = this.PropPut(0x000000c0, []interface{}{rhs})
}

func (this *Trendline) Period() int32 {
	retVal, _ := this.PropGet(0x000000b8, nil)
	return retVal.LValVal()
}

func (this *Trendline) SetPeriod(rhs int32)  {
	_ = this.PropPut(0x000000b8, []interface{}{rhs})
}

func (this *Trendline) Select() ole.Variant {
	retVal, _ := this.Call(0x000000eb, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Trendline) Type() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *Trendline) SetType(rhs int32)  {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

func (this *Trendline) Backward2() float64 {
	retVal, _ := this.PropGet(0x00000a5a, nil)
	return retVal.DblValVal()
}

func (this *Trendline) SetBackward2(rhs float64)  {
	_ = this.PropPut(0x00000a5a, []interface{}{rhs})
}

func (this *Trendline) Forward2() float64 {
	retVal, _ := this.PropGet(0x00000a5b, nil)
	return retVal.DblValVal()
}

func (this *Trendline) SetForward2(rhs float64)  {
	_ = this.PropPut(0x00000a5b, []interface{}{rhs})
}

func (this *Trendline) Format() *ChartFormat {
	retVal, _ := this.PropGet(0x00000074, nil)
	return NewChartFormat(retVal.IDispatch(), false, true)
}

