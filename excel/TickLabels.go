package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000208C9-0000-0000-C000-000000000046
var IID_TickLabels = syscall.GUID{0x000208C9, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TickLabels struct {
	ole.OleClient
}

func NewTickLabels(pDisp *win32.IDispatch, addRef bool, scoped bool) *TickLabels {
	p := &TickLabels{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TickLabelsFromVar(v ole.Variant) *TickLabels {
	return NewTickLabels(v.PdispValVal(), false, false)
}

func (this *TickLabels) IID() *syscall.GUID {
	return &IID_TickLabels
}

func (this *TickLabels) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TickLabels) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *TickLabels) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *TickLabels) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *TickLabels) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *TickLabels) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *TickLabels) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *TickLabels) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *TickLabels) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *TickLabels) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *TickLabels) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *TickLabels) Delete() ole.Variant {
	retVal := this.Call(0x00000075, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *TickLabels) Font() *Font {
	retVal := this.PropGet(0x00000092, nil)
	return NewFont(retVal.PdispValVal(), false, true)
}

func (this *TickLabels) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TickLabels) NumberFormat() string {
	retVal := this.PropGet(0x000000c1, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TickLabels) SetNumberFormat(rhs string)  {
	retVal := this.PropPut(0x000000c1, []interface{}{rhs})
	_= retVal
}

func (this *TickLabels) NumberFormatLinked() bool {
	retVal := this.PropGet(0x000000c2, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TickLabels) SetNumberFormatLinked(rhs bool)  {
	retVal := this.PropPut(0x000000c2, []interface{}{rhs})
	_= retVal
}

func (this *TickLabels) NumberFormatLocal() ole.Variant {
	retVal := this.PropGet(0x00000449, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *TickLabels) SetNumberFormatLocal(rhs interface{})  {
	retVal := this.PropPut(0x00000449, []interface{}{rhs})
	_= retVal
}

func (this *TickLabels) Orientation() int32 {
	retVal := this.PropGet(0x00000086, nil)
	return retVal.LValVal()
}

func (this *TickLabels) SetOrientation(rhs int32)  {
	retVal := this.PropPut(0x00000086, []interface{}{rhs})
	_= retVal
}

func (this *TickLabels) Select() ole.Variant {
	retVal := this.Call(0x000000eb, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *TickLabels) ReadingOrder() int32 {
	retVal := this.PropGet(0x000003cf, nil)
	return retVal.LValVal()
}

func (this *TickLabels) SetReadingOrder(rhs int32)  {
	retVal := this.PropPut(0x000003cf, []interface{}{rhs})
	_= retVal
}

func (this *TickLabels) AutoScaleFont() ole.Variant {
	retVal := this.PropGet(0x000005f5, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *TickLabels) SetAutoScaleFont(rhs interface{})  {
	retVal := this.PropPut(0x000005f5, []interface{}{rhs})
	_= retVal
}

func (this *TickLabels) Depth() int32 {
	retVal := this.PropGet(0x00000762, nil)
	return retVal.LValVal()
}

func (this *TickLabels) Offset() int32 {
	retVal := this.PropGet(0x000000fe, nil)
	return retVal.LValVal()
}

func (this *TickLabels) SetOffset(rhs int32)  {
	retVal := this.PropPut(0x000000fe, []interface{}{rhs})
	_= retVal
}

func (this *TickLabels) Alignment() int32 {
	retVal := this.PropGet(0x000001c5, nil)
	return retVal.LValVal()
}

func (this *TickLabels) SetAlignment(rhs int32)  {
	retVal := this.PropPut(0x000001c5, []interface{}{rhs})
	_= retVal
}

func (this *TickLabels) MultiLevel() bool {
	retVal := this.PropGet(0x00000a5d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TickLabels) SetMultiLevel(rhs bool)  {
	retVal := this.PropPut(0x00000a5d, []interface{}{rhs})
	_= retVal
}

func (this *TickLabels) Format() *ChartFormat {
	retVal := this.PropGet(0x00000074, nil)
	return NewChartFormat(retVal.PdispValVal(), false, true)
}
