package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244BC-0000-0000-C000-000000000046
var IID_SparkVerticalAxis = syscall.GUID{0x000244BC, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type SparkVerticalAxis struct {
	ole.OleClient
}

func NewSparkVerticalAxis(pDisp *win32.IDispatch, addRef bool, scoped bool) *SparkVerticalAxis {
	p := &SparkVerticalAxis{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SparkVerticalAxisFromVar(v ole.Variant) *SparkVerticalAxis {
	return NewSparkVerticalAxis(v.PdispValVal(), false, false)
}

func (this *SparkVerticalAxis) IID() *syscall.GUID {
	return &IID_SparkVerticalAxis
}

func (this *SparkVerticalAxis) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SparkVerticalAxis) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *SparkVerticalAxis) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *SparkVerticalAxis) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *SparkVerticalAxis) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *SparkVerticalAxis) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *SparkVerticalAxis) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *SparkVerticalAxis) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *SparkVerticalAxis) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *SparkVerticalAxis) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *SparkVerticalAxis) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *SparkVerticalAxis) MinScaleType() int32 {
	retVal := this.PropGet(0x00000b95, nil)
	return retVal.LValVal()
}

func (this *SparkVerticalAxis) SetMinScaleType(rhs int32)  {
	retVal := this.PropPut(0x00000b95, []interface{}{rhs})
	_= retVal
}

func (this *SparkVerticalAxis) CustomMinScaleValue() ole.Variant {
	retVal := this.PropGet(0x00000b96, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *SparkVerticalAxis) SetCustomMinScaleValue(rhs interface{})  {
	retVal := this.PropPut(0x00000b96, []interface{}{rhs})
	_= retVal
}

func (this *SparkVerticalAxis) MaxScaleType() int32 {
	retVal := this.PropGet(0x00000b97, nil)
	return retVal.LValVal()
}

func (this *SparkVerticalAxis) SetMaxScaleType(rhs int32)  {
	retVal := this.PropPut(0x00000b97, []interface{}{rhs})
	_= retVal
}

func (this *SparkVerticalAxis) CustomMaxScaleValue() ole.Variant {
	retVal := this.PropGet(0x00000b98, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *SparkVerticalAxis) SetCustomMaxScaleValue(rhs interface{})  {
	retVal := this.PropPut(0x00000b98, []interface{}{rhs})
	_= retVal
}

