package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244BB-0000-0000-C000-000000000046
var IID_SparkHorizontalAxis = syscall.GUID{0x000244BB, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type SparkHorizontalAxis struct {
	ole.OleClient
}

func NewSparkHorizontalAxis(pDisp *win32.IDispatch, addRef bool, scoped bool) *SparkHorizontalAxis {
	p := &SparkHorizontalAxis{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SparkHorizontalAxisFromVar(v ole.Variant) *SparkHorizontalAxis {
	return NewSparkHorizontalAxis(v.PdispValVal(), false, false)
}

func (this *SparkHorizontalAxis) IID() *syscall.GUID {
	return &IID_SparkHorizontalAxis
}

func (this *SparkHorizontalAxis) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SparkHorizontalAxis) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *SparkHorizontalAxis) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *SparkHorizontalAxis) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *SparkHorizontalAxis) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *SparkHorizontalAxis) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *SparkHorizontalAxis) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *SparkHorizontalAxis) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *SparkHorizontalAxis) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *SparkHorizontalAxis) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *SparkHorizontalAxis) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *SparkHorizontalAxis) Axis() *SparkColor {
	retVal := this.PropGet(0x00000043, nil)
	return NewSparkColor(retVal.PdispValVal(), false, true)
}

func (this *SparkHorizontalAxis) IsDateAxis() bool {
	retVal := this.PropGet(0x00000b93, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SparkHorizontalAxis) RightToLeftPlotOrder() bool {
	retVal := this.PropGet(0x00000b94, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SparkHorizontalAxis) SetRightToLeftPlotOrder(rhs bool)  {
	retVal := this.PropPut(0x00000b94, []interface{}{rhs})
	_= retVal
}

