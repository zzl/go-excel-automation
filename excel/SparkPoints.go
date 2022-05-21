package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244B8-0000-0000-C000-000000000046
var IID_SparkPoints = syscall.GUID{0x000244B8, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type SparkPoints struct {
	ole.OleClient
}

func NewSparkPoints(pDisp *win32.IDispatch, addRef bool, scoped bool) *SparkPoints {
	p := &SparkPoints{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SparkPointsFromVar(v ole.Variant) *SparkPoints {
	return NewSparkPoints(v.PdispValVal(), false, false)
}

func (this *SparkPoints) IID() *syscall.GUID {
	return &IID_SparkPoints
}

func (this *SparkPoints) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SparkPoints) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *SparkPoints) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *SparkPoints) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *SparkPoints) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *SparkPoints) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *SparkPoints) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *SparkPoints) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *SparkPoints) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *SparkPoints) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *SparkPoints) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *SparkPoints) Negative() *SparkColor {
	retVal := this.PropGet(0x00000b8b, nil)
	return NewSparkColor(retVal.PdispValVal(), false, true)
}

func (this *SparkPoints) Markers() *SparkColor {
	retVal := this.PropGet(0x00000b8c, nil)
	return NewSparkColor(retVal.PdispValVal(), false, true)
}

func (this *SparkPoints) Highpoint() *SparkColor {
	retVal := this.PropGet(0x00000b8d, nil)
	return NewSparkColor(retVal.PdispValVal(), false, true)
}

func (this *SparkPoints) Lowpoint() *SparkColor {
	retVal := this.PropGet(0x00000b8e, nil)
	return NewSparkColor(retVal.PdispValVal(), false, true)
}

func (this *SparkPoints) Firstpoint() *SparkColor {
	retVal := this.PropGet(0x00000b8f, nil)
	return NewSparkColor(retVal.PdispValVal(), false, true)
}

func (this *SparkPoints) Lastpoint() *SparkColor {
	retVal := this.PropGet(0x00000b90, nil)
	return NewSparkColor(retVal.PdispValVal(), false, true)
}

