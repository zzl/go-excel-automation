package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244B2-0000-0000-C000-000000000046
var IID_ChartFormat = syscall.GUID{0x000244B2, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ChartFormat struct {
	ole.OleClient
}

func NewChartFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *ChartFormat {
	p := &ChartFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ChartFormatFromVar(v ole.Variant) *ChartFormat {
	return NewChartFormat(v.PdispValVal(), false, false)
}

func (this *ChartFormat) IID() *syscall.GUID {
	return &IID_ChartFormat
}

func (this *ChartFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ChartFormat) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *ChartFormat) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ChartFormat) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ChartFormat) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *ChartFormat) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *ChartFormat) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *ChartFormat) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *ChartFormat) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *ChartFormat) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ChartFormat) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ChartFormat) Fill() *FillFormat {
	retVal := this.PropGet(0x0000067f, nil)
	return NewFillFormat(retVal.PdispValVal(), false, true)
}

func (this *ChartFormat) Glow() *ole.DispatchClass {
	retVal := this.PropGet(0x00000a67, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ChartFormat) Line() *LineFormat {
	retVal := this.PropGet(0x00000331, nil)
	return NewLineFormat(retVal.PdispValVal(), false, true)
}

func (this *ChartFormat) PictureFormat() *PictureFormat {
	retVal := this.PropGet(0x0000065f, nil)
	return NewPictureFormat(retVal.PdispValVal(), false, true)
}

func (this *ChartFormat) Shadow() *ShadowFormat {
	retVal := this.PropGet(0x00000067, nil)
	return NewShadowFormat(retVal.PdispValVal(), false, true)
}

func (this *ChartFormat) SoftEdge() *ole.DispatchClass {
	retVal := this.PropGet(0x00000a66, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ChartFormat) TextFrame2() *TextFrame2 {
	retVal := this.PropGet(0x00000a63, nil)
	return NewTextFrame2(retVal.PdispValVal(), false, true)
}

func (this *ChartFormat) ThreeD() *ThreeDFormat {
	retVal := this.PropGet(0x000006a7, nil)
	return NewThreeDFormat(retVal.PdispValVal(), false, true)
}

