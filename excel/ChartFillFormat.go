package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024435-0000-0000-C000-000000000046
var IID_ChartFillFormat = syscall.GUID{0x00024435, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ChartFillFormat struct {
	ole.OleClient
}

func NewChartFillFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *ChartFillFormat {
	p := &ChartFillFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ChartFillFormatFromVar(v ole.Variant) *ChartFillFormat {
	return NewChartFillFormat(v.PdispValVal(), false, false)
}

func (this *ChartFillFormat) IID() *syscall.GUID {
	return &IID_ChartFillFormat
}

func (this *ChartFillFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ChartFillFormat) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *ChartFillFormat) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ChartFillFormat) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ChartFillFormat) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *ChartFillFormat) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *ChartFillFormat) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *ChartFillFormat) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *ChartFillFormat) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *ChartFillFormat) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ChartFillFormat) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ChartFillFormat) OneColorGradient(style int32, variant int32, degree float32)  {
	retVal := this.Call(0x00000655, []interface{}{style, variant, degree})
	_= retVal
}

func (this *ChartFillFormat) TwoColorGradient(style int32, variant int32)  {
	retVal := this.Call(0x00000658, []interface{}{style, variant})
	_= retVal
}

func (this *ChartFillFormat) PresetTextured(presetTexture int32)  {
	retVal := this.Call(0x00000659, []interface{}{presetTexture})
	_= retVal
}

func (this *ChartFillFormat) Solid()  {
	retVal := this.Call(0x0000065b, nil)
	_= retVal
}

func (this *ChartFillFormat) Patterned(pattern int32)  {
	retVal := this.Call(0x0000065c, []interface{}{pattern})
	_= retVal
}

var ChartFillFormat_UserPicture_OptArgs= []string{
	"PictureFile", "PictureFormat", "PictureStackUnit", "PicturePlacement", 
}

func (this *ChartFillFormat) UserPicture(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ChartFillFormat_UserPicture_OptArgs, optArgs)
	retVal := this.Call(0x0000065d, nil, optArgs...)
	_= retVal
}

func (this *ChartFillFormat) UserTextured(textureFile string)  {
	retVal := this.Call(0x00000662, []interface{}{textureFile})
	_= retVal
}

func (this *ChartFillFormat) PresetGradient(style int32, variant int32, presetGradientType int32)  {
	retVal := this.Call(0x00000664, []interface{}{style, variant, presetGradientType})
	_= retVal
}

func (this *ChartFillFormat) BackColor() *ChartColorFormat {
	retVal := this.PropGet(0x00000666, nil)
	return NewChartColorFormat(retVal.PdispValVal(), false, true)
}

func (this *ChartFillFormat) ForeColor() *ChartColorFormat {
	retVal := this.PropGet(0x00000667, nil)
	return NewChartColorFormat(retVal.PdispValVal(), false, true)
}

func (this *ChartFillFormat) GradientColorType() int32 {
	retVal := this.PropGet(0x00000668, nil)
	return retVal.LValVal()
}

func (this *ChartFillFormat) GradientDegree() float32 {
	retVal := this.PropGet(0x00000669, nil)
	return retVal.FltValVal()
}

func (this *ChartFillFormat) GradientStyle() int32 {
	retVal := this.PropGet(0x0000066a, nil)
	return retVal.LValVal()
}

func (this *ChartFillFormat) GradientVariant() int32 {
	retVal := this.PropGet(0x0000066b, nil)
	return retVal.LValVal()
}

func (this *ChartFillFormat) Pattern() int32 {
	retVal := this.PropGet(0x0000005f, nil)
	return retVal.LValVal()
}

func (this *ChartFillFormat) PresetGradientType() int32 {
	retVal := this.PropGet(0x00000665, nil)
	return retVal.LValVal()
}

func (this *ChartFillFormat) PresetTexture() int32 {
	retVal := this.PropGet(0x0000065a, nil)
	return retVal.LValVal()
}

func (this *ChartFillFormat) TextureName() string {
	retVal := this.PropGet(0x0000066c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ChartFillFormat) TextureType() int32 {
	retVal := this.PropGet(0x0000066d, nil)
	return retVal.LValVal()
}

func (this *ChartFillFormat) Type() int32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *ChartFillFormat) Visible() int32 {
	retVal := this.PropGet(0x0000022e, nil)
	return retVal.LValVal()
}

func (this *ChartFillFormat) SetVisible(rhs int32)  {
	retVal := this.PropPut(0x0000022e, []interface{}{rhs})
	_= retVal
}

