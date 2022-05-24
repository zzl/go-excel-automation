package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000C0314-0000-0000-C000-000000000046
var IID_FillFormat = syscall.GUID{0x000C0314, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type FillFormat struct {
	ole.OleClient
}

func NewFillFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *FillFormat {
	 if pDisp == nil {
		return nil;
	}
	p := &FillFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FillFormatFromVar(v ole.Variant) *FillFormat {
	return NewFillFormat(v.IDispatch(), false, false)
}

func (this *FillFormat) IID() *syscall.GUID {
	return &IID_FillFormat
}

func (this *FillFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *FillFormat) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x60020000, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *FillFormat) Creator() int32 {
	retVal, _ := this.PropGet(0x60020001, nil)
	return retVal.LValVal()
}

func (this *FillFormat) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000001, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *FillFormat) Background()  {
	retVal, _ := this.Call(0x0000000a, nil)
	_= retVal
}

func (this *FillFormat) OneColorGradient(style int32, variant int32, degree float32)  {
	retVal, _ := this.Call(0x0000000b, []interface{}{style, variant, degree})
	_= retVal
}

func (this *FillFormat) Patterned(pattern int32)  {
	retVal, _ := this.Call(0x0000000c, []interface{}{pattern})
	_= retVal
}

func (this *FillFormat) PresetGradient(style int32, variant int32, presetGradientType int32)  {
	retVal, _ := this.Call(0x0000000d, []interface{}{style, variant, presetGradientType})
	_= retVal
}

func (this *FillFormat) PresetTextured(presetTexture int32)  {
	retVal, _ := this.Call(0x0000000e, []interface{}{presetTexture})
	_= retVal
}

func (this *FillFormat) Solid()  {
	retVal, _ := this.Call(0x0000000f, nil)
	_= retVal
}

func (this *FillFormat) TwoColorGradient(style int32, variant int32)  {
	retVal, _ := this.Call(0x00000010, []interface{}{style, variant})
	_= retVal
}

func (this *FillFormat) UserPicture(pictureFile string)  {
	retVal, _ := this.Call(0x00000011, []interface{}{pictureFile})
	_= retVal
}

func (this *FillFormat) UserTextured(textureFile string)  {
	retVal, _ := this.Call(0x00000012, []interface{}{textureFile})
	_= retVal
}

func (this *FillFormat) BackColor() *ColorFormat {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewColorFormat(retVal.IDispatch(), false, true)
}

func (this *FillFormat) SetBackColor(rhs *ColorFormat)  {
	_ = this.PropPut(0x00000064, []interface{}{rhs})
}

func (this *FillFormat) ForeColor() *ColorFormat {
	retVal, _ := this.PropGet(0x00000065, nil)
	return NewColorFormat(retVal.IDispatch(), false, true)
}

func (this *FillFormat) SetForeColor(rhs *ColorFormat)  {
	_ = this.PropPut(0x00000065, []interface{}{rhs})
}

func (this *FillFormat) GradientColorType() int32 {
	retVal, _ := this.PropGet(0x00000066, nil)
	return retVal.LValVal()
}

func (this *FillFormat) GradientDegree() float32 {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.FltValVal()
}

func (this *FillFormat) GradientStyle() int32 {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *FillFormat) GradientVariant() int32 {
	retVal, _ := this.PropGet(0x00000069, nil)
	return retVal.LValVal()
}

func (this *FillFormat) Pattern() int32 {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return retVal.LValVal()
}

func (this *FillFormat) PresetGradientType() int32 {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return retVal.LValVal()
}

func (this *FillFormat) PresetTexture() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *FillFormat) TextureName() string {
	retVal, _ := this.PropGet(0x0000006d, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *FillFormat) TextureType() int32 {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return retVal.LValVal()
}

func (this *FillFormat) Transparency() float32 {
	retVal, _ := this.PropGet(0x0000006f, nil)
	return retVal.FltValVal()
}

func (this *FillFormat) SetTransparency(rhs float32)  {
	_ = this.PropPut(0x0000006f, []interface{}{rhs})
}

func (this *FillFormat) Type() int32 {
	retVal, _ := this.PropGet(0x00000070, nil)
	return retVal.LValVal()
}

func (this *FillFormat) Visible() int32 {
	retVal, _ := this.PropGet(0x00000071, nil)
	return retVal.LValVal()
}

func (this *FillFormat) SetVisible(rhs int32)  {
	_ = this.PropPut(0x00000071, []interface{}{rhs})
}

func (this *FillFormat) GradientStops() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000072, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *FillFormat) TextureOffsetX() float32 {
	retVal, _ := this.PropGet(0x00000073, nil)
	return retVal.FltValVal()
}

func (this *FillFormat) SetTextureOffsetX(rhs float32)  {
	_ = this.PropPut(0x00000073, []interface{}{rhs})
}

func (this *FillFormat) TextureOffsetY() float32 {
	retVal, _ := this.PropGet(0x00000074, nil)
	return retVal.FltValVal()
}

func (this *FillFormat) SetTextureOffsetY(rhs float32)  {
	_ = this.PropPut(0x00000074, []interface{}{rhs})
}

func (this *FillFormat) TextureAlignment() int32 {
	retVal, _ := this.PropGet(0x00000075, nil)
	return retVal.LValVal()
}

func (this *FillFormat) SetTextureAlignment(rhs int32)  {
	_ = this.PropPut(0x00000075, []interface{}{rhs})
}

func (this *FillFormat) TextureHorizontalScale() float32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.FltValVal()
}

func (this *FillFormat) SetTextureHorizontalScale(rhs float32)  {
	_ = this.PropPut(0x00000076, []interface{}{rhs})
}

func (this *FillFormat) TextureVerticalScale() float32 {
	retVal, _ := this.PropGet(0x00000077, nil)
	return retVal.FltValVal()
}

func (this *FillFormat) SetTextureVerticalScale(rhs float32)  {
	_ = this.PropPut(0x00000077, []interface{}{rhs})
}

func (this *FillFormat) TextureTile() int32 {
	retVal, _ := this.PropGet(0x00000078, nil)
	return retVal.LValVal()
}

func (this *FillFormat) SetTextureTile(rhs int32)  {
	_ = this.PropPut(0x00000078, []interface{}{rhs})
}

func (this *FillFormat) RotateWithObject() int32 {
	retVal, _ := this.PropGet(0x00000079, nil)
	return retVal.LValVal()
}

func (this *FillFormat) SetRotateWithObject(rhs int32)  {
	_ = this.PropPut(0x00000079, []interface{}{rhs})
}

func (this *FillFormat) PictureEffects() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *FillFormat) GradientAngle() float32 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.FltValVal()
}

func (this *FillFormat) SetGradientAngle(rhs float32)  {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

