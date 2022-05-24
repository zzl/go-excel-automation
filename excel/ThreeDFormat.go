package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000C0321-0000-0000-C000-000000000046
var IID_ThreeDFormat = syscall.GUID{0x000C0321, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ThreeDFormat struct {
	ole.OleClient
}

func NewThreeDFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *ThreeDFormat {
	 if pDisp == nil {
		return nil;
	}
	p := &ThreeDFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ThreeDFormatFromVar(v ole.Variant) *ThreeDFormat {
	return NewThreeDFormat(v.IDispatch(), false, false)
}

func (this *ThreeDFormat) IID() *syscall.GUID {
	return &IID_ThreeDFormat
}

func (this *ThreeDFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ThreeDFormat) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x60020000, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ThreeDFormat) Creator() int32 {
	retVal, _ := this.PropGet(0x60020001, nil)
	return retVal.LValVal()
}

func (this *ThreeDFormat) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000001, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ThreeDFormat) IncrementRotationX(increment float32)  {
	retVal, _ := this.Call(0x0000000a, []interface{}{increment})
	_= retVal
}

func (this *ThreeDFormat) IncrementRotationY(increment float32)  {
	retVal, _ := this.Call(0x0000000b, []interface{}{increment})
	_= retVal
}

func (this *ThreeDFormat) ResetRotation()  {
	retVal, _ := this.Call(0x0000000c, nil)
	_= retVal
}

func (this *ThreeDFormat) SetThreeDFormat(presetThreeDFormat int32)  {
	retVal, _ := this.Call(0x0000000d, []interface{}{presetThreeDFormat})
	_= retVal
}

func (this *ThreeDFormat) SetExtrusionDirection(presetExtrusionDirection int32)  {
	retVal, _ := this.Call(0x0000000e, []interface{}{presetExtrusionDirection})
	_= retVal
}

func (this *ThreeDFormat) Depth() float32 {
	retVal, _ := this.PropGet(0x00000064, nil)
	return retVal.FltValVal()
}

func (this *ThreeDFormat) SetDepth(rhs float32)  {
	_ = this.PropPut(0x00000064, []interface{}{rhs})
}

func (this *ThreeDFormat) ExtrusionColor() *ColorFormat {
	retVal, _ := this.PropGet(0x00000065, nil)
	return NewColorFormat(retVal.IDispatch(), false, true)
}

func (this *ThreeDFormat) ExtrusionColorType() int32 {
	retVal, _ := this.PropGet(0x00000066, nil)
	return retVal.LValVal()
}

func (this *ThreeDFormat) SetExtrusionColorType(rhs int32)  {
	_ = this.PropPut(0x00000066, []interface{}{rhs})
}

func (this *ThreeDFormat) Perspective() int32 {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *ThreeDFormat) SetPerspective(rhs int32)  {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *ThreeDFormat) PresetExtrusionDirection() int32 {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *ThreeDFormat) PresetLightingDirection() int32 {
	retVal, _ := this.PropGet(0x00000069, nil)
	return retVal.LValVal()
}

func (this *ThreeDFormat) SetPresetLightingDirection(rhs int32)  {
	_ = this.PropPut(0x00000069, []interface{}{rhs})
}

func (this *ThreeDFormat) PresetLightingSoftness() int32 {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return retVal.LValVal()
}

func (this *ThreeDFormat) SetPresetLightingSoftness(rhs int32)  {
	_ = this.PropPut(0x0000006a, []interface{}{rhs})
}

func (this *ThreeDFormat) PresetMaterial() int32 {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return retVal.LValVal()
}

func (this *ThreeDFormat) SetPresetMaterial(rhs int32)  {
	_ = this.PropPut(0x0000006b, []interface{}{rhs})
}

func (this *ThreeDFormat) PresetThreeDFormat() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *ThreeDFormat) RotationX() float32 {
	retVal, _ := this.PropGet(0x0000006d, nil)
	return retVal.FltValVal()
}

func (this *ThreeDFormat) SetRotationX(rhs float32)  {
	_ = this.PropPut(0x0000006d, []interface{}{rhs})
}

func (this *ThreeDFormat) RotationY() float32 {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return retVal.FltValVal()
}

func (this *ThreeDFormat) SetRotationY(rhs float32)  {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *ThreeDFormat) Visible() int32 {
	retVal, _ := this.PropGet(0x0000006f, nil)
	return retVal.LValVal()
}

func (this *ThreeDFormat) SetVisible(rhs int32)  {
	_ = this.PropPut(0x0000006f, []interface{}{rhs})
}

func (this *ThreeDFormat) SetPresetCamera(presetCamera int32)  {
	retVal, _ := this.Call(0x0000000f, []interface{}{presetCamera})
	_= retVal
}

func (this *ThreeDFormat) IncrementRotationZ(increment float32)  {
	retVal, _ := this.Call(0x00000010, []interface{}{increment})
	_= retVal
}

func (this *ThreeDFormat) IncrementRotationHorizontal(increment float32)  {
	retVal, _ := this.Call(0x00000011, []interface{}{increment})
	_= retVal
}

func (this *ThreeDFormat) IncrementRotationVertical(increment float32)  {
	retVal, _ := this.Call(0x00000012, []interface{}{increment})
	_= retVal
}

func (this *ThreeDFormat) PresetLighting() int32 {
	retVal, _ := this.PropGet(0x00000070, nil)
	return retVal.LValVal()
}

func (this *ThreeDFormat) SetPresetLighting(rhs int32)  {
	_ = this.PropPut(0x00000070, []interface{}{rhs})
}

func (this *ThreeDFormat) Z() float32 {
	retVal, _ := this.PropGet(0x00000071, nil)
	return retVal.FltValVal()
}

func (this *ThreeDFormat) SetZ(rhs float32)  {
	_ = this.PropPut(0x00000071, []interface{}{rhs})
}

func (this *ThreeDFormat) BevelTopType() int32 {
	retVal, _ := this.PropGet(0x00000072, nil)
	return retVal.LValVal()
}

func (this *ThreeDFormat) SetBevelTopType(rhs int32)  {
	_ = this.PropPut(0x00000072, []interface{}{rhs})
}

func (this *ThreeDFormat) BevelTopInset() float32 {
	retVal, _ := this.PropGet(0x00000073, nil)
	return retVal.FltValVal()
}

func (this *ThreeDFormat) SetBevelTopInset(rhs float32)  {
	_ = this.PropPut(0x00000073, []interface{}{rhs})
}

func (this *ThreeDFormat) BevelTopDepth() float32 {
	retVal, _ := this.PropGet(0x00000074, nil)
	return retVal.FltValVal()
}

func (this *ThreeDFormat) SetBevelTopDepth(rhs float32)  {
	_ = this.PropPut(0x00000074, []interface{}{rhs})
}

func (this *ThreeDFormat) BevelBottomType() int32 {
	retVal, _ := this.PropGet(0x00000075, nil)
	return retVal.LValVal()
}

func (this *ThreeDFormat) SetBevelBottomType(rhs int32)  {
	_ = this.PropPut(0x00000075, []interface{}{rhs})
}

func (this *ThreeDFormat) BevelBottomInset() float32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.FltValVal()
}

func (this *ThreeDFormat) SetBevelBottomInset(rhs float32)  {
	_ = this.PropPut(0x00000076, []interface{}{rhs})
}

func (this *ThreeDFormat) BevelBottomDepth() float32 {
	retVal, _ := this.PropGet(0x00000077, nil)
	return retVal.FltValVal()
}

func (this *ThreeDFormat) SetBevelBottomDepth(rhs float32)  {
	_ = this.PropPut(0x00000077, []interface{}{rhs})
}

func (this *ThreeDFormat) PresetCamera() int32 {
	retVal, _ := this.PropGet(0x00000078, nil)
	return retVal.LValVal()
}

func (this *ThreeDFormat) RotationZ() float32 {
	retVal, _ := this.PropGet(0x00000079, nil)
	return retVal.FltValVal()
}

func (this *ThreeDFormat) SetRotationZ(rhs float32)  {
	_ = this.PropPut(0x00000079, []interface{}{rhs})
}

func (this *ThreeDFormat) ContourWidth() float32 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.FltValVal()
}

func (this *ThreeDFormat) SetContourWidth(rhs float32)  {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *ThreeDFormat) ContourColor() *ColorFormat {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return NewColorFormat(retVal.IDispatch(), false, true)
}

func (this *ThreeDFormat) FieldOfView() float32 {
	retVal, _ := this.PropGet(0x0000007c, nil)
	return retVal.FltValVal()
}

func (this *ThreeDFormat) SetFieldOfView(rhs float32)  {
	_ = this.PropPut(0x0000007c, []interface{}{rhs})
}

func (this *ThreeDFormat) ProjectText() int32 {
	retVal, _ := this.PropGet(0x0000007d, nil)
	return retVal.LValVal()
}

func (this *ThreeDFormat) SetProjectText(rhs int32)  {
	_ = this.PropPut(0x0000007d, []interface{}{rhs})
}

func (this *ThreeDFormat) LightAngle() float32 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.FltValVal()
}

func (this *ThreeDFormat) SetLightAngle(rhs float32)  {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

