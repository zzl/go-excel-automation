package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
)

// 000C031F-0000-0000-C000-000000000046
var IID_TextEffectFormat = syscall.GUID{0x000C031F, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TextEffectFormat struct {
	ole.OleClient
}

func NewTextEffectFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *TextEffectFormat {
	if pDisp == nil {
		return nil
	}
	p := &TextEffectFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TextEffectFormatFromVar(v ole.Variant) *TextEffectFormat {
	return NewTextEffectFormat(v.IDispatch(), false, false)
}

func (this *TextEffectFormat) IID() *syscall.GUID {
	return &IID_TextEffectFormat
}

func (this *TextEffectFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TextEffectFormat) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x60020000, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TextEffectFormat) Creator() int32 {
	retVal, _ := this.PropGet(0x60020001, nil)
	return retVal.LValVal()
}

func (this *TextEffectFormat) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000001, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TextEffectFormat) ToggleVerticalText() {
	retVal, _ := this.Call(0x0000000a, nil)
	_ = retVal
}

func (this *TextEffectFormat) Alignment() int32 {
	retVal, _ := this.PropGet(0x00000064, nil)
	return retVal.LValVal()
}

func (this *TextEffectFormat) SetAlignment(rhs int32) {
	_ = this.PropPut(0x00000064, []interface{}{rhs})
}

func (this *TextEffectFormat) FontBold() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *TextEffectFormat) SetFontBold(rhs int32) {
	_ = this.PropPut(0x00000065, []interface{}{rhs})
}

func (this *TextEffectFormat) FontItalic() int32 {
	retVal, _ := this.PropGet(0x00000066, nil)
	return retVal.LValVal()
}

func (this *TextEffectFormat) SetFontItalic(rhs int32) {
	_ = this.PropPut(0x00000066, []interface{}{rhs})
}

func (this *TextEffectFormat) FontName() string {
	retVal, _ := this.PropGet(0x00000067, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TextEffectFormat) SetFontName(rhs string) {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *TextEffectFormat) FontSize() float32 {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.FltValVal()
}

func (this *TextEffectFormat) SetFontSize(rhs float32) {
	_ = this.PropPut(0x00000068, []interface{}{rhs})
}

func (this *TextEffectFormat) KernedPairs() int32 {
	retVal, _ := this.PropGet(0x00000069, nil)
	return retVal.LValVal()
}

func (this *TextEffectFormat) SetKernedPairs(rhs int32) {
	_ = this.PropPut(0x00000069, []interface{}{rhs})
}

func (this *TextEffectFormat) NormalizedHeight() int32 {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return retVal.LValVal()
}

func (this *TextEffectFormat) SetNormalizedHeight(rhs int32) {
	_ = this.PropPut(0x0000006a, []interface{}{rhs})
}

func (this *TextEffectFormat) PresetShape() int32 {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return retVal.LValVal()
}

func (this *TextEffectFormat) SetPresetShape(rhs int32) {
	_ = this.PropPut(0x0000006b, []interface{}{rhs})
}

func (this *TextEffectFormat) PresetTextEffect() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *TextEffectFormat) SetPresetTextEffect(rhs int32) {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

func (this *TextEffectFormat) RotatedChars() int32 {
	retVal, _ := this.PropGet(0x0000006d, nil)
	return retVal.LValVal()
}

func (this *TextEffectFormat) SetRotatedChars(rhs int32) {
	_ = this.PropPut(0x0000006d, []interface{}{rhs})
}

func (this *TextEffectFormat) Text() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TextEffectFormat) SetText(rhs string) {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *TextEffectFormat) Tracking() float32 {
	retVal, _ := this.PropGet(0x0000006f, nil)
	return retVal.FltValVal()
}

func (this *TextEffectFormat) SetTracking(rhs float32) {
	_ = this.PropPut(0x0000006f, []interface{}{rhs})
}
