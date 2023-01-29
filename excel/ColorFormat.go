package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
)

// 000C0312-0000-0000-C000-000000000046
var IID_ColorFormat = syscall.GUID{0x000C0312, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ColorFormat struct {
	ole.OleClient
}

func NewColorFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *ColorFormat {
	if pDisp == nil {
		return nil
	}
	p := &ColorFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ColorFormatFromVar(v ole.Variant) *ColorFormat {
	return NewColorFormat(v.IDispatch(), false, false)
}

func (this *ColorFormat) IID() *syscall.GUID {
	return &IID_ColorFormat
}

func (this *ColorFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ColorFormat) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x60020000, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ColorFormat) Creator() int32 {
	retVal, _ := this.PropGet(0x60020001, nil)
	return retVal.LValVal()
}

func (this *ColorFormat) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000001, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ColorFormat) RGB() int32 {
	retVal, _ := this.PropGet(0x00000000, nil)
	return retVal.LValVal()
}

func (this *ColorFormat) SetRGB(rhs int32) {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *ColorFormat) SchemeColor() int32 {
	retVal, _ := this.PropGet(0x00000064, nil)
	return retVal.LValVal()
}

func (this *ColorFormat) SetSchemeColor(rhs int32) {
	_ = this.PropPut(0x00000064, []interface{}{rhs})
}

func (this *ColorFormat) Type() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *ColorFormat) TintAndShade() float32 {
	retVal, _ := this.PropGet(0x00000066, nil)
	return retVal.FltValVal()
}

func (this *ColorFormat) SetTintAndShade(rhs float32) {
	_ = this.PropPut(0x00000066, []interface{}{rhs})
}

func (this *ColorFormat) ObjectThemeColor() int32 {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *ColorFormat) SetObjectThemeColor(rhs int32) {
	_ = this.PropPut(0x00000068, []interface{}{rhs})
}

func (this *ColorFormat) Brightness() float32 {
	retVal, _ := this.PropGet(0x00000069, nil)
	return retVal.FltValVal()
}

func (this *ColorFormat) SetBrightness(rhs float32) {
	_ = this.PropPut(0x00000069, []interface{}{rhs})
}
