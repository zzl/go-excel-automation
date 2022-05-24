package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000C0398-0000-0000-C000-000000000046
var IID_TextFrame2 = syscall.GUID{0x000C0398, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TextFrame2 struct {
	ole.OleClient
}

func NewTextFrame2(pDisp *win32.IDispatch, addRef bool, scoped bool) *TextFrame2 {
	 if pDisp == nil {
		return nil;
	}
	p := &TextFrame2{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TextFrame2FromVar(v ole.Variant) *TextFrame2 {
	return NewTextFrame2(v.IDispatch(), false, false)
}

func (this *TextFrame2) IID() *syscall.GUID {
	return &IID_TextFrame2
}

func (this *TextFrame2) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TextFrame2) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x60020000, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TextFrame2) Creator() int32 {
	retVal, _ := this.PropGet(0x60020001, nil)
	return retVal.LValVal()
}

func (this *TextFrame2) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000001, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TextFrame2) MarginBottom() float32 {
	retVal, _ := this.PropGet(0x00000064, nil)
	return retVal.FltValVal()
}

func (this *TextFrame2) SetMarginBottom(rhs float32)  {
	_ = this.PropPut(0x00000064, []interface{}{rhs})
}

func (this *TextFrame2) MarginLeft() float32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.FltValVal()
}

func (this *TextFrame2) SetMarginLeft(rhs float32)  {
	_ = this.PropPut(0x00000065, []interface{}{rhs})
}

func (this *TextFrame2) MarginRight() float32 {
	retVal, _ := this.PropGet(0x00000066, nil)
	return retVal.FltValVal()
}

func (this *TextFrame2) SetMarginRight(rhs float32)  {
	_ = this.PropPut(0x00000066, []interface{}{rhs})
}

func (this *TextFrame2) MarginTop() float32 {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.FltValVal()
}

func (this *TextFrame2) SetMarginTop(rhs float32)  {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *TextFrame2) Orientation() int32 {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *TextFrame2) SetOrientation(rhs int32)  {
	_ = this.PropPut(0x00000068, []interface{}{rhs})
}

func (this *TextFrame2) HorizontalAnchor() int32 {
	retVal, _ := this.PropGet(0x00000069, nil)
	return retVal.LValVal()
}

func (this *TextFrame2) SetHorizontalAnchor(rhs int32)  {
	_ = this.PropPut(0x00000069, []interface{}{rhs})
}

func (this *TextFrame2) VerticalAnchor() int32 {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return retVal.LValVal()
}

func (this *TextFrame2) SetVerticalAnchor(rhs int32)  {
	_ = this.PropPut(0x0000006a, []interface{}{rhs})
}

func (this *TextFrame2) PathFormat() int32 {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return retVal.LValVal()
}

func (this *TextFrame2) SetPathFormat(rhs int32)  {
	_ = this.PropPut(0x0000006b, []interface{}{rhs})
}

func (this *TextFrame2) WarpFormat() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *TextFrame2) SetWarpFormat(rhs int32)  {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

func (this *TextFrame2) WordArtformat() int32 {
	retVal, _ := this.PropGet(0x0000006d, nil)
	return retVal.LValVal()
}

func (this *TextFrame2) SetWordArtformat(rhs int32)  {
	_ = this.PropPut(0x0000006d, []interface{}{rhs})
}

func (this *TextFrame2) WordWrap() int32 {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return retVal.LValVal()
}

func (this *TextFrame2) SetWordWrap(rhs int32)  {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *TextFrame2) AutoSize() int32 {
	retVal, _ := this.PropGet(0x0000006f, nil)
	return retVal.LValVal()
}

func (this *TextFrame2) SetAutoSize(rhs int32)  {
	_ = this.PropPut(0x0000006f, []interface{}{rhs})
}

func (this *TextFrame2) ThreeD() *ThreeDFormat {
	retVal, _ := this.PropGet(0x00000070, nil)
	return NewThreeDFormat(retVal.IDispatch(), false, true)
}

func (this *TextFrame2) HasText() int32 {
	retVal, _ := this.PropGet(0x00000071, nil)
	return retVal.LValVal()
}

func (this *TextFrame2) TextRange() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000072, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TextFrame2) Column() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000073, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TextFrame2) Ruler() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000074, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TextFrame2) DeleteText()  {
	retVal, _ := this.Call(0x00000075, nil)
	_= retVal
}

func (this *TextFrame2) NoTextRotation() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *TextFrame2) SetNoTextRotation(rhs int32)  {
	_ = this.PropPut(0x00000076, []interface{}{rhs})
}

