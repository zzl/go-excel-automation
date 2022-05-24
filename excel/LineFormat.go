package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000C0317-0000-0000-C000-000000000046
var IID_LineFormat = syscall.GUID{0x000C0317, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type LineFormat struct {
	ole.OleClient
}

func NewLineFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *LineFormat {
	 if pDisp == nil {
		return nil;
	}
	p := &LineFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func LineFormatFromVar(v ole.Variant) *LineFormat {
	return NewLineFormat(v.IDispatch(), false, false)
}

func (this *LineFormat) IID() *syscall.GUID {
	return &IID_LineFormat
}

func (this *LineFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *LineFormat) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x60020000, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *LineFormat) Creator() int32 {
	retVal, _ := this.PropGet(0x60020001, nil)
	return retVal.LValVal()
}

func (this *LineFormat) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000001, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *LineFormat) BackColor() *ColorFormat {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewColorFormat(retVal.IDispatch(), false, true)
}

func (this *LineFormat) SetBackColor(rhs *ColorFormat)  {
	_ = this.PropPut(0x00000064, []interface{}{rhs})
}

func (this *LineFormat) BeginArrowheadLength() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *LineFormat) SetBeginArrowheadLength(rhs int32)  {
	_ = this.PropPut(0x00000065, []interface{}{rhs})
}

func (this *LineFormat) BeginArrowheadStyle() int32 {
	retVal, _ := this.PropGet(0x00000066, nil)
	return retVal.LValVal()
}

func (this *LineFormat) SetBeginArrowheadStyle(rhs int32)  {
	_ = this.PropPut(0x00000066, []interface{}{rhs})
}

func (this *LineFormat) BeginArrowheadWidth() int32 {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *LineFormat) SetBeginArrowheadWidth(rhs int32)  {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *LineFormat) DashStyle() int32 {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *LineFormat) SetDashStyle(rhs int32)  {
	_ = this.PropPut(0x00000068, []interface{}{rhs})
}

func (this *LineFormat) EndArrowheadLength() int32 {
	retVal, _ := this.PropGet(0x00000069, nil)
	return retVal.LValVal()
}

func (this *LineFormat) SetEndArrowheadLength(rhs int32)  {
	_ = this.PropPut(0x00000069, []interface{}{rhs})
}

func (this *LineFormat) EndArrowheadStyle() int32 {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return retVal.LValVal()
}

func (this *LineFormat) SetEndArrowheadStyle(rhs int32)  {
	_ = this.PropPut(0x0000006a, []interface{}{rhs})
}

func (this *LineFormat) EndArrowheadWidth() int32 {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return retVal.LValVal()
}

func (this *LineFormat) SetEndArrowheadWidth(rhs int32)  {
	_ = this.PropPut(0x0000006b, []interface{}{rhs})
}

func (this *LineFormat) ForeColor() *ColorFormat {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return NewColorFormat(retVal.IDispatch(), false, true)
}

func (this *LineFormat) SetForeColor(rhs *ColorFormat)  {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

func (this *LineFormat) Pattern() int32 {
	retVal, _ := this.PropGet(0x0000006d, nil)
	return retVal.LValVal()
}

func (this *LineFormat) SetPattern(rhs int32)  {
	_ = this.PropPut(0x0000006d, []interface{}{rhs})
}

func (this *LineFormat) Style() int32 {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return retVal.LValVal()
}

func (this *LineFormat) SetStyle(rhs int32)  {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *LineFormat) Transparency() float32 {
	retVal, _ := this.PropGet(0x0000006f, nil)
	return retVal.FltValVal()
}

func (this *LineFormat) SetTransparency(rhs float32)  {
	_ = this.PropPut(0x0000006f, []interface{}{rhs})
}

func (this *LineFormat) Visible() int32 {
	retVal, _ := this.PropGet(0x00000070, nil)
	return retVal.LValVal()
}

func (this *LineFormat) SetVisible(rhs int32)  {
	_ = this.PropPut(0x00000070, []interface{}{rhs})
}

func (this *LineFormat) Weight() float32 {
	retVal, _ := this.PropGet(0x00000071, nil)
	return retVal.FltValVal()
}

func (this *LineFormat) SetWeight(rhs float32)  {
	_ = this.PropPut(0x00000071, []interface{}{rhs})
}

func (this *LineFormat) InsetPen() int32 {
	retVal, _ := this.PropGet(0x00000072, nil)
	return retVal.LValVal()
}

func (this *LineFormat) SetInsetPen(rhs int32)  {
	_ = this.PropPut(0x00000072, []interface{}{rhs})
}

