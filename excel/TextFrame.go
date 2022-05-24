package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002443D-0000-0000-C000-000000000046
var IID_TextFrame = syscall.GUID{0x0002443D, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TextFrame struct {
	ole.OleClient
}

func NewTextFrame(pDisp *win32.IDispatch, addRef bool, scoped bool) *TextFrame {
	 if pDisp == nil {
		return nil;
	}
	p := &TextFrame{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TextFrameFromVar(v ole.Variant) *TextFrame {
	return NewTextFrame(v.IDispatch(), false, false)
}

func (this *TextFrame) IID() *syscall.GUID {
	return &IID_TextFrame
}

func (this *TextFrame) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TextFrame) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *TextFrame) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *TextFrame) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *TextFrame) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *TextFrame) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *TextFrame) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *TextFrame) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *TextFrame) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *TextFrame) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *TextFrame) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TextFrame) MarginBottom() float32 {
	retVal, _ := this.PropGet(0x000006d1, nil)
	return retVal.FltValVal()
}

func (this *TextFrame) SetMarginBottom(rhs float32)  {
	_ = this.PropPut(0x000006d1, []interface{}{rhs})
}

func (this *TextFrame) MarginLeft() float32 {
	retVal, _ := this.PropGet(0x000006d2, nil)
	return retVal.FltValVal()
}

func (this *TextFrame) SetMarginLeft(rhs float32)  {
	_ = this.PropPut(0x000006d2, []interface{}{rhs})
}

func (this *TextFrame) MarginRight() float32 {
	retVal, _ := this.PropGet(0x000006d3, nil)
	return retVal.FltValVal()
}

func (this *TextFrame) SetMarginRight(rhs float32)  {
	_ = this.PropPut(0x000006d3, []interface{}{rhs})
}

func (this *TextFrame) MarginTop() float32 {
	retVal, _ := this.PropGet(0x000006d4, nil)
	return retVal.FltValVal()
}

func (this *TextFrame) SetMarginTop(rhs float32)  {
	_ = this.PropPut(0x000006d4, []interface{}{rhs})
}

func (this *TextFrame) Orientation() int32 {
	retVal, _ := this.PropGet(0x00000086, nil)
	return retVal.LValVal()
}

func (this *TextFrame) SetOrientation(rhs int32)  {
	_ = this.PropPut(0x00000086, []interface{}{rhs})
}

var TextFrame_Characters_OptArgs= []string{
	"Start", "Length", 
}

func (this *TextFrame) Characters(optArgs ...interface{}) *Characters {
	optArgs = ole.ProcessOptArgs(TextFrame_Characters_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000025b, nil, optArgs...)
	return NewCharacters(retVal.IDispatch(), false, true)
}

func (this *TextFrame) HorizontalAlignment() int32 {
	retVal, _ := this.PropGet(0x00000088, nil)
	return retVal.LValVal()
}

func (this *TextFrame) SetHorizontalAlignment(rhs int32)  {
	_ = this.PropPut(0x00000088, []interface{}{rhs})
}

func (this *TextFrame) VerticalAlignment() int32 {
	retVal, _ := this.PropGet(0x00000089, nil)
	return retVal.LValVal()
}

func (this *TextFrame) SetVerticalAlignment(rhs int32)  {
	_ = this.PropPut(0x00000089, []interface{}{rhs})
}

func (this *TextFrame) AutoSize() bool {
	retVal, _ := this.PropGet(0x00000266, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TextFrame) SetAutoSize(rhs bool)  {
	_ = this.PropPut(0x00000266, []interface{}{rhs})
}

func (this *TextFrame) ReadingOrder() int32 {
	retVal, _ := this.PropGet(0x000003cf, nil)
	return retVal.LValVal()
}

func (this *TextFrame) SetReadingOrder(rhs int32)  {
	_ = this.PropPut(0x000003cf, []interface{}{rhs})
}

func (this *TextFrame) AutoMargins() bool {
	retVal, _ := this.PropGet(0x000006d5, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TextFrame) SetAutoMargins(rhs bool)  {
	_ = this.PropPut(0x000006d5, []interface{}{rhs})
}

func (this *TextFrame) VerticalOverflow() int32 {
	retVal, _ := this.PropGet(0x00000b6a, nil)
	return retVal.LValVal()
}

func (this *TextFrame) SetVerticalOverflow(rhs int32)  {
	_ = this.PropPut(0x00000b6a, []interface{}{rhs})
}

func (this *TextFrame) HorizontalOverflow() int32 {
	retVal, _ := this.PropGet(0x00000b6b, nil)
	return retVal.LValVal()
}

func (this *TextFrame) SetHorizontalOverflow(rhs int32)  {
	_ = this.PropPut(0x00000b6b, []interface{}{rhs})
}

