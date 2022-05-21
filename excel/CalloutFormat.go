package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000C0311-0000-0000-C000-000000000046
var IID_CalloutFormat = syscall.GUID{0x000C0311, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type CalloutFormat struct {
	ole.OleClient
}

func NewCalloutFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *CalloutFormat {
	p := &CalloutFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CalloutFormatFromVar(v ole.Variant) *CalloutFormat {
	return NewCalloutFormat(v.PdispValVal(), false, false)
}

func (this *CalloutFormat) IID() *syscall.GUID {
	return &IID_CalloutFormat
}

func (this *CalloutFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CalloutFormat) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x60020000, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *CalloutFormat) Creator() int32 {
	retVal := this.PropGet(0x60020001, nil)
	return retVal.LValVal()
}

func (this *CalloutFormat) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000001, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *CalloutFormat) AutomaticLength()  {
	retVal := this.Call(0x0000000a, nil)
	_= retVal
}

func (this *CalloutFormat) CustomDrop(drop float32)  {
	retVal := this.Call(0x0000000b, []interface{}{drop})
	_= retVal
}

func (this *CalloutFormat) CustomLength(length float32)  {
	retVal := this.Call(0x0000000c, []interface{}{length})
	_= retVal
}

func (this *CalloutFormat) PresetDrop(dropType int32)  {
	retVal := this.Call(0x0000000d, []interface{}{dropType})
	_= retVal
}

func (this *CalloutFormat) Accent() int32 {
	retVal := this.PropGet(0x00000064, nil)
	return retVal.LValVal()
}

func (this *CalloutFormat) SetAccent(rhs int32)  {
	retVal := this.PropPut(0x00000064, []interface{}{rhs})
	_= retVal
}

func (this *CalloutFormat) Angle() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *CalloutFormat) SetAngle(rhs int32)  {
	retVal := this.PropPut(0x00000065, []interface{}{rhs})
	_= retVal
}

func (this *CalloutFormat) AutoAttach() int32 {
	retVal := this.PropGet(0x00000066, nil)
	return retVal.LValVal()
}

func (this *CalloutFormat) SetAutoAttach(rhs int32)  {
	retVal := this.PropPut(0x00000066, []interface{}{rhs})
	_= retVal
}

func (this *CalloutFormat) AutoLength() int32 {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *CalloutFormat) Border() int32 {
	retVal := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *CalloutFormat) SetBorder(rhs int32)  {
	retVal := this.PropPut(0x00000068, []interface{}{rhs})
	_= retVal
}

func (this *CalloutFormat) Drop() float32 {
	retVal := this.PropGet(0x00000069, nil)
	return retVal.FltValVal()
}

func (this *CalloutFormat) DropType() int32 {
	retVal := this.PropGet(0x0000006a, nil)
	return retVal.LValVal()
}

func (this *CalloutFormat) Gap() float32 {
	retVal := this.PropGet(0x0000006b, nil)
	return retVal.FltValVal()
}

func (this *CalloutFormat) SetGap(rhs float32)  {
	retVal := this.PropPut(0x0000006b, []interface{}{rhs})
	_= retVal
}

func (this *CalloutFormat) Length() float32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.FltValVal()
}

func (this *CalloutFormat) Type() int32 {
	retVal := this.PropGet(0x0000006d, nil)
	return retVal.LValVal()
}

func (this *CalloutFormat) SetType(rhs int32)  {
	retVal := this.PropPut(0x0000006d, []interface{}{rhs})
	_= retVal
}

