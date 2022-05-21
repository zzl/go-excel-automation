package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000C0310-0000-0000-C000-000000000046
var IID_Adjustments = syscall.GUID{0x000C0310, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Adjustments struct {
	ole.OleClient
}

func NewAdjustments(pDisp *win32.IDispatch, addRef bool, scoped bool) *Adjustments {
	p := &Adjustments{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func AdjustmentsFromVar(v ole.Variant) *Adjustments {
	return NewAdjustments(v.PdispValVal(), false, false)
}

func (this *Adjustments) IID() *syscall.GUID {
	return &IID_Adjustments
}

func (this *Adjustments) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Adjustments) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x60020000, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Adjustments) Creator() int32 {
	retVal := this.PropGet(0x60020001, nil)
	return retVal.LValVal()
}

func (this *Adjustments) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000001, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Adjustments) Count() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Adjustments) Item(index int32) float32 {
	retVal := this.PropGet(0x00000000, []interface{}{index})
	return retVal.FltValVal()
}

func (this *Adjustments) SetItem(index int32, rhs float32)  {
	retVal := this.PropPut(0x00000000, []interface{}{index, rhs})
	_= retVal
}

