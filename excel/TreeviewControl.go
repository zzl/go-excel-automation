package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
)

// 0002444B-0000-0000-C000-000000000046
var IID_TreeviewControl = syscall.GUID{0x0002444B, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TreeviewControl struct {
	ole.OleClient
}

func NewTreeviewControl(pDisp *win32.IDispatch, addRef bool, scoped bool) *TreeviewControl {
	if pDisp == nil {
		return nil
	}
	p := &TreeviewControl{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TreeviewControlFromVar(v ole.Variant) *TreeviewControl {
	return NewTreeviewControl(v.IDispatch(), false, false)
}

func (this *TreeviewControl) IID() *syscall.GUID {
	return &IID_TreeviewControl
}

func (this *TreeviewControl) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TreeviewControl) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *TreeviewControl) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *TreeviewControl) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TreeviewControl) Hidden() ole.Variant {
	retVal, _ := this.PropGet(0x0000010c, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *TreeviewControl) SetHidden(rhs interface{}) {
	_ = this.PropPut(0x0000010c, []interface{}{rhs})
}

func (this *TreeviewControl) Drilled() ole.Variant {
	retVal, _ := this.PropGet(0x0000077d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *TreeviewControl) SetDrilled(rhs interface{}) {
	_ = this.PropPut(0x0000077d, []interface{}{rhs})
}
