package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// A43788C1-D91B-11D3-8F39-00C04F3651B8
var IID_IRTDUpdateEvent = syscall.GUID{0xA43788C1, 0xD91B, 0x11D3, 
	[8]byte{0x8F, 0x39, 0x00, 0xC0, 0x4F, 0x36, 0x51, 0xB8}}

type IRTDUpdateEvent struct {
	ole.OleClient
}

func NewIRTDUpdateEvent(pDisp *win32.IDispatch, addRef bool, scoped bool) *IRTDUpdateEvent {
	 if pDisp == nil {
		return nil;
	}
	p := &IRTDUpdateEvent{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func IRTDUpdateEventFromVar(v ole.Variant) *IRTDUpdateEvent {
	return NewIRTDUpdateEvent(v.IDispatch(), false, false)
}

func (this *IRTDUpdateEvent) IID() *syscall.GUID {
	return &IID_IRTDUpdateEvent
}

func (this *IRTDUpdateEvent) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *IRTDUpdateEvent) UpdateNotify()  {
	retVal, _ := this.Call(0x0000000a, nil)
	_= retVal
}

func (this *IRTDUpdateEvent) HeartbeatInterval() int32 {
	retVal, _ := this.PropGet(0x0000000b, nil)
	return retVal.LValVal()
}

func (this *IRTDUpdateEvent) SetHeartbeatInterval(rhs int32)  {
	_ = this.PropPut(0x0000000b, []interface{}{rhs})
}

func (this *IRTDUpdateEvent) Disconnect()  {
	retVal, _ := this.Call(0x0000000c, nil)
	_= retVal
}

