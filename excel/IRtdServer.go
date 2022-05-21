package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// EC0E6191-DB51-11D3-8F3E-00C04F3651B8
var IID_IRtdServer = syscall.GUID{0xEC0E6191, 0xDB51, 0x11D3, 
	[8]byte{0x8F, 0x3E, 0x00, 0xC0, 0x4F, 0x36, 0x51, 0xB8}}

type IRtdServer struct {
	ole.OleClient
}

func NewIRtdServer(pDisp *win32.IDispatch, addRef bool, scoped bool) *IRtdServer {
	p := &IRtdServer{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func IRtdServerFromVar(v ole.Variant) *IRtdServer {
	return NewIRtdServer(v.PdispValVal(), false, false)
}

func (this *IRtdServer) IID() *syscall.GUID {
	return &IID_IRtdServer
}

func (this *IRtdServer) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *IRtdServer) ServerStart(callbackObject *IRTDUpdateEvent) int32 {
	retVal := this.Call(0x0000000a, []interface{}{callbackObject})
	return retVal.LValVal()
}

func (this *IRtdServer) ConnectData(topicID int32, strings **win32.SAFEARRAY, getNewValues *win32.VARIANT_BOOL) ole.Variant {
	retVal := this.Call(0x0000000b, []interface{}{topicID, strings, getNewValues})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *IRtdServer) RefreshData(topicCount *int32) *win32.SAFEARRAY {
	retVal := this.Call(0x0000000c, []interface{}{topicCount})
	return retVal.ParrayVal()
}

func (this *IRtdServer) DisconnectData(topicID int32)  {
	retVal := this.Call(0x0000000d, []interface{}{topicID})
	_= retVal
}

func (this *IRtdServer) Heartbeat() int32 {
	retVal := this.Call(0x0000000e, nil)
	return retVal.LValVal()
}

func (this *IRtdServer) ServerTerminate()  {
	retVal := this.Call(0x0000000f, nil)
	_= retVal
}

