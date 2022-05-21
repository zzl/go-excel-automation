package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024466-0000-0000-C000-000000000046
var IID_Speech = syscall.GUID{0x00024466, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Speech struct {
	ole.OleClient
}

func NewSpeech(pDisp *win32.IDispatch, addRef bool, scoped bool) *Speech {
	p := &Speech{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SpeechFromVar(v ole.Variant) *Speech {
	return NewSpeech(v.PdispValVal(), false, false)
}

func (this *Speech) IID() *syscall.GUID {
	return &IID_Speech
}

func (this *Speech) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Speech) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Speech) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Speech) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Speech) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Speech) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Speech) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Speech) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

var Speech_Speak_OptArgs= []string{
	"SpeakAsync", "SpeakXML", "Purge", 
}

func (this *Speech) Speak(text string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Speech_Speak_OptArgs, optArgs)
	retVal := this.Call(0x000007e1, []interface{}{text}, optArgs...)
	_= retVal
}

func (this *Speech) Direction() int32 {
	retVal := this.PropGet(0x000000a8, nil)
	return retVal.LValVal()
}

func (this *Speech) SetDirection(rhs int32)  {
	retVal := this.PropPut(0x000000a8, []interface{}{rhs})
	_= retVal
}

func (this *Speech) SpeakCellOnEnter() bool {
	retVal := this.PropGet(0x000008bb, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Speech) SetSpeakCellOnEnter(rhs bool)  {
	retVal := this.PropPut(0x000008bb, []interface{}{rhs})
	_= retVal
}

