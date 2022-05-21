package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244AC-0000-0000-C000-000000000046
var IID_Research = syscall.GUID{0x000244AC, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Research struct {
	ole.OleClient
}

func NewResearch(pDisp *win32.IDispatch, addRef bool, scoped bool) *Research {
	p := &Research{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ResearchFromVar(v ole.Variant) *Research {
	return NewResearch(v.PdispValVal(), false, false)
}

func (this *Research) IID() *syscall.GUID {
	return &IID_Research
}

func (this *Research) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Research) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Research) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Research) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Research) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Research) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Research) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Research) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Research) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Research) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Research) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Research_Query_OptArgs= []string{
	"QueryString", "QueryLanguage", "UseSelection", "LaunchQuery", 
}

func (this *Research) Query(serviceID string, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Research_Query_OptArgs, optArgs)
	retVal := this.Call(0x00000abf, []interface{}{serviceID}, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Research) IsResearchService(serviceID string) bool {
	retVal := this.Call(0x00000ac5, []interface{}{serviceID})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Research) SetLanguagePair(languageFrom int32, languageTo int32) ole.Variant {
	retVal := this.Call(0x00000ac6, []interface{}{languageFrom, languageTo})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

