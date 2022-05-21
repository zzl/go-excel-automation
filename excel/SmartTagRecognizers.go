package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024463-0000-0000-C000-000000000046
var IID_SmartTagRecognizers = syscall.GUID{0x00024463, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type SmartTagRecognizers struct {
	ole.OleClient
}

func NewSmartTagRecognizers(pDisp *win32.IDispatch, addRef bool, scoped bool) *SmartTagRecognizers {
	p := &SmartTagRecognizers{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SmartTagRecognizersFromVar(v ole.Variant) *SmartTagRecognizers {
	return NewSmartTagRecognizers(v.PdispValVal(), false, false)
}

func (this *SmartTagRecognizers) IID() *syscall.GUID {
	return &IID_SmartTagRecognizers
}

func (this *SmartTagRecognizers) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SmartTagRecognizers) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *SmartTagRecognizers) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *SmartTagRecognizers) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *SmartTagRecognizers) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *SmartTagRecognizers) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *SmartTagRecognizers) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *SmartTagRecognizers) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *SmartTagRecognizers) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *SmartTagRecognizers) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *SmartTagRecognizers) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *SmartTagRecognizers) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *SmartTagRecognizers) Item(index interface{}) *SmartTagRecognizer {
	retVal := this.PropGet(0x000000aa, []interface{}{index})
	return NewSmartTagRecognizer(retVal.PdispValVal(), false, true)
}

func (this *SmartTagRecognizers) Default_(index interface{}) *SmartTagRecognizer {
	retVal := this.PropGet(0x00000000, []interface{}{index})
	return NewSmartTagRecognizer(retVal.PdispValVal(), false, true)
}

func (this *SmartTagRecognizers) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *SmartTagRecognizers) ForEach(action func(item *SmartTagRecognizer) bool) {
	pEnum := this.NewEnum_()
	var pEnumVar *win32.IEnumVARIANT
	pEnum.QueryInterface(&win32.IID_IEnumVARIANT, unsafe.Pointer(&pEnumVar))
	defer pEnumVar.Release();
	for {
		var c uint32
		var v ole.Variant
		pEnumVar.Next(1, (*win32.VARIANT)(&v), &c)
		if c == 0 {
			break
		}
		pItem := (*SmartTagRecognizer)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *SmartTagRecognizers) Recognize() bool {
	retVal := this.PropGet(0x000008a9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SmartTagRecognizers) SetRecognize(rhs bool)  {
	retVal := this.PropPut(0x000008a9, []interface{}{rhs})
	_= retVal
}

