package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244AA-0000-0000-C000-000000000046
var IID_SortFields = syscall.GUID{0x000244AA, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type SortFields struct {
	ole.OleClient
}

func NewSortFields(pDisp *win32.IDispatch, addRef bool, scoped bool) *SortFields {
	p := &SortFields{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SortFieldsFromVar(v ole.Variant) *SortFields {
	return NewSortFields(v.PdispValVal(), false, false)
}

func (this *SortFields) IID() *syscall.GUID {
	return &IID_SortFields
}

func (this *SortFields) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SortFields) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *SortFields) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *SortFields) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *SortFields) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *SortFields) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *SortFields) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *SortFields) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *SortFields) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *SortFields) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *SortFields) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var SortFields_Add_OptArgs= []string{
	"SortOn", "Order", "CustomOrder", "DataOption", 
}

func (this *SortFields) Add(key *Range, optArgs ...interface{}) *SortField {
	optArgs = ole.ProcessOptArgs(SortFields_Add_OptArgs, optArgs)
	retVal := this.Call(0x000000b5, []interface{}{key}, optArgs...)
	return NewSortField(retVal.PdispValVal(), false, true)
}

func (this *SortFields) Item(index interface{}) *SortField {
	retVal := this.PropGet(0x000000aa, []interface{}{index})
	return NewSortField(retVal.PdispValVal(), false, true)
}

func (this *SortFields) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *SortFields) Clear()  {
	retVal := this.Call(0x0000006f, nil)
	_= retVal
}

func (this *SortFields) Default_(index interface{}) *SortField {
	retVal := this.PropGet(0x00000000, []interface{}{index})
	return NewSortField(retVal.PdispValVal(), false, true)
}

func (this *SortFields) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *SortFields) ForEach(action func(item *SortField) bool) {
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
		pItem := (*SortField)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

