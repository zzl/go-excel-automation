package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024424-0000-0000-C000-000000000046
var IID_FormatConditions = syscall.GUID{0x00024424, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type FormatConditions struct {
	ole.OleClient
}

func NewFormatConditions(pDisp *win32.IDispatch, addRef bool, scoped bool) *FormatConditions {
	p := &FormatConditions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FormatConditionsFromVar(v ole.Variant) *FormatConditions {
	return NewFormatConditions(v.PdispValVal(), false, false)
}

func (this *FormatConditions) IID() *syscall.GUID {
	return &IID_FormatConditions
}

func (this *FormatConditions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *FormatConditions) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *FormatConditions) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *FormatConditions) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *FormatConditions) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *FormatConditions) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *FormatConditions) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *FormatConditions) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *FormatConditions) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *FormatConditions) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *FormatConditions) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *FormatConditions) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *FormatConditions) Item(index interface{}) *ole.DispatchClass {
	retVal := this.Call(0x000000aa, []interface{}{index})
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var FormatConditions_Add_OptArgs= []string{
	"Operator", "Formula1", "Formula2", "String", 
	"TextOperator", "DateOperator", "ScopeType", 
}

func (this *FormatConditions) Add(type_ int32, optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(FormatConditions_Add_OptArgs, optArgs)
	retVal := this.Call(0x000000b5, []interface{}{type_}, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *FormatConditions) Default_(index interface{}) *ole.DispatchClass {
	retVal := this.PropGet(0x00000000, []interface{}{index})
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *FormatConditions) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *FormatConditions) ForEach(action func(item *ole.DispatchClass) bool) {
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
		pItem := (*ole.DispatchClass)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *FormatConditions) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

func (this *FormatConditions) AddColorScale(colorScaleType int32) *ole.DispatchClass {
	retVal := this.Call(0x00000a38, []interface{}{colorScaleType})
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *FormatConditions) AddDatabar() *ole.DispatchClass {
	retVal := this.Call(0x00000a3a, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *FormatConditions) AddIconSetCondition() *ole.DispatchClass {
	retVal := this.Call(0x00000a3b, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *FormatConditions) AddTop10() *ole.DispatchClass {
	retVal := this.Call(0x00000a3c, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *FormatConditions) AddAboveAverage() *ole.DispatchClass {
	retVal := this.Call(0x00000a3d, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *FormatConditions) AddUniqueValues() *ole.DispatchClass {
	retVal := this.Call(0x00000a3e, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

