package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024420-0000-0000-C000-000000000046
var IID_CalculatedFields = syscall.GUID{0x00024420, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type CalculatedFields struct {
	ole.OleClient
}

func NewCalculatedFields(pDisp *win32.IDispatch, addRef bool, scoped bool) *CalculatedFields {
	 if pDisp == nil {
		return nil;
	}
	p := &CalculatedFields{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CalculatedFieldsFromVar(v ole.Variant) *CalculatedFields {
	return NewCalculatedFields(v.IDispatch(), false, false)
}

func (this *CalculatedFields) IID() *syscall.GUID {
	return &IID_CalculatedFields
}

func (this *CalculatedFields) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CalculatedFields) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *CalculatedFields) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *CalculatedFields) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *CalculatedFields) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *CalculatedFields) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *CalculatedFields) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *CalculatedFields) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *CalculatedFields) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *CalculatedFields) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *CalculatedFields) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *CalculatedFields) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *CalculatedFields) Add_(name string, formula string) *PivotField {
	retVal, _ := this.Call(0x00000825, []interface{}{name, formula})
	return NewPivotField(retVal.IDispatch(), false, true)
}

func (this *CalculatedFields) Item(index interface{}) *PivotField {
	retVal, _ := this.Call(0x000000aa, []interface{}{index})
	return NewPivotField(retVal.IDispatch(), false, true)
}

func (this *CalculatedFields) Default_(field interface{}) *PivotField {
	retVal, _ := this.PropGet(0x00000000, []interface{}{field})
	return NewPivotField(retVal.IDispatch(), false, true)
}

func (this *CalculatedFields) NewEnum_() *com.UnknownClass {
	retVal, _ := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *CalculatedFields) ForEach(action func(item *PivotField) bool) {
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
		pItem := (*PivotField)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

var CalculatedFields_Add_OptArgs= []string{
	"UseStandardFormula", 
}

func (this *CalculatedFields) Add(name string, formula string, optArgs ...interface{}) *PivotField {
	optArgs = ole.ProcessOptArgs(CalculatedFields_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b5, []interface{}{name, formula}, optArgs...)
	return NewPivotField(retVal.IDispatch(), false, true)
}

