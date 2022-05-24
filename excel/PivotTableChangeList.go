package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244C1-0000-0000-C000-000000000046
var IID_PivotTableChangeList = syscall.GUID{0x000244C1, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type PivotTableChangeList struct {
	ole.OleClient
}

func NewPivotTableChangeList(pDisp *win32.IDispatch, addRef bool, scoped bool) *PivotTableChangeList {
	 if pDisp == nil {
		return nil;
	}
	p := &PivotTableChangeList{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PivotTableChangeListFromVar(v ole.Variant) *PivotTableChangeList {
	return NewPivotTableChangeList(v.IDispatch(), false, false)
}

func (this *PivotTableChangeList) IID() *syscall.GUID {
	return &IID_PivotTableChangeList
}

func (this *PivotTableChangeList) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *PivotTableChangeList) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *PivotTableChangeList) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *PivotTableChangeList) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *PivotTableChangeList) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *PivotTableChangeList) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *PivotTableChangeList) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *PivotTableChangeList) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *PivotTableChangeList) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *PivotTableChangeList) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *PivotTableChangeList) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *PivotTableChangeList) Default_(index interface{}) *ValueChange {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return NewValueChange(retVal.IDispatch(), false, true)
}

func (this *PivotTableChangeList) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *PivotTableChangeList) ForEach(action func(item *ValueChange) bool) {
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
		pItem := (*ValueChange)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *PivotTableChangeList) Item(index interface{}) *ValueChange {
	retVal, _ := this.PropGet(0x000000aa, []interface{}{index})
	return NewValueChange(retVal.IDispatch(), false, true)
}

func (this *PivotTableChangeList) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

var PivotTableChangeList_Add_OptArgs= []string{
	"AllocationValue", "AllocationMethod", "AllocationWeightExpression", 
}

func (this *PivotTableChangeList) Add(tuple string, value float64, optArgs ...interface{}) *ValueChange {
	optArgs = ole.ProcessOptArgs(PivotTableChangeList_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b5, []interface{}{tuple, value}, optArgs...)
	return NewValueChange(retVal.IDispatch(), false, true)
}

