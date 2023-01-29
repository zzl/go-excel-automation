package excel

import (
	"github.com/zzl/go-win32api/v2/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024484-0000-0000-C000-000000000046
var IID_PivotFilters = syscall.GUID{0x00024484, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type PivotFilters struct {
	ole.OleClient
}

func NewPivotFilters(pDisp *win32.IDispatch, addRef bool, scoped bool) *PivotFilters {
	if pDisp == nil {
		return nil
	}
	p := &PivotFilters{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PivotFiltersFromVar(v ole.Variant) *PivotFilters {
	return NewPivotFilters(v.IDispatch(), false, false)
}

func (this *PivotFilters) IID() *syscall.GUID {
	return &IID_PivotFilters
}

func (this *PivotFilters) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *PivotFilters) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *PivotFilters) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *PivotFilters) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *PivotFilters) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *PivotFilters) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *PivotFilters) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *PivotFilters) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *PivotFilters) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *PivotFilters) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *PivotFilters) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *PivotFilters) Default_(index interface{}) *PivotFilter {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return NewPivotFilter(retVal.IDispatch(), false, true)
}

func (this *PivotFilters) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *PivotFilters) ForEach(action func(item *PivotFilter) bool) {
	pEnum := this.NewEnum_()
	var pEnumVar *win32.IEnumVARIANT
	pEnum.QueryInterface(&win32.IID_IEnumVARIANT, unsafe.Pointer(&pEnumVar))
	defer pEnumVar.Release()
	for {
		var c uint32
		var v ole.Variant
		pEnumVar.Next(1, (*win32.VARIANT)(&v), &c)
		if c == 0 {
			break
		}
		pItem := (*PivotFilter)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *PivotFilters) Item(index interface{}) *PivotFilter {
	retVal, _ := this.PropGet(0x000000aa, []interface{}{index})
	return NewPivotFilter(retVal.IDispatch(), false, true)
}

func (this *PivotFilters) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

var PivotFilters_Add_OptArgs = []string{
	"DataField", "Value1", "Value2", "Order",
	"Name", "Description", "MemberPropertyField",
}

func (this *PivotFilters) Add(type_ int32, optArgs ...interface{}) *PivotFilter {
	optArgs = ole.ProcessOptArgs(PivotFilters_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b5, []interface{}{type_}, optArgs...)
	return NewPivotFilter(retVal.IDispatch(), false, true)
}

