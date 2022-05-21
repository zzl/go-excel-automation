package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024421-0000-0000-C000-000000000046
var IID_CalculatedItems = syscall.GUID{0x00024421, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type CalculatedItems struct {
	ole.OleClient
}

func NewCalculatedItems(pDisp *win32.IDispatch, addRef bool, scoped bool) *CalculatedItems {
	p := &CalculatedItems{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CalculatedItemsFromVar(v ole.Variant) *CalculatedItems {
	return NewCalculatedItems(v.PdispValVal(), false, false)
}

func (this *CalculatedItems) IID() *syscall.GUID {
	return &IID_CalculatedItems
}

func (this *CalculatedItems) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CalculatedItems) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *CalculatedItems) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *CalculatedItems) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *CalculatedItems) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *CalculatedItems) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *CalculatedItems) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *CalculatedItems) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *CalculatedItems) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *CalculatedItems) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *CalculatedItems) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *CalculatedItems) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *CalculatedItems) Add_(name string, formula string) *PivotItem {
	retVal := this.Call(0x00000825, []interface{}{name, formula})
	return NewPivotItem(retVal.PdispValVal(), false, true)
}

func (this *CalculatedItems) Item(index interface{}) *PivotItem {
	retVal := this.Call(0x000000aa, []interface{}{index})
	return NewPivotItem(retVal.PdispValVal(), false, true)
}

func (this *CalculatedItems) Default_(field interface{}) *PivotItem {
	retVal := this.PropGet(0x00000000, []interface{}{field})
	return NewPivotItem(retVal.PdispValVal(), false, true)
}

func (this *CalculatedItems) NewEnum_() *com.UnknownClass {
	retVal := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *CalculatedItems) ForEach(action func(item *PivotItem) bool) {
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
		pItem := (*PivotItem)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

var CalculatedItems_Add_OptArgs= []string{
	"UseStandardFormula", 
}

func (this *CalculatedItems) Add(name string, formula string, optArgs ...interface{}) *PivotItem {
	optArgs = ole.ProcessOptArgs(CalculatedItems_Add_OptArgs, optArgs)
	retVal := this.Call(0x000000b5, []interface{}{name, formula}, optArgs...)
	return NewPivotItem(retVal.PdispValVal(), false, true)
}

