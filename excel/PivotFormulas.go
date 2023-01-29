package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 0002441F-0000-0000-C000-000000000046
var IID_PivotFormulas = syscall.GUID{0x0002441F, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type PivotFormulas struct {
	ole.OleClient
}

func NewPivotFormulas(pDisp *win32.IDispatch, addRef bool, scoped bool) *PivotFormulas {
	if pDisp == nil {
		return nil
	}
	p := &PivotFormulas{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PivotFormulasFromVar(v ole.Variant) *PivotFormulas {
	return NewPivotFormulas(v.IDispatch(), false, false)
}

func (this *PivotFormulas) IID() *syscall.GUID {
	return &IID_PivotFormulas
}

func (this *PivotFormulas) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *PivotFormulas) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *PivotFormulas) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *PivotFormulas) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *PivotFormulas) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *PivotFormulas) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *PivotFormulas) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *PivotFormulas) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *PivotFormulas) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *PivotFormulas) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *PivotFormulas) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *PivotFormulas) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *PivotFormulas) Add_(formula string) *PivotFormula {
	retVal, _ := this.Call(0x00000825, []interface{}{formula})
	return NewPivotFormula(retVal.IDispatch(), false, true)
}

func (this *PivotFormulas) Item(index interface{}) *PivotFormula {
	retVal, _ := this.Call(0x000000aa, []interface{}{index})
	return NewPivotFormula(retVal.IDispatch(), false, true)
}

func (this *PivotFormulas) Default_(index interface{}) *PivotFormula {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return NewPivotFormula(retVal.IDispatch(), false, true)
}

func (this *PivotFormulas) NewEnum_() *com.UnknownClass {
	retVal, _ := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *PivotFormulas) ForEach(action func(item *PivotFormula) bool) {
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
		pItem := (*PivotFormula)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

var PivotFormulas_Add_OptArgs = []string{
	"UseStandardFormula",
}

func (this *PivotFormulas) Add(formula string, optArgs ...interface{}) *PivotFormula {
	optArgs = ole.ProcessOptArgs(PivotFormulas_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b5, []interface{}{formula}, optArgs...)
	return NewPivotFormula(retVal.IDispatch(), false, true)
}
