package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00024494-0000-0000-C000-000000000046
var IID_ColorScaleCriteria = syscall.GUID{0x00024494, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ColorScaleCriteria struct {
	ole.OleClient
}

func NewColorScaleCriteria(pDisp *win32.IDispatch, addRef bool, scoped bool) *ColorScaleCriteria {
	if pDisp == nil {
		return nil
	}
	p := &ColorScaleCriteria{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ColorScaleCriteriaFromVar(v ole.Variant) *ColorScaleCriteria {
	return NewColorScaleCriteria(v.IDispatch(), false, false)
}

func (this *ColorScaleCriteria) IID() *syscall.GUID {
	return &IID_ColorScaleCriteria
}

func (this *ColorScaleCriteria) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ColorScaleCriteria) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *ColorScaleCriteria) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ColorScaleCriteria) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ColorScaleCriteria) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *ColorScaleCriteria) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *ColorScaleCriteria) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *ColorScaleCriteria) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *ColorScaleCriteria) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *ColorScaleCriteria) Default_(index interface{}) *ColorScaleCriterion {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return NewColorScaleCriterion(retVal.IDispatch(), false, true)
}

func (this *ColorScaleCriteria) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *ColorScaleCriteria) ForEach(action func(item *ColorScaleCriterion) bool) {
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
		pItem := (*ColorScaleCriterion)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *ColorScaleCriteria) Item(index interface{}) *ColorScaleCriterion {
	retVal, _ := this.PropGet(0x000000aa, []interface{}{index})
	return NewColorScaleCriterion(retVal.IDispatch(), false, true)
}
