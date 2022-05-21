package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244C5-0000-0000-C000-000000000046
var IID_SlicerCacheLevels = syscall.GUID{0x000244C5, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type SlicerCacheLevels struct {
	ole.OleClient
}

func NewSlicerCacheLevels(pDisp *win32.IDispatch, addRef bool, scoped bool) *SlicerCacheLevels {
	p := &SlicerCacheLevels{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SlicerCacheLevelsFromVar(v ole.Variant) *SlicerCacheLevels {
	return NewSlicerCacheLevels(v.PdispValVal(), false, false)
}

func (this *SlicerCacheLevels) IID() *syscall.GUID {
	return &IID_SlicerCacheLevels
}

func (this *SlicerCacheLevels) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SlicerCacheLevels) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *SlicerCacheLevels) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *SlicerCacheLevels) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *SlicerCacheLevels) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *SlicerCacheLevels) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *SlicerCacheLevels) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *SlicerCacheLevels) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *SlicerCacheLevels) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *SlicerCacheLevels) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *SlicerCacheLevels) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *SlicerCacheLevels) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

var SlicerCacheLevels_Item_OptArgs= []string{
	"Level", 
}

func (this *SlicerCacheLevels) Item(optArgs ...interface{}) *SlicerCacheLevel {
	optArgs = ole.ProcessOptArgs(SlicerCacheLevels_Item_OptArgs, optArgs)
	retVal := this.PropGet(0x000000aa, nil, optArgs...)
	return NewSlicerCacheLevel(retVal.PdispValVal(), false, true)
}

var SlicerCacheLevels_Default__OptArgs= []string{
	"Level", 
}

func (this *SlicerCacheLevels) Default_(optArgs ...interface{}) *SlicerCacheLevel {
	optArgs = ole.ProcessOptArgs(SlicerCacheLevels_Default__OptArgs, optArgs)
	retVal := this.PropGet(0x00000000, nil, optArgs...)
	return NewSlicerCacheLevel(retVal.PdispValVal(), false, true)
}

func (this *SlicerCacheLevels) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *SlicerCacheLevels) ForEach(action func(item *SlicerCacheLevel) bool) {
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
		pItem := (*SlicerCacheLevel)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

