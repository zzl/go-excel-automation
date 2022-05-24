package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244C6-0000-0000-C000-000000000046
var IID_SlicerCacheLevel = syscall.GUID{0x000244C6, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type SlicerCacheLevel struct {
	ole.OleClient
}

func NewSlicerCacheLevel(pDisp *win32.IDispatch, addRef bool, scoped bool) *SlicerCacheLevel {
	 if pDisp == nil {
		return nil;
	}
	p := &SlicerCacheLevel{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SlicerCacheLevelFromVar(v ole.Variant) *SlicerCacheLevel {
	return NewSlicerCacheLevel(v.IDispatch(), false, false)
}

func (this *SlicerCacheLevel) IID() *syscall.GUID {
	return &IID_SlicerCacheLevel
}

func (this *SlicerCacheLevel) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SlicerCacheLevel) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *SlicerCacheLevel) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *SlicerCacheLevel) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *SlicerCacheLevel) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *SlicerCacheLevel) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *SlicerCacheLevel) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *SlicerCacheLevel) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *SlicerCacheLevel) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *SlicerCacheLevel) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *SlicerCacheLevel) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *SlicerCacheLevel) SlicerItems() *SlicerItems {
	retVal, _ := this.PropGet(0x00000ba1, nil)
	return NewSlicerItems(retVal.IDispatch(), false, true)
}

func (this *SlicerCacheLevel) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *SlicerCacheLevel) Ordinal() int32 {
	retVal, _ := this.PropGet(0x00000ba5, nil)
	return retVal.LValVal()
}

func (this *SlicerCacheLevel) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SlicerCacheLevel) CrossFilterType() int32 {
	retVal, _ := this.PropGet(0x00000ba2, nil)
	return retVal.LValVal()
}

func (this *SlicerCacheLevel) SetCrossFilterType(rhs int32)  {
	_ = this.PropPut(0x00000ba2, []interface{}{rhs})
}

func (this *SlicerCacheLevel) SortItems() int32 {
	retVal, _ := this.PropGet(0x00000ba3, nil)
	return retVal.LValVal()
}

func (this *SlicerCacheLevel) SetSortItems(rhs int32)  {
	_ = this.PropPut(0x00000ba3, []interface{}{rhs})
}

func (this *SlicerCacheLevel) VisibleSlicerItemsList() ole.Variant {
	retVal, _ := this.PropGet(0x00000ba0, nil)
	com.AddToScope(retVal)
	return *retVal
}

