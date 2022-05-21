package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244C9-0000-0000-C000-000000000046
var IID_SlicerItem = syscall.GUID{0x000244C9, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type SlicerItem struct {
	ole.OleClient
}

func NewSlicerItem(pDisp *win32.IDispatch, addRef bool, scoped bool) *SlicerItem {
	p := &SlicerItem{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SlicerItemFromVar(v ole.Variant) *SlicerItem {
	return NewSlicerItem(v.PdispValVal(), false, false)
}

func (this *SlicerItem) IID() *syscall.GUID {
	return &IID_SlicerItem
}

func (this *SlicerItem) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SlicerItem) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *SlicerItem) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *SlicerItem) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *SlicerItem) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *SlicerItem) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *SlicerItem) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *SlicerItem) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *SlicerItem) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *SlicerItem) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *SlicerItem) Parent() *SlicerCache {
	retVal := this.PropGet(0x00000096, nil)
	return NewSlicerCache(retVal.PdispValVal(), false, true)
}

func (this *SlicerItem) Caption() string {
	retVal := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SlicerItem) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SlicerItem) SourceName() ole.Variant {
	retVal := this.PropGet(0x000002d1, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *SlicerItem) SourceNameStandard() string {
	retVal := this.PropGet(0x00000864, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SlicerItem) Value() string {
	retVal := this.PropGet(0x00000006, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SlicerItem) Selected() bool {
	retVal := this.PropGet(0x00000463, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SlicerItem) SetSelected(rhs bool)  {
	retVal := this.PropPut(0x00000463, []interface{}{rhs})
	_= retVal
}

func (this *SlicerItem) HasData() bool {
	retVal := this.PropGet(0x00000bad, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

