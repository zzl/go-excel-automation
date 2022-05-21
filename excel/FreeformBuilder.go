package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002443F-0000-0000-C000-000000000046
var IID_FreeformBuilder = syscall.GUID{0x0002443F, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type FreeformBuilder struct {
	ole.OleClient
}

func NewFreeformBuilder(pDisp *win32.IDispatch, addRef bool, scoped bool) *FreeformBuilder {
	p := &FreeformBuilder{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FreeformBuilderFromVar(v ole.Variant) *FreeformBuilder {
	return NewFreeformBuilder(v.PdispValVal(), false, false)
}

func (this *FreeformBuilder) IID() *syscall.GUID {
	return &IID_FreeformBuilder
}

func (this *FreeformBuilder) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *FreeformBuilder) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *FreeformBuilder) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *FreeformBuilder) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *FreeformBuilder) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *FreeformBuilder) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *FreeformBuilder) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *FreeformBuilder) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *FreeformBuilder) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *FreeformBuilder) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *FreeformBuilder) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var FreeformBuilder_AddNodes_OptArgs= []string{
	"X2", "Y2", "X3", "Y3", 
}

func (this *FreeformBuilder) AddNodes(segmentType int32, editingType int32, x1 float32, y1 float32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(FreeformBuilder_AddNodes_OptArgs, optArgs)
	retVal := this.Call(0x000006e2, []interface{}{segmentType, editingType, x1, y1}, optArgs...)
	_= retVal
}

func (this *FreeformBuilder) ConvertToShape() *Shape {
	retVal := this.Call(0x000006e6, nil)
	return NewShape(retVal.PdispValVal(), false, true)
}

