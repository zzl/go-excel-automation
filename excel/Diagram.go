package excel

import (
	"github.com/zzl/go-win32api/v2/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002446F-0000-0000-C000-000000000046
var IID_Diagram = syscall.GUID{0x0002446F, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Diagram struct {
	ole.OleClient
}

func NewDiagram(pDisp *win32.IDispatch, addRef bool, scoped bool) *Diagram {
	if pDisp == nil {
		return nil
	}
	p := &Diagram{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DiagramFromVar(v ole.Variant) *Diagram {
	return NewDiagram(v.IDispatch(), false, false)
}

func (this *Diagram) IID() *syscall.GUID {
	return &IID_Diagram
}

func (this *Diagram) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Diagram) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Diagram) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Diagram) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Diagram) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Diagram) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Diagram) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Diagram) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Diagram) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Diagram) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Diagram) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Diagram) Nodes() *DiagramNodes {
	retVal, _ := this.PropGet(0x000006a5, nil)
	return NewDiagramNodes(retVal.IDispatch(), false, true)
}

func (this *Diagram) Type() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *Diagram) AutoLayout() int32 {
	retVal, _ := this.PropGet(0x000008c3, nil)
	return retVal.LValVal()
}

func (this *Diagram) SetAutoLayout(rhs int32) {
	_ = this.PropPut(0x000008c3, []interface{}{rhs})
}

func (this *Diagram) Reverse() int32 {
	retVal, _ := this.PropGet(0x000008c4, nil)
	return retVal.LValVal()
}

func (this *Diagram) SetReverse(rhs int32) {
	_ = this.PropPut(0x000008c4, []interface{}{rhs})
}

func (this *Diagram) AutoFormat() int32 {
	retVal, _ := this.PropGet(0x00000072, nil)
	return retVal.LValVal()
}

func (this *Diagram) SetAutoFormat(rhs int32) {
	_ = this.PropPut(0x00000072, []interface{}{rhs})
}

func (this *Diagram) Convert(type_ int32) {
	retVal, _ := this.Call(0x00000416, []interface{}{type_})
	_ = retVal
}

func (this *Diagram) FitText() {
	retVal, _ := this.Call(0x00000900, nil)
	_ = retVal
}
