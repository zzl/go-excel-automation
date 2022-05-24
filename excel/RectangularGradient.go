package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000244B0-0000-0000-C000-000000000046
var IID_RectangularGradient = syscall.GUID{0x000244B0, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type RectangularGradient struct {
	ole.OleClient
}

func NewRectangularGradient(pDisp *win32.IDispatch, addRef bool, scoped bool) *RectangularGradient {
	 if pDisp == nil {
		return nil;
	}
	p := &RectangularGradient{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func RectangularGradientFromVar(v ole.Variant) *RectangularGradient {
	return NewRectangularGradient(v.IDispatch(), false, false)
}

func (this *RectangularGradient) IID() *syscall.GUID {
	return &IID_RectangularGradient
}

func (this *RectangularGradient) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *RectangularGradient) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *RectangularGradient) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *RectangularGradient) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *RectangularGradient) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *RectangularGradient) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *RectangularGradient) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *RectangularGradient) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *RectangularGradient) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *RectangularGradient) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *RectangularGradient) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *RectangularGradient) ColorStops() *ColorStops {
	retVal, _ := this.PropGet(0x00000ac9, nil)
	return NewColorStops(retVal.IDispatch(), false, true)
}

func (this *RectangularGradient) RectangleTop() float64 {
	retVal, _ := this.PropGet(0x00000aca, nil)
	return retVal.DblValVal()
}

func (this *RectangularGradient) SetRectangleTop(rhs float64)  {
	_ = this.PropPut(0x00000aca, []interface{}{rhs})
}

func (this *RectangularGradient) RectangleBottom() float64 {
	retVal, _ := this.PropGet(0x00000acb, nil)
	return retVal.DblValVal()
}

func (this *RectangularGradient) SetRectangleBottom(rhs float64)  {
	_ = this.PropPut(0x00000acb, []interface{}{rhs})
}

func (this *RectangularGradient) RectangleLeft() float64 {
	retVal, _ := this.PropGet(0x00000acc, nil)
	return retVal.DblValVal()
}

func (this *RectangularGradient) SetRectangleLeft(rhs float64)  {
	_ = this.PropPut(0x00000acc, []interface{}{rhs})
}

func (this *RectangularGradient) RectangleRight() float64 {
	retVal, _ := this.PropGet(0x00000acd, nil)
	return retVal.DblValVal()
}

func (this *RectangularGradient) SetRectangleRight(rhs float64)  {
	_ = this.PropPut(0x00000acd, []interface{}{rhs})
}

