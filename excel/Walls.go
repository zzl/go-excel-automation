package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000208C8-0000-0000-C000-000000000046
var IID_Walls = syscall.GUID{0x000208C8, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Walls struct {
	ole.OleClient
}

func NewWalls(pDisp *win32.IDispatch, addRef bool, scoped bool) *Walls {
	 if pDisp == nil {
		return nil;
	}
	p := &Walls{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func WallsFromVar(v ole.Variant) *Walls {
	return NewWalls(v.IDispatch(), false, false)
}

func (this *Walls) IID() *syscall.GUID {
	return &IID_Walls
}

func (this *Walls) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Walls) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Walls) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Walls) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Walls) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Walls) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Walls) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Walls) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Walls) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Walls) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Walls) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Walls) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Walls) Select() ole.Variant {
	retVal, _ := this.Call(0x000000eb, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Walls) Border() *Border {
	retVal, _ := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.IDispatch(), false, true)
}

func (this *Walls) ClearFormats() ole.Variant {
	retVal, _ := this.Call(0x00000070, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Walls) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *Walls) Fill() *ChartFillFormat {
	retVal, _ := this.PropGet(0x0000067f, nil)
	return NewChartFillFormat(retVal.IDispatch(), false, true)
}

func (this *Walls) PictureType() ole.Variant {
	retVal, _ := this.PropGet(0x000000a1, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Walls) SetPictureType(rhs interface{})  {
	_ = this.PropPut(0x000000a1, []interface{}{rhs})
}

func (this *Walls) Paste()  {
	retVal, _ := this.Call(0x000000d3, nil)
	_= retVal
}

func (this *Walls) PictureUnit() ole.Variant {
	retVal, _ := this.PropGet(0x000000a2, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Walls) SetPictureUnit(rhs interface{})  {
	_ = this.PropPut(0x000000a2, []interface{}{rhs})
}

func (this *Walls) Thickness() int32 {
	retVal, _ := this.PropGet(0x00000a5c, nil)
	return retVal.LValVal()
}

func (this *Walls) SetThickness(rhs int32)  {
	_ = this.PropPut(0x00000a5c, []interface{}{rhs})
}

func (this *Walls) Format() *ChartFormat {
	retVal, _ := this.PropGet(0x00000074, nil)
	return NewChartFormat(retVal.IDispatch(), false, true)
}

