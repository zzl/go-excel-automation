package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024459-0000-0000-C000-000000000046
var IID_Graphic = syscall.GUID{0x00024459, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Graphic struct {
	ole.OleClient
}

func NewGraphic(pDisp *win32.IDispatch, addRef bool, scoped bool) *Graphic {
	p := &Graphic{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func GraphicFromVar(v ole.Variant) *Graphic {
	return NewGraphic(v.PdispValVal(), false, false)
}

func (this *Graphic) IID() *syscall.GUID {
	return &IID_Graphic
}

func (this *Graphic) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Graphic) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Graphic) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Graphic) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Graphic) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Graphic) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Graphic) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Graphic) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Graphic) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Graphic) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Graphic) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Graphic) Brightness() float32 {
	retVal := this.PropGet(0x00000892, nil)
	return retVal.FltValVal()
}

func (this *Graphic) SetBrightness(rhs float32)  {
	retVal := this.PropPut(0x00000892, []interface{}{rhs})
	_= retVal
}

func (this *Graphic) ColorType() int32 {
	retVal := this.PropGet(0x00000893, nil)
	return retVal.LValVal()
}

func (this *Graphic) SetColorType(rhs int32)  {
	retVal := this.PropPut(0x00000893, []interface{}{rhs})
	_= retVal
}

func (this *Graphic) Contrast() float32 {
	retVal := this.PropGet(0x00000894, nil)
	return retVal.FltValVal()
}

func (this *Graphic) SetContrast(rhs float32)  {
	retVal := this.PropPut(0x00000894, []interface{}{rhs})
	_= retVal
}

func (this *Graphic) CropBottom() float32 {
	retVal := this.PropGet(0x00000895, nil)
	return retVal.FltValVal()
}

func (this *Graphic) SetCropBottom(rhs float32)  {
	retVal := this.PropPut(0x00000895, []interface{}{rhs})
	_= retVal
}

func (this *Graphic) CropLeft() float32 {
	retVal := this.PropGet(0x00000896, nil)
	return retVal.FltValVal()
}

func (this *Graphic) SetCropLeft(rhs float32)  {
	retVal := this.PropPut(0x00000896, []interface{}{rhs})
	_= retVal
}

func (this *Graphic) CropRight() float32 {
	retVal := this.PropGet(0x00000897, nil)
	return retVal.FltValVal()
}

func (this *Graphic) SetCropRight(rhs float32)  {
	retVal := this.PropPut(0x00000897, []interface{}{rhs})
	_= retVal
}

func (this *Graphic) CropTop() float32 {
	retVal := this.PropGet(0x00000898, nil)
	return retVal.FltValVal()
}

func (this *Graphic) SetCropTop(rhs float32)  {
	retVal := this.PropPut(0x00000898, []interface{}{rhs})
	_= retVal
}

func (this *Graphic) Filename() string {
	retVal := this.PropGet(0x00000587, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Graphic) SetFilename(rhs string)  {
	retVal := this.PropPut(0x00000587, []interface{}{rhs})
	_= retVal
}

func (this *Graphic) Height() float32 {
	retVal := this.PropGet(0x0000007b, nil)
	return retVal.FltValVal()
}

func (this *Graphic) SetHeight(rhs float32)  {
	retVal := this.PropPut(0x0000007b, []interface{}{rhs})
	_= retVal
}

func (this *Graphic) LockAspectRatio() int32 {
	retVal := this.PropGet(0x000006a4, nil)
	return retVal.LValVal()
}

func (this *Graphic) SetLockAspectRatio(rhs int32)  {
	retVal := this.PropPut(0x000006a4, []interface{}{rhs})
	_= retVal
}

func (this *Graphic) Width() float32 {
	retVal := this.PropGet(0x0000007a, nil)
	return retVal.FltValVal()
}

func (this *Graphic) SetWidth(rhs float32)  {
	retVal := this.PropPut(0x0000007a, []interface{}{rhs})
	_= retVal
}

