package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020855-0000-0000-C000-000000000046
var IID_Borders = syscall.GUID{0x00020855, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Borders struct {
	ole.OleClient
}

func NewBorders(pDisp *win32.IDispatch, addRef bool, scoped bool) *Borders {
	 if pDisp == nil {
		return nil;
	}
	p := &Borders{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func BordersFromVar(v ole.Variant) *Borders {
	return NewBorders(v.IDispatch(), false, false)
}

func (this *Borders) IID() *syscall.GUID {
	return &IID_Borders
}

func (this *Borders) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Borders) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Borders) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Borders) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Borders) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Borders) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Borders) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Borders) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Borders) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Borders) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Borders) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Borders) Color() ole.Variant {
	retVal, _ := this.PropGet(0x00000063, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Borders) SetColor(rhs interface{})  {
	_ = this.PropPut(0x00000063, []interface{}{rhs})
}

func (this *Borders) ColorIndex() ole.Variant {
	retVal, _ := this.PropGet(0x00000061, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Borders) SetColorIndex(rhs interface{})  {
	_ = this.PropPut(0x00000061, []interface{}{rhs})
}

func (this *Borders) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *Borders) Item(index int32) *Border {
	retVal, _ := this.PropGet(0x000000aa, []interface{}{index})
	return NewBorder(retVal.IDispatch(), false, true)
}

func (this *Borders) LineStyle() ole.Variant {
	retVal, _ := this.PropGet(0x00000077, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Borders) SetLineStyle(rhs interface{})  {
	_ = this.PropPut(0x00000077, []interface{}{rhs})
}

func (this *Borders) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Borders) ForEach(action func(item *Border) bool) {
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
		pItem := (*Border)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Borders) Value() ole.Variant {
	retVal, _ := this.PropGet(0x00000006, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Borders) SetValue(rhs interface{})  {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *Borders) Weight() ole.Variant {
	retVal, _ := this.PropGet(0x00000078, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Borders) SetWeight(rhs interface{})  {
	_ = this.PropPut(0x00000078, []interface{}{rhs})
}

func (this *Borders) Default_(index int32) *Border {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return NewBorder(retVal.IDispatch(), false, true)
}

func (this *Borders) ThemeColor() ole.Variant {
	retVal, _ := this.PropGet(0x0000093d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Borders) SetThemeColor(rhs interface{})  {
	_ = this.PropPut(0x0000093d, []interface{}{rhs})
}

func (this *Borders) TintAndShade() ole.Variant {
	retVal, _ := this.PropGet(0x0000093e, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Borders) SetTintAndShade(rhs interface{})  {
	_ = this.PropPut(0x0000093e, []interface{}{rhs})
}

