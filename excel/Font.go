package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002084D-0000-0000-C000-000000000046
var IID_Font = syscall.GUID{0x0002084D, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Font struct {
	ole.OleClient
}

func NewFont(pDisp *win32.IDispatch, addRef bool, scoped bool) *Font {
	p := &Font{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FontFromVar(v ole.Variant) *Font {
	return NewFont(v.PdispValVal(), false, false)
}

func (this *Font) IID() *syscall.GUID {
	return &IID_Font
}

func (this *Font) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Font) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Font) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Font) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Font) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Font) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Font) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Font) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Font) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Font) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Font) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Font) Background() ole.Variant {
	retVal := this.PropGet(0x000000b4, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Font) SetBackground(rhs interface{})  {
	retVal := this.PropPut(0x000000b4, []interface{}{rhs})
	_= retVal
}

func (this *Font) Bold() ole.Variant {
	retVal := this.PropGet(0x00000060, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Font) SetBold(rhs interface{})  {
	retVal := this.PropPut(0x00000060, []interface{}{rhs})
	_= retVal
}

func (this *Font) Color() ole.Variant {
	retVal := this.PropGet(0x00000063, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Font) SetColor(rhs interface{})  {
	retVal := this.PropPut(0x00000063, []interface{}{rhs})
	_= retVal
}

func (this *Font) ColorIndex() ole.Variant {
	retVal := this.PropGet(0x00000061, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Font) SetColorIndex(rhs interface{})  {
	retVal := this.PropPut(0x00000061, []interface{}{rhs})
	_= retVal
}

func (this *Font) FontStyle() ole.Variant {
	retVal := this.PropGet(0x000000b1, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Font) SetFontStyle(rhs interface{})  {
	retVal := this.PropPut(0x000000b1, []interface{}{rhs})
	_= retVal
}

func (this *Font) Italic() ole.Variant {
	retVal := this.PropGet(0x00000065, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Font) SetItalic(rhs interface{})  {
	retVal := this.PropPut(0x00000065, []interface{}{rhs})
	_= retVal
}

func (this *Font) Name() ole.Variant {
	retVal := this.PropGet(0x0000006e, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Font) SetName(rhs interface{})  {
	retVal := this.PropPut(0x0000006e, []interface{}{rhs})
	_= retVal
}

func (this *Font) OutlineFont() ole.Variant {
	retVal := this.PropGet(0x000000dd, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Font) SetOutlineFont(rhs interface{})  {
	retVal := this.PropPut(0x000000dd, []interface{}{rhs})
	_= retVal
}

func (this *Font) Shadow() ole.Variant {
	retVal := this.PropGet(0x00000067, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Font) SetShadow(rhs interface{})  {
	retVal := this.PropPut(0x00000067, []interface{}{rhs})
	_= retVal
}

func (this *Font) Size() ole.Variant {
	retVal := this.PropGet(0x00000068, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Font) SetSize(rhs interface{})  {
	retVal := this.PropPut(0x00000068, []interface{}{rhs})
	_= retVal
}

func (this *Font) Strikethrough() ole.Variant {
	retVal := this.PropGet(0x00000069, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Font) SetStrikethrough(rhs interface{})  {
	retVal := this.PropPut(0x00000069, []interface{}{rhs})
	_= retVal
}

func (this *Font) Subscript() ole.Variant {
	retVal := this.PropGet(0x000000b3, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Font) SetSubscript(rhs interface{})  {
	retVal := this.PropPut(0x000000b3, []interface{}{rhs})
	_= retVal
}

func (this *Font) Superscript() ole.Variant {
	retVal := this.PropGet(0x000000b2, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Font) SetSuperscript(rhs interface{})  {
	retVal := this.PropPut(0x000000b2, []interface{}{rhs})
	_= retVal
}

func (this *Font) Underline() ole.Variant {
	retVal := this.PropGet(0x0000006a, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Font) SetUnderline(rhs interface{})  {
	retVal := this.PropPut(0x0000006a, []interface{}{rhs})
	_= retVal
}

func (this *Font) ThemeColor() ole.Variant {
	retVal := this.PropGet(0x0000093d, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Font) SetThemeColor(rhs interface{})  {
	retVal := this.PropPut(0x0000093d, []interface{}{rhs})
	_= retVal
}

func (this *Font) TintAndShade() ole.Variant {
	retVal := this.PropGet(0x0000093e, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Font) SetTintAndShade(rhs interface{})  {
	retVal := this.PropPut(0x0000093e, []interface{}{rhs})
	_= retVal
}

func (this *Font) ThemeFont() int32 {
	retVal := this.PropGet(0x0000093f, nil)
	return retVal.LValVal()
}

func (this *Font) SetThemeFont(rhs int32)  {
	retVal := this.PropPut(0x0000093f, []interface{}{rhs})
	_= retVal
}

