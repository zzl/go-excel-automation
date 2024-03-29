package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
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
	if pDisp == nil {
		return nil
	}
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
	return NewFont(v.IDispatch(), false, false)
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

func (this *Font) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Font) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Font) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Font) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Font) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Font) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Font) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Font) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Font) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Font) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Font) Background() ole.Variant {
	retVal, _ := this.PropGet(0x000000b4, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Font) SetBackground(rhs interface{}) {
	_ = this.PropPut(0x000000b4, []interface{}{rhs})
}

func (this *Font) Bold() ole.Variant {
	retVal, _ := this.PropGet(0x00000060, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Font) SetBold(rhs interface{}) {
	_ = this.PropPut(0x00000060, []interface{}{rhs})
}

func (this *Font) Color() ole.Variant {
	retVal, _ := this.PropGet(0x00000063, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Font) SetColor(rhs interface{}) {
	_ = this.PropPut(0x00000063, []interface{}{rhs})
}

func (this *Font) ColorIndex() ole.Variant {
	retVal, _ := this.PropGet(0x00000061, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Font) SetColorIndex(rhs interface{}) {
	_ = this.PropPut(0x00000061, []interface{}{rhs})
}

func (this *Font) FontStyle() ole.Variant {
	retVal, _ := this.PropGet(0x000000b1, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Font) SetFontStyle(rhs interface{}) {
	_ = this.PropPut(0x000000b1, []interface{}{rhs})
}

func (this *Font) Italic() ole.Variant {
	retVal, _ := this.PropGet(0x00000065, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Font) SetItalic(rhs interface{}) {
	_ = this.PropPut(0x00000065, []interface{}{rhs})
}

func (this *Font) Name() ole.Variant {
	retVal, _ := this.PropGet(0x0000006e, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Font) SetName(rhs interface{}) {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *Font) OutlineFont() ole.Variant {
	retVal, _ := this.PropGet(0x000000dd, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Font) SetOutlineFont(rhs interface{}) {
	_ = this.PropPut(0x000000dd, []interface{}{rhs})
}

func (this *Font) Shadow() ole.Variant {
	retVal, _ := this.PropGet(0x00000067, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Font) SetShadow(rhs interface{}) {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *Font) Size() ole.Variant {
	retVal, _ := this.PropGet(0x00000068, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Font) SetSize(rhs interface{}) {
	_ = this.PropPut(0x00000068, []interface{}{rhs})
}

func (this *Font) Strikethrough() ole.Variant {
	retVal, _ := this.PropGet(0x00000069, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Font) SetStrikethrough(rhs interface{}) {
	_ = this.PropPut(0x00000069, []interface{}{rhs})
}

func (this *Font) Subscript() ole.Variant {
	retVal, _ := this.PropGet(0x000000b3, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Font) SetSubscript(rhs interface{}) {
	_ = this.PropPut(0x000000b3, []interface{}{rhs})
}

func (this *Font) Superscript() ole.Variant {
	retVal, _ := this.PropGet(0x000000b2, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Font) SetSuperscript(rhs interface{}) {
	_ = this.PropPut(0x000000b2, []interface{}{rhs})
}

func (this *Font) Underline() ole.Variant {
	retVal, _ := this.PropGet(0x0000006a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Font) SetUnderline(rhs interface{}) {
	_ = this.PropPut(0x0000006a, []interface{}{rhs})
}

func (this *Font) ThemeColor() ole.Variant {
	retVal, _ := this.PropGet(0x0000093d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Font) SetThemeColor(rhs interface{}) {
	_ = this.PropPut(0x0000093d, []interface{}{rhs})
}

func (this *Font) TintAndShade() ole.Variant {
	retVal, _ := this.PropGet(0x0000093e, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Font) SetTintAndShade(rhs interface{}) {
	_ = this.PropPut(0x0000093e, []interface{}{rhs})
}

func (this *Font) ThemeFont() int32 {
	retVal, _ := this.PropGet(0x0000093f, nil)
	return retVal.LValVal()
}

func (this *Font) SetThemeFont(rhs int32) {
	_ = this.PropPut(0x0000093f, []interface{}{rhs})
}
