package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024465-0000-0000-C000-000000000046
var IID_SpellingOptions = syscall.GUID{0x00024465, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type SpellingOptions struct {
	ole.OleClient
}

func NewSpellingOptions(pDisp *win32.IDispatch, addRef bool, scoped bool) *SpellingOptions {
	 if pDisp == nil {
		return nil;
	}
	p := &SpellingOptions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SpellingOptionsFromVar(v ole.Variant) *SpellingOptions {
	return NewSpellingOptions(v.IDispatch(), false, false)
}

func (this *SpellingOptions) IID() *syscall.GUID {
	return &IID_SpellingOptions
}

func (this *SpellingOptions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SpellingOptions) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *SpellingOptions) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *SpellingOptions) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *SpellingOptions) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *SpellingOptions) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *SpellingOptions) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *SpellingOptions) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *SpellingOptions) DictLang() int32 {
	retVal, _ := this.PropGet(0x000008ac, nil)
	return retVal.LValVal()
}

func (this *SpellingOptions) SetDictLang(rhs int32)  {
	_ = this.PropPut(0x000008ac, []interface{}{rhs})
}

func (this *SpellingOptions) UserDict() string {
	retVal, _ := this.PropGet(0x000008ad, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SpellingOptions) SetUserDict(rhs string)  {
	_ = this.PropPut(0x000008ad, []interface{}{rhs})
}

func (this *SpellingOptions) IgnoreCaps() bool {
	retVal, _ := this.PropGet(0x000008ae, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SpellingOptions) SetIgnoreCaps(rhs bool)  {
	_ = this.PropPut(0x000008ae, []interface{}{rhs})
}

func (this *SpellingOptions) SuggestMainOnly() bool {
	retVal, _ := this.PropGet(0x000008af, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SpellingOptions) SetSuggestMainOnly(rhs bool)  {
	_ = this.PropPut(0x000008af, []interface{}{rhs})
}

func (this *SpellingOptions) IgnoreMixedDigits() bool {
	retVal, _ := this.PropGet(0x000008b0, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SpellingOptions) SetIgnoreMixedDigits(rhs bool)  {
	_ = this.PropPut(0x000008b0, []interface{}{rhs})
}

func (this *SpellingOptions) IgnoreFileNames() bool {
	retVal, _ := this.PropGet(0x000008b1, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SpellingOptions) SetIgnoreFileNames(rhs bool)  {
	_ = this.PropPut(0x000008b1, []interface{}{rhs})
}

func (this *SpellingOptions) GermanPostReform() bool {
	retVal, _ := this.PropGet(0x000008b2, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SpellingOptions) SetGermanPostReform(rhs bool)  {
	_ = this.PropPut(0x000008b2, []interface{}{rhs})
}

func (this *SpellingOptions) KoreanCombineAux() bool {
	retVal, _ := this.PropGet(0x000008b3, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SpellingOptions) SetKoreanCombineAux(rhs bool)  {
	_ = this.PropPut(0x000008b3, []interface{}{rhs})
}

func (this *SpellingOptions) KoreanUseAutoChangeList() bool {
	retVal, _ := this.PropGet(0x000008b4, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SpellingOptions) SetKoreanUseAutoChangeList(rhs bool)  {
	_ = this.PropPut(0x000008b4, []interface{}{rhs})
}

func (this *SpellingOptions) KoreanProcessCompound() bool {
	retVal, _ := this.PropGet(0x000008b5, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SpellingOptions) SetKoreanProcessCompound(rhs bool)  {
	_ = this.PropPut(0x000008b5, []interface{}{rhs})
}

func (this *SpellingOptions) HebrewModes() int32 {
	retVal, _ := this.PropGet(0x000008b6, nil)
	return retVal.LValVal()
}

func (this *SpellingOptions) SetHebrewModes(rhs int32)  {
	_ = this.PropPut(0x000008b6, []interface{}{rhs})
}

func (this *SpellingOptions) ArabicModes() int32 {
	retVal, _ := this.PropGet(0x000008b7, nil)
	return retVal.LValVal()
}

func (this *SpellingOptions) SetArabicModes(rhs int32)  {
	_ = this.PropPut(0x000008b7, []interface{}{rhs})
}

func (this *SpellingOptions) ArabicStrictAlefHamza() bool {
	retVal, _ := this.PropGet(0x00000b74, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SpellingOptions) SetArabicStrictAlefHamza(rhs bool)  {
	_ = this.PropPut(0x00000b74, []interface{}{rhs})
}

func (this *SpellingOptions) ArabicStrictFinalYaa() bool {
	retVal, _ := this.PropGet(0x00000b75, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SpellingOptions) SetArabicStrictFinalYaa(rhs bool)  {
	_ = this.PropPut(0x00000b75, []interface{}{rhs})
}

func (this *SpellingOptions) ArabicStrictTaaMarboota() bool {
	retVal, _ := this.PropGet(0x00000b76, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SpellingOptions) SetArabicStrictTaaMarboota(rhs bool)  {
	_ = this.PropPut(0x00000b76, []interface{}{rhs})
}

func (this *SpellingOptions) RussianStrictE() bool {
	retVal, _ := this.PropGet(0x00000b77, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SpellingOptions) SetRussianStrictE(rhs bool)  {
	_ = this.PropPut(0x00000b77, []interface{}{rhs})
}

func (this *SpellingOptions) SpanishModes() int32 {
	retVal, _ := this.PropGet(0x00000b78, nil)
	return retVal.LValVal()
}

func (this *SpellingOptions) SetSpanishModes(rhs int32)  {
	_ = this.PropPut(0x00000b78, []interface{}{rhs})
}

func (this *SpellingOptions) PortugalReform() int32 {
	retVal, _ := this.PropGet(0x00000b79, nil)
	return retVal.LValVal()
}

func (this *SpellingOptions) SetPortugalReform(rhs int32)  {
	_ = this.PropPut(0x00000b79, []interface{}{rhs})
}

func (this *SpellingOptions) BrazilReform() int32 {
	retVal, _ := this.PropGet(0x00000b7a, nil)
	return retVal.LValVal()
}

func (this *SpellingOptions) SetBrazilReform(rhs int32)  {
	_ = this.PropPut(0x00000b7a, []interface{}{rhs})
}

