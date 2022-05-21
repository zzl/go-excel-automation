package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024447-0000-0000-C000-000000000046
var IID_Phonetics = syscall.GUID{0x00024447, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Phonetics struct {
	ole.OleClient
}

func NewPhonetics(pDisp *win32.IDispatch, addRef bool, scoped bool) *Phonetics {
	p := &Phonetics{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PhoneticsFromVar(v ole.Variant) *Phonetics {
	return NewPhonetics(v.PdispValVal(), false, false)
}

func (this *Phonetics) IID() *syscall.GUID {
	return &IID_Phonetics
}

func (this *Phonetics) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Phonetics) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Phonetics) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Phonetics) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Phonetics) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Phonetics) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Phonetics) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Phonetics) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Phonetics) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Phonetics) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Phonetics) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Phonetics) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *Phonetics) Start() int32 {
	retVal := this.PropGet(0x00000260, nil)
	return retVal.LValVal()
}

func (this *Phonetics) Length() int32 {
	retVal := this.PropGet(0x00000261, nil)
	return retVal.LValVal()
}

func (this *Phonetics) Visible() bool {
	retVal := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Phonetics) SetVisible(rhs bool)  {
	retVal := this.PropPut(0x0000022e, []interface{}{rhs})
	_= retVal
}

func (this *Phonetics) CharacterType() int32 {
	retVal := this.PropGet(0x0000068a, nil)
	return retVal.LValVal()
}

func (this *Phonetics) SetCharacterType(rhs int32)  {
	retVal := this.PropPut(0x0000068a, []interface{}{rhs})
	_= retVal
}

func (this *Phonetics) Alignment() int32 {
	retVal := this.PropGet(0x000001c5, nil)
	return retVal.LValVal()
}

func (this *Phonetics) SetAlignment(rhs int32)  {
	retVal := this.PropPut(0x000001c5, []interface{}{rhs})
	_= retVal
}

func (this *Phonetics) Font() *Font {
	retVal := this.PropGet(0x00000092, nil)
	return NewFont(retVal.PdispValVal(), false, true)
}

func (this *Phonetics) Item(index int32) *ole.DispatchClass {
	retVal := this.PropGet(0x000000aa, []interface{}{index})
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Phonetics) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

func (this *Phonetics) Add(start int32, length int32, text string)  {
	retVal := this.Call(0x000000b5, []interface{}{start, length, text})
	_= retVal
}

func (this *Phonetics) Text() string {
	retVal := this.PropGet(0x0000008a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Phonetics) SetText(rhs string)  {
	retVal := this.PropPut(0x0000008a, []interface{}{rhs})
	_= retVal
}

func (this *Phonetics) Default_(index int32) *ole.DispatchClass {
	retVal := this.PropGet(0x00000000, []interface{}{index})
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Phonetics) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Phonetics) ForEach(action func(item *ole.DispatchClass) bool) {
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
		pItem := (*ole.DispatchClass)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

