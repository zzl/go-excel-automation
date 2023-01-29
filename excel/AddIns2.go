package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 000244B5-0000-0000-C000-000000000046
var IID_AddIns2 = syscall.GUID{0x000244B5, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type AddIns2 struct {
	ole.OleClient
}

func NewAddIns2(pDisp *win32.IDispatch, addRef bool, scoped bool) *AddIns2 {
	if pDisp == nil {
		return nil
	}
	p := &AddIns2{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func AddIns2FromVar(v ole.Variant) *AddIns2 {
	return NewAddIns2(v.IDispatch(), false, false)
}

func (this *AddIns2) IID() *syscall.GUID {
	return &IID_AddIns2
}

func (this *AddIns2) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *AddIns2) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *AddIns2) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *AddIns2) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *AddIns2) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *AddIns2) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *AddIns2) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *AddIns2) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *AddIns2) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *AddIns2) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *AddIns2) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var AddIns2_Add_OptArgs = []string{
	"CopyFile",
}

func (this *AddIns2) Add(filename string, optArgs ...interface{}) *AddIn {
	optArgs = ole.ProcessOptArgs(AddIns2_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b5, []interface{}{filename}, optArgs...)
	return NewAddIn(retVal.IDispatch(), false, true)
}

func (this *AddIns2) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *AddIns2) Item(index interface{}) *AddIn {
	retVal, _ := this.PropGet(0x000000aa, []interface{}{index})
	return NewAddIn(retVal.IDispatch(), false, true)
}

func (this *AddIns2) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *AddIns2) ForEach(action func(item *AddIn) bool) {
	pEnum := this.NewEnum_()
	var pEnumVar *win32.IEnumVARIANT
	pEnum.QueryInterface(&win32.IID_IEnumVARIANT, unsafe.Pointer(&pEnumVar))
	defer pEnumVar.Release()
	for {
		var c uint32
		var v ole.Variant
		pEnumVar.Next(1, (*win32.VARIANT)(&v), &c)
		if c == 0 {
			break
		}
		pItem := (*AddIn)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *AddIns2) Default_(index interface{}) *AddIn {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return NewAddIn(retVal.IDispatch(), false, true)
}
