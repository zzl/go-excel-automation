package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 0002442D-0000-0000-C000-000000000046
var IID_ODBCErrors = syscall.GUID{0x0002442D, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ODBCErrors struct {
	ole.OleClient
}

func NewODBCErrors(pDisp *win32.IDispatch, addRef bool, scoped bool) *ODBCErrors {
	if pDisp == nil {
		return nil
	}
	p := &ODBCErrors{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ODBCErrorsFromVar(v ole.Variant) *ODBCErrors {
	return NewODBCErrors(v.IDispatch(), false, false)
}

func (this *ODBCErrors) IID() *syscall.GUID {
	return &IID_ODBCErrors
}

func (this *ODBCErrors) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ODBCErrors) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *ODBCErrors) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ODBCErrors) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ODBCErrors) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *ODBCErrors) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *ODBCErrors) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *ODBCErrors) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *ODBCErrors) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ODBCErrors) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ODBCErrors) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ODBCErrors) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *ODBCErrors) Item(index int32) *ODBCError {
	retVal, _ := this.Call(0x000000aa, []interface{}{index})
	return NewODBCError(retVal.IDispatch(), false, true)
}

func (this *ODBCErrors) Default_(index int32) *ODBCError {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return NewODBCError(retVal.IDispatch(), false, true)
}

func (this *ODBCErrors) NewEnum_() *com.UnknownClass {
	retVal, _ := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *ODBCErrors) ForEach(action func(item *ODBCError) bool) {
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
		pItem := (*ODBCError)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}
