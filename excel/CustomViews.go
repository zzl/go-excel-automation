package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024422-0000-0000-C000-000000000046
var IID_CustomViews = syscall.GUID{0x00024422, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type CustomViews struct {
	ole.OleClient
}

func NewCustomViews(pDisp *win32.IDispatch, addRef bool, scoped bool) *CustomViews {
	 if pDisp == nil {
		return nil;
	}
	p := &CustomViews{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CustomViewsFromVar(v ole.Variant) *CustomViews {
	return NewCustomViews(v.IDispatch(), false, false)
}

func (this *CustomViews) IID() *syscall.GUID {
	return &IID_CustomViews
}

func (this *CustomViews) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CustomViews) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *CustomViews) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *CustomViews) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *CustomViews) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *CustomViews) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *CustomViews) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *CustomViews) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *CustomViews) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *CustomViews) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *CustomViews) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *CustomViews) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *CustomViews) Item(viewName interface{}) *CustomView {
	retVal, _ := this.Call(0x000000aa, []interface{}{viewName})
	return NewCustomView(retVal.IDispatch(), false, true)
}

var CustomViews_Add_OptArgs= []string{
	"PrintSettings", "RowColSettings", 
}

func (this *CustomViews) Add(viewName string, optArgs ...interface{}) *CustomView {
	optArgs = ole.ProcessOptArgs(CustomViews_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b5, []interface{}{viewName}, optArgs...)
	return NewCustomView(retVal.IDispatch(), false, true)
}

func (this *CustomViews) Default_(viewName interface{}) *CustomView {
	retVal, _ := this.PropGet(0x00000000, []interface{}{viewName})
	return NewCustomView(retVal.IDispatch(), false, true)
}

func (this *CustomViews) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *CustomViews) ForEach(action func(item *CustomView) bool) {
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
		pItem := (*CustomView)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

