package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024470-0000-0000-C000-000000000046
var IID_ListObjects = syscall.GUID{0x00024470, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ListObjects struct {
	ole.OleClient
}

func NewListObjects(pDisp *win32.IDispatch, addRef bool, scoped bool) *ListObjects {
	 if pDisp == nil {
		return nil;
	}
	p := &ListObjects{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ListObjectsFromVar(v ole.Variant) *ListObjects {
	return NewListObjects(v.IDispatch(), false, false)
}

func (this *ListObjects) IID() *syscall.GUID {
	return &IID_ListObjects
}

func (this *ListObjects) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ListObjects) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *ListObjects) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ListObjects) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ListObjects) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *ListObjects) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *ListObjects) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *ListObjects) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *ListObjects) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ListObjects) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ListObjects) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var ListObjects_Add__OptArgs= []string{
	"SourceType", "Source", "LinkSource", "XlListObjectHasHeaders", "Destination", 
}

func (this *ListObjects) Add_(optArgs ...interface{}) *ListObject {
	optArgs = ole.ProcessOptArgs(ListObjects_Add__OptArgs, optArgs)
	retVal, _ := this.Call(0x00000825, nil, optArgs...)
	return NewListObject(retVal.IDispatch(), false, true)
}

func (this *ListObjects) Default_(index interface{}) *ListObject {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return NewListObject(retVal.IDispatch(), false, true)
}

func (this *ListObjects) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *ListObjects) ForEach(action func(item *ListObject) bool) {
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
		pItem := (*ListObject)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *ListObjects) Item(index interface{}) *ListObject {
	retVal, _ := this.PropGet(0x000000aa, []interface{}{index})
	return NewListObject(retVal.IDispatch(), false, true)
}

func (this *ListObjects) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

var ListObjects_Add_OptArgs= []string{
	"SourceType", "Source", "LinkSource", "XlListObjectHasHeaders", 
	"Destination", "TableStyleName", 
}

func (this *ListObjects) Add(optArgs ...interface{}) *ListObject {
	optArgs = ole.ProcessOptArgs(ListObjects_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b5, nil, optArgs...)
	return NewListObject(retVal.IDispatch(), false, true)
}

