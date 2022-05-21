package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002447C-0000-0000-C000-000000000046
var IID_XmlMaps = syscall.GUID{0x0002447C, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type XmlMaps struct {
	ole.OleClient
}

func NewXmlMaps(pDisp *win32.IDispatch, addRef bool, scoped bool) *XmlMaps {
	p := &XmlMaps{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func XmlMapsFromVar(v ole.Variant) *XmlMaps {
	return NewXmlMaps(v.PdispValVal(), false, false)
}

func (this *XmlMaps) IID() *syscall.GUID {
	return &IID_XmlMaps
}

func (this *XmlMaps) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *XmlMaps) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *XmlMaps) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *XmlMaps) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *XmlMaps) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *XmlMaps) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *XmlMaps) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *XmlMaps) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *XmlMaps) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *XmlMaps) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *XmlMaps) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var XmlMaps_Add_OptArgs= []string{
	"RootElementName", 
}

func (this *XmlMaps) Add(schema string, optArgs ...interface{}) *XmlMap {
	optArgs = ole.ProcessOptArgs(XmlMaps_Add_OptArgs, optArgs)
	retVal := this.Call(0x000000b5, []interface{}{schema}, optArgs...)
	return NewXmlMap(retVal.PdispValVal(), false, true)
}

func (this *XmlMaps) Default_(index interface{}) *XmlMap {
	retVal := this.PropGet(0x00000000, []interface{}{index})
	return NewXmlMap(retVal.PdispValVal(), false, true)
}

func (this *XmlMaps) Item(index interface{}) *XmlMap {
	retVal := this.PropGet(0x000000aa, []interface{}{index})
	return NewXmlMap(retVal.PdispValVal(), false, true)
}

func (this *XmlMaps) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *XmlMaps) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *XmlMaps) ForEach(action func(item *XmlMap) bool) {
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
		pItem := (*XmlMap)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

