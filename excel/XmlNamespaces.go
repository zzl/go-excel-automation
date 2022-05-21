package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024477-0000-0000-C000-000000000046
var IID_XmlNamespaces = syscall.GUID{0x00024477, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type XmlNamespaces struct {
	ole.OleClient
}

func NewXmlNamespaces(pDisp *win32.IDispatch, addRef bool, scoped bool) *XmlNamespaces {
	p := &XmlNamespaces{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func XmlNamespacesFromVar(v ole.Variant) *XmlNamespaces {
	return NewXmlNamespaces(v.PdispValVal(), false, false)
}

func (this *XmlNamespaces) IID() *syscall.GUID {
	return &IID_XmlNamespaces
}

func (this *XmlNamespaces) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *XmlNamespaces) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *XmlNamespaces) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *XmlNamespaces) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *XmlNamespaces) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *XmlNamespaces) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *XmlNamespaces) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *XmlNamespaces) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *XmlNamespaces) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *XmlNamespaces) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *XmlNamespaces) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *XmlNamespaces) Default_(index interface{}) *XmlNamespace {
	retVal := this.PropGet(0x00000000, []interface{}{index})
	return NewXmlNamespace(retVal.PdispValVal(), false, true)
}

func (this *XmlNamespaces) Item(index interface{}) *XmlNamespace {
	retVal := this.PropGet(0x000000aa, []interface{}{index})
	return NewXmlNamespace(retVal.PdispValVal(), false, true)
}

func (this *XmlNamespaces) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *XmlNamespaces) Value() string {
	retVal := this.PropGet(0x00000006, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XmlNamespaces) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *XmlNamespaces) ForEach(action func(item *XmlNamespace) bool) {
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
		pItem := (*XmlNamespace)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

var XmlNamespaces_InstallManifest_OptArgs= []string{
	"InstallForAllUsers", 
}

func (this *XmlNamespaces) InstallManifest(path string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(XmlNamespaces_InstallManifest_OptArgs, optArgs)
	retVal := this.Call(0x00000917, []interface{}{path}, optArgs...)
	_= retVal
}
