package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00024443-0000-0000-C000-000000000046
var IID_PublishObjects = syscall.GUID{0x00024443, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type PublishObjects struct {
	ole.OleClient
}

func NewPublishObjects(pDisp *win32.IDispatch, addRef bool, scoped bool) *PublishObjects {
	if pDisp == nil {
		return nil
	}
	p := &PublishObjects{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PublishObjectsFromVar(v ole.Variant) *PublishObjects {
	return NewPublishObjects(v.IDispatch(), false, false)
}

func (this *PublishObjects) IID() *syscall.GUID {
	return &IID_PublishObjects
}

func (this *PublishObjects) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *PublishObjects) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *PublishObjects) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *PublishObjects) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *PublishObjects) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *PublishObjects) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *PublishObjects) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *PublishObjects) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *PublishObjects) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *PublishObjects) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *PublishObjects) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var PublishObjects_Add_OptArgs = []string{
	"Sheet", "Source", "HtmlType", "DivID", "Title",
}

func (this *PublishObjects) Add(sourceType int32, filename string, optArgs ...interface{}) *PublishObject {
	optArgs = ole.ProcessOptArgs(PublishObjects_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b5, []interface{}{sourceType, filename}, optArgs...)
	return NewPublishObject(retVal.IDispatch(), false, true)
}

func (this *PublishObjects) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *PublishObjects) Item(index interface{}) *PublishObject {
	retVal, _ := this.PropGet(0x000000aa, []interface{}{index})
	return NewPublishObject(retVal.IDispatch(), false, true)
}

func (this *PublishObjects) Default_(index interface{}) *PublishObject {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return NewPublishObject(retVal.IDispatch(), false, true)
}

func (this *PublishObjects) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *PublishObjects) ForEach(action func(item *PublishObject) bool) {
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
		pItem := (*PublishObject)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *PublishObjects) Delete() {
	retVal, _ := this.Call(0x00000075, nil)
	_ = retVal
}

func (this *PublishObjects) Publish() {
	retVal, _ := this.Call(0x00000767, nil)
	_ = retVal
}
