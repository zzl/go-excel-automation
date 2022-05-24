package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024430-0000-0000-C000-000000000046
var IID_Hyperlinks = syscall.GUID{0x00024430, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Hyperlinks struct {
	ole.OleClient
}

func NewHyperlinks(pDisp *win32.IDispatch, addRef bool, scoped bool) *Hyperlinks {
	 if pDisp == nil {
		return nil;
	}
	p := &Hyperlinks{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func HyperlinksFromVar(v ole.Variant) *Hyperlinks {
	return NewHyperlinks(v.IDispatch(), false, false)
}

func (this *Hyperlinks) IID() *syscall.GUID {
	return &IID_Hyperlinks
}

func (this *Hyperlinks) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Hyperlinks) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Hyperlinks) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Hyperlinks) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Hyperlinks) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Hyperlinks) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Hyperlinks) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Hyperlinks) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Hyperlinks) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Hyperlinks) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Hyperlinks) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Hyperlinks_Add_OptArgs= []string{
	"SubAddress", "ScreenTip", "TextToDisplay", 
}

func (this *Hyperlinks) Add(anchor *win32.IUnknown, address string, optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Hyperlinks_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b5, []interface{}{anchor, address}, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Hyperlinks) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *Hyperlinks) Item(index interface{}) *Hyperlink {
	retVal, _ := this.PropGet(0x000000aa, []interface{}{index})
	return NewHyperlink(retVal.IDispatch(), false, true)
}

func (this *Hyperlinks) Default_(index interface{}) *Hyperlink {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return NewHyperlink(retVal.IDispatch(), false, true)
}

func (this *Hyperlinks) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Hyperlinks) ForEach(action func(item *Hyperlink) bool) {
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
		pItem := (*Hyperlink)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Hyperlinks) Delete()  {
	retVal, _ := this.Call(0x00000075, nil)
	_= retVal
}

