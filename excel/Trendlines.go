package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000208BD-0000-0000-C000-000000000046
var IID_Trendlines = syscall.GUID{0x000208BD, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Trendlines struct {
	ole.OleClient
}

func NewTrendlines(pDisp *win32.IDispatch, addRef bool, scoped bool) *Trendlines {
	p := &Trendlines{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TrendlinesFromVar(v ole.Variant) *Trendlines {
	return NewTrendlines(v.PdispValVal(), false, false)
}

func (this *Trendlines) IID() *syscall.GUID {
	return &IID_Trendlines
}

func (this *Trendlines) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Trendlines) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Trendlines) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Trendlines) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Trendlines) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Trendlines) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Trendlines) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Trendlines) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Trendlines) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Trendlines) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Trendlines) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Trendlines_Add_OptArgs= []string{
	"Order", "Period", "Forward", "Backward", 
	"Intercept", "DisplayEquation", "DisplayRSquared", "Name", 
}

func (this *Trendlines) Add(type_ int32, optArgs ...interface{}) *Trendline {
	optArgs = ole.ProcessOptArgs(Trendlines_Add_OptArgs, optArgs)
	retVal := this.Call(0x000000b5, []interface{}{type_}, optArgs...)
	return NewTrendline(retVal.PdispValVal(), false, true)
}

func (this *Trendlines) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

var Trendlines_Item_OptArgs= []string{
	"Index", 
}

func (this *Trendlines) Item(optArgs ...interface{}) *Trendline {
	optArgs = ole.ProcessOptArgs(Trendlines_Item_OptArgs, optArgs)
	retVal := this.Call(0x000000aa, nil, optArgs...)
	return NewTrendline(retVal.PdispValVal(), false, true)
}

func (this *Trendlines) NewEnum_() *com.UnknownClass {
	retVal := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Trendlines) ForEach(action func(item *Trendline) bool) {
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
		pItem := (*Trendline)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

var Trendlines_Default__OptArgs= []string{
	"Index", 
}

func (this *Trendlines) Default_(optArgs ...interface{}) *Trendline {
	optArgs = ole.ProcessOptArgs(Trendlines_Default__OptArgs, optArgs)
	retVal := this.Call(0x00000000, nil, optArgs...)
	return NewTrendline(retVal.PdispValVal(), false, true)
}
