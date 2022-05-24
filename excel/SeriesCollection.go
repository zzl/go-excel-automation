package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002086C-0000-0000-C000-000000000046
var IID_SeriesCollection = syscall.GUID{0x0002086C, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type SeriesCollection struct {
	ole.OleClient
}

func NewSeriesCollection(pDisp *win32.IDispatch, addRef bool, scoped bool) *SeriesCollection {
	 if pDisp == nil {
		return nil;
	}
	p := &SeriesCollection{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SeriesCollectionFromVar(v ole.Variant) *SeriesCollection {
	return NewSeriesCollection(v.IDispatch(), false, false)
}

func (this *SeriesCollection) IID() *syscall.GUID {
	return &IID_SeriesCollection
}

func (this *SeriesCollection) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SeriesCollection) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *SeriesCollection) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *SeriesCollection) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *SeriesCollection) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *SeriesCollection) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *SeriesCollection) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *SeriesCollection) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *SeriesCollection) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *SeriesCollection) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *SeriesCollection) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var SeriesCollection_Add_OptArgs= []string{
	"Rowcol", "SeriesLabels", "CategoryLabels", "Replace", 
}

func (this *SeriesCollection) Add(source interface{}, optArgs ...interface{}) *Series {
	optArgs = ole.ProcessOptArgs(SeriesCollection_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b5, []interface{}{source}, optArgs...)
	return NewSeries(retVal.IDispatch(), false, true)
}

func (this *SeriesCollection) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

var SeriesCollection_Extend_OptArgs= []string{
	"Rowcol", "CategoryLabels", 
}

func (this *SeriesCollection) Extend(source interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(SeriesCollection_Extend_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000e3, []interface{}{source}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *SeriesCollection) Item(index interface{}) *Series {
	retVal, _ := this.Call(0x000000aa, []interface{}{index})
	return NewSeries(retVal.IDispatch(), false, true)
}

func (this *SeriesCollection) NewEnum_() *com.UnknownClass {
	retVal, _ := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *SeriesCollection) ForEach(action func(item *Series) bool) {
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
		pItem := (*Series)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

var SeriesCollection_Paste_OptArgs= []string{
	"Rowcol", "SeriesLabels", "CategoryLabels", "Replace", "NewSeries", 
}

func (this *SeriesCollection) Paste(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(SeriesCollection_Paste_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d3, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *SeriesCollection) NewSeries() *Series {
	retVal, _ := this.Call(0x0000045d, nil)
	return NewSeries(retVal.IDispatch(), false, true)
}

func (this *SeriesCollection) Default_(index interface{}) *Series {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewSeries(retVal.IDispatch(), false, true)
}

