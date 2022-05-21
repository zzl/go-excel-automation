package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002086D-0000-0000-C000-000000000046
var IID_Charts = syscall.GUID{0x0002086D, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Charts struct {
	ole.OleClient
}

func NewCharts(pDisp *win32.IDispatch, addRef bool, scoped bool) *Charts {
	p := &Charts{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ChartsFromVar(v ole.Variant) *Charts {
	return NewCharts(v.PdispValVal(), false, false)
}

func (this *Charts) IID() *syscall.GUID {
	return &IID_Charts
}

func (this *Charts) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Charts) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Charts) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Charts) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Charts) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Charts) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Charts) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Charts) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Charts) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Charts) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Charts) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Charts_Add_OptArgs= []string{
	"Before", "After", "Count", 
}

func (this *Charts) Add(optArgs ...interface{}) *Chart {
	optArgs = ole.ProcessOptArgs(Charts_Add_OptArgs, optArgs)
	retVal := this.Call(0x000000b5, nil, optArgs...)
	return NewChart(retVal.PdispValVal(), false, true)
}

var Charts_Copy_OptArgs= []string{
	"Before", "After", 
}

func (this *Charts) Copy(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Charts_Copy_OptArgs, optArgs)
	retVal := this.Call(0x00000227, nil, optArgs...)
	_= retVal
}

func (this *Charts) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *Charts) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

func (this *Charts) Dummy7_()  {
	retVal := this.Call(0x00010007, nil)
	_= retVal
}

func (this *Charts) Item(index interface{}) *ole.DispatchClass {
	retVal := this.PropGet(0x000000aa, []interface{}{index})
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Charts_Move_OptArgs= []string{
	"Before", "After", 
}

func (this *Charts) Move(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Charts_Move_OptArgs, optArgs)
	retVal := this.Call(0x0000027d, nil, optArgs...)
	_= retVal
}

func (this *Charts) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Charts) ForEach(action func(item *ole.DispatchClass) bool) {
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
		pItem := (*ole.DispatchClass)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

var Charts_PrintOut___OptArgs= []string{
	"From", "To", "Copies", "Preview", 
	"ActivePrinter", "PrintToFile", "Collate", 
}

func (this *Charts) PrintOut__(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Charts_PrintOut___OptArgs, optArgs)
	retVal := this.Call(0x00000389, nil, optArgs...)
	_= retVal
}

var Charts_PrintPreview_OptArgs= []string{
	"EnableChanges", 
}

func (this *Charts) PrintPreview(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Charts_PrintPreview_OptArgs, optArgs)
	retVal := this.Call(0x00000119, nil, optArgs...)
	_= retVal
}

var Charts_Select_OptArgs= []string{
	"Replace", 
}

func (this *Charts) Select(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Charts_Select_OptArgs, optArgs)
	retVal := this.Call(0x000000eb, nil, optArgs...)
	_= retVal
}

func (this *Charts) HPageBreaks() *HPageBreaks {
	retVal := this.PropGet(0x0000058a, nil)
	return NewHPageBreaks(retVal.PdispValVal(), false, true)
}

func (this *Charts) VPageBreaks() *VPageBreaks {
	retVal := this.PropGet(0x0000058b, nil)
	return NewVPageBreaks(retVal.PdispValVal(), false, true)
}

func (this *Charts) Visible() ole.Variant {
	retVal := this.PropGet(0x0000022e, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Charts) SetVisible(rhs interface{})  {
	retVal := this.PropPut(0x0000022e, []interface{}{rhs})
	_= retVal
}

func (this *Charts) Default_(index interface{}) *ole.DispatchClass {
	retVal := this.PropGet(0x00000000, []interface{}{index})
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Charts_PrintOut__OptArgs= []string{
	"From", "To", "Copies", "Preview", 
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName", 
}

func (this *Charts) PrintOut_(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Charts_PrintOut__OptArgs, optArgs)
	retVal := this.Call(0x000006ec, nil, optArgs...)
	_= retVal
}

var Charts_PrintOut_OptArgs= []string{
	"From", "To", "Copies", "Preview", 
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName", 
}

func (this *Charts) PrintOut(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Charts_PrintOut_OptArgs, optArgs)
	retVal := this.Call(0x00000939, nil, optArgs...)
	_= retVal
}

