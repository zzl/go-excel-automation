package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 000208B1-0000-0000-C000-000000000046
var IID_Worksheets = syscall.GUID{0x000208B1, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Worksheets struct {
	ole.OleClient
}

func NewWorksheets(pDisp *win32.IDispatch, addRef bool, scoped bool) *Worksheets {
	if pDisp == nil {
		return nil
	}
	p := &Worksheets{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func WorksheetsFromVar(v ole.Variant) *Worksheets {
	return NewWorksheets(v.IDispatch(), false, false)
}

func (this *Worksheets) IID() *syscall.GUID {
	return &IID_Worksheets
}

func (this *Worksheets) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Worksheets) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Worksheets) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Worksheets) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Worksheets) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Worksheets) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Worksheets) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Worksheets) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Worksheets) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Worksheets) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Worksheets) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Worksheets_Add_OptArgs = []string{
	"Before", "After", "Count", "Type",
}

func (this *Worksheets) Add(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Worksheets_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b5, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Worksheets_Copy_OptArgs = []string{
	"Before", "After",
}

func (this *Worksheets) Copy(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Worksheets_Copy_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000227, nil, optArgs...)
	_ = retVal
}

func (this *Worksheets) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *Worksheets) Delete() {
	retVal, _ := this.Call(0x00000075, nil)
	_ = retVal
}

var Worksheets_FillAcrossSheets_OptArgs = []string{
	"Type",
}

func (this *Worksheets) FillAcrossSheets(range_ *Range, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Worksheets_FillAcrossSheets_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001d5, []interface{}{range_}, optArgs...)
	_ = retVal
}

func (this *Worksheets) Item(index interface{}) *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000000aa, []interface{}{index})
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Worksheets_Move_OptArgs = []string{
	"Before", "After",
}

func (this *Worksheets) Move(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Worksheets_Move_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000027d, nil, optArgs...)
	_ = retVal
}

func (this *Worksheets) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Worksheets) ForEach(action func(item *ole.DispatchClass) bool) {
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
		pItem := (*ole.DispatchClass)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

var Worksheets_PrintOut___OptArgs = []string{
	"From", "To", "Copies", "Preview",
	"ActivePrinter", "PrintToFile", "Collate",
}

func (this *Worksheets) PrintOut__(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Worksheets_PrintOut___OptArgs, optArgs)
	retVal, _ := this.Call(0x00000389, nil, optArgs...)
	_ = retVal
}

var Worksheets_PrintPreview_OptArgs = []string{
	"EnableChanges",
}

func (this *Worksheets) PrintPreview(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Worksheets_PrintPreview_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000119, nil, optArgs...)
	_ = retVal
}

var Worksheets_Select_OptArgs = []string{
	"Replace",
}

func (this *Worksheets) Select(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Worksheets_Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000eb, nil, optArgs...)
	_ = retVal
}

func (this *Worksheets) HPageBreaks() *HPageBreaks {
	retVal, _ := this.PropGet(0x0000058a, nil)
	return NewHPageBreaks(retVal.IDispatch(), false, true)
}

func (this *Worksheets) VPageBreaks() *VPageBreaks {
	retVal, _ := this.PropGet(0x0000058b, nil)
	return NewVPageBreaks(retVal.IDispatch(), false, true)
}

func (this *Worksheets) Visible() ole.Variant {
	retVal, _ := this.PropGet(0x0000022e, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Worksheets) SetVisible(rhs interface{}) {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

func (this *Worksheets) Default_(index interface{}) *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Worksheets_PrintOut__OptArgs = []string{
	"From", "To", "Copies", "Preview",
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName",
}

func (this *Worksheets) PrintOut_(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Worksheets_PrintOut__OptArgs, optArgs)
	retVal, _ := this.Call(0x000006ec, nil, optArgs...)
	_ = retVal
}

var Worksheets_PrintOut_OptArgs = []string{
	"From", "To", "Copies", "Preview",
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName", "IgnorePrintAreas",
}

func (this *Worksheets) PrintOut(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Worksheets_PrintOut_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000939, nil, optArgs...)
	_ = retVal
}
