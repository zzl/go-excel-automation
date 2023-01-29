package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 000208B0-0000-0000-C000-000000000046
var IID_DialogSheets = syscall.GUID{0x000208B0, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type DialogSheets struct {
	ole.OleClient
}

func NewDialogSheets(pDisp *win32.IDispatch, addRef bool, scoped bool) *DialogSheets {
	if pDisp == nil {
		return nil
	}
	p := &DialogSheets{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DialogSheetsFromVar(v ole.Variant) *DialogSheets {
	return NewDialogSheets(v.IDispatch(), false, false)
}

func (this *DialogSheets) IID() *syscall.GUID {
	return &IID_DialogSheets
}

func (this *DialogSheets) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DialogSheets) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *DialogSheets) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *DialogSheets) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *DialogSheets) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *DialogSheets) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *DialogSheets) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *DialogSheets) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *DialogSheets) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *DialogSheets) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *DialogSheets) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var DialogSheets_Add_OptArgs = []string{
	"Before", "After", "Count",
}

func (this *DialogSheets) Add(optArgs ...interface{}) *DialogSheet {
	optArgs = ole.ProcessOptArgs(DialogSheets_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b5, nil, optArgs...)
	return NewDialogSheet(retVal.IDispatch(), false, true)
}

var DialogSheets_Copy_OptArgs = []string{
	"Before", "After",
}

func (this *DialogSheets) Copy(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(DialogSheets_Copy_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000227, nil, optArgs...)
	_ = retVal
}

func (this *DialogSheets) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *DialogSheets) Delete() {
	retVal, _ := this.Call(0x00000075, nil)
	_ = retVal
}

func (this *DialogSheets) Dummy7_() {
	retVal, _ := this.Call(0x00010007, nil)
	_ = retVal
}

func (this *DialogSheets) Item(index interface{}) *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000000aa, []interface{}{index})
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var DialogSheets_Move_OptArgs = []string{
	"Before", "After",
}

func (this *DialogSheets) Move(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(DialogSheets_Move_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000027d, nil, optArgs...)
	_ = retVal
}

func (this *DialogSheets) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *DialogSheets) ForEach(action func(item *ole.DispatchClass) bool) {
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

var DialogSheets_PrintOut___OptArgs = []string{
	"From", "To", "Copies", "Preview",
	"ActivePrinter", "PrintToFile", "Collate",
}

func (this *DialogSheets) PrintOut__(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(DialogSheets_PrintOut___OptArgs, optArgs)
	retVal, _ := this.Call(0x00000389, nil, optArgs...)
	_ = retVal
}

var DialogSheets_PrintPreview_OptArgs = []string{
	"EnableChanges",
}

func (this *DialogSheets) PrintPreview(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(DialogSheets_PrintPreview_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000119, nil, optArgs...)
	_ = retVal
}

var DialogSheets_Select_OptArgs = []string{
	"Replace",
}

func (this *DialogSheets) Select(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(DialogSheets_Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000eb, nil, optArgs...)
	_ = retVal
}

func (this *DialogSheets) HPageBreaks() *HPageBreaks {
	retVal, _ := this.PropGet(0x0000058a, nil)
	return NewHPageBreaks(retVal.IDispatch(), false, true)
}

func (this *DialogSheets) VPageBreaks() *VPageBreaks {
	retVal, _ := this.PropGet(0x0000058b, nil)
	return NewVPageBreaks(retVal.IDispatch(), false, true)
}

func (this *DialogSheets) Visible() ole.Variant {
	retVal, _ := this.PropGet(0x0000022e, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DialogSheets) SetVisible(rhs interface{}) {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

func (this *DialogSheets) Default_(index interface{}) *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var DialogSheets_PrintOut__OptArgs = []string{
	"From", "To", "Copies", "Preview",
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName",
}

func (this *DialogSheets) PrintOut_(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(DialogSheets_PrintOut__OptArgs, optArgs)
	retVal, _ := this.Call(0x000006ec, nil, optArgs...)
	_ = retVal
}

var DialogSheets_PrintOut_OptArgs = []string{
	"From", "To", "Copies", "Preview",
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName",
}

func (this *DialogSheets) PrintOut(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(DialogSheets_PrintOut_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000939, nil, optArgs...)
	_ = retVal
}
