package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000208D7-0000-0000-C000-000000000046
var IID_Sheets = syscall.GUID{0x000208D7, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Sheets struct {
	ole.OleClient
}

func NewSheets(pDisp *win32.IDispatch, addRef bool, scoped bool) *Sheets {
	p := &Sheets{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SheetsFromVar(v ole.Variant) *Sheets {
	return NewSheets(v.PdispValVal(), false, false)
}

func (this *Sheets) IID() *syscall.GUID {
	return &IID_Sheets
}

func (this *Sheets) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Sheets) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Sheets) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Sheets) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Sheets_Add_OptArgs= []string{
	"Before", "After", "Count", "Type", 
}

func (this *Sheets) Add(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Sheets_Add_OptArgs, optArgs)
	retVal := this.Call(0x000000b5, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Sheets_Copy_OptArgs= []string{
	"Before", "After", 
}

func (this *Sheets) Copy(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Sheets_Copy_OptArgs, optArgs)
	retVal := this.Call(0x00000227, nil, optArgs...)
	_= retVal
}

func (this *Sheets) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *Sheets) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

func (this *Sheets) FillAcrossSheets(range_ *Range, type_ int32)  {
	retVal := this.Call(0x000001d5, []interface{}{range_, type_})
	_= retVal
}

func (this *Sheets) Item(index interface{}) *ole.DispatchClass {
	retVal := this.PropGet(0x000000aa, []interface{}{index})
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Sheets_Move_OptArgs= []string{
	"Before", "After", 
}

func (this *Sheets) Move(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Sheets_Move_OptArgs, optArgs)
	retVal := this.Call(0x0000027d, nil, optArgs...)
	_= retVal
}

func (this *Sheets) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Sheets) ForEach(action func(item *ole.DispatchClass) bool) {
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

var Sheets_PrintOut___OptArgs= []string{
	"From", "To", "Copies", "Preview", 
	"ActivePrinter", "PrintToFile", "Collate", 
}

func (this *Sheets) PrintOut__(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Sheets_PrintOut___OptArgs, optArgs)
	retVal := this.Call(0x00000389, nil, optArgs...)
	_= retVal
}

var Sheets_PrintPreview_OptArgs= []string{
	"EnableChanges", 
}

func (this *Sheets) PrintPreview(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Sheets_PrintPreview_OptArgs, optArgs)
	retVal := this.Call(0x00000119, nil, optArgs...)
	_= retVal
}

var Sheets_Select_OptArgs= []string{
	"Replace", 
}

func (this *Sheets) Select(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Sheets_Select_OptArgs, optArgs)
	retVal := this.Call(0x000000eb, nil, optArgs...)
	_= retVal
}

func (this *Sheets) HPageBreaks() *HPageBreaks {
	retVal := this.PropGet(0x0000058a, nil)
	return NewHPageBreaks(retVal.PdispValVal(), false, true)
}

func (this *Sheets) VPageBreaks() *VPageBreaks {
	retVal := this.PropGet(0x0000058b, nil)
	return NewVPageBreaks(retVal.PdispValVal(), false, true)
}

func (this *Sheets) Visible() ole.Variant {
	retVal := this.PropGet(0x0000022e, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Sheets) SetVisible(rhs interface{})  {
	retVal := this.PropPut(0x0000022e, []interface{}{rhs})
	_= retVal
}

func (this *Sheets) Default_(index interface{}) *ole.DispatchClass {
	retVal := this.PropGet(0x00000000, []interface{}{index})
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var Sheets_PrintOut__OptArgs= []string{
	"From", "To", "Copies", "Preview", 
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName", 
}

func (this *Sheets) PrintOut_(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Sheets_PrintOut__OptArgs, optArgs)
	retVal := this.Call(0x000006ec, nil, optArgs...)
	_= retVal
}

var Sheets_PrintOut_OptArgs= []string{
	"From", "To", "Copies", "Preview", 
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName", "IgnorePrintAreas", 
}

func (this *Sheets) PrintOut(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Sheets_PrintOut_OptArgs, optArgs)
	retVal := this.Call(0x00000939, nil, optArgs...)
	_= retVal
}

