package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000208DB-0000-0000-C000-000000000046
var IID_Workbooks = syscall.GUID{0x000208DB, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Workbooks struct {
	ole.OleClient
}

func NewWorkbooks(pDisp *win32.IDispatch, addRef bool, scoped bool) *Workbooks {
	 if pDisp == nil {
		return nil;
	}
	p := &Workbooks{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func WorkbooksFromVar(v ole.Variant) *Workbooks {
	return NewWorkbooks(v.IDispatch(), false, false)
}

func (this *Workbooks) IID() *syscall.GUID {
	return &IID_Workbooks
}

func (this *Workbooks) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Workbooks) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Workbooks) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Workbooks) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Workbooks_Add_OptArgs= []string{
	"Template", 
}

func (this *Workbooks) Add(optArgs ...interface{}) *Workbook {
	optArgs = ole.ProcessOptArgs(Workbooks_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b5, nil, optArgs...)
	return NewWorkbook(retVal.IDispatch(), false, true)
}

func (this *Workbooks) Close()  {
	retVal, _ := this.Call(0x00000115, nil)
	_= retVal
}

func (this *Workbooks) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *Workbooks) Item(index interface{}) *Workbook {
	retVal, _ := this.PropGet(0x000000aa, []interface{}{index})
	return NewWorkbook(retVal.IDispatch(), false, true)
}

func (this *Workbooks) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Workbooks) ForEach(action func(item *Workbook) bool) {
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
		pItem := (*Workbook)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

var Workbooks_Open__OptArgs= []string{
	"UpdateLinks", "ReadOnly", "Format", "Password", 
	"WriteResPassword", "IgnoreReadOnlyRecommended", "Origin", "Delimiter", 
	"Editable", "Notify", "Converter", "AddToMru", 
}

func (this *Workbooks) Open_(filename string, optArgs ...interface{}) *Workbook {
	optArgs = ole.ProcessOptArgs(Workbooks_Open__OptArgs, optArgs)
	retVal, _ := this.Call(0x000002aa, []interface{}{filename}, optArgs...)
	return NewWorkbook(retVal.IDispatch(), false, true)
}

var Workbooks_OpenText___OptArgs= []string{
	"Origin", "StartRow", "DataType", "TextQualifier", 
	"ConsecutiveDelimiter", "Tab", "Semicolon", "Comma", 
	"Space", "Other", "OtherChar", "FieldInfo", "TextVisualLayout", 
}

func (this *Workbooks) OpenText__(filename string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Workbooks_OpenText___OptArgs, optArgs)
	retVal, _ := this.Call(0x000002ab, []interface{}{filename}, optArgs...)
	_= retVal
}

func (this *Workbooks) Default_(index interface{}) *Workbook {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return NewWorkbook(retVal.IDispatch(), false, true)
}

var Workbooks_OpenText__OptArgs= []string{
	"Origin", "StartRow", "DataType", "TextQualifier", 
	"ConsecutiveDelimiter", "Tab", "Semicolon", "Comma", 
	"Space", "Other", "OtherChar", "FieldInfo", 
	"TextVisualLayout", "DecimalSeparator", "ThousandsSeparator", 
}

func (this *Workbooks) OpenText_(filename string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Workbooks_OpenText__OptArgs, optArgs)
	retVal, _ := this.Call(0x000006ed, []interface{}{filename}, optArgs...)
	_= retVal
}

var Workbooks_Open_OptArgs= []string{
	"UpdateLinks", "ReadOnly", "Format", "Password", 
	"WriteResPassword", "IgnoreReadOnlyRecommended", "Origin", "Delimiter", 
	"Editable", "Notify", "Converter", "AddToMru", 
	"Local", "CorruptLoad", 
}

func (this *Workbooks) Open(filename string, optArgs ...interface{}) *Workbook {
	optArgs = ole.ProcessOptArgs(Workbooks_Open_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000783, []interface{}{filename}, optArgs...)
	return NewWorkbook(retVal.IDispatch(), false, true)
}

var Workbooks_OpenText_OptArgs= []string{
	"Origin", "StartRow", "DataType", "TextQualifier", 
	"ConsecutiveDelimiter", "Tab", "Semicolon", "Comma", 
	"Space", "Other", "OtherChar", "FieldInfo", 
	"TextVisualLayout", "DecimalSeparator", "ThousandsSeparator", "TrailingMinusNumbers", "Local", 
}

func (this *Workbooks) OpenText(filename string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Workbooks_OpenText_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000784, []interface{}{filename}, optArgs...)
	_= retVal
}

var Workbooks_OpenDatabase_OptArgs= []string{
	"CommandText", "CommandType", "BackgroundQuery", "ImportDataAs", 
}

func (this *Workbooks) OpenDatabase(filename string, optArgs ...interface{}) *Workbook {
	optArgs = ole.ProcessOptArgs(Workbooks_OpenDatabase_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000813, []interface{}{filename}, optArgs...)
	return NewWorkbook(retVal.IDispatch(), false, true)
}

func (this *Workbooks) CheckOut(filename string)  {
	retVal, _ := this.Call(0x00000815, []interface{}{filename})
	_= retVal
}

func (this *Workbooks) CanCheckOut(filename string) bool {
	retVal, _ := this.Call(0x00000816, []interface{}{filename})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var Workbooks_OpenXML__OptArgs= []string{
	"Stylesheets", 
}

func (this *Workbooks) OpenXML_(filename string, optArgs ...interface{}) *Workbook {
	optArgs = ole.ProcessOptArgs(Workbooks_OpenXML__OptArgs, optArgs)
	retVal, _ := this.Call(0x00000817, []interface{}{filename}, optArgs...)
	return NewWorkbook(retVal.IDispatch(), false, true)
}

var Workbooks_OpenXML_OptArgs= []string{
	"Stylesheets", "LoadOption", 
}

func (this *Workbooks) OpenXML(filename string, optArgs ...interface{}) *Workbook {
	optArgs = ole.ProcessOptArgs(Workbooks_OpenXML_OptArgs, optArgs)
	retVal, _ := this.Call(0x000008e8, []interface{}{filename}, optArgs...)
	return NewWorkbook(retVal.IDispatch(), false, true)
}

