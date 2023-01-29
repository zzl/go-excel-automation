package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 000208AE-0000-0000-C000-000000000046
var IID_Modules = syscall.GUID{0x000208AE, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Modules struct {
	ole.OleClient
}

func NewModules(pDisp *win32.IDispatch, addRef bool, scoped bool) *Modules {
	if pDisp == nil {
		return nil
	}
	p := &Modules{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ModulesFromVar(v ole.Variant) *Modules {
	return NewModules(v.IDispatch(), false, false)
}

func (this *Modules) IID() *syscall.GUID {
	return &IID_Modules
}

func (this *Modules) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Modules) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Modules) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Modules) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Modules) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Modules) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Modules) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Modules) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Modules) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Modules) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Modules) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Modules_Add_OptArgs = []string{
	"Before", "After", "Count",
}

func (this *Modules) Add(optArgs ...interface{}) *Module {
	optArgs = ole.ProcessOptArgs(Modules_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b5, nil, optArgs...)
	return NewModule(retVal.IDispatch(), false, true)
}

var Modules_Copy_OptArgs = []string{
	"Before", "After",
}

func (this *Modules) Copy(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Modules_Copy_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000227, nil, optArgs...)
	_ = retVal
}

func (this *Modules) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *Modules) Delete() {
	retVal, _ := this.Call(0x00000075, nil)
	_ = retVal
}

func (this *Modules) Dummy7_() {
	retVal, _ := this.Call(0x00010007, nil)
	_ = retVal
}

func (this *Modules) Item(index interface{}) *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000000aa, []interface{}{index})
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Modules_Move_OptArgs = []string{
	"Before", "After",
}

func (this *Modules) Move(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Modules_Move_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000027d, nil, optArgs...)
	_ = retVal
}

func (this *Modules) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Modules) ForEach(action func(item *ole.DispatchClass) bool) {
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

var Modules_PrintOut___OptArgs = []string{
	"From", "To", "Copies", "Preview",
	"ActivePrinter", "PrintToFile", "Collate",
}

func (this *Modules) PrintOut__(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Modules_PrintOut___OptArgs, optArgs)
	retVal, _ := this.Call(0x00000389, nil, optArgs...)
	_ = retVal
}

func (this *Modules) Dummy12_() {
	retVal, _ := this.Call(0x0001000c, nil)
	_ = retVal
}

var Modules_Select_OptArgs = []string{
	"Replace",
}

func (this *Modules) Select(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Modules_Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000eb, nil, optArgs...)
	_ = retVal
}

func (this *Modules) HPageBreaks() *HPageBreaks {
	retVal, _ := this.PropGet(0x0000058a, nil)
	return NewHPageBreaks(retVal.IDispatch(), false, true)
}

func (this *Modules) VPageBreaks() *VPageBreaks {
	retVal, _ := this.PropGet(0x0000058b, nil)
	return NewVPageBreaks(retVal.IDispatch(), false, true)
}

func (this *Modules) Visible() ole.Variant {
	retVal, _ := this.PropGet(0x0000022e, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Modules) SetVisible(rhs interface{}) {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

func (this *Modules) Default_(index interface{}) *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000000, []interface{}{index})
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Modules_PrintOut__OptArgs = []string{
	"From", "To", "Copies", "Preview",
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName",
}

func (this *Modules) PrintOut_(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Modules_PrintOut__OptArgs, optArgs)
	retVal, _ := this.Call(0x000006ec, nil, optArgs...)
	_ = retVal
}

var Modules_PrintOut_OptArgs = []string{
	"From", "To", "Copies", "Preview",
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName", "IgnorePrintAreas",
}

func (this *Modules) PrintOut(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Modules_PrintOut_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000939, nil, optArgs...)
	_ = retVal
}
