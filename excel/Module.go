package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 000208AD-0000-0000-C000-000000000046
var IID_Module = syscall.GUID{0x000208AD, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Module struct {
	ole.OleClient
}

func NewModule(pDisp *win32.IDispatch, addRef bool, scoped bool) *Module {
	if pDisp == nil {
		return nil
	}
	p := &Module{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ModuleFromVar(v ole.Variant) *Module {
	return NewModule(v.IDispatch(), false, false)
}

func (this *Module) IID() *syscall.GUID {
	return &IID_Module
}

func (this *Module) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Module) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Module) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Module) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Module) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Module) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Module) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Module) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Module) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Module) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Module) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Module) Activate() {
	retVal, _ := this.Call(0x00000130, nil)
	_ = retVal
}

var Module_Copy_OptArgs = []string{
	"Before", "After",
}

func (this *Module) Copy(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Module_Copy_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000227, nil, optArgs...)
	_ = retVal
}

func (this *Module) Delete() {
	retVal, _ := this.Call(0x00000075, nil)
	_ = retVal
}

func (this *Module) CodeName() string {
	retVal, _ := this.PropGet(0x0000055d, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Module) CodeName_() string {
	retVal, _ := this.PropGet(-2147418112, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Module) SetCodeName_(rhs string) {
	_ = this.PropPut(-2147418112, []interface{}{rhs})
}

func (this *Module) Index() int32 {
	retVal, _ := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

var Module_Move_OptArgs = []string{
	"Before", "After",
}

func (this *Module) Move(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Module_Move_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000027d, nil, optArgs...)
	_ = retVal
}

func (this *Module) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Module) SetName(rhs string) {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *Module) Next() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000001f6, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Module) OnDoubleClick() string {
	retVal, _ := this.PropGet(0x00000274, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Module) SetOnDoubleClick(rhs string) {
	_ = this.PropPut(0x00000274, []interface{}{rhs})
}

func (this *Module) OnSheetActivate() string {
	retVal, _ := this.PropGet(0x00000407, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Module) SetOnSheetActivate(rhs string) {
	_ = this.PropPut(0x00000407, []interface{}{rhs})
}

func (this *Module) OnSheetDeactivate() string {
	retVal, _ := this.PropGet(0x00000439, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Module) SetOnSheetDeactivate(rhs string) {
	_ = this.PropPut(0x00000439, []interface{}{rhs})
}

func (this *Module) PageSetup() *PageSetup {
	retVal, _ := this.PropGet(0x000003e6, nil)
	return NewPageSetup(retVal.IDispatch(), false, true)
}

func (this *Module) Previous() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000001f7, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Module_PrintOut___OptArgs = []string{
	"From", "To", "Copies", "Preview",
	"ActivePrinter", "PrintToFile", "Collate",
}

func (this *Module) PrintOut__(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Module_PrintOut___OptArgs, optArgs)
	retVal, _ := this.Call(0x00000389, nil, optArgs...)
	_ = retVal
}

func (this *Module) Dummy18_() {
	retVal, _ := this.Call(0x00010012, nil)
	_ = retVal
}

var Module_Protect__OptArgs = []string{
	"Password", "DrawingObjects", "Contents", "Scenarios", "UserInterfaceOnly",
}

func (this *Module) Protect_(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Module_Protect__OptArgs, optArgs)
	retVal, _ := this.Call(0x0000011a, nil, optArgs...)
	_ = retVal
}

func (this *Module) ProtectContents() bool {
	retVal, _ := this.PropGet(0x00000124, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Module) Dummy21_() {
	retVal, _ := this.Call(0x00010015, nil)
	_ = retVal
}

func (this *Module) ProtectionMode() bool {
	retVal, _ := this.PropGet(0x00000487, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Module) Dummy23_() {
	retVal, _ := this.Call(0x00010017, nil)
	_ = retVal
}

var Module_SaveAs__OptArgs = []string{
	"FileFormat", "Password", "WriteResPassword", "ReadOnlyRecommended",
	"CreateBackup", "AddToMru", "TextCodepage", "TextVisualLayout",
}

func (this *Module) SaveAs_(filename string, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Module_SaveAs__OptArgs, optArgs)
	retVal, _ := this.Call(0x0000011c, []interface{}{filename}, optArgs...)
	_ = retVal
}

var Module_Select_OptArgs = []string{
	"Replace",
}

func (this *Module) Select(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Module_Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000eb, nil, optArgs...)
	_ = retVal
}

var Module_Unprotect_OptArgs = []string{
	"Password",
}

func (this *Module) Unprotect(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Module_Unprotect_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000011d, nil, optArgs...)
	_ = retVal
}

func (this *Module) Visible() int32 {
	retVal, _ := this.PropGet(0x0000022e, nil)
	return retVal.LValVal()
}

func (this *Module) SetVisible(rhs int32) {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

func (this *Module) Shapes() *Shapes {
	retVal, _ := this.PropGet(0x00000561, nil)
	return NewShapes(retVal.IDispatch(), false, true)
}

var Module_InsertFile_OptArgs = []string{
	"Merge",
}

func (this *Module) InsertFile(filename interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Module_InsertFile_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000248, []interface{}{filename}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Module_SaveAs_OptArgs = []string{
	"FileFormat", "Password", "WriteResPassword", "ReadOnlyRecommended",
	"CreateBackup", "AddToMru", "TextCodepage", "TextVisualLayout",
}

func (this *Module) SaveAs(filename string, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Module_SaveAs_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000785, []interface{}{filename}, optArgs...)
	_ = retVal
}

var Module_Protect_OptArgs = []string{
	"Password", "DrawingObjects", "Contents", "Scenarios", "UserInterfaceOnly",
}

func (this *Module) Protect(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Module_Protect_OptArgs, optArgs)
	retVal, _ := this.Call(0x000007ed, nil, optArgs...)
	_ = retVal
}

var Module_PrintOut__OptArgs = []string{
	"From", "To", "Copies", "Preview",
	"ActivePrinter", "PrintToFile", "Collate",
}

func (this *Module) PrintOut_(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Module_PrintOut__OptArgs, optArgs)
	retVal, _ := this.Call(0x000006ec, nil, optArgs...)
	_ = retVal
}

var Module_PrintOut_OptArgs = []string{
	"From", "To", "Copies", "Preview",
	"ActivePrinter", "PrintToFile", "Collate",
}

func (this *Module) PrintOut(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Module_PrintOut_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000939, nil, optArgs...)
	_ = retVal
}
