package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
)

// 000208DA-0000-0000-C000-000000000046
var IID_Workbook_ = syscall.GUID{0x000208DA, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Workbook_ struct {
	ole.OleClient
}

func NewWorkbook_(pDisp *win32.IDispatch, addRef bool, scoped bool) *Workbook_ {
	if pDisp == nil {
		return nil
	}
	p := &Workbook_{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func Workbook_FromVar(v ole.Variant) *Workbook_ {
	return NewWorkbook_(v.IDispatch(), false, false)
}

func (this *Workbook_) IID() *syscall.GUID {
	return &IID_Workbook_
}

func (this *Workbook_) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Workbook_) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Workbook_) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Workbook_) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Workbook_) AcceptLabelsInFormulas() bool {
	retVal, _ := this.PropGet(0x000005a1, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetAcceptLabelsInFormulas(rhs bool) {
	_ = this.PropPut(0x000005a1, []interface{}{rhs})
}

func (this *Workbook_) Activate() {
	retVal, _ := this.Call(0x00000130, nil)
	_ = retVal
}

func (this *Workbook_) ActiveChart() *Chart {
	retVal, _ := this.PropGet(0x000000b7, nil)
	return NewChart(retVal.IDispatch(), false, true)
}

func (this *Workbook_) ActiveSheet() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000133, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Workbook_) Author() string {
	retVal, _ := this.PropGet(0x0000023e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Workbook_) SetAuthor(rhs string) {
	_ = this.PropPut(0x0000023e, []interface{}{rhs})
}

func (this *Workbook_) AutoUpdateFrequency() int32 {
	retVal, _ := this.PropGet(0x000005a2, nil)
	return retVal.LValVal()
}

func (this *Workbook_) SetAutoUpdateFrequency(rhs int32) {
	_ = this.PropPut(0x000005a2, []interface{}{rhs})
}

func (this *Workbook_) AutoUpdateSaveChanges() bool {
	retVal, _ := this.PropGet(0x000005a3, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetAutoUpdateSaveChanges(rhs bool) {
	_ = this.PropPut(0x000005a3, []interface{}{rhs})
}

func (this *Workbook_) ChangeHistoryDuration() int32 {
	retVal, _ := this.PropGet(0x000005a4, nil)
	return retVal.LValVal()
}

func (this *Workbook_) SetChangeHistoryDuration(rhs int32) {
	_ = this.PropPut(0x000005a4, []interface{}{rhs})
}

func (this *Workbook_) BuiltinDocumentProperties() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000498, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Workbook__ChangeFileAccess_OptArgs = []string{
	"WritePassword", "Notify",
}

func (this *Workbook_) ChangeFileAccess(mode int32, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__ChangeFileAccess_OptArgs, optArgs)
	retVal, _ := this.Call(0x000003dd, []interface{}{mode}, optArgs...)
	_ = retVal
}

var Workbook__ChangeLink_OptArgs = []string{
	"Type",
}

func (this *Workbook_) ChangeLink(name string, newName string, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__ChangeLink_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000322, []interface{}{name, newName}, optArgs...)
	_ = retVal
}

func (this *Workbook_) Charts() *Sheets {
	retVal, _ := this.PropGet(0x00000079, nil)
	return NewSheets(retVal.IDispatch(), false, true)
}

var Workbook__Close_OptArgs = []string{
	"SaveChanges", "Filename", "RouteWorkbook",
}

func (this *Workbook_) Close(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__Close_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000115, nil, optArgs...)
	_ = retVal
}

func (this *Workbook_) CodeName() string {
	retVal, _ := this.PropGet(0x0000055d, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Workbook_) CodeName_() string {
	retVal, _ := this.PropGet(-2147418112, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Workbook_) SetCodeName_(rhs string) {
	_ = this.PropPut(-2147418112, []interface{}{rhs})
}

var Workbook__Colors_OptArgs = []string{
	"Index",
}

func (this *Workbook_) Colors(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Workbook__Colors_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000011e, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Workbook__SetColors_OptArgs = []string{
	"Index",
}

func (this *Workbook_) SetColors(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__SetColors_OptArgs, optArgs)
	_ = this.PropPut(0x0000011e, nil, optArgs...)
}

func (this *Workbook_) CommandBars() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000059f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Workbook_) Comments() string {
	retVal, _ := this.PropGet(0x0000023f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Workbook_) SetComments(rhs string) {
	_ = this.PropPut(0x0000023f, []interface{}{rhs})
}

func (this *Workbook_) ConflictResolution() int32 {
	retVal, _ := this.PropGet(0x00000497, nil)
	return retVal.LValVal()
}

func (this *Workbook_) SetConflictResolution(rhs int32) {
	_ = this.PropPut(0x00000497, []interface{}{rhs})
}

func (this *Workbook_) Container() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000004a6, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Workbook_) CreateBackup() bool {
	retVal, _ := this.PropGet(0x0000011f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) CustomDocumentProperties() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000499, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Workbook_) Date1904() bool {
	retVal, _ := this.PropGet(0x00000193, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetDate1904(rhs bool) {
	_ = this.PropPut(0x00000193, []interface{}{rhs})
}

func (this *Workbook_) DeleteNumberFormat(numberFormat string) {
	retVal, _ := this.Call(0x0000018d, []interface{}{numberFormat})
	_ = retVal
}

func (this *Workbook_) DialogSheets() *Sheets {
	retVal, _ := this.PropGet(0x000002fc, nil)
	return NewSheets(retVal.IDispatch(), false, true)
}

func (this *Workbook_) DisplayDrawingObjects() int32 {
	retVal, _ := this.PropGet(0x00000194, nil)
	return retVal.LValVal()
}

func (this *Workbook_) SetDisplayDrawingObjects(rhs int32) {
	_ = this.PropPut(0x00000194, []interface{}{rhs})
}

func (this *Workbook_) ExclusiveAccess() bool {
	retVal, _ := this.Call(0x00000490, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) FileFormat() int32 {
	retVal, _ := this.PropGet(0x00000120, nil)
	return retVal.LValVal()
}

func (this *Workbook_) ForwardMailer() {
	retVal, _ := this.Call(0x000003cd, nil)
	_ = retVal
}

func (this *Workbook_) FullName() string {
	retVal, _ := this.PropGet(0x00000121, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Workbook_) HasMailer() bool {
	retVal, _ := this.PropGet(0x000003d0, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetHasMailer(rhs bool) {
	_ = this.PropPut(0x000003d0, []interface{}{rhs})
}

func (this *Workbook_) HasPassword() bool {
	retVal, _ := this.PropGet(0x00000122, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) HasRoutingSlip() bool {
	retVal, _ := this.PropGet(0x000003b6, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetHasRoutingSlip(rhs bool) {
	_ = this.PropPut(0x000003b6, []interface{}{rhs})
}

func (this *Workbook_) IsAddin() bool {
	retVal, _ := this.PropGet(0x000005a5, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetIsAddin(rhs bool) {
	_ = this.PropPut(0x000005a5, []interface{}{rhs})
}

func (this *Workbook_) Keywords() string {
	retVal, _ := this.PropGet(0x00000241, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Workbook_) SetKeywords(rhs string) {
	_ = this.PropPut(0x00000241, []interface{}{rhs})
}

var Workbook__LinkInfo_OptArgs = []string{
	"Type", "EditionRef",
}

func (this *Workbook_) LinkInfo(name string, linkInfo int32, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Workbook__LinkInfo_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000327, []interface{}{name, linkInfo}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Workbook__LinkSources_OptArgs = []string{
	"Type",
}

func (this *Workbook_) LinkSources(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Workbook__LinkSources_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000328, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Workbook_) Mailer() *Mailer {
	retVal, _ := this.PropGet(0x000003d3, nil)
	return NewMailer(retVal.IDispatch(), false, true)
}

func (this *Workbook_) MergeWorkbook(filename interface{}) {
	retVal, _ := this.Call(0x000005a6, []interface{}{filename})
	_ = retVal
}

func (this *Workbook_) Modules() *Sheets {
	retVal, _ := this.PropGet(0x00000246, nil)
	return NewSheets(retVal.IDispatch(), false, true)
}

func (this *Workbook_) MultiUserEditing() bool {
	retVal, _ := this.PropGet(0x00000491, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Workbook_) Names() *Names {
	retVal, _ := this.PropGet(0x000001ba, nil)
	return NewNames(retVal.IDispatch(), false, true)
}

func (this *Workbook_) NewWindow() *Window {
	retVal, _ := this.Call(0x00000118, nil)
	return NewWindow(retVal.IDispatch(), false, true)
}

func (this *Workbook_) OnSave() string {
	retVal, _ := this.PropGet(0x0000049a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Workbook_) SetOnSave(rhs string) {
	_ = this.PropPut(0x0000049a, []interface{}{rhs})
}

func (this *Workbook_) OnSheetActivate() string {
	retVal, _ := this.PropGet(0x00000407, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Workbook_) SetOnSheetActivate(rhs string) {
	_ = this.PropPut(0x00000407, []interface{}{rhs})
}

func (this *Workbook_) OnSheetDeactivate() string {
	retVal, _ := this.PropGet(0x00000439, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Workbook_) SetOnSheetDeactivate(rhs string) {
	_ = this.PropPut(0x00000439, []interface{}{rhs})
}

var Workbook__OpenLinks_OptArgs = []string{
	"ReadOnly", "Type",
}

func (this *Workbook_) OpenLinks(name string, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__OpenLinks_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000323, []interface{}{name}, optArgs...)
	_ = retVal
}

func (this *Workbook_) Path() string {
	retVal, _ := this.PropGet(0x00000123, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Workbook_) PersonalViewListSettings() bool {
	retVal, _ := this.PropGet(0x000005a7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetPersonalViewListSettings(rhs bool) {
	_ = this.PropPut(0x000005a7, []interface{}{rhs})
}

func (this *Workbook_) PersonalViewPrintSettings() bool {
	retVal, _ := this.PropGet(0x000005a8, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetPersonalViewPrintSettings(rhs bool) {
	_ = this.PropPut(0x000005a8, []interface{}{rhs})
}

func (this *Workbook_) PivotCaches() *PivotCaches {
	retVal, _ := this.Call(0x000005a9, nil)
	return NewPivotCaches(retVal.IDispatch(), false, true)
}

var Workbook__Post_OptArgs = []string{
	"DestName",
}

func (this *Workbook_) Post(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__Post_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000048e, nil, optArgs...)
	_ = retVal
}

func (this *Workbook_) PrecisionAsDisplayed() bool {
	retVal, _ := this.PropGet(0x00000195, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetPrecisionAsDisplayed(rhs bool) {
	_ = this.PropPut(0x00000195, []interface{}{rhs})
}

var Workbook__PrintOut___OptArgs = []string{
	"From", "To", "Copies", "Preview",
	"ActivePrinter", "PrintToFile", "Collate",
}

func (this *Workbook_) PrintOut__(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__PrintOut___OptArgs, optArgs)
	retVal, _ := this.Call(0x00000389, nil, optArgs...)
	_ = retVal
}

var Workbook__PrintPreview_OptArgs = []string{
	"EnableChanges",
}

func (this *Workbook_) PrintPreview(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__PrintPreview_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000119, nil, optArgs...)
	_ = retVal
}

var Workbook__Protect__OptArgs = []string{
	"Password", "Structure", "Windows",
}

func (this *Workbook_) Protect_(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__Protect__OptArgs, optArgs)
	retVal, _ := this.Call(0x0000011a, nil, optArgs...)
	_ = retVal
}

var Workbook__ProtectSharing__OptArgs = []string{
	"Filename", "Password", "WriteResPassword", "ReadOnlyRecommended",
	"CreateBackup", "SharingPassword",
}

func (this *Workbook_) ProtectSharing_(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__ProtectSharing__OptArgs, optArgs)
	retVal, _ := this.Call(0x000005aa, nil, optArgs...)
	_ = retVal
}

func (this *Workbook_) ProtectStructure() bool {
	retVal, _ := this.PropGet(0x0000024c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) ProtectWindows() bool {
	retVal, _ := this.PropGet(0x00000127, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) ReadOnly() bool {
	retVal, _ := this.PropGet(0x00000128, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) ReadOnlyRecommended_() bool {
	retVal, _ := this.PropGet(0x00000129, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) RefreshAll() {
	retVal, _ := this.Call(0x000005ac, nil)
	_ = retVal
}

func (this *Workbook_) Reply() {
	retVal, _ := this.Call(0x000003d1, nil)
	_ = retVal
}

func (this *Workbook_) ReplyAll() {
	retVal, _ := this.Call(0x000003d2, nil)
	_ = retVal
}

func (this *Workbook_) RemoveUser(index int32) {
	retVal, _ := this.Call(0x000005ad, []interface{}{index})
	_ = retVal
}

func (this *Workbook_) RevisionNumber() int32 {
	retVal, _ := this.PropGet(0x00000494, nil)
	return retVal.LValVal()
}

func (this *Workbook_) Route() {
	retVal, _ := this.Call(0x000003b2, nil)
	_ = retVal
}

func (this *Workbook_) Routed() bool {
	retVal, _ := this.PropGet(0x000003b7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) RoutingSlip() *RoutingSlip {
	retVal, _ := this.PropGet(0x000003b5, nil)
	return NewRoutingSlip(retVal.IDispatch(), false, true)
}

func (this *Workbook_) RunAutoMacros(which int32) {
	retVal, _ := this.Call(0x0000027a, []interface{}{which})
	_ = retVal
}

func (this *Workbook_) Save() {
	retVal, _ := this.Call(0x0000011b, nil)
	_ = retVal
}

var Workbook__SaveAs__OptArgs = []string{
	"Filename", "FileFormat", "Password", "WriteResPassword",
	"ReadOnlyRecommended", "CreateBackup", "AccessMode", "ConflictResolution",
	"AddToMru", "TextCodepage", "TextVisualLayout",
}

func (this *Workbook_) SaveAs_(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__SaveAs__OptArgs, optArgs)
	retVal, _ := this.Call(0x0000011c, nil, optArgs...)
	_ = retVal
}

var Workbook__SaveCopyAs_OptArgs = []string{
	"Filename",
}

func (this *Workbook_) SaveCopyAs(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__SaveCopyAs_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000af, nil, optArgs...)
	_ = retVal
}

func (this *Workbook_) Saved() bool {
	retVal, _ := this.PropGet(0x0000012a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetSaved(rhs bool) {
	_ = this.PropPut(0x0000012a, []interface{}{rhs})
}

func (this *Workbook_) SaveLinkValues() bool {
	retVal, _ := this.PropGet(0x00000196, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetSaveLinkValues(rhs bool) {
	_ = this.PropPut(0x00000196, []interface{}{rhs})
}

var Workbook__SendMail_OptArgs = []string{
	"Subject", "ReturnReceipt",
}

func (this *Workbook_) SendMail(recipients interface{}, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__SendMail_OptArgs, optArgs)
	retVal, _ := this.Call(0x000003b3, []interface{}{recipients}, optArgs...)
	_ = retVal
}

var Workbook__SendMailer_OptArgs = []string{
	"FileFormat", "Priority",
}

func (this *Workbook_) SendMailer(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__SendMailer_OptArgs, optArgs)
	retVal, _ := this.Call(0x000003d4, nil, optArgs...)
	_ = retVal
}

var Workbook__SetLinkOnData_OptArgs = []string{
	"Procedure",
}

func (this *Workbook_) SetLinkOnData(name string, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__SetLinkOnData_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000329, []interface{}{name}, optArgs...)
	_ = retVal
}

func (this *Workbook_) Sheets() *Sheets {
	retVal, _ := this.PropGet(0x000001e5, nil)
	return NewSheets(retVal.IDispatch(), false, true)
}

func (this *Workbook_) ShowConflictHistory() bool {
	retVal, _ := this.PropGet(0x00000493, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetShowConflictHistory(rhs bool) {
	_ = this.PropPut(0x00000493, []interface{}{rhs})
}

func (this *Workbook_) Styles() *Styles {
	retVal, _ := this.PropGet(0x000001ed, nil)
	return NewStyles(retVal.IDispatch(), false, true)
}

func (this *Workbook_) Subject() string {
	retVal, _ := this.PropGet(0x000003b9, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Workbook_) SetSubject(rhs string) {
	_ = this.PropPut(0x000003b9, []interface{}{rhs})
}

func (this *Workbook_) Title() string {
	retVal, _ := this.PropGet(0x000000c7, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Workbook_) SetTitle(rhs string) {
	_ = this.PropPut(0x000000c7, []interface{}{rhs})
}

var Workbook__Unprotect_OptArgs = []string{
	"Password",
}

func (this *Workbook_) Unprotect(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__Unprotect_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000011d, nil, optArgs...)
	_ = retVal
}

var Workbook__UnprotectSharing_OptArgs = []string{
	"SharingPassword",
}

func (this *Workbook_) UnprotectSharing(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__UnprotectSharing_OptArgs, optArgs)
	retVal, _ := this.Call(0x000005af, nil, optArgs...)
	_ = retVal
}

func (this *Workbook_) UpdateFromFile() {
	retVal, _ := this.Call(0x000003e3, nil)
	_ = retVal
}

var Workbook__UpdateLink_OptArgs = []string{
	"Name", "Type",
}

func (this *Workbook_) UpdateLink(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__UpdateLink_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000324, nil, optArgs...)
	_ = retVal
}

func (this *Workbook_) UpdateRemoteReferences() bool {
	retVal, _ := this.PropGet(0x0000019b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetUpdateRemoteReferences(rhs bool) {
	_ = this.PropPut(0x0000019b, []interface{}{rhs})
}

func (this *Workbook_) UserControl() bool {
	retVal, _ := this.PropGet(0x000004ba, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetUserControl(rhs bool) {
	_ = this.PropPut(0x000004ba, []interface{}{rhs})
}

func (this *Workbook_) UserStatus() ole.Variant {
	retVal, _ := this.PropGet(0x00000495, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Workbook_) CustomViews() *CustomViews {
	retVal, _ := this.PropGet(0x000005b0, nil)
	return NewCustomViews(retVal.IDispatch(), false, true)
}

func (this *Workbook_) Windows() *Windows {
	retVal, _ := this.PropGet(0x000001ae, nil)
	return NewWindows(retVal.IDispatch(), false, true)
}

func (this *Workbook_) Worksheets() *Sheets {
	retVal, _ := this.PropGet(0x000001ee, nil)
	return NewSheets(retVal.IDispatch(), false, true)
}

func (this *Workbook_) WriteReserved() bool {
	retVal, _ := this.PropGet(0x0000012b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) WriteReservedBy() string {
	retVal, _ := this.PropGet(0x0000012c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Workbook_) Excel4IntlMacroSheets() *Sheets {
	retVal, _ := this.PropGet(0x00000245, nil)
	return NewSheets(retVal.IDispatch(), false, true)
}

func (this *Workbook_) Excel4MacroSheets() *Sheets {
	retVal, _ := this.PropGet(0x00000243, nil)
	return NewSheets(retVal.IDispatch(), false, true)
}

func (this *Workbook_) TemplateRemoveExtData() bool {
	retVal, _ := this.PropGet(0x000005b1, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetTemplateRemoveExtData(rhs bool) {
	_ = this.PropPut(0x000005b1, []interface{}{rhs})
}

var Workbook__HighlightChangesOptions_OptArgs = []string{
	"When", "Who", "Where",
}

func (this *Workbook_) HighlightChangesOptions(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__HighlightChangesOptions_OptArgs, optArgs)
	retVal, _ := this.Call(0x000005b2, nil, optArgs...)
	_ = retVal
}

func (this *Workbook_) HighlightChangesOnScreen() bool {
	retVal, _ := this.PropGet(0x000005b5, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetHighlightChangesOnScreen(rhs bool) {
	_ = this.PropPut(0x000005b5, []interface{}{rhs})
}

func (this *Workbook_) KeepChangeHistory() bool {
	retVal, _ := this.PropGet(0x000005b6, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetKeepChangeHistory(rhs bool) {
	_ = this.PropPut(0x000005b6, []interface{}{rhs})
}

func (this *Workbook_) ListChangesOnNewSheet() bool {
	retVal, _ := this.PropGet(0x000005b7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetListChangesOnNewSheet(rhs bool) {
	_ = this.PropPut(0x000005b7, []interface{}{rhs})
}

var Workbook__PurgeChangeHistoryNow_OptArgs = []string{
	"SharingPassword",
}

func (this *Workbook_) PurgeChangeHistoryNow(days int32, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__PurgeChangeHistoryNow_OptArgs, optArgs)
	retVal, _ := this.Call(0x000005b8, []interface{}{days}, optArgs...)
	_ = retVal
}

var Workbook__AcceptAllChanges_OptArgs = []string{
	"When", "Who", "Where",
}

func (this *Workbook_) AcceptAllChanges(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__AcceptAllChanges_OptArgs, optArgs)
	retVal, _ := this.Call(0x000005ba, nil, optArgs...)
	_ = retVal
}

var Workbook__RejectAllChanges_OptArgs = []string{
	"When", "Who", "Where",
}

func (this *Workbook_) RejectAllChanges(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__RejectAllChanges_OptArgs, optArgs)
	retVal, _ := this.Call(0x000005bb, nil, optArgs...)
	_ = retVal
}

var Workbook__PivotTableWizard_OptArgs = []string{
	"SourceType", "SourceData", "TableDestination", "TableName",
	"RowGrand", "ColumnGrand", "SaveData", "HasAutoFormat",
	"AutoPage", "Reserved", "BackgroundQuery", "OptimizeCache",
	"PageFieldOrder", "PageFieldWrapCount", "ReadData", "Connection",
}

func (this *Workbook_) PivotTableWizard(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__PivotTableWizard_OptArgs, optArgs)
	retVal, _ := this.Call(0x000002ac, nil, optArgs...)
	_ = retVal
}

func (this *Workbook_) ResetColors() {
	retVal, _ := this.Call(0x000005bc, nil)
	_ = retVal
}

func (this *Workbook_) VBProject() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000005bd, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Workbook__FollowHyperlink_OptArgs = []string{
	"SubAddress", "NewWindow", "AddHistory", "ExtraInfo",
	"Method", "HeaderInfo",
}

func (this *Workbook_) FollowHyperlink(address string, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__FollowHyperlink_OptArgs, optArgs)
	retVal, _ := this.Call(0x000005be, []interface{}{address}, optArgs...)
	_ = retVal
}

func (this *Workbook_) AddToFavorites() {
	retVal, _ := this.Call(0x000005c4, nil)
	_ = retVal
}

func (this *Workbook_) IsInplace() bool {
	retVal, _ := this.PropGet(0x000006e9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var Workbook__PrintOut__OptArgs = []string{
	"From", "To", "Copies", "Preview",
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName",
}

func (this *Workbook_) PrintOut_(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__PrintOut__OptArgs, optArgs)
	retVal, _ := this.Call(0x000006ec, nil, optArgs...)
	_ = retVal
}

func (this *Workbook_) WebPagePreview() {
	retVal, _ := this.Call(0x0000071a, nil)
	_ = retVal
}

func (this *Workbook_) PublishObjects() *PublishObjects {
	retVal, _ := this.PropGet(0x0000071b, nil)
	return NewPublishObjects(retVal.IDispatch(), false, true)
}

func (this *Workbook_) WebOptions() *WebOptions {
	retVal, _ := this.PropGet(0x0000071c, nil)
	return NewWebOptions(retVal.IDispatch(), false, true)
}

func (this *Workbook_) ReloadAs(encoding int32) {
	retVal, _ := this.Call(0x0000071d, []interface{}{encoding})
	_ = retVal
}

func (this *Workbook_) HTMLProject() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000071f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Workbook_) EnvelopeVisible() bool {
	retVal, _ := this.PropGet(0x00000720, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetEnvelopeVisible(rhs bool) {
	_ = this.PropPut(0x00000720, []interface{}{rhs})
}

func (this *Workbook_) CalculationVersion() int32 {
	retVal, _ := this.PropGet(0x0000070e, nil)
	return retVal.LValVal()
}

func (this *Workbook_) Dummy17(calcid int32) {
	retVal, _ := this.Call(0x000007fc, []interface{}{calcid})
	_ = retVal
}

func (this *Workbook_) Sblt(s string) {
	retVal, _ := this.Call(0x00000722, []interface{}{s})
	_ = retVal
}

func (this *Workbook_) VBASigned() bool {
	retVal, _ := this.PropGet(0x00000724, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) ShowPivotTableFieldList() bool {
	retVal, _ := this.PropGet(0x000007fe, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetShowPivotTableFieldList(rhs bool) {
	_ = this.PropPut(0x000007fe, []interface{}{rhs})
}

func (this *Workbook_) UpdateLinks() int32 {
	retVal, _ := this.PropGet(0x00000360, nil)
	return retVal.LValVal()
}

func (this *Workbook_) SetUpdateLinks(rhs int32) {
	_ = this.PropPut(0x00000360, []interface{}{rhs})
}

func (this *Workbook_) BreakLink(name string, type_ int32) {
	retVal, _ := this.Call(0x000007ff, []interface{}{name, type_})
	_ = retVal
}

func (this *Workbook_) Dummy16() {
	retVal, _ := this.Call(0x00000800, nil)
	_ = retVal
}

var Workbook__SaveAs_OptArgs = []string{
	"Filename", "FileFormat", "Password", "WriteResPassword",
	"ReadOnlyRecommended", "CreateBackup", "AccessMode", "ConflictResolution",
	"AddToMru", "TextCodepage", "TextVisualLayout", "Local",
}

func (this *Workbook_) SaveAs(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__SaveAs_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000785, nil, optArgs...)
	_ = retVal
}

func (this *Workbook_) EnableAutoRecover() bool {
	retVal, _ := this.PropGet(0x00000801, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetEnableAutoRecover(rhs bool) {
	_ = this.PropPut(0x00000801, []interface{}{rhs})
}

func (this *Workbook_) RemovePersonalInformation() bool {
	retVal, _ := this.PropGet(0x00000802, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetRemovePersonalInformation(rhs bool) {
	_ = this.PropPut(0x00000802, []interface{}{rhs})
}

func (this *Workbook_) FullNameURLEncoded() string {
	retVal, _ := this.PropGet(0x00000787, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var Workbook__CheckIn_OptArgs = []string{
	"SaveChanges", "Comments", "MakePublic",
}

func (this *Workbook_) CheckIn(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__CheckIn_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000803, nil, optArgs...)
	_ = retVal
}

func (this *Workbook_) CanCheckIn() bool {
	retVal, _ := this.Call(0x00000805, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var Workbook__SendForReview_OptArgs = []string{
	"Recipients", "Subject", "ShowMessage", "IncludeAttachment",
}

func (this *Workbook_) SendForReview(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__SendForReview_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000806, nil, optArgs...)
	_ = retVal
}

var Workbook__ReplyWithChanges_OptArgs = []string{
	"ShowMessage",
}

func (this *Workbook_) ReplyWithChanges(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__ReplyWithChanges_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000809, nil, optArgs...)
	_ = retVal
}

func (this *Workbook_) EndReview() {
	retVal, _ := this.Call(0x0000080a, nil)
	_ = retVal
}

func (this *Workbook_) Password() string {
	retVal, _ := this.PropGet(0x000001ad, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Workbook_) SetPassword(rhs string) {
	_ = this.PropPut(0x000001ad, []interface{}{rhs})
}

func (this *Workbook_) WritePassword() string {
	retVal, _ := this.PropGet(0x00000468, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Workbook_) SetWritePassword(rhs string) {
	_ = this.PropPut(0x00000468, []interface{}{rhs})
}

func (this *Workbook_) PasswordEncryptionProvider() string {
	retVal, _ := this.PropGet(0x0000080b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Workbook_) PasswordEncryptionAlgorithm() string {
	retVal, _ := this.PropGet(0x0000080c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Workbook_) PasswordEncryptionKeyLength() int32 {
	retVal, _ := this.PropGet(0x0000080d, nil)
	return retVal.LValVal()
}

var Workbook__SetPasswordEncryptionOptions_OptArgs = []string{
	"PasswordEncryptionProvider", "PasswordEncryptionAlgorithm", "PasswordEncryptionKeyLength", "PasswordEncryptionFileProperties",
}

func (this *Workbook_) SetPasswordEncryptionOptions(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__SetPasswordEncryptionOptions_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000080e, nil, optArgs...)
	_ = retVal
}

func (this *Workbook_) PasswordEncryptionFileProperties() bool {
	retVal, _ := this.PropGet(0x0000080f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) ReadOnlyRecommended() bool {
	retVal, _ := this.PropGet(0x000007d5, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetReadOnlyRecommended(rhs bool) {
	_ = this.PropPut(0x000007d5, []interface{}{rhs})
}

var Workbook__Protect_OptArgs = []string{
	"Password", "Structure", "Windows",
}

func (this *Workbook_) Protect(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__Protect_OptArgs, optArgs)
	retVal, _ := this.Call(0x000007ed, nil, optArgs...)
	_ = retVal
}

func (this *Workbook_) SmartTagOptions() *SmartTagOptions {
	retVal, _ := this.PropGet(0x00000810, nil)
	return NewSmartTagOptions(retVal.IDispatch(), false, true)
}

func (this *Workbook_) RecheckSmartTags() {
	retVal, _ := this.Call(0x00000811, nil)
	_ = retVal
}

func (this *Workbook_) Permission() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000008d8, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Workbook_) SharedWorkspace() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000008d9, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Workbook_) Sync() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000008da, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Workbook__SendFaxOverInternet_OptArgs = []string{
	"Recipients", "Subject", "ShowMessage",
}

func (this *Workbook_) SendFaxOverInternet(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__SendFaxOverInternet_OptArgs, optArgs)
	retVal, _ := this.Call(0x000008db, nil, optArgs...)
	_ = retVal
}

func (this *Workbook_) XmlNamespaces() *XmlNamespaces {
	retVal, _ := this.PropGet(0x000008dc, nil)
	return NewXmlNamespaces(retVal.IDispatch(), false, true)
}

func (this *Workbook_) XmlMaps() *XmlMaps {
	retVal, _ := this.PropGet(0x000008dd, nil)
	return NewXmlMaps(retVal.IDispatch(), false, true)
}

var Workbook__XmlImport_OptArgs = []string{
	"Overwrite", "Destination",
}

func (this *Workbook_) XmlImport(url string, importMap **XmlMap, optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Workbook__XmlImport_OptArgs, optArgs)
	retVal, _ := this.Call(0x000008de, []interface{}{url, importMap}, optArgs...)
	return retVal.LValVal()
}

func (this *Workbook_) SmartDocument() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000008e1, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Workbook_) DocumentLibraryVersions() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000008e2, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Workbook_) InactiveListBorderVisible() bool {
	retVal, _ := this.PropGet(0x000008e3, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetInactiveListBorderVisible(rhs bool) {
	_ = this.PropPut(0x000008e3, []interface{}{rhs})
}

func (this *Workbook_) DisplayInkComments() bool {
	retVal, _ := this.PropGet(0x000008e4, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetDisplayInkComments(rhs bool) {
	_ = this.PropPut(0x000008e4, []interface{}{rhs})
}

var Workbook__XmlImportXml_OptArgs = []string{
	"Overwrite", "Destination",
}

func (this *Workbook_) XmlImportXml(data string, importMap **XmlMap, optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Workbook__XmlImportXml_OptArgs, optArgs)
	retVal, _ := this.Call(0x000008e5, []interface{}{data, importMap}, optArgs...)
	return retVal.LValVal()
}

func (this *Workbook_) SaveAsXMLData(filename string, map_ *XmlMap) {
	retVal, _ := this.Call(0x000008e6, []interface{}{filename, map_})
	_ = retVal
}

func (this *Workbook_) ToggleFormsDesign() {
	retVal, _ := this.Call(0x000008e7, nil)
	_ = retVal
}

func (this *Workbook_) ContentTypeProperties() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000009d0, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Workbook_) Connections() *Connections {
	retVal, _ := this.PropGet(0x000009d1, nil)
	return NewConnections(retVal.IDispatch(), false, true)
}

func (this *Workbook_) RemoveDocumentInformation(removeDocInfoType int32) {
	retVal, _ := this.Call(0x000009d2, []interface{}{removeDocInfoType})
	_ = retVal
}

func (this *Workbook_) Signatures() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000009d4, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Workbook__CheckInWithVersion_OptArgs = []string{
	"SaveChanges", "Comments", "MakePublic", "VersionType",
}

func (this *Workbook_) CheckInWithVersion(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__CheckInWithVersion_OptArgs, optArgs)
	retVal, _ := this.Call(0x000009d5, nil, optArgs...)
	_ = retVal
}

func (this *Workbook_) ServerPolicy() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000009d7, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Workbook_) LockServerFile() {
	retVal, _ := this.Call(0x000009d8, nil)
	_ = retVal
}

func (this *Workbook_) DocumentInspectors() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000009d9, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Workbook_) GetWorkflowTasks() *ole.DispatchClass {
	retVal, _ := this.Call(0x000009da, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Workbook_) GetWorkflowTemplates() *ole.DispatchClass {
	retVal, _ := this.Call(0x000009db, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Workbook__PrintOut_OptArgs = []string{
	"From", "To", "Copies", "Preview",
	"ActivePrinter", "PrintToFile", "Collate", "PrToFileName", "IgnorePrintAreas",
}

func (this *Workbook_) PrintOut(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__PrintOut_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000939, nil, optArgs...)
	_ = retVal
}

func (this *Workbook_) ServerViewableItems() *ServerViewableItems {
	retVal, _ := this.PropGet(0x000009dc, nil)
	return NewServerViewableItems(retVal.IDispatch(), false, true)
}

func (this *Workbook_) TableStyles() *TableStyles {
	retVal, _ := this.PropGet(0x000009dd, nil)
	return NewTableStyles(retVal.IDispatch(), false, true)
}

func (this *Workbook_) DefaultTableStyle() ole.Variant {
	retVal, _ := this.PropGet(0x000009de, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Workbook_) SetDefaultTableStyle(rhs interface{}) {
	_ = this.PropPut(0x000009de, []interface{}{rhs})
}

func (this *Workbook_) DefaultPivotTableStyle() ole.Variant {
	retVal, _ := this.PropGet(0x000009df, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Workbook_) SetDefaultPivotTableStyle(rhs interface{}) {
	_ = this.PropPut(0x000009df, []interface{}{rhs})
}

func (this *Workbook_) CheckCompatibility() bool {
	retVal, _ := this.PropGet(0x000009e0, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetCheckCompatibility(rhs bool) {
	_ = this.PropPut(0x000009e0, []interface{}{rhs})
}

func (this *Workbook_) HasVBProject() bool {
	retVal, _ := this.PropGet(0x000009e1, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) CustomXMLParts() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000009e2, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Workbook_) Final() bool {
	retVal, _ := this.PropGet(0x000009e3, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetFinal(rhs bool) {
	_ = this.PropPut(0x000009e3, []interface{}{rhs})
}

func (this *Workbook_) Research() *Research {
	retVal, _ := this.PropGet(0x000009e4, nil)
	return NewResearch(retVal.IDispatch(), false, true)
}

func (this *Workbook_) Theme() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000009e5, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Workbook_) ApplyTheme(filename string) {
	retVal, _ := this.Call(0x000009e6, []interface{}{filename})
	_ = retVal
}

func (this *Workbook_) Excel8CompatibilityMode() bool {
	retVal, _ := this.PropGet(0x000009e7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) ConnectionsDisabled() bool {
	retVal, _ := this.PropGet(0x000009e8, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) EnableConnections() {
	retVal, _ := this.Call(0x000009e9, nil)
	_ = retVal
}

func (this *Workbook_) ShowPivotChartActiveFields() bool {
	retVal, _ := this.PropGet(0x000009ea, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetShowPivotChartActiveFields(rhs bool) {
	_ = this.PropPut(0x000009ea, []interface{}{rhs})
}

var Workbook__ExportAsFixedFormat_OptArgs = []string{
	"Filename", "Quality", "IncludeDocProperties", "IgnorePrintAreas",
	"From", "To", "OpenAfterPublish", "FixedFormatExtClassPtr",
}

func (this *Workbook_) ExportAsFixedFormat(type_ int32, optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__ExportAsFixedFormat_OptArgs, optArgs)
	retVal, _ := this.Call(0x000009bd, []interface{}{type_}, optArgs...)
	_ = retVal
}

func (this *Workbook_) IconSets() *IconSets {
	retVal, _ := this.PropGet(0x000009eb, nil)
	return NewIconSets(retVal.IDispatch(), false, true)
}

func (this *Workbook_) EncryptionProvider() string {
	retVal, _ := this.PropGet(0x000009ec, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Workbook_) SetEncryptionProvider(rhs string) {
	_ = this.PropPut(0x000009ec, []interface{}{rhs})
}

func (this *Workbook_) DoNotPromptForConvert() bool {
	retVal, _ := this.PropGet(0x000009ed, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetDoNotPromptForConvert(rhs bool) {
	_ = this.PropPut(0x000009ed, []interface{}{rhs})
}

func (this *Workbook_) ForceFullCalculation() bool {
	retVal, _ := this.PropGet(0x000009ee, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Workbook_) SetForceFullCalculation(rhs bool) {
	_ = this.PropPut(0x000009ee, []interface{}{rhs})
}

var Workbook__ProtectSharing_OptArgs = []string{
	"Filename", "Password", "WriteResPassword", "ReadOnlyRecommended",
	"CreateBackup", "SharingPassword", "FileFormat",
}

func (this *Workbook_) ProtectSharing(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(Workbook__ProtectSharing_OptArgs, optArgs)
	retVal, _ := this.Call(0x000009ef, nil, optArgs...)
	_ = retVal
}

func (this *Workbook_) SlicerCaches() *SlicerCaches {
	retVal, _ := this.PropGet(0x00000b32, nil)
	return NewSlicerCaches(retVal.IDispatch(), false, true)
}

func (this *Workbook_) ActiveSlicer() *Slicer {
	retVal, _ := this.PropGet(0x00000b33, nil)
	return NewSlicer(retVal.IDispatch(), false, true)
}

func (this *Workbook_) DefaultSlicerStyle() ole.Variant {
	retVal, _ := this.PropGet(0x00000b34, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Workbook_) SetDefaultSlicerStyle(rhs interface{}) {
	_ = this.PropPut(0x00000b34, []interface{}{rhs})
}

func (this *Workbook_) Dummy26() {
	retVal, _ := this.Call(0x00000b35, nil)
	_ = retVal
}

func (this *Workbook_) Dummy27() {
	retVal, _ := this.Call(0x00000b36, nil)
	_ = retVal
}

func (this *Workbook_) AccuracyVersion() int32 {
	retVal, _ := this.PropGet(0x00000b37, nil)
	return retVal.LValVal()
}

func (this *Workbook_) SetAccuracyVersion(rhs int32) {
	_ = this.PropPut(0x00000b37, []interface{}{rhs})
}
