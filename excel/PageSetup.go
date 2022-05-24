package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000208B4-0000-0000-C000-000000000046
var IID_PageSetup = syscall.GUID{0x000208B4, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type PageSetup struct {
	ole.OleClient
}

func NewPageSetup(pDisp *win32.IDispatch, addRef bool, scoped bool) *PageSetup {
	 if pDisp == nil {
		return nil;
	}
	p := &PageSetup{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PageSetupFromVar(v ole.Variant) *PageSetup {
	return NewPageSetup(v.IDispatch(), false, false)
}

func (this *PageSetup) IID() *syscall.GUID {
	return &IID_PageSetup
}

func (this *PageSetup) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *PageSetup) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *PageSetup) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *PageSetup) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *PageSetup) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *PageSetup) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *PageSetup) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *PageSetup) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *PageSetup) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *PageSetup) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *PageSetup) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *PageSetup) BlackAndWhite() bool {
	retVal, _ := this.PropGet(0x000003f1, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PageSetup) SetBlackAndWhite(rhs bool)  {
	_ = this.PropPut(0x000003f1, []interface{}{rhs})
}

func (this *PageSetup) BottomMargin() float64 {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return retVal.DblValVal()
}

func (this *PageSetup) SetBottomMargin(rhs float64)  {
	_ = this.PropPut(0x000003ea, []interface{}{rhs})
}

func (this *PageSetup) CenterFooter() string {
	retVal, _ := this.PropGet(0x000003f2, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PageSetup) SetCenterFooter(rhs string)  {
	_ = this.PropPut(0x000003f2, []interface{}{rhs})
}

func (this *PageSetup) CenterHeader() string {
	retVal, _ := this.PropGet(0x000003f3, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PageSetup) SetCenterHeader(rhs string)  {
	_ = this.PropPut(0x000003f3, []interface{}{rhs})
}

func (this *PageSetup) CenterHorizontally() bool {
	retVal, _ := this.PropGet(0x000003ed, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PageSetup) SetCenterHorizontally(rhs bool)  {
	_ = this.PropPut(0x000003ed, []interface{}{rhs})
}

func (this *PageSetup) CenterVertically() bool {
	retVal, _ := this.PropGet(0x000003ee, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PageSetup) SetCenterVertically(rhs bool)  {
	_ = this.PropPut(0x000003ee, []interface{}{rhs})
}

func (this *PageSetup) ChartSize() int32 {
	retVal, _ := this.PropGet(0x000003f4, nil)
	return retVal.LValVal()
}

func (this *PageSetup) SetChartSize(rhs int32)  {
	_ = this.PropPut(0x000003f4, []interface{}{rhs})
}

func (this *PageSetup) Draft() bool {
	retVal, _ := this.PropGet(0x000003fc, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PageSetup) SetDraft(rhs bool)  {
	_ = this.PropPut(0x000003fc, []interface{}{rhs})
}

func (this *PageSetup) FirstPageNumber() int32 {
	retVal, _ := this.PropGet(0x000003f0, nil)
	return retVal.LValVal()
}

func (this *PageSetup) SetFirstPageNumber(rhs int32)  {
	_ = this.PropPut(0x000003f0, []interface{}{rhs})
}

func (this *PageSetup) FitToPagesTall() ole.Variant {
	retVal, _ := this.PropGet(0x000003f5, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PageSetup) SetFitToPagesTall(rhs interface{})  {
	_ = this.PropPut(0x000003f5, []interface{}{rhs})
}

func (this *PageSetup) FitToPagesWide() ole.Variant {
	retVal, _ := this.PropGet(0x000003f6, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PageSetup) SetFitToPagesWide(rhs interface{})  {
	_ = this.PropPut(0x000003f6, []interface{}{rhs})
}

func (this *PageSetup) FooterMargin() float64 {
	retVal, _ := this.PropGet(0x000003f7, nil)
	return retVal.DblValVal()
}

func (this *PageSetup) SetFooterMargin(rhs float64)  {
	_ = this.PropPut(0x000003f7, []interface{}{rhs})
}

func (this *PageSetup) HeaderMargin() float64 {
	retVal, _ := this.PropGet(0x000003f8, nil)
	return retVal.DblValVal()
}

func (this *PageSetup) SetHeaderMargin(rhs float64)  {
	_ = this.PropPut(0x000003f8, []interface{}{rhs})
}

func (this *PageSetup) LeftFooter() string {
	retVal, _ := this.PropGet(0x000003f9, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PageSetup) SetLeftFooter(rhs string)  {
	_ = this.PropPut(0x000003f9, []interface{}{rhs})
}

func (this *PageSetup) LeftHeader() string {
	retVal, _ := this.PropGet(0x000003fa, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PageSetup) SetLeftHeader(rhs string)  {
	_ = this.PropPut(0x000003fa, []interface{}{rhs})
}

func (this *PageSetup) LeftMargin() float64 {
	retVal, _ := this.PropGet(0x000003e7, nil)
	return retVal.DblValVal()
}

func (this *PageSetup) SetLeftMargin(rhs float64)  {
	_ = this.PropPut(0x000003e7, []interface{}{rhs})
}

func (this *PageSetup) Order() int32 {
	retVal, _ := this.PropGet(0x000000c0, nil)
	return retVal.LValVal()
}

func (this *PageSetup) SetOrder(rhs int32)  {
	_ = this.PropPut(0x000000c0, []interface{}{rhs})
}

func (this *PageSetup) Orientation() int32 {
	retVal, _ := this.PropGet(0x00000086, nil)
	return retVal.LValVal()
}

func (this *PageSetup) SetOrientation(rhs int32)  {
	_ = this.PropPut(0x00000086, []interface{}{rhs})
}

func (this *PageSetup) PaperSize() int32 {
	retVal, _ := this.PropGet(0x000003ef, nil)
	return retVal.LValVal()
}

func (this *PageSetup) SetPaperSize(rhs int32)  {
	_ = this.PropPut(0x000003ef, []interface{}{rhs})
}

func (this *PageSetup) PrintArea() string {
	retVal, _ := this.PropGet(0x000003fb, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PageSetup) SetPrintArea(rhs string)  {
	_ = this.PropPut(0x000003fb, []interface{}{rhs})
}

func (this *PageSetup) PrintGridlines() bool {
	retVal, _ := this.PropGet(0x000003ec, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PageSetup) SetPrintGridlines(rhs bool)  {
	_ = this.PropPut(0x000003ec, []interface{}{rhs})
}

func (this *PageSetup) PrintHeadings() bool {
	retVal, _ := this.PropGet(0x000003eb, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PageSetup) SetPrintHeadings(rhs bool)  {
	_ = this.PropPut(0x000003eb, []interface{}{rhs})
}

func (this *PageSetup) PrintNotes() bool {
	retVal, _ := this.PropGet(0x000003fd, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PageSetup) SetPrintNotes(rhs bool)  {
	_ = this.PropPut(0x000003fd, []interface{}{rhs})
}

var PageSetup_PrintQuality_OptArgs= []string{
	"Index", 
}

func (this *PageSetup) PrintQuality(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(PageSetup_PrintQuality_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x000003fe, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var PageSetup_SetPrintQuality_OptArgs= []string{
	"Index", 
}

func (this *PageSetup) SetPrintQuality(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(PageSetup_SetPrintQuality_OptArgs, optArgs)
	_ = this.PropPut(0x000003fe, nil, optArgs...)
}

func (this *PageSetup) PrintTitleColumns() string {
	retVal, _ := this.PropGet(0x000003ff, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PageSetup) SetPrintTitleColumns(rhs string)  {
	_ = this.PropPut(0x000003ff, []interface{}{rhs})
}

func (this *PageSetup) PrintTitleRows() string {
	retVal, _ := this.PropGet(0x00000400, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PageSetup) SetPrintTitleRows(rhs string)  {
	_ = this.PropPut(0x00000400, []interface{}{rhs})
}

func (this *PageSetup) RightFooter() string {
	retVal, _ := this.PropGet(0x00000401, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PageSetup) SetRightFooter(rhs string)  {
	_ = this.PropPut(0x00000401, []interface{}{rhs})
}

func (this *PageSetup) RightHeader() string {
	retVal, _ := this.PropGet(0x00000402, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PageSetup) SetRightHeader(rhs string)  {
	_ = this.PropPut(0x00000402, []interface{}{rhs})
}

func (this *PageSetup) RightMargin() float64 {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return retVal.DblValVal()
}

func (this *PageSetup) SetRightMargin(rhs float64)  {
	_ = this.PropPut(0x000003e8, []interface{}{rhs})
}

func (this *PageSetup) TopMargin() float64 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.DblValVal()
}

func (this *PageSetup) SetTopMargin(rhs float64)  {
	_ = this.PropPut(0x000003e9, []interface{}{rhs})
}

func (this *PageSetup) Zoom() ole.Variant {
	retVal, _ := this.PropGet(0x00000297, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PageSetup) SetZoom(rhs interface{})  {
	_ = this.PropPut(0x00000297, []interface{}{rhs})
}

func (this *PageSetup) PrintComments() int32 {
	retVal, _ := this.PropGet(0x000005f4, nil)
	return retVal.LValVal()
}

func (this *PageSetup) SetPrintComments(rhs int32)  {
	_ = this.PropPut(0x000005f4, []interface{}{rhs})
}

func (this *PageSetup) PrintErrors() int32 {
	retVal, _ := this.PropGet(0x00000865, nil)
	return retVal.LValVal()
}

func (this *PageSetup) SetPrintErrors(rhs int32)  {
	_ = this.PropPut(0x00000865, []interface{}{rhs})
}

func (this *PageSetup) CenterHeaderPicture() *Graphic {
	retVal, _ := this.PropGet(0x00000866, nil)
	return NewGraphic(retVal.IDispatch(), false, true)
}

func (this *PageSetup) CenterFooterPicture() *Graphic {
	retVal, _ := this.PropGet(0x00000867, nil)
	return NewGraphic(retVal.IDispatch(), false, true)
}

func (this *PageSetup) LeftHeaderPicture() *Graphic {
	retVal, _ := this.PropGet(0x00000868, nil)
	return NewGraphic(retVal.IDispatch(), false, true)
}

func (this *PageSetup) LeftFooterPicture() *Graphic {
	retVal, _ := this.PropGet(0x00000869, nil)
	return NewGraphic(retVal.IDispatch(), false, true)
}

func (this *PageSetup) RightHeaderPicture() *Graphic {
	retVal, _ := this.PropGet(0x0000086a, nil)
	return NewGraphic(retVal.IDispatch(), false, true)
}

func (this *PageSetup) RightFooterPicture() *Graphic {
	retVal, _ := this.PropGet(0x0000086b, nil)
	return NewGraphic(retVal.IDispatch(), false, true)
}

func (this *PageSetup) OddAndEvenPagesHeaderFooter() bool {
	retVal, _ := this.PropGet(0x00000a28, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PageSetup) SetOddAndEvenPagesHeaderFooter(rhs bool)  {
	_ = this.PropPut(0x00000a28, []interface{}{rhs})
}

func (this *PageSetup) DifferentFirstPageHeaderFooter() bool {
	retVal, _ := this.PropGet(0x00000a29, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PageSetup) SetDifferentFirstPageHeaderFooter(rhs bool)  {
	_ = this.PropPut(0x00000a29, []interface{}{rhs})
}

func (this *PageSetup) ScaleWithDocHeaderFooter() bool {
	retVal, _ := this.PropGet(0x00000a2a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PageSetup) SetScaleWithDocHeaderFooter(rhs bool)  {
	_ = this.PropPut(0x00000a2a, []interface{}{rhs})
}

func (this *PageSetup) AlignMarginsHeaderFooter() bool {
	retVal, _ := this.PropGet(0x00000a2b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PageSetup) SetAlignMarginsHeaderFooter(rhs bool)  {
	_ = this.PropPut(0x00000a2b, []interface{}{rhs})
}

func (this *PageSetup) Pages() *Pages {
	retVal, _ := this.PropGet(0x00000a2c, nil)
	return NewPages(retVal.IDispatch(), false, true)
}

func (this *PageSetup) EvenPage() *Page {
	retVal, _ := this.PropGet(0x00000a2d, nil)
	return NewPage(retVal.IDispatch(), false, true)
}

func (this *PageSetup) FirstPage() *Page {
	retVal, _ := this.PropGet(0x00000a2e, nil)
	return NewPage(retVal.IDispatch(), false, true)
}

