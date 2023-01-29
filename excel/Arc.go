package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 000208A0-0000-0000-C000-000000000046
var IID_Arc = syscall.GUID{0x000208A0, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Arc struct {
	ole.OleClient
}

func NewArc(pDisp *win32.IDispatch, addRef bool, scoped bool) *Arc {
	if pDisp == nil {
		return nil
	}
	p := &Arc{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ArcFromVar(v ole.Variant) *Arc {
	return NewArc(v.IDispatch(), false, false)
}

func (this *Arc) IID() *syscall.GUID {
	return &IID_Arc
}

func (this *Arc) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Arc) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *Arc) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Arc) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Arc) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *Arc) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *Arc) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *Arc) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *Arc) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Arc) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Arc) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Arc) BottomRightCell() *Range {
	retVal, _ := this.PropGet(0x00000267, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Arc) BringToFront() ole.Variant {
	retVal, _ := this.Call(0x0000025a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Arc) Copy() ole.Variant {
	retVal, _ := this.Call(0x00000227, nil)
	com.AddToScope(retVal)
	return *retVal
}

var Arc_CopyPicture_OptArgs = []string{
	"Appearance", "Format",
}

func (this *Arc) CopyPicture(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Arc_CopyPicture_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d5, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Arc) Cut() ole.Variant {
	retVal, _ := this.Call(0x00000235, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Arc) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Arc) Duplicate() *ole.DispatchClass {
	retVal, _ := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Arc) Enabled() bool {
	retVal, _ := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Arc) SetEnabled(rhs bool) {
	_ = this.PropPut(0x00000258, []interface{}{rhs})
}

func (this *Arc) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *Arc) SetHeight(rhs float64) {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *Arc) Index() int32 {
	retVal, _ := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *Arc) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *Arc) SetLeft(rhs float64) {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *Arc) Locked() bool {
	retVal, _ := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Arc) SetLocked(rhs bool) {
	_ = this.PropPut(0x0000010d, []interface{}{rhs})
}

func (this *Arc) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Arc) SetName(rhs string) {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *Arc) OnAction() string {
	retVal, _ := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Arc) SetOnAction(rhs string) {
	_ = this.PropPut(0x00000254, []interface{}{rhs})
}

func (this *Arc) Placement() ole.Variant {
	retVal, _ := this.PropGet(0x00000269, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Arc) SetPlacement(rhs interface{}) {
	_ = this.PropPut(0x00000269, []interface{}{rhs})
}

func (this *Arc) PrintObject() bool {
	retVal, _ := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Arc) SetPrintObject(rhs bool) {
	_ = this.PropPut(0x0000026a, []interface{}{rhs})
}

var Arc_Select_OptArgs = []string{
	"Replace",
}

func (this *Arc) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Arc_Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000eb, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Arc) SendToBack() ole.Variant {
	retVal, _ := this.Call(0x0000025d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Arc) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *Arc) SetTop(rhs float64) {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *Arc) TopLeftCell() *Range {
	retVal, _ := this.PropGet(0x0000026c, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Arc) Visible() bool {
	retVal, _ := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Arc) SetVisible(rhs bool) {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

func (this *Arc) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *Arc) SetWidth(rhs float64) {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *Arc) ZOrder() int32 {
	retVal, _ := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *Arc) ShapeRange() *ShapeRange {
	retVal, _ := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.IDispatch(), false, true)
}

func (this *Arc) AddIndent() bool {
	retVal, _ := this.PropGet(0x00000427, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Arc) SetAddIndent(rhs bool) {
	_ = this.PropPut(0x00000427, []interface{}{rhs})
}

func (this *Arc) AutoScaleFont() ole.Variant {
	retVal, _ := this.PropGet(0x000005f5, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Arc) SetAutoScaleFont(rhs interface{}) {
	_ = this.PropPut(0x000005f5, []interface{}{rhs})
}

func (this *Arc) AutoSize() bool {
	retVal, _ := this.PropGet(0x00000266, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Arc) SetAutoSize(rhs bool) {
	_ = this.PropPut(0x00000266, []interface{}{rhs})
}

func (this *Arc) Caption() string {
	retVal, _ := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Arc) SetCaption(rhs string) {
	_ = this.PropPut(0x0000008b, []interface{}{rhs})
}

var Arc_Characters_OptArgs = []string{
	"Start", "Length",
}

func (this *Arc) Characters(optArgs ...interface{}) *Characters {
	optArgs = ole.ProcessOptArgs(Arc_Characters_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000025b, nil, optArgs...)
	return NewCharacters(retVal.IDispatch(), false, true)
}

var Arc_CheckSpelling_OptArgs = []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang",
}

func (this *Arc) CheckSpelling(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Arc_CheckSpelling_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001f9, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Arc) Font() *Font {
	retVal, _ := this.PropGet(0x00000092, nil)
	return NewFont(retVal.IDispatch(), false, true)
}

func (this *Arc) Formula() string {
	retVal, _ := this.PropGet(0x00000105, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Arc) SetFormula(rhs string) {
	_ = this.PropPut(0x00000105, []interface{}{rhs})
}

func (this *Arc) HorizontalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x00000088, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Arc) SetHorizontalAlignment(rhs interface{}) {
	_ = this.PropPut(0x00000088, []interface{}{rhs})
}

func (this *Arc) LockedText() bool {
	retVal, _ := this.PropGet(0x00000268, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Arc) SetLockedText(rhs bool) {
	_ = this.PropPut(0x00000268, []interface{}{rhs})
}

func (this *Arc) Orientation() ole.Variant {
	retVal, _ := this.PropGet(0x00000086, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Arc) SetOrientation(rhs interface{}) {
	_ = this.PropPut(0x00000086, []interface{}{rhs})
}

func (this *Arc) Text() string {
	retVal, _ := this.PropGet(0x0000008a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Arc) SetText(rhs string) {
	_ = this.PropPut(0x0000008a, []interface{}{rhs})
}

func (this *Arc) VerticalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x00000089, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Arc) SetVerticalAlignment(rhs interface{}) {
	_ = this.PropPut(0x00000089, []interface{}{rhs})
}

func (this *Arc) ReadingOrder() int32 {
	retVal, _ := this.PropGet(0x000003cf, nil)
	return retVal.LValVal()
}

func (this *Arc) SetReadingOrder(rhs int32) {
	_ = this.PropPut(0x000003cf, []interface{}{rhs})
}

func (this *Arc) Border() *Border {
	retVal, _ := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.IDispatch(), false, true)
}

func (this *Arc) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *Arc) Dummy43_() {
	retVal, _ := this.Call(0x0001002b, nil)
	_ = retVal
}
