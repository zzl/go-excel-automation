package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 00020898-0000-0000-C000-000000000046
var IID_GroupObject = syscall.GUID{0x00020898, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type GroupObject struct {
	ole.OleClient
}

func NewGroupObject(pDisp *win32.IDispatch, addRef bool, scoped bool) *GroupObject {
	if pDisp == nil {
		return nil
	}
	p := &GroupObject{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func GroupObjectFromVar(v ole.Variant) *GroupObject {
	return NewGroupObject(v.IDispatch(), false, false)
}

func (this *GroupObject) IID() *syscall.GUID {
	return &IID_GroupObject
}

func (this *GroupObject) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *GroupObject) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_ = retVal
}

func (this *GroupObject) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *GroupObject) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *GroupObject) GetTypeInfoCount(pctinfo *uint32) {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_ = retVal
}

func (this *GroupObject) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_ = retVal
}

func (this *GroupObject) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_ = retVal
}

func (this *GroupObject) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_ = retVal
}

func (this *GroupObject) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *GroupObject) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *GroupObject) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *GroupObject) BottomRightCell() *Range {
	retVal, _ := this.PropGet(0x00000267, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *GroupObject) BringToFront() ole.Variant {
	retVal, _ := this.Call(0x0000025a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObject) Copy() ole.Variant {
	retVal, _ := this.Call(0x00000227, nil)
	com.AddToScope(retVal)
	return *retVal
}

var GroupObject_CopyPicture_OptArgs = []string{
	"Appearance", "Format",
}

func (this *GroupObject) CopyPicture(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(GroupObject_CopyPicture_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d5, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObject) Cut() ole.Variant {
	retVal, _ := this.Call(0x00000235, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObject) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObject) Duplicate() *ole.DispatchClass {
	retVal, _ := this.Call(0x0000040f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *GroupObject) Enabled() bool {
	retVal, _ := this.PropGet(0x00000258, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObject) SetEnabled(rhs bool) {
	_ = this.PropPut(0x00000258, []interface{}{rhs})
}

func (this *GroupObject) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *GroupObject) SetHeight(rhs float64) {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *GroupObject) Index() int32 {
	retVal, _ := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *GroupObject) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *GroupObject) SetLeft(rhs float64) {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *GroupObject) Locked() bool {
	retVal, _ := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObject) SetLocked(rhs bool) {
	_ = this.PropPut(0x0000010d, []interface{}{rhs})
}

func (this *GroupObject) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *GroupObject) SetName(rhs string) {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *GroupObject) OnAction() string {
	retVal, _ := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *GroupObject) SetOnAction(rhs string) {
	_ = this.PropPut(0x00000254, []interface{}{rhs})
}

func (this *GroupObject) Placement() ole.Variant {
	retVal, _ := this.PropGet(0x00000269, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObject) SetPlacement(rhs interface{}) {
	_ = this.PropPut(0x00000269, []interface{}{rhs})
}

func (this *GroupObject) PrintObject() bool {
	retVal, _ := this.PropGet(0x0000026a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObject) SetPrintObject(rhs bool) {
	_ = this.PropPut(0x0000026a, []interface{}{rhs})
}

var GroupObject_Select_OptArgs = []string{
	"Replace",
}

func (this *GroupObject) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(GroupObject_Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000eb, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObject) SendToBack() ole.Variant {
	retVal, _ := this.Call(0x0000025d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObject) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *GroupObject) SetTop(rhs float64) {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *GroupObject) TopLeftCell() *Range {
	retVal, _ := this.PropGet(0x0000026c, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *GroupObject) Visible() bool {
	retVal, _ := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObject) SetVisible(rhs bool) {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

func (this *GroupObject) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *GroupObject) SetWidth(rhs float64) {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *GroupObject) ZOrder() int32 {
	retVal, _ := this.PropGet(0x0000026e, nil)
	return retVal.LValVal()
}

func (this *GroupObject) ShapeRange() *ShapeRange {
	retVal, _ := this.PropGet(0x000005f8, nil)
	return NewShapeRange(retVal.IDispatch(), false, true)
}

func (this *GroupObject) Dummy27_() {
	retVal, _ := this.Call(0x0001001b, nil)
	_ = retVal
}

func (this *GroupObject) Dummy28_() {
	retVal, _ := this.Call(0x0001001c, nil)
	_ = retVal
}

func (this *GroupObject) AddIndent() bool {
	retVal, _ := this.PropGet(0x00000427, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObject) SetAddIndent(rhs bool) {
	_ = this.PropPut(0x00000427, []interface{}{rhs})
}

func (this *GroupObject) Dummy30_() {
	retVal, _ := this.Call(0x0001001e, nil)
	_ = retVal
}

func (this *GroupObject) ArrowHeadLength() ole.Variant {
	retVal, _ := this.PropGet(0x00000263, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObject) SetArrowHeadLength(rhs interface{}) {
	_ = this.PropPut(0x00000263, []interface{}{rhs})
}

func (this *GroupObject) ArrowHeadStyle() ole.Variant {
	retVal, _ := this.PropGet(0x00000264, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObject) SetArrowHeadStyle(rhs interface{}) {
	_ = this.PropPut(0x00000264, []interface{}{rhs})
}

func (this *GroupObject) ArrowHeadWidth() ole.Variant {
	retVal, _ := this.PropGet(0x00000265, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObject) SetArrowHeadWidth(rhs interface{}) {
	_ = this.PropPut(0x00000265, []interface{}{rhs})
}

func (this *GroupObject) AutoSize() bool {
	retVal, _ := this.PropGet(0x00000266, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObject) SetAutoSize(rhs bool) {
	_ = this.PropPut(0x00000266, []interface{}{rhs})
}

func (this *GroupObject) Border() *Border {
	retVal, _ := this.PropGet(0x00000080, nil)
	return NewBorder(retVal.IDispatch(), false, true)
}

func (this *GroupObject) Dummy36_() {
	retVal, _ := this.Call(0x00010024, nil)
	_ = retVal
}

func (this *GroupObject) Dummy37_() {
	retVal, _ := this.Call(0x00010025, nil)
	_ = retVal
}

func (this *GroupObject) Dummy38_() {
	retVal, _ := this.Call(0x00010026, nil)
	_ = retVal
}

var GroupObject_CheckSpelling_OptArgs = []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "SpellLang",
}

func (this *GroupObject) CheckSpelling(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(GroupObject_CheckSpelling_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001f9, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObject) Default_() int32 {
	retVal, _ := this.PropGet(0x00000000, nil)
	return retVal.LValVal()
}

func (this *GroupObject) SetDefault_(rhs int32) {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *GroupObject) Dummy41_() {
	retVal, _ := this.Call(0x00010029, nil)
	_ = retVal
}

func (this *GroupObject) Dummy42_() {
	retVal, _ := this.Call(0x0001002a, nil)
	_ = retVal
}

func (this *GroupObject) Dummy43_() {
	retVal, _ := this.Call(0x0001002b, nil)
	_ = retVal
}

func (this *GroupObject) Dummy44_() {
	retVal, _ := this.Call(0x0001002c, nil)
	_ = retVal
}

func (this *GroupObject) Dummy45_() {
	retVal, _ := this.Call(0x0001002d, nil)
	_ = retVal
}

func (this *GroupObject) Font() *Font {
	retVal, _ := this.PropGet(0x00000092, nil)
	return NewFont(retVal.IDispatch(), false, true)
}

func (this *GroupObject) Dummy47_() {
	retVal, _ := this.Call(0x0001002f, nil)
	_ = retVal
}

func (this *GroupObject) Dummy48_() {
	retVal, _ := this.Call(0x00010030, nil)
	_ = retVal
}

func (this *GroupObject) HorizontalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x00000088, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObject) SetHorizontalAlignment(rhs interface{}) {
	_ = this.PropPut(0x00000088, []interface{}{rhs})
}

func (this *GroupObject) Dummy50_() {
	retVal, _ := this.Call(0x00010032, nil)
	_ = retVal
}

func (this *GroupObject) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *GroupObject) Dummy52_() {
	retVal, _ := this.Call(0x00010034, nil)
	_ = retVal
}

func (this *GroupObject) Dummy53_() {
	retVal, _ := this.Call(0x00010035, nil)
	_ = retVal
}

func (this *GroupObject) Dummy54_() {
	retVal, _ := this.Call(0x00010036, nil)
	_ = retVal
}

func (this *GroupObject) Dummy55_() {
	retVal, _ := this.Call(0x00010037, nil)
	_ = retVal
}

func (this *GroupObject) Dummy56_() {
	retVal, _ := this.Call(0x00010038, nil)
	_ = retVal
}

func (this *GroupObject) Dummy57_() {
	retVal, _ := this.Call(0x00010039, nil)
	_ = retVal
}

func (this *GroupObject) Dummy58_() {
	retVal, _ := this.Call(0x0001003a, nil)
	_ = retVal
}

func (this *GroupObject) Dummy59_() {
	retVal, _ := this.Call(0x0001003b, nil)
	_ = retVal
}

func (this *GroupObject) Dummy60_() {
	retVal, _ := this.Call(0x0001003c, nil)
	_ = retVal
}

func (this *GroupObject) Dummy61_() {
	retVal, _ := this.Call(0x0001003d, nil)
	_ = retVal
}

func (this *GroupObject) Dummy62_() {
	retVal, _ := this.Call(0x0001003e, nil)
	_ = retVal
}

func (this *GroupObject) Dummy63_() {
	retVal, _ := this.Call(0x0001003f, nil)
	_ = retVal
}

func (this *GroupObject) Orientation() ole.Variant {
	retVal, _ := this.PropGet(0x00000086, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObject) SetOrientation(rhs interface{}) {
	_ = this.PropPut(0x00000086, []interface{}{rhs})
}

func (this *GroupObject) Dummy65_() {
	retVal, _ := this.Call(0x00010041, nil)
	_ = retVal
}

func (this *GroupObject) Dummy66_() {
	retVal, _ := this.Call(0x00010042, nil)
	_ = retVal
}

func (this *GroupObject) Dummy67_() {
	retVal, _ := this.Call(0x00010043, nil)
	_ = retVal
}

func (this *GroupObject) Dummy68_() {
	retVal, _ := this.Call(0x00010044, nil)
	_ = retVal
}

func (this *GroupObject) RoundedCorners() bool {
	retVal, _ := this.PropGet(0x0000026b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObject) SetRoundedCorners(rhs bool) {
	_ = this.PropPut(0x0000026b, []interface{}{rhs})
}

func (this *GroupObject) Dummy70_() {
	retVal, _ := this.Call(0x00010046, nil)
	_ = retVal
}

func (this *GroupObject) Shadow() bool {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *GroupObject) SetShadow(rhs bool) {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *GroupObject) Dummy72_() {
	retVal, _ := this.Call(0x00010048, nil)
	_ = retVal
}

func (this *GroupObject) Dummy73_() {
	retVal, _ := this.Call(0x00010049, nil)
	_ = retVal
}

func (this *GroupObject) Ungroup() *ole.DispatchClass {
	retVal, _ := this.Call(0x000000f4, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *GroupObject) Dummy75_() {
	retVal, _ := this.Call(0x0001004b, nil)
	_ = retVal
}

func (this *GroupObject) VerticalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x00000089, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *GroupObject) SetVerticalAlignment(rhs interface{}) {
	_ = this.PropPut(0x00000089, []interface{}{rhs})
}

func (this *GroupObject) Dummy77_() {
	retVal, _ := this.Call(0x0001004d, nil)
	_ = retVal
}

func (this *GroupObject) ReadingOrder() int32 {
	retVal, _ := this.PropGet(0x000003cf, nil)
	return retVal.LValVal()
}

func (this *GroupObject) SetReadingOrder(rhs int32) {
	_ = this.PropPut(0x000003cf, []interface{}{rhs})
}
