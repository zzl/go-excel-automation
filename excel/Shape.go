package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024439-0000-0000-C000-000000000046
var IID_Shape = syscall.GUID{0x00024439, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Shape struct {
	ole.OleClient
}

func NewShape(pDisp *win32.IDispatch, addRef bool, scoped bool) *Shape {
	p := &Shape{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ShapeFromVar(v ole.Variant) *Shape {
	return NewShape(v.PdispValVal(), false, false)
}

func (this *Shape) IID() *syscall.GUID {
	return &IID_Shape
}

func (this *Shape) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Shape) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Shape) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Shape) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Shape) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Shape) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Shape) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Shape) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Shape) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Shape) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Shape) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Shape) Apply()  {
	retVal := this.Call(0x0000068b, nil)
	_= retVal
}

func (this *Shape) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

func (this *Shape) Duplicate() *Shape {
	retVal := this.Call(0x0000040f, nil)
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *Shape) Flip(flipCmd int32)  {
	retVal := this.Call(0x0000068c, []interface{}{flipCmd})
	_= retVal
}

func (this *Shape) IncrementLeft(increment float32)  {
	retVal := this.Call(0x0000068e, []interface{}{increment})
	_= retVal
}

func (this *Shape) IncrementRotation(increment float32)  {
	retVal := this.Call(0x00000690, []interface{}{increment})
	_= retVal
}

func (this *Shape) IncrementTop(increment float32)  {
	retVal := this.Call(0x00000691, []interface{}{increment})
	_= retVal
}

func (this *Shape) PickUp()  {
	retVal := this.Call(0x00000692, nil)
	_= retVal
}

func (this *Shape) RerouteConnections()  {
	retVal := this.Call(0x00000693, nil)
	_= retVal
}

var Shape_ScaleHeight_OptArgs= []string{
	"Scale", 
}

func (this *Shape) ScaleHeight(factor float32, relativeToOriginalSize int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Shape_ScaleHeight_OptArgs, optArgs)
	retVal := this.Call(0x00000694, []interface{}{factor, relativeToOriginalSize}, optArgs...)
	_= retVal
}

var Shape_ScaleWidth_OptArgs= []string{
	"Scale", 
}

func (this *Shape) ScaleWidth(factor float32, relativeToOriginalSize int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Shape_ScaleWidth_OptArgs, optArgs)
	retVal := this.Call(0x00000698, []interface{}{factor, relativeToOriginalSize}, optArgs...)
	_= retVal
}

var Shape_Select_OptArgs= []string{
	"Replace", 
}

func (this *Shape) Select(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Shape_Select_OptArgs, optArgs)
	retVal := this.Call(0x000000eb, nil, optArgs...)
	_= retVal
}

func (this *Shape) SetShapesDefaultProperties()  {
	retVal := this.Call(0x00000699, nil)
	_= retVal
}

func (this *Shape) Ungroup() *ShapeRange {
	retVal := this.Call(0x000000f4, nil)
	return NewShapeRange(retVal.PdispValVal(), false, true)
}

func (this *Shape) ZOrder(zorderCmd int32)  {
	retVal := this.Call(0x0000026e, []interface{}{zorderCmd})
	_= retVal
}

func (this *Shape) Adjustments() *Adjustments {
	retVal := this.PropGet(0x0000069b, nil)
	return NewAdjustments(retVal.PdispValVal(), false, true)
}

func (this *Shape) TextFrame() *TextFrame {
	retVal := this.PropGet(0x0000069c, nil)
	return NewTextFrame(retVal.PdispValVal(), false, true)
}

func (this *Shape) AutoShapeType() int32 {
	retVal := this.PropGet(0x0000069d, nil)
	return retVal.LValVal()
}

func (this *Shape) SetAutoShapeType(rhs int32)  {
	retVal := this.PropPut(0x0000069d, []interface{}{rhs})
	_= retVal
}

func (this *Shape) Callout() *CalloutFormat {
	retVal := this.PropGet(0x0000069e, nil)
	return NewCalloutFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) ConnectionSiteCount() int32 {
	retVal := this.PropGet(0x0000069f, nil)
	return retVal.LValVal()
}

func (this *Shape) Connector() int32 {
	retVal := this.PropGet(0x000006a0, nil)
	return retVal.LValVal()
}

func (this *Shape) ConnectorFormat() *ConnectorFormat {
	retVal := this.PropGet(0x000006a1, nil)
	return NewConnectorFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) Fill() *FillFormat {
	retVal := this.PropGet(0x0000067f, nil)
	return NewFillFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) GroupItems() *GroupShapes {
	retVal := this.PropGet(0x000006a2, nil)
	return NewGroupShapes(retVal.PdispValVal(), false, true)
}

func (this *Shape) Height() float32 {
	retVal := this.PropGet(0x0000007b, nil)
	return retVal.FltValVal()
}

func (this *Shape) SetHeight(rhs float32)  {
	retVal := this.PropPut(0x0000007b, []interface{}{rhs})
	_= retVal
}

func (this *Shape) HorizontalFlip() int32 {
	retVal := this.PropGet(0x000006a3, nil)
	return retVal.LValVal()
}

func (this *Shape) Left() float32 {
	retVal := this.PropGet(0x0000007f, nil)
	return retVal.FltValVal()
}

func (this *Shape) SetLeft(rhs float32)  {
	retVal := this.PropPut(0x0000007f, []interface{}{rhs})
	_= retVal
}

func (this *Shape) Line() *LineFormat {
	retVal := this.PropGet(0x00000331, nil)
	return NewLineFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) LockAspectRatio() int32 {
	retVal := this.PropGet(0x000006a4, nil)
	return retVal.LValVal()
}

func (this *Shape) SetLockAspectRatio(rhs int32)  {
	retVal := this.PropPut(0x000006a4, []interface{}{rhs})
	_= retVal
}

func (this *Shape) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Shape) SetName(rhs string)  {
	retVal := this.PropPut(0x0000006e, []interface{}{rhs})
	_= retVal
}

func (this *Shape) Nodes() *ShapeNodes {
	retVal := this.PropGet(0x000006a5, nil)
	return NewShapeNodes(retVal.PdispValVal(), false, true)
}

func (this *Shape) Rotation() float32 {
	retVal := this.PropGet(0x0000003b, nil)
	return retVal.FltValVal()
}

func (this *Shape) SetRotation(rhs float32)  {
	retVal := this.PropPut(0x0000003b, []interface{}{rhs})
	_= retVal
}

func (this *Shape) PictureFormat() *PictureFormat {
	retVal := this.PropGet(0x0000065f, nil)
	return NewPictureFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) Shadow() *ShadowFormat {
	retVal := this.PropGet(0x00000067, nil)
	return NewShadowFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) TextEffect() *TextEffectFormat {
	retVal := this.PropGet(0x000006a6, nil)
	return NewTextEffectFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) ThreeD() *ThreeDFormat {
	retVal := this.PropGet(0x000006a7, nil)
	return NewThreeDFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) Top() float32 {
	retVal := this.PropGet(0x0000007e, nil)
	return retVal.FltValVal()
}

func (this *Shape) SetTop(rhs float32)  {
	retVal := this.PropPut(0x0000007e, []interface{}{rhs})
	_= retVal
}

func (this *Shape) Type() int32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *Shape) VerticalFlip() int32 {
	retVal := this.PropGet(0x000006a8, nil)
	return retVal.LValVal()
}

func (this *Shape) Vertices() ole.Variant {
	retVal := this.PropGet(0x0000026d, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Shape) Visible() int32 {
	retVal := this.PropGet(0x0000022e, nil)
	return retVal.LValVal()
}

func (this *Shape) SetVisible(rhs int32)  {
	retVal := this.PropPut(0x0000022e, []interface{}{rhs})
	_= retVal
}

func (this *Shape) Width() float32 {
	retVal := this.PropGet(0x0000007a, nil)
	return retVal.FltValVal()
}

func (this *Shape) SetWidth(rhs float32)  {
	retVal := this.PropPut(0x0000007a, []interface{}{rhs})
	_= retVal
}

func (this *Shape) ZOrderPosition() int32 {
	retVal := this.PropGet(0x000006a9, nil)
	return retVal.LValVal()
}

func (this *Shape) Hyperlink() *Hyperlink {
	retVal := this.PropGet(0x000006aa, nil)
	return NewHyperlink(retVal.PdispValVal(), false, true)
}

func (this *Shape) BlackWhiteMode() int32 {
	retVal := this.PropGet(0x000006ab, nil)
	return retVal.LValVal()
}

func (this *Shape) SetBlackWhiteMode(rhs int32)  {
	retVal := this.PropPut(0x000006ab, []interface{}{rhs})
	_= retVal
}

func (this *Shape) DrawingObject() *ole.DispatchClass {
	retVal := this.PropGet(0x000006ac, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Shape) OnAction() string {
	retVal := this.PropGet(0x00000254, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Shape) SetOnAction(rhs string)  {
	retVal := this.PropPut(0x00000254, []interface{}{rhs})
	_= retVal
}

func (this *Shape) Locked() bool {
	retVal := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Shape) SetLocked(rhs bool)  {
	retVal := this.PropPut(0x0000010d, []interface{}{rhs})
	_= retVal
}

func (this *Shape) TopLeftCell() *Range {
	retVal := this.PropGet(0x0000026c, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Shape) BottomRightCell() *Range {
	retVal := this.PropGet(0x00000267, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Shape) Placement() int32 {
	retVal := this.PropGet(0x00000269, nil)
	return retVal.LValVal()
}

func (this *Shape) SetPlacement(rhs int32)  {
	retVal := this.PropPut(0x00000269, []interface{}{rhs})
	_= retVal
}

func (this *Shape) Copy()  {
	retVal := this.Call(0x00000227, nil)
	_= retVal
}

func (this *Shape) Cut()  {
	retVal := this.Call(0x00000235, nil)
	_= retVal
}

var Shape_CopyPicture_OptArgs= []string{
	"Appearance", "Format", 
}

func (this *Shape) CopyPicture(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Shape_CopyPicture_OptArgs, optArgs)
	retVal := this.Call(0x000000d5, nil, optArgs...)
	_= retVal
}

func (this *Shape) ControlFormat() *ControlFormat {
	retVal := this.PropGet(0x000006ad, nil)
	return NewControlFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) LinkFormat() *LinkFormat {
	retVal := this.PropGet(0x000006ae, nil)
	return NewLinkFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) OLEFormat() *OLEFormat {
	retVal := this.PropGet(0x000006af, nil)
	return NewOLEFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) FormControlType() int32 {
	retVal := this.PropGet(0x000006b0, nil)
	return retVal.LValVal()
}

func (this *Shape) AlternativeText() string {
	retVal := this.PropGet(0x00000763, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Shape) SetAlternativeText(rhs string)  {
	retVal := this.PropPut(0x00000763, []interface{}{rhs})
	_= retVal
}

func (this *Shape) Script() *ole.DispatchClass {
	retVal := this.PropGet(0x00000764, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Shape) DiagramNode() *DiagramNode {
	retVal := this.PropGet(0x00000875, nil)
	return NewDiagramNode(retVal.PdispValVal(), false, true)
}

func (this *Shape) HasDiagramNode() int32 {
	retVal := this.PropGet(0x00000876, nil)
	return retVal.LValVal()
}

func (this *Shape) Diagram() *Diagram {
	retVal := this.PropGet(0x00000877, nil)
	return NewDiagram(retVal.PdispValVal(), false, true)
}

func (this *Shape) HasDiagram() int32 {
	retVal := this.PropGet(0x00000878, nil)
	return retVal.LValVal()
}

func (this *Shape) Child() int32 {
	retVal := this.PropGet(0x00000879, nil)
	return retVal.LValVal()
}

func (this *Shape) ParentGroup() *Shape {
	retVal := this.PropGet(0x0000087a, nil)
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *Shape) CanvasItems() *ole.DispatchClass {
	retVal := this.PropGet(0x0000087b, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Shape) ID() int32 {
	retVal := this.PropGet(0x0000023a, nil)
	return retVal.LValVal()
}

func (this *Shape) CanvasCropLeft(increment float32)  {
	retVal := this.Call(0x0000087c, []interface{}{increment})
	_= retVal
}

func (this *Shape) CanvasCropTop(increment float32)  {
	retVal := this.Call(0x0000087d, []interface{}{increment})
	_= retVal
}

func (this *Shape) CanvasCropRight(increment float32)  {
	retVal := this.Call(0x0000087e, []interface{}{increment})
	_= retVal
}

func (this *Shape) CanvasCropBottom(increment float32)  {
	retVal := this.Call(0x0000087f, []interface{}{increment})
	_= retVal
}

func (this *Shape) Chart() *Chart {
	retVal := this.PropGet(0x00000007, nil)
	return NewChart(retVal.PdispValVal(), false, true)
}

func (this *Shape) HasChart() int32 {
	retVal := this.PropGet(0x00000a62, nil)
	return retVal.LValVal()
}

func (this *Shape) TextFrame2() *TextFrame2 {
	retVal := this.PropGet(0x00000a63, nil)
	return NewTextFrame2(retVal.PdispValVal(), false, true)
}

func (this *Shape) ShapeStyle() int32 {
	retVal := this.PropGet(0x00000a64, nil)
	return retVal.LValVal()
}

func (this *Shape) SetShapeStyle(rhs int32)  {
	retVal := this.PropPut(0x00000a64, []interface{}{rhs})
	_= retVal
}

func (this *Shape) BackgroundStyle() int32 {
	retVal := this.PropGet(0x00000a65, nil)
	return retVal.LValVal()
}

func (this *Shape) SetBackgroundStyle(rhs int32)  {
	retVal := this.PropPut(0x00000a65, []interface{}{rhs})
	_= retVal
}

func (this *Shape) SoftEdge() *ole.DispatchClass {
	retVal := this.PropGet(0x00000a66, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Shape) Glow() *ole.DispatchClass {
	retVal := this.PropGet(0x00000a67, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Shape) Reflection() *ole.DispatchClass {
	retVal := this.PropGet(0x00000a68, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Shape) HasSmartArt() int32 {
	retVal := this.PropGet(0x00000b66, nil)
	return retVal.LValVal()
}

func (this *Shape) SmartArt() *ole.DispatchClass {
	retVal := this.PropGet(0x00000b67, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Shape) Title() string {
	retVal := this.PropGet(0x000000c7, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Shape) SetTitle(rhs string)  {
	retVal := this.PropPut(0x000000c7, []interface{}{rhs})
	_= retVal
}

