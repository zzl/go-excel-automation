package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002443B-0000-0000-C000-000000000046
var IID_ShapeRange = syscall.GUID{0x0002443B, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ShapeRange struct {
	ole.OleClient
}

func NewShapeRange(pDisp *win32.IDispatch, addRef bool, scoped bool) *ShapeRange {
	p := &ShapeRange{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ShapeRangeFromVar(v ole.Variant) *ShapeRange {
	return NewShapeRange(v.PdispValVal(), false, false)
}

func (this *ShapeRange) IID() *syscall.GUID {
	return &IID_ShapeRange
}

func (this *ShapeRange) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ShapeRange) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *ShapeRange) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *ShapeRange) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *ShapeRange) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *ShapeRange) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *ShapeRange) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *ShapeRange) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *ShapeRange) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *ShapeRange) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ShapeRange) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) Item(index interface{}) *Shape {
	retVal := this.Call(0x000000aa, []interface{}{index})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *ShapeRange) Default_(index interface{}) *Shape {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *ShapeRange) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *ShapeRange) ForEach(action func(item *Shape) bool) {
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
		pItem := (*Shape)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *ShapeRange) Align(alignCmd int32, relativeTo int32)  {
	retVal := this.Call(0x000006cc, []interface{}{alignCmd, relativeTo})
	_= retVal
}

func (this *ShapeRange) Apply()  {
	retVal := this.Call(0x0000068b, nil)
	_= retVal
}

func (this *ShapeRange) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

func (this *ShapeRange) Distribute(distributeCmd int32, relativeTo int32)  {
	retVal := this.Call(0x000006ce, []interface{}{distributeCmd, relativeTo})
	_= retVal
}

func (this *ShapeRange) Duplicate() *ShapeRange {
	retVal := this.Call(0x0000040f, nil)
	return NewShapeRange(retVal.PdispValVal(), false, true)
}

func (this *ShapeRange) Flip(flipCmd int32)  {
	retVal := this.Call(0x0000068c, []interface{}{flipCmd})
	_= retVal
}

func (this *ShapeRange) IncrementLeft(increment float32)  {
	retVal := this.Call(0x0000068e, []interface{}{increment})
	_= retVal
}

func (this *ShapeRange) IncrementRotation(increment float32)  {
	retVal := this.Call(0x00000690, []interface{}{increment})
	_= retVal
}

func (this *ShapeRange) IncrementTop(increment float32)  {
	retVal := this.Call(0x00000691, []interface{}{increment})
	_= retVal
}

func (this *ShapeRange) Group() *Shape {
	retVal := this.Call(0x0000002e, nil)
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *ShapeRange) PickUp()  {
	retVal := this.Call(0x00000692, nil)
	_= retVal
}

func (this *ShapeRange) RerouteConnections()  {
	retVal := this.Call(0x00000693, nil)
	_= retVal
}

func (this *ShapeRange) Regroup() *Shape {
	retVal := this.Call(0x000006d0, nil)
	return NewShape(retVal.PdispValVal(), false, true)
}

var ShapeRange_ScaleHeight_OptArgs= []string{
	"Scale", 
}

func (this *ShapeRange) ScaleHeight(factor float32, relativeToOriginalSize int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ShapeRange_ScaleHeight_OptArgs, optArgs)
	retVal := this.Call(0x00000694, []interface{}{factor, relativeToOriginalSize}, optArgs...)
	_= retVal
}

var ShapeRange_ScaleWidth_OptArgs= []string{
	"Scale", 
}

func (this *ShapeRange) ScaleWidth(factor float32, relativeToOriginalSize int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ShapeRange_ScaleWidth_OptArgs, optArgs)
	retVal := this.Call(0x00000698, []interface{}{factor, relativeToOriginalSize}, optArgs...)
	_= retVal
}

var ShapeRange_Select_OptArgs= []string{
	"Replace", 
}

func (this *ShapeRange) Select(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ShapeRange_Select_OptArgs, optArgs)
	retVal := this.Call(0x000000eb, nil, optArgs...)
	_= retVal
}

func (this *ShapeRange) SetShapesDefaultProperties()  {
	retVal := this.Call(0x00000699, nil)
	_= retVal
}

func (this *ShapeRange) Ungroup() *ShapeRange {
	retVal := this.Call(0x000000f4, nil)
	return NewShapeRange(retVal.PdispValVal(), false, true)
}

func (this *ShapeRange) ZOrder(zorderCmd int32)  {
	retVal := this.Call(0x0000026e, []interface{}{zorderCmd})
	_= retVal
}

func (this *ShapeRange) Adjustments() *Adjustments {
	retVal := this.PropGet(0x0000069b, nil)
	return NewAdjustments(retVal.PdispValVal(), false, true)
}

func (this *ShapeRange) TextFrame() *TextFrame {
	retVal := this.PropGet(0x0000069c, nil)
	return NewTextFrame(retVal.PdispValVal(), false, true)
}

func (this *ShapeRange) AutoShapeType() int32 {
	retVal := this.PropGet(0x0000069d, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) SetAutoShapeType(rhs int32)  {
	retVal := this.PropPut(0x0000069d, []interface{}{rhs})
	_= retVal
}

func (this *ShapeRange) Callout() *CalloutFormat {
	retVal := this.PropGet(0x0000069e, nil)
	return NewCalloutFormat(retVal.PdispValVal(), false, true)
}

func (this *ShapeRange) ConnectionSiteCount() int32 {
	retVal := this.PropGet(0x0000069f, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) Connector() int32 {
	retVal := this.PropGet(0x000006a0, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) ConnectorFormat() *ConnectorFormat {
	retVal := this.PropGet(0x000006a1, nil)
	return NewConnectorFormat(retVal.PdispValVal(), false, true)
}

func (this *ShapeRange) Fill() *FillFormat {
	retVal := this.PropGet(0x0000067f, nil)
	return NewFillFormat(retVal.PdispValVal(), false, true)
}

func (this *ShapeRange) GroupItems() *GroupShapes {
	retVal := this.PropGet(0x000006a2, nil)
	return NewGroupShapes(retVal.PdispValVal(), false, true)
}

func (this *ShapeRange) Height() float32 {
	retVal := this.PropGet(0x0000007b, nil)
	return retVal.FltValVal()
}

func (this *ShapeRange) SetHeight(rhs float32)  {
	retVal := this.PropPut(0x0000007b, []interface{}{rhs})
	_= retVal
}

func (this *ShapeRange) HorizontalFlip() int32 {
	retVal := this.PropGet(0x000006a3, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) Left() float32 {
	retVal := this.PropGet(0x0000007f, nil)
	return retVal.FltValVal()
}

func (this *ShapeRange) SetLeft(rhs float32)  {
	retVal := this.PropPut(0x0000007f, []interface{}{rhs})
	_= retVal
}

func (this *ShapeRange) Line() *LineFormat {
	retVal := this.PropGet(0x00000331, nil)
	return NewLineFormat(retVal.PdispValVal(), false, true)
}

func (this *ShapeRange) LockAspectRatio() int32 {
	retVal := this.PropGet(0x000006a4, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) SetLockAspectRatio(rhs int32)  {
	retVal := this.PropPut(0x000006a4, []interface{}{rhs})
	_= retVal
}

func (this *ShapeRange) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ShapeRange) SetName(rhs string)  {
	retVal := this.PropPut(0x0000006e, []interface{}{rhs})
	_= retVal
}

func (this *ShapeRange) Nodes() *ShapeNodes {
	retVal := this.PropGet(0x000006a5, nil)
	return NewShapeNodes(retVal.PdispValVal(), false, true)
}

func (this *ShapeRange) Rotation() float32 {
	retVal := this.PropGet(0x0000003b, nil)
	return retVal.FltValVal()
}

func (this *ShapeRange) SetRotation(rhs float32)  {
	retVal := this.PropPut(0x0000003b, []interface{}{rhs})
	_= retVal
}

func (this *ShapeRange) PictureFormat() *PictureFormat {
	retVal := this.PropGet(0x0000065f, nil)
	return NewPictureFormat(retVal.PdispValVal(), false, true)
}

func (this *ShapeRange) Shadow() *ShadowFormat {
	retVal := this.PropGet(0x00000067, nil)
	return NewShadowFormat(retVal.PdispValVal(), false, true)
}

func (this *ShapeRange) TextEffect() *TextEffectFormat {
	retVal := this.PropGet(0x000006a6, nil)
	return NewTextEffectFormat(retVal.PdispValVal(), false, true)
}

func (this *ShapeRange) ThreeD() *ThreeDFormat {
	retVal := this.PropGet(0x000006a7, nil)
	return NewThreeDFormat(retVal.PdispValVal(), false, true)
}

func (this *ShapeRange) Top() float32 {
	retVal := this.PropGet(0x0000007e, nil)
	return retVal.FltValVal()
}

func (this *ShapeRange) SetTop(rhs float32)  {
	retVal := this.PropPut(0x0000007e, []interface{}{rhs})
	_= retVal
}

func (this *ShapeRange) Type() int32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) VerticalFlip() int32 {
	retVal := this.PropGet(0x000006a8, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) Vertices() ole.Variant {
	retVal := this.PropGet(0x0000026d, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ShapeRange) Visible() int32 {
	retVal := this.PropGet(0x0000022e, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) SetVisible(rhs int32)  {
	retVal := this.PropPut(0x0000022e, []interface{}{rhs})
	_= retVal
}

func (this *ShapeRange) Width() float32 {
	retVal := this.PropGet(0x0000007a, nil)
	return retVal.FltValVal()
}

func (this *ShapeRange) SetWidth(rhs float32)  {
	retVal := this.PropPut(0x0000007a, []interface{}{rhs})
	_= retVal
}

func (this *ShapeRange) ZOrderPosition() int32 {
	retVal := this.PropGet(0x000006a9, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) BlackWhiteMode() int32 {
	retVal := this.PropGet(0x000006ab, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) SetBlackWhiteMode(rhs int32)  {
	retVal := this.PropPut(0x000006ab, []interface{}{rhs})
	_= retVal
}

func (this *ShapeRange) AlternativeText() string {
	retVal := this.PropGet(0x00000763, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ShapeRange) SetAlternativeText(rhs string)  {
	retVal := this.PropPut(0x00000763, []interface{}{rhs})
	_= retVal
}

func (this *ShapeRange) DiagramNode() *DiagramNode {
	retVal := this.PropGet(0x00000875, nil)
	return NewDiagramNode(retVal.PdispValVal(), false, true)
}

func (this *ShapeRange) HasDiagramNode() int32 {
	retVal := this.PropGet(0x00000876, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) Diagram() *Diagram {
	retVal := this.PropGet(0x00000877, nil)
	return NewDiagram(retVal.PdispValVal(), false, true)
}

func (this *ShapeRange) HasDiagram() int32 {
	retVal := this.PropGet(0x00000878, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) Child() int32 {
	retVal := this.PropGet(0x00000879, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) ParentGroup() *Shape {
	retVal := this.PropGet(0x0000087a, nil)
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *ShapeRange) CanvasItems() *ole.DispatchClass {
	retVal := this.PropGet(0x0000087b, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ShapeRange) ID() int32 {
	retVal := this.PropGet(0x0000023a, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) CanvasCropLeft(increment float32)  {
	retVal := this.Call(0x0000087c, []interface{}{increment})
	_= retVal
}

func (this *ShapeRange) CanvasCropTop(increment float32)  {
	retVal := this.Call(0x0000087d, []interface{}{increment})
	_= retVal
}

func (this *ShapeRange) CanvasCropRight(increment float32)  {
	retVal := this.Call(0x0000087e, []interface{}{increment})
	_= retVal
}

func (this *ShapeRange) CanvasCropBottom(increment float32)  {
	retVal := this.Call(0x0000087f, []interface{}{increment})
	_= retVal
}

func (this *ShapeRange) Chart() *Chart {
	retVal := this.PropGet(0x00000007, nil)
	return NewChart(retVal.PdispValVal(), false, true)
}

func (this *ShapeRange) HasChart() int32 {
	retVal := this.PropGet(0x00000a62, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) TextFrame2() *TextFrame2 {
	retVal := this.PropGet(0x00000a63, nil)
	return NewTextFrame2(retVal.PdispValVal(), false, true)
}

func (this *ShapeRange) ShapeStyle() int32 {
	retVal := this.PropGet(0x00000a64, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) SetShapeStyle(rhs int32)  {
	retVal := this.PropPut(0x00000a64, []interface{}{rhs})
	_= retVal
}

func (this *ShapeRange) BackgroundStyle() int32 {
	retVal := this.PropGet(0x00000a65, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) SetBackgroundStyle(rhs int32)  {
	retVal := this.PropPut(0x00000a65, []interface{}{rhs})
	_= retVal
}

func (this *ShapeRange) SoftEdge() *ole.DispatchClass {
	retVal := this.PropGet(0x00000a66, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ShapeRange) Glow() *ole.DispatchClass {
	retVal := this.PropGet(0x00000a67, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ShapeRange) Reflection() *ole.DispatchClass {
	retVal := this.PropGet(0x00000a68, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ShapeRange) Title() string {
	retVal := this.PropGet(0x000000c7, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ShapeRange) SetTitle(rhs string)  {
	retVal := this.PropPut(0x000000c7, []interface{}{rhs})
	_= retVal
}

