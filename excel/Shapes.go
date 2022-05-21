package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002443A-0000-0000-C000-000000000046
var IID_Shapes = syscall.GUID{0x0002443A, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Shapes struct {
	ole.OleClient
}

func NewShapes(pDisp *win32.IDispatch, addRef bool, scoped bool) *Shapes {
	p := &Shapes{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ShapesFromVar(v ole.Variant) *Shapes {
	return NewShapes(v.PdispValVal(), false, false)
}

func (this *Shapes) IID() *syscall.GUID {
	return &IID_Shapes
}

func (this *Shapes) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Shapes) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Shapes) AddRef() uint32 {
	retVal := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Shapes) Release() uint32 {
	retVal := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Shapes) GetTypeInfoCount(pctinfo *uint32)  {
	retVal := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Shapes) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Shapes) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Shapes) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Shapes) Application() *Application {
	retVal := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Shapes) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Shapes) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Shapes) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *Shapes) Item(index interface{}) *Shape {
	retVal := this.Call(0x000000aa, []interface{}{index})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *Shapes) Default_(index interface{}) *Shape {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *Shapes) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Shapes) ForEach(action func(item *Shape) bool) {
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

func (this *Shapes) AddCallout(type_ int32, left float32, top float32, width float32, height float32) *Shape {
	retVal := this.Call(0x000006b1, []interface{}{type_, left, top, width, height})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *Shapes) AddConnector(type_ int32, beginX float32, beginY float32, endX float32, endY float32) *Shape {
	retVal := this.Call(0x000006b2, []interface{}{type_, beginX, beginY, endX, endY})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *Shapes) AddCurve(safeArrayOfPoints interface{}) *Shape {
	retVal := this.Call(0x000006b7, []interface{}{safeArrayOfPoints})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *Shapes) AddLabel(orientation int32, left float32, top float32, width float32, height float32) *Shape {
	retVal := this.Call(0x000006b9, []interface{}{orientation, left, top, width, height})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *Shapes) AddLine(beginX float32, beginY float32, endX float32, endY float32) *Shape {
	retVal := this.Call(0x000006ba, []interface{}{beginX, beginY, endX, endY})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *Shapes) AddPicture(filename string, linkToFile int32, saveWithDocument int32, left float32, top float32, width float32, height float32) *Shape {
	retVal := this.Call(0x000006bb, []interface{}{filename, linkToFile, saveWithDocument, left, top, width, height})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *Shapes) AddPolyline(safeArrayOfPoints interface{}) *Shape {
	retVal := this.Call(0x000006be, []interface{}{safeArrayOfPoints})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *Shapes) AddShape(type_ int32, left float32, top float32, width float32, height float32) *Shape {
	retVal := this.Call(0x000006bf, []interface{}{type_, left, top, width, height})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *Shapes) AddTextEffect(presetTextEffect int32, text string, fontName string, fontSize float32, fontBold int32, fontItalic int32, left float32, top float32) *Shape {
	retVal := this.Call(0x000006c0, []interface{}{presetTextEffect, text, fontName, fontSize, fontBold, fontItalic, left, top})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *Shapes) AddTextbox(orientation int32, left float32, top float32, width float32, height float32) *Shape {
	retVal := this.Call(0x000006c6, []interface{}{orientation, left, top, width, height})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *Shapes) BuildFreeform(editingType int32, x1 float32, y1 float32) *FreeformBuilder {
	retVal := this.Call(0x000006c7, []interface{}{editingType, x1, y1})
	return NewFreeformBuilder(retVal.PdispValVal(), false, true)
}

func (this *Shapes) Range(index interface{}) *ShapeRange {
	retVal := this.PropGet(0x000000c5, []interface{}{index})
	return NewShapeRange(retVal.PdispValVal(), false, true)
}

func (this *Shapes) SelectAll()  {
	retVal := this.Call(0x000006c9, nil)
	_= retVal
}

func (this *Shapes) AddFormControl(type_ int32, left int32, top int32, width int32, height int32) *Shape {
	retVal := this.Call(0x000006ca, []interface{}{type_, left, top, width, height})
	return NewShape(retVal.PdispValVal(), false, true)
}

var Shapes_AddOLEObject_OptArgs= []string{
	"ClassType", "Filename", "Link", "DisplayAsIcon", 
	"IconFileName", "IconIndex", "IconLabel", "Left", 
	"Top", "Width", "Height", 
}

func (this *Shapes) AddOLEObject(optArgs ...interface{}) *Shape {
	optArgs = ole.ProcessOptArgs(Shapes_AddOLEObject_OptArgs, optArgs)
	retVal := this.Call(0x000006cb, nil, optArgs...)
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *Shapes) AddDiagram(type_ int32, left float32, top float32, width float32, height float32) *Shape {
	retVal := this.Call(0x00000880, []interface{}{type_, left, top, width, height})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *Shapes) AddCanvas(left float32, top float32, width float32, height float32) *Shape {
	retVal := this.Call(0x00000881, []interface{}{left, top, width, height})
	return NewShape(retVal.PdispValVal(), false, true)
}

var Shapes_AddChart_OptArgs= []string{
	"XlChartType", "Left", "Top", "Width", "Height", 
}

func (this *Shapes) AddChart(optArgs ...interface{}) *Shape {
	optArgs = ole.ProcessOptArgs(Shapes_AddChart_OptArgs, optArgs)
	retVal := this.Call(0x00000a69, nil, optArgs...)
	return NewShape(retVal.PdispValVal(), false, true)
}

var Shapes_AddSmartArt_OptArgs= []string{
	"Left", "Top", "Width", "Height", 
}

func (this *Shapes) AddSmartArt(layout *ole.DispatchClass, optArgs ...interface{}) *Shape {
	optArgs = ole.ProcessOptArgs(Shapes_AddSmartArt_OptArgs, optArgs)
	retVal := this.Call(0x00000b68, []interface{}{layout}, optArgs...)
	return NewShape(retVal.PdispValVal(), false, true)
}

