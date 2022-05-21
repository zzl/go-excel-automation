package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000C0319-0000-0000-C000-000000000046
var IID_ShapeNodes = syscall.GUID{0x000C0319, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ShapeNodes struct {
	ole.OleClient
}

func NewShapeNodes(pDisp *win32.IDispatch, addRef bool, scoped bool) *ShapeNodes {
	p := &ShapeNodes{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ShapeNodesFromVar(v ole.Variant) *ShapeNodes {
	return NewShapeNodes(v.PdispValVal(), false, false)
}

func (this *ShapeNodes) IID() *syscall.GUID {
	return &IID_ShapeNodes
}

func (this *ShapeNodes) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ShapeNodes) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x60020000, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ShapeNodes) Creator() int32 {
	retVal := this.PropGet(0x60020001, nil)
	return retVal.LValVal()
}

func (this *ShapeNodes) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000001, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ShapeNodes) Count() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *ShapeNodes) Item(index interface{}) *ShapeNode {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewShapeNode(retVal.PdispValVal(), false, true)
}

func (this *ShapeNodes) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *ShapeNodes) ForEach(action func(item *ShapeNode) bool) {
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
		pItem := (*ShapeNode)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *ShapeNodes) Delete(index int32)  {
	retVal := this.Call(0x0000000b, []interface{}{index})
	_= retVal
}

func (this *ShapeNodes) Insert(index int32, segmentType int32, editingType int32, x1 float32, y1 float32, x2 float32, y2 float32, x3 float32, y3 float32)  {
	retVal := this.Call(0x0000000c, []interface{}{index, segmentType, editingType, x1, y1, x2, y2, x3, y3})
	_= retVal
}

func (this *ShapeNodes) SetEditingType(index int32, editingType int32)  {
	retVal := this.Call(0x0000000d, []interface{}{index, editingType})
	_= retVal
}

func (this *ShapeNodes) SetPosition(index int32, x1 float32, y1 float32)  {
	retVal := this.Call(0x0000000e, []interface{}{index, x1, y1})
	_= retVal
}

func (this *ShapeNodes) SetSegmentType(index int32, segmentType int32)  {
	retVal := this.Call(0x0000000f, []interface{}{index, segmentType})
	_= retVal
}

