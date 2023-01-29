package excel

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
	"unsafe"
)

// 000C036F-0000-0000-C000-000000000046
var IID_DiagramNodeChildren = syscall.GUID{0x000C036F, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type DiagramNodeChildren struct {
	ole.OleClient
}

func NewDiagramNodeChildren(pDisp *win32.IDispatch, addRef bool, scoped bool) *DiagramNodeChildren {
	if pDisp == nil {
		return nil
	}
	p := &DiagramNodeChildren{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DiagramNodeChildrenFromVar(v ole.Variant) *DiagramNodeChildren {
	return NewDiagramNodeChildren(v.IDispatch(), false, false)
}

func (this *DiagramNodeChildren) IID() *syscall.GUID {
	return &IID_DiagramNodeChildren
}

func (this *DiagramNodeChildren) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DiagramNodeChildren) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x60020000, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DiagramNodeChildren) Creator() int32 {
	retVal, _ := this.PropGet(0x60020001, nil)
	return retVal.LValVal()
}

func (this *DiagramNodeChildren) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *DiagramNodeChildren) ForEach(action func(item *DiagramNode) bool) {
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
		pItem := (*DiagramNode)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *DiagramNodeChildren) Item(index interface{}) *DiagramNode {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewDiagramNode(retVal.IDispatch(), false, true)
}

var DiagramNodeChildren_AddNode_OptArgs = []string{
	"Index", "nodeType",
}

func (this *DiagramNodeChildren) AddNode(optArgs ...interface{}) *DiagramNode {
	optArgs = ole.ProcessOptArgs(DiagramNodeChildren_AddNode_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000a, nil, optArgs...)
	return NewDiagramNode(retVal.IDispatch(), false, true)
}

func (this *DiagramNodeChildren) SelectAll() {
	retVal, _ := this.Call(0x0000000b, nil)
	_ = retVal
}

func (this *DiagramNodeChildren) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000064, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DiagramNodeChildren) Count() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *DiagramNodeChildren) FirstChild() *DiagramNode {
	retVal, _ := this.PropGet(0x00000067, nil)
	return NewDiagramNode(retVal.IDispatch(), false, true)
}

func (this *DiagramNodeChildren) LastChild() *DiagramNode {
	retVal, _ := this.PropGet(0x00000068, nil)
	return NewDiagramNode(retVal.IDispatch(), false, true)
}
