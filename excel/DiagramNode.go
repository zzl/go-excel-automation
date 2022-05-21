package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000C0370-0000-0000-C000-000000000046
var IID_DiagramNode = syscall.GUID{0x000C0370, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type DiagramNode struct {
	ole.OleClient
}

func NewDiagramNode(pDisp *win32.IDispatch, addRef bool, scoped bool) *DiagramNode {
	p := &DiagramNode{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DiagramNodeFromVar(v ole.Variant) *DiagramNode {
	return NewDiagramNode(v.PdispValVal(), false, false)
}

func (this *DiagramNode) IID() *syscall.GUID {
	return &IID_DiagramNode
}

func (this *DiagramNode) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DiagramNode) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x60020000, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *DiagramNode) Creator() int32 {
	retVal := this.PropGet(0x60020001, nil)
	return retVal.LValVal()
}

func (this *DiagramNode) AddNode(pos int32, nodeType int32) *DiagramNode {
	retVal := this.Call(0x0000000a, []interface{}{pos, nodeType})
	return NewDiagramNode(retVal.PdispValVal(), false, true)
}

func (this *DiagramNode) Delete()  {
	retVal := this.Call(0x0000000b, nil)
	_= retVal
}

func (this *DiagramNode) MoveNode(pTargetNode *DiagramNode, pos int32)  {
	retVal := this.Call(0x0000000c, []interface{}{pTargetNode, pos})
	_= retVal
}

func (this *DiagramNode) ReplaceNode(pTargetNode *DiagramNode)  {
	retVal := this.Call(0x0000000d, []interface{}{pTargetNode})
	_= retVal
}

func (this *DiagramNode) SwapNode(pTargetNode *DiagramNode, swapChildren bool)  {
	retVal := this.Call(0x0000000e, []interface{}{pTargetNode, swapChildren})
	_= retVal
}

func (this *DiagramNode) CloneNode(copyChildren bool, pTargetNode *DiagramNode, pos int32) *DiagramNode {
	retVal := this.Call(0x0000000f, []interface{}{copyChildren, pTargetNode, pos})
	return NewDiagramNode(retVal.PdispValVal(), false, true)
}

func (this *DiagramNode) TransferChildren(pReceivingNode *DiagramNode)  {
	retVal := this.Call(0x00000010, []interface{}{pReceivingNode})
	_= retVal
}

func (this *DiagramNode) NextNode() *DiagramNode {
	retVal := this.Call(0x00000011, nil)
	return NewDiagramNode(retVal.PdispValVal(), false, true)
}

func (this *DiagramNode) PrevNode() *DiagramNode {
	retVal := this.Call(0x00000012, nil)
	return NewDiagramNode(retVal.PdispValVal(), false, true)
}

func (this *DiagramNode) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000064, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *DiagramNode) Children() *DiagramNodeChildren {
	retVal := this.PropGet(0x00000065, nil)
	return NewDiagramNodeChildren(retVal.PdispValVal(), false, true)
}

func (this *DiagramNode) Shape() *Shape {
	retVal := this.PropGet(0x00000066, nil)
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *DiagramNode) Root() *DiagramNode {
	retVal := this.PropGet(0x00000067, nil)
	return NewDiagramNode(retVal.PdispValVal(), false, true)
}

func (this *DiagramNode) Diagram() *ole.DispatchClass {
	retVal := this.PropGet(0x00000068, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *DiagramNode) Layout() int32 {
	retVal := this.PropGet(0x00000069, nil)
	return retVal.LValVal()
}

func (this *DiagramNode) SetLayout(rhs int32)  {
	retVal := this.PropPut(0x00000069, []interface{}{rhs})
	_= retVal
}

func (this *DiagramNode) TextShape() *Shape {
	retVal := this.PropGet(0x0000006a, nil)
	return NewShape(retVal.PdispValVal(), false, true)
}

