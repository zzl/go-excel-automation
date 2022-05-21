package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000C0318-0000-0000-C000-000000000046
var IID_ShapeNode = syscall.GUID{0x000C0318, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ShapeNode struct {
	ole.OleClient
}

func NewShapeNode(pDisp *win32.IDispatch, addRef bool, scoped bool) *ShapeNode {
	p := &ShapeNode{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ShapeNodeFromVar(v ole.Variant) *ShapeNode {
	return NewShapeNode(v.PdispValVal(), false, false)
}

func (this *ShapeNode) IID() *syscall.GUID {
	return &IID_ShapeNode
}

func (this *ShapeNode) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ShapeNode) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x60020000, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ShapeNode) Creator() int32 {
	retVal := this.PropGet(0x60020001, nil)
	return retVal.LValVal()
}

func (this *ShapeNode) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000001, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ShapeNode) EditingType() int32 {
	retVal := this.PropGet(0x00000064, nil)
	return retVal.LValVal()
}

func (this *ShapeNode) Points() ole.Variant {
	retVal := this.PropGet(0x00000065, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ShapeNode) SegmentType() int32 {
	retVal := this.PropGet(0x00000066, nil)
	return retVal.LValVal()
}

