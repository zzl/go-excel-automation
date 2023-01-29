package excel

import (
	"github.com/zzl/go-win32api/v2/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00024444-0000-0000-C000-000000000046
var IID_PublishObject = syscall.GUID{0x00024444, 0x0000, 0x0000,
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type PublishObject struct {
	ole.OleClient
}

func NewPublishObject(pDisp *win32.IDispatch, addRef bool, scoped bool) *PublishObject {
	if pDisp == nil {
		return nil
	}
	p := &PublishObject{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PublishObjectFromVar(v ole.Variant) *PublishObject {
	return NewPublishObject(v.IDispatch(), false, false)
}

func (this *PublishObject) IID() *syscall.GUID {
	return &IID_PublishObject
}

func (this *PublishObject) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *PublishObject) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *PublishObject) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *PublishObject) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *PublishObject) Delete() {
	retVal, _ := this.Call(0x00000075, nil)
	_ = retVal
}

var PublishObject_Publish_OptArgs = []string{
	"Create",
}

func (this *PublishObject) Publish(optArgs ...interface{}) {
	optArgs = ole.ProcessOptArgs(PublishObject_Publish_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000767, nil, optArgs...)
	_ = retVal
}

func (this *PublishObject) DivID() string {
	retVal, _ := this.PropGet(0x00000766, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PublishObject) Sheet() string {
	retVal, _ := this.PropGet(0x000002ef, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PublishObject) SourceType() int32 {
	retVal, _ := this.PropGet(0x000002ad, nil)
	return retVal.LValVal()
}

func (this *PublishObject) Source() string {
	retVal, _ := this.PropGet(0x000000de, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PublishObject) HtmlType() int32 {
	retVal, _ := this.PropGet(0x00000765, nil)
	return retVal.LValVal()
}

func (this *PublishObject) SetHtmlType(rhs int32) {
	_ = this.PropPut(0x00000765, []interface{}{rhs})
}

func (this *PublishObject) Title() string {
	retVal, _ := this.PropGet(0x000000c7, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PublishObject) SetTitle(rhs string) {
	_ = this.PropPut(0x000000c7, []interface{}{rhs})
}

func (this *PublishObject) Filename() string {
	retVal, _ := this.PropGet(0x00000587, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PublishObject) SetFilename(rhs string) {
	_ = this.PropPut(0x00000587, []interface{}{rhs})
}

func (this *PublishObject) AutoRepublish() bool {
	retVal, _ := this.PropGet(0x00000882, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PublishObject) SetAutoRepublish(rhs bool) {
	_ = this.PropPut(0x00000882, []interface{}{rhs})
}

