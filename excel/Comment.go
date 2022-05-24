package excel

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00024427-0000-0000-C000-000000000046
var IID_Comment = syscall.GUID{0x00024427, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Comment struct {
	ole.OleClient
}

func NewComment(pDisp *win32.IDispatch, addRef bool, scoped bool) *Comment {
	 if pDisp == nil {
		return nil;
	}
	p := &Comment{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CommentFromVar(v ole.Variant) *Comment {
	return NewComment(v.IDispatch(), false, false)
}

func (this *Comment) IID() *syscall.GUID {
	return &IID_Comment
}

func (this *Comment) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Comment) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer)  {
	retVal, _ := this.Call(0x60000000, []interface{}{riid, ppvObj})
	_= retVal
}

func (this *Comment) AddRef() uint32 {
	retVal, _ := this.Call(0x60000001, nil)
	return retVal.UintValVal()
}

func (this *Comment) Release() uint32 {
	retVal, _ := this.Call(0x60000002, nil)
	return retVal.UintValVal()
}

func (this *Comment) GetTypeInfoCount(pctinfo *uint32)  {
	retVal, _ := this.Call(0x60010000, []interface{}{pctinfo})
	_= retVal
}

func (this *Comment) GetTypeInfo(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer)  {
	retVal, _ := this.Call(0x60010001, []interface{}{itinfo, lcid, pptinfo})
	_= retVal
}

func (this *Comment) GetIDsOfNames(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32)  {
	retVal, _ := this.Call(0x60010002, []interface{}{riid, rgszNames, cNames, lcid, rgdispid})
	_= retVal
}

func (this *Comment) Invoke(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32)  {
	retVal, _ := this.Call(0x60010003, []interface{}{dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr})
	_= retVal
}

func (this *Comment) Application() *Application {
	retVal, _ := this.PropGet(0x00000094, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Comment) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Comment) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Comment) Author() string {
	retVal, _ := this.PropGet(0x0000023e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Comment) Shape() *Shape {
	retVal, _ := this.PropGet(0x0000062e, nil)
	return NewShape(retVal.IDispatch(), false, true)
}

func (this *Comment) Visible() bool {
	retVal, _ := this.PropGet(0x0000022e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Comment) SetVisible(rhs bool)  {
	_ = this.PropPut(0x0000022e, []interface{}{rhs})
}

var Comment_Text_OptArgs= []string{
	"Text", "Start", "Overwrite", 
}

func (this *Comment) Text(optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(Comment_Text_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000008a, nil, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Comment) Delete()  {
	retVal, _ := this.Call(0x00000075, nil)
	_= retVal
}

func (this *Comment) Next() *Comment {
	retVal, _ := this.Call(0x000001f6, nil)
	return NewComment(retVal.IDispatch(), false, true)
}

func (this *Comment) Previous() *Comment {
	retVal, _ := this.Call(0x000001f7, nil)
	return NewComment(retVal.IDispatch(), false, true)
}

